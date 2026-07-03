#!/usr/bin/env python3
"""
Gestión Stock Fiscal | GS — Backend Flask + SQLite
Grupo Simpa SA  |  Versión 1.0
─────────────────────────────────────────────
Instrucciones rápidas:
  1. Ejecutar install.bat  (una sola vez)
  2. Ejecutar start.bat    (cada vez que querés iniciar el servidor)
  3. Acceder desde el browser: http://IP_DEL_SERVIDOR:5000
"""
import os, sys, json, sqlite3, shutil
from datetime import datetime, timedelta
from functools import wraps
from flask import (Flask, render_template, request, jsonify,
                   session, redirect, url_for, send_file, g)
from werkzeug.security import generate_password_hash, check_password_hash

# ── Dependencias opcionales ────────────────────────────────────────────────
# Se usa xlsxwriter (no openpyxl) para generar el Excel: openpyxl >= 3.x ya no
# escribe la tabla de "shared strings" estándar de Office (usa "inline strings"
# y deja [Content_Types].xml al final del .xlsx). Ese archivo es 100% válido y
# Excel lo abre sin problema, pero el lector de SAP es más estricto y lo
# rechaza. xlsxwriter genera la misma estructura "clásica" que produce Excel,
# así el .xlsx que sale del portal ya es compatible con SAP sin tener que
# abrirlo y volver a guardarlo manualmente.
try:
    import xlsxwriter
    XLSX_OK = True
except ImportError:
    XLSX_OK = False
    print("⚠  xlsxwriter no instalado — exportación Excel deshabilitada")

try:
    from apscheduler.schedulers.background import BackgroundScheduler
    SCHED_OK = True
except ImportError:
    SCHED_OK = False
    print("⚠  apscheduler no instalado — envío automático deshabilitado")

# ── Setup de Flask ─────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET', 'simpa_stock_bl_2026_s3cr3t')
app.permanent_session_lifetime = timedelta(hours=8)

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
DB_FILE   = os.path.join(BASE_DIR, 'gestion_stock_bl.db')
CFG_FILE  = os.path.join(BASE_DIR, 'portal_config.json')

# Busca DB_DatosMaestros.csv en la misma carpeta o un nivel arriba
MAESTROS_CSV = None
for _c in [os.path.join(BASE_DIR, 'DB_DatosMaestros.csv'),
           os.path.join(BASE_DIR, '..', 'DB_DatosMaestros.csv')]:
    if os.path.isfile(_c):
        MAESTROS_CSV = os.path.abspath(_c); break

# ── Configuración ──────────────────────────────────────────────────────────
DEFAULT_CFG = {
    "carpeta_arca": "",              # Ruta UNC o letra de unidad de la carpeta SAP
    "label_arca":   "Carpeta SAP",   # Etiqueta visible
    "hora_envio":   "17:00",         # HH:MM  — hora del envío automático
    "envio_auto":   True,            # True = habilitar scheduler
    "aduana_def":   "0001",          # Aduana por defecto
    "empresa":      "Grupo Simpa SA"
}

def load_cfg():
    if os.path.isfile(CFG_FILE):
        with open(CFG_FILE, encoding='utf-8') as f:
            cfg = json.load(f)
        for k, v in DEFAULT_CFG.items():
            cfg.setdefault(k, v)
        return cfg
    return DEFAULT_CFG.copy()

def save_cfg(cfg):
    with open(CFG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)

# ── Base de datos — esquema DDL ────────────────────────────────────────────
DDL = """
PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS usuarios (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    username      TEXT    UNIQUE NOT NULL COLLATE NOCASE,
    nombre        TEXT    NOT NULL,
    password_hash TEXT    NOT NULL,
    -- admin: todo | comex: carga BL | deposito: ingreso/ubicación | lectura: solo ver
    rol           TEXT    NOT NULL CHECK(rol IN ('admin','comex','deposito','lectura')),
    activo        INTEGER NOT NULL DEFAULT 1,
    created_at    TEXT    DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS bl (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    estado           TEXT    NOT NULL DEFAULT 'PRECARGA',
    aduana           TEXT    NOT NULL DEFAULT '0001',
    numero_bl        TEXT    UNIQUE NOT NULL,
    id_comex         TEXT,
    tlat             TEXT,
    cuit             TEXT,
    pais             TEXT,
    manifiesto       TEXT,
    fecha            TEXT,
    hora             TEXT,
    cond_merc        TEXT,
    ubicacion        TEXT,
    cond_imo         INTEGER DEFAULT 0,
    num_imo          TEXT,
    impedimento      INTEGER DEFAULT 0,
    tipo_impedimento TEXT,
    desc_impedimento TEXT,
    observaciones    TEXT,
    created_by       INTEGER REFERENCES usuarios(id),
    created_at       TEXT    DEFAULT (datetime('now','localtime')),
    updated_by       INTEGER REFERENCES usuarios(id),
    updated_at       TEXT    DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS mercaderia (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    bl_id     INTEGER NOT NULL REFERENCES bl(id) ON DELETE CASCADE,
    embalaje  TEXT,
    condicion TEXT DEFAULT 'Buena',
    cantidad  REAL,
    peso      REAL,
    oc        TEXT,
    pos_oc    TEXT,
    ee        TEXT,
    pos_ee    TEXT,
    despacho  TEXT
);

CREATE TABLE IF NOT EXISTS contenedor (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    bl_id    INTEGER NOT NULL REFERENCES bl(id) ON DELETE CASCADE,
    numero   TEXT,
    longitud TEXT,
    tipo     TEXT,
    bultos   INTEGER
);

CREATE TABLE IF NOT EXISTS contenedor_vacio (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    numero   TEXT,
    longitud TEXT,
    fecha    TEXT,
    hora     TEXT
);

-- Locks por sección: 'header' (comex) y 'deposito' son COMPATIBLES entre sí.
-- 'full' (admin) bloquea todo.
CREATE TABLE IF NOT EXISTS locks (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    bl_id      INTEGER NOT NULL REFERENCES bl(id) ON DELETE CASCADE,
    user_id    INTEGER NOT NULL REFERENCES usuarios(id),
    section    TEXT    NOT NULL DEFAULT 'full',
    expires_at TEXT    NOT NULL
);

CREATE TABLE IF NOT EXISTS audit_log (
    id      INTEGER PRIMARY KEY AUTOINCREMENT,
    tabla   TEXT,
    reg_id  INTEGER,
    accion  TEXT,
    user_id INTEGER,
    datos   TEXT,
    ts      TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS export_log (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    archivo   TEXT,
    ruta      TEXT,
    registros INTEGER,
    user_id   INTEGER,
    ts        TEXT DEFAULT (datetime('now','localtime')),
    estado    TEXT,
    mensaje   TEXT
);
"""

def init_db():
    con = sqlite3.connect(DB_FILE)
    con.executescript(DDL)
    # ── Migraciones incrementales (columnas agregadas después de la v1) ──────
    migraciones = [
        "ALTER TABLE mercaderia ADD COLUMN condicion TEXT DEFAULT 'Buena'",
    ]
    for sql in migraciones:
        try:
            con.execute(sql); con.commit()
        except Exception:
            pass   # columna ya existe → OK
    if con.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0] == 0:
        con.execute(
            "INSERT INTO usuarios(username,nombre,password_hash,rol) VALUES(?,?,?,?)",
            ('admin', 'Administrador', generate_password_hash('admin123'), 'admin')
        )
        print("  ✓ Usuario admin creado → usuario: admin | contraseña: admin123")
    con.commit(); con.close()
    print(f"  ✓ Base de datos: {DB_FILE}")

def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB_FILE)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA journal_mode=WAL")
        g.db.execute("PRAGMA foreign_keys=ON")
    return g.db

@app.teardown_appcontext
def close_db(_=None):
    db = g.pop('db', None)
    if db: db.close()

# ── Auth helpers ───────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def wrapper(*a, **kw):
        if 'uid' not in session:
            return (jsonify({'error': 'No autenticado'}), 401) \
                   if request.is_json else redirect(url_for('login_page'))
        return f(*a, **kw)
    return wrapper

def roles(*allowed):
    """Decorador: solo deja pasar si el rol del usuario está en allowed"""
    def deco(f):
        @wraps(f)
        def wrapper(*a, **kw):
            if session.get('rol') not in allowed:
                return jsonify({'error': 'Sin permiso para esta operación'}), 403
            return f(*a, **kw)
        return wrapper
    return deco

def me():
    return {
        'id':      session.get('uid'),
        'username': session.get('usr'),
        'nombre':  session.get('nombre'),
        'rol':     session.get('rol')
    }

# ── Locks concurrentes ─────────────────────────────────────────────────────
# Tabla de compatibilidad:
#   header  + header   → CONFLICTO   (dos comex editando mismo BL)
#   header  + deposito → COMPATIBLE  (comex y deposito pueden editar simultáneo)
#   deposito + deposito → CONFLICTO
#   full    + cualquiera → CONFLICTO (admin bloquea todo)
INCOMPAT = {
    'header':   ('header', 'full'),
    'deposito': ('deposito', 'full'),
    'full':     ('header', 'deposito', 'full'),
}

def lock_acquire(db, bl_id, uid, section):
    # Limpiar locks vencidos
    db.execute("DELETE FROM locks WHERE expires_at < datetime('now','localtime')")
    # Ver si hay conflicto
    incomp = INCOMPAT.get(section, ('full',))
    ph = ','.join('?' * len(incomp))
    conflict = db.execute(
        f"SELECT l.section, u.nombre FROM locks l "
        f"JOIN usuarios u ON l.user_id=u.id "
        f"WHERE l.bl_id=? AND l.user_id!=? AND l.section IN ({ph}) "
        f"AND l.expires_at > datetime('now','localtime')",
        (bl_id, uid, *incomp)
    ).fetchone()
    if conflict:
        return False, f"BL en edición por {conflict['nombre']}"
    # Renovar o crear lock
    db.execute("DELETE FROM locks WHERE bl_id=? AND user_id=? AND section=?",
               (bl_id, uid, section))
    exp = (datetime.now() + timedelta(minutes=10)).strftime('%Y-%m-%d %H:%M:%S')
    db.execute("INSERT INTO locks(bl_id,user_id,section,expires_at) VALUES(?,?,?,?)",
               (bl_id, uid, section, exp))
    db.commit()
    return True, 'OK'

def lock_release(db, bl_id, uid):
    db.execute("DELETE FROM locks WHERE bl_id=? AND user_id=?", (bl_id, uid))
    db.commit()

def lock_info(db, bl_id):
    db.execute("DELETE FROM locks WHERE expires_at < datetime('now','localtime')")
    rows = db.execute(
        "SELECT l.section, u.nombre, l.expires_at FROM locks l "
        "JOIN usuarios u ON l.user_id=u.id "
        "WHERE l.bl_id=? AND l.expires_at > datetime('now','localtime')",
        (bl_id,)
    ).fetchall()
    return [dict(r) for r in rows]

# ── Páginas ────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    if 'uid' not in session: return redirect(url_for('login_page'))
    return render_template('index.html', user=me())

@app.route('/login')
def login_page():
    if 'uid' in session: return redirect('/')
    return render_template('login.html')

# ── Auth API ───────────────────────────────────────────────────────────────
@app.route('/api/auth/login', methods=['POST'])
def api_login():
    d = request.get_json(force=True)
    db = get_db()
    u = db.execute("SELECT * FROM usuarios WHERE username=? AND activo=1",
                   (d.get('username', '').strip(),)).fetchone()
    if u and check_password_hash(u['password_hash'], d.get('password', '')):
        session.permanent = True
        session.update(uid=u['id'], usr=u['username'],
                       nombre=u['nombre'], rol=u['rol'])
        return jsonify({'ok': True, 'rol': u['rol'], 'nombre': u['nombre']})
    return jsonify({'ok': False, 'error': 'Usuario o contraseña incorrectos'}), 401

@app.route('/api/auth/logout', methods=['POST'])
def api_logout():
    session.clear(); return jsonify({'ok': True})

@app.route('/api/auth/me')
@login_required
def api_me():
    return jsonify(me())

# ── BL API ─────────────────────────────────────────────────────────────────
@app.route('/api/bl')
@login_required
def api_bl_list():
    db = get_db()
    bls = db.execute("""
        SELECT b.*,
               c.nombre AS created_by_nombre,
               u.nombre AS updated_by_nombre,
               (SELECT COUNT(*) FROM mercaderia WHERE bl_id = b.id) cant_merc,
               (SELECT COUNT(*) FROM contenedor  WHERE bl_id = b.id) cant_cont
        FROM bl b
        LEFT JOIN usuarios c ON b.created_by = c.id
        LEFT JOIN usuarios u ON b.updated_by = u.id
        ORDER BY b.updated_at DESC
    """).fetchall()
    # Locks activos para todos los BLs
    all_locks = db.execute("""
        SELECT l.bl_id, l.section, u.nombre locked_by
        FROM locks l JOIN usuarios u ON l.user_id = u.id
        WHERE l.expires_at > datetime('now','localtime')
    """).fetchall()
    lmap = {}
    for l in all_locks:
        lmap.setdefault(l['bl_id'], []).append(dict(l))
    result = []
    for b in bls:
        d = dict(b); d['locks'] = lmap.get(d['id'], [])
        result.append(d)
    return jsonify(result)

@app.route('/api/bl/<int:bid>')
@login_required
def api_bl_get(bid):
    db = get_db()
    b = db.execute("SELECT * FROM bl WHERE id=?", (bid,)).fetchone()
    if not b: return jsonify({'error': 'No encontrado'}), 404
    mercs = db.execute("SELECT * FROM mercaderia WHERE bl_id=? ORDER BY id", (bid,)).fetchall()
    conts = db.execute("SELECT * FROM contenedor  WHERE bl_id=? ORDER BY id", (bid,)).fetchall()
    return jsonify({
        'bl':          dict(b),
        'mercaderias': [dict(m) for m in mercs],
        'contenedores': [dict(c) for c in conts],
        'locks':       lock_info(db, bid)
    })

@app.route('/api/bl', methods=['POST'])
@login_required
@roles('admin', 'comex')
def api_bl_create():
    d = request.get_json(force=True)
    db = get_db()
    try:
        cur = db.execute("""
            INSERT INTO bl(estado,aduana,numero_bl,id_comex,tlat,cuit,pais,manifiesto,
                           fecha,hora,cond_merc,ubicacion,cond_imo,num_imo,
                           impedimento,tipo_impedimento,desc_impedimento,observaciones,
                           created_by,updated_by)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            d.get('estado', 'PRECARGA'), d.get('aduana', '0001'),
            d.get('numero_bl', '').upper().strip(),
            d.get('id_comex'), d.get('tlat', '').upper().strip(),
            d.get('cuit'), d.get('pais'), d.get('manifiesto'),
            d.get('fecha'), d.get('hora'), d.get('cond_merc'), d.get('ubicacion'),
            1 if d.get('cond_imo') else 0, d.get('num_imo'),
            1 if d.get('impedimento') else 0, d.get('tipo_impedimento'),
            d.get('desc_impedimento'), d.get('observaciones'),
            session['uid'], session['uid']
        ))
        bid = cur.lastrowid
        _insert_mercs(db, bid, d.get('mercaderias', []))
        _insert_conts(db, bid, d.get('contenedores', []))
        db.commit()
        _audit(db, 'bl', bid, 'CREATE', d)
        return jsonify({'ok': True, 'id': bid}), 201
    except sqlite3.IntegrityError:
        return jsonify({'error': f"El BL '{d.get('numero_bl')}' ya existe"}), 409

@app.route('/api/bl/<int:bid>', methods=['PUT'])
@login_required
@roles('admin', 'comex', 'deposito')
def api_bl_update(bid):
    d = request.get_json(force=True)
    db = get_db()
    rol = session['rol']
    section = 'full' if rol == 'admin' else ('header' if rol == 'comex' else 'deposito')
    ok, msg = lock_acquire(db, bid, session['uid'], section)
    if not ok: return jsonify({'error': msg}), 423  # 423 = Locked

    if rol == 'deposito':
        # Depósito puede cambiar: estado, fecha/hora de ingreso, ubicación, observaciones + mercaderías
        db.execute("""
            UPDATE bl SET estado=?, fecha=COALESCE(?,fecha), hora=COALESCE(?,hora),
                          ubicacion=?, observaciones=?,
                          updated_by=?, updated_at=datetime('now','localtime')
            WHERE id=?
        """, (d.get('estado'),
              d.get('fecha') or None, d.get('hora') or None,
              d.get('ubicacion'), d.get('observaciones'),
              session['uid'], bid))
        if 'mercaderias' in d:
            db.execute("DELETE FROM mercaderia WHERE bl_id=?", (bid,))
            _insert_mercs(db, bid, d['mercaderias'])
        if 'contenedores' in d:
            db.execute("DELETE FROM contenedor WHERE bl_id=?", (bid,))
            _insert_conts(db, bid, d['contenedores'])

    elif rol == 'comex':
        # Comex edita campos de cabecera del BL
        db.execute("""
            UPDATE bl SET estado=?,aduana=?,numero_bl=?,id_comex=?,tlat=?,cuit=?,
                          pais=?,manifiesto=?,fecha=?,hora=?,
                          cond_imo=?,num_imo=?,impedimento=?,tipo_impedimento=?,
                          desc_impedimento=?,
                          updated_by=?, updated_at=datetime('now','localtime')
            WHERE id=?
        """, (
            d.get('estado'), d.get('aduana'), d.get('numero_bl', '').upper().strip(),
            d.get('id_comex'), d.get('tlat', '').upper().strip(),
            d.get('cuit'), d.get('pais'), d.get('manifiesto'),
            d.get('fecha'), d.get('hora'),
            1 if d.get('cond_imo') else 0, d.get('num_imo'),
            1 if d.get('impedimento') else 0, d.get('tipo_impedimento'),
            d.get('desc_impedimento'), session['uid'], bid
        ))
        if 'mercaderias' in d:
            db.execute("DELETE FROM mercaderia WHERE bl_id=?", (bid,))
            _insert_mercs(db, bid, d['mercaderias'])
        if 'contenedores' in d:
            db.execute("DELETE FROM contenedor WHERE bl_id=?", (bid,))
            _insert_conts(db, bid, d['contenedores'])

    else:  # admin — acceso completo
        db.execute("""
            UPDATE bl SET estado=?,aduana=?,numero_bl=?,id_comex=?,tlat=?,cuit=?,
                          pais=?,manifiesto=?,fecha=?,hora=?,cond_merc=?,ubicacion=?,
                          cond_imo=?,num_imo=?,impedimento=?,tipo_impedimento=?,
                          desc_impedimento=?,observaciones=?,
                          updated_by=?, updated_at=datetime('now','localtime')
            WHERE id=?
        """, (
            d.get('estado'), d.get('aduana'), d.get('numero_bl', '').upper().strip(),
            d.get('id_comex'), d.get('tlat', '').upper().strip(),
            d.get('cuit'), d.get('pais'), d.get('manifiesto'),
            d.get('fecha'), d.get('hora'), d.get('cond_merc'), d.get('ubicacion'),
            1 if d.get('cond_imo') else 0, d.get('num_imo'),
            1 if d.get('impedimento') else 0, d.get('tipo_impedimento'),
            d.get('desc_impedimento'), d.get('observaciones'),
            session['uid'], bid
        ))
        if 'mercaderias' in d:
            db.execute("DELETE FROM mercaderia WHERE bl_id=?", (bid,))
            _insert_mercs(db, bid, d['mercaderias'])
        if 'contenedores' in d:
            db.execute("DELETE FROM contenedor WHERE bl_id=?", (bid,))
            _insert_conts(db, bid, d['contenedores'])

    db.commit()
    lock_release(db, bid, session['uid'])
    _audit(db, 'bl', bid, 'UPDATE', d)
    return jsonify({'ok': True})

@app.route('/api/bl/<int:bid>', methods=['DELETE'])
@login_required
@roles('admin')
def api_bl_delete(bid):
    db = get_db()
    db.execute("DELETE FROM bl WHERE id=?", (bid,))
    _audit(db, 'bl', bid, 'DELETE', {})
    return jsonify({'ok': True})

@app.route('/api/bl/<int:bid>/lock', methods=['DELETE'])
@login_required
def api_lock_release(bid):
    lock_release(get_db(), bid, session['uid'])
    return jsonify({'ok': True})

def _insert_mercs(db, bid, rows):
    for m in rows:
        db.execute("""
            INSERT INTO mercaderia(bl_id,embalaje,condicion,cantidad,peso,oc,pos_oc,ee,pos_ee,despacho)
            VALUES(?,?,?,?,?,?,?,?,?,?)
        """, (bid, m.get('embalaje'), m.get('condicion', 'Buena'),
              m.get('cantidad'), m.get('peso'),
              m.get('oc'), m.get('pos_oc'), m.get('ee'), m.get('pos_ee'), m.get('despacho')))

def _insert_conts(db, bid, rows):
    for c in rows:
        db.execute("""
            INSERT INTO contenedor(bl_id,numero,longitud,tipo,bultos)
            VALUES(?,?,?,?,?)
        """, (bid, c.get('numero'), c.get('longitud'), c.get('tipo'), c.get('bultos')))

def _audit(db, tabla, rid, accion, datos):
    db.execute("INSERT INTO audit_log(tabla,reg_id,accion,user_id,datos) VALUES(?,?,?,?,?)",
               (tabla, rid, accion, session.get('uid'),
                json.dumps(datos, ensure_ascii=False)))
    db.commit()

# ── Contenedor vacío API ───────────────────────────────────────────────────
@app.route('/api/contenedor-vacio')
@login_required
def api_cv_list():
    rows = get_db().execute(
        "SELECT * FROM contenedor_vacio ORDER BY fecha DESC, id DESC"
    ).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/contenedor-vacio', methods=['POST'])
@login_required
@roles('admin', 'deposito')
def api_cv_create():
    d = request.get_json(force=True)
    db = get_db()
    cur = db.execute(
        "INSERT INTO contenedor_vacio(numero,longitud,fecha,hora) VALUES(?,?,?,?)",
        (d.get('numero'), d.get('longitud'), d.get('fecha'), d.get('hora'))
    )
    db.commit()
    return jsonify({'ok': True, 'id': cur.lastrowid}), 201

@app.route('/api/contenedor-vacio/<int:cid>', methods=['PUT'])
@login_required
@roles('admin', 'deposito')
def api_cv_update(cid):
    d = request.get_json(force=True)
    db = get_db()
    db.execute(
        "UPDATE contenedor_vacio SET numero=?, longitud=?, fecha=?, hora=? WHERE id=?",
        (d.get('numero'), d.get('longitud'), d.get('fecha'), d.get('hora'), cid)
    )
    db.commit()
    return jsonify({'ok': True})

@app.route('/api/contenedor-vacio/<int:cid>', methods=['DELETE'])
@login_required
@roles('admin', 'deposito')
def api_cv_delete(cid):
    db = get_db()
    db.execute("DELETE FROM contenedor_vacio WHERE id=?", (cid,))
    db.commit()
    return jsonify({'ok': True})

# ── Usuarios API (solo admin) ──────────────────────────────────────────────
@app.route('/api/usuarios')
@login_required
@roles('admin')
def api_users_list():
    rows = get_db().execute(
        "SELECT id,username,nombre,rol,activo,created_at FROM usuarios ORDER BY nombre"
    ).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/usuarios', methods=['POST'])
@login_required
@roles('admin')
def api_users_create():
    d = request.get_json(force=True)
    try:
        get_db().execute(
            "INSERT INTO usuarios(username,nombre,password_hash,rol) VALUES(?,?,?,?)",
            (d['username'].lower().strip(), d['nombre'],
             generate_password_hash(d['password']), d['rol'])
        )
        get_db().commit()
        return jsonify({'ok': True}), 201
    except sqlite3.IntegrityError:
        return jsonify({'error': 'Nombre de usuario ya existe'}), 409

@app.route('/api/usuarios/<int:uid>', methods=['PUT'])
@login_required
@roles('admin')
def api_users_update(uid):
    d = request.get_json(force=True)
    db = get_db()
    if d.get('password'):
        db.execute(
            "UPDATE usuarios SET nombre=?,rol=?,activo=?,password_hash=? WHERE id=?",
            (d['nombre'], d['rol'], 1 if d.get('activo', True) else 0,
             generate_password_hash(d['password']), uid)
        )
    else:
        db.execute(
            "UPDATE usuarios SET nombre=?,rol=?,activo=? WHERE id=?",
            (d['nombre'], d['rol'], 1 if d.get('activo', True) else 0, uid)
        )
    db.commit()
    return jsonify({'ok': True})

# ── Configuración API ──────────────────────────────────────────────────────
@app.route('/api/config')
@login_required
def api_cfg_get():
    cfg = load_cfg()
    if session.get('rol') != 'admin':
        cfg.pop('carpeta_arca', None)  # Ocultar ruta SAP a no-admins
    return jsonify(cfg)

@app.route('/api/config', methods=['PUT'])
@login_required
@roles('admin')
def api_cfg_set():
    cfg = load_cfg()
    cfg.update(request.get_json(force=True))
    save_cfg(cfg)
    _restart_scheduler()  # Aplicar nueva hora de envío
    return jsonify({'ok': True})

# ── Estadísticas dashboard ─────────────────────────────────────────────────
@app.route('/api/stats')
@login_required
def api_stats():
    db = get_db()
    estados = db.execute(
        "SELECT estado, COUNT(*) n FROM bl GROUP BY estado"
    ).fetchall()
    recientes = db.execute("""
        SELECT b.numero_bl, b.estado, b.updated_at, u.nombre usuario
        FROM bl b LEFT JOIN usuarios u ON b.updated_by=u.id
        ORDER BY b.updated_at DESC LIMIT 8
    """).fetchall()
    ultimo_exp = db.execute(
        "SELECT ts, archivo, estado, registros FROM export_log ORDER BY ts DESC LIMIT 1"
    ).fetchone()
    cfg = load_cfg()
    return jsonify({
        'por_estado':   {r['estado']: r['n'] for r in estados},
        'total':        sum(r['n'] for r in estados),
        'recientes':    [dict(r) for r in recientes],
        'ultimo_export': dict(ultimo_exp) if ultimo_exp else None,
        'hora_envio':   cfg.get('hora_envio', '17:00'),
        'envio_auto':   cfg.get('envio_auto', True)
    })

# ── Datos maestros (países, embalajes, aduanas) ────────────────────────────
_maestros_cache = None

def _load_maestros():
    global _maestros_cache
    if _maestros_cache:
        return _maestros_cache
    paises = []; embalajes = []; aduanas = []
    if MAESTROS_CSV:
        section = None
        with open(MAESTROS_CSV, encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line.startswith('#SECTION:'):
                    section = line.split(':')[1]; continue
                if not line or line.startswith('codigo'):
                    continue
                parts = line.split(',', 1)
                if len(parts) == 2:
                    cod = parts[0].strip()
                    nom = parts[1].strip().strip('"')
                    if section == 'PAISES':    paises.append({'codigo': cod, 'nombre': nom})
                    elif section == 'EMBALAJES': embalajes.append({'codigo': cod, 'nombre': nom})
                    elif section == 'ADUANAS':   aduanas.append({'codigo': cod, 'nombre': nom})
        print(f"  ✓ Maestros cargados: {len(paises)} países, "
              f"{len(embalajes)} embalajes, {len(aduanas)} aduanas")
    _maestros_cache = {'paises': paises, 'embalajes': embalajes, 'aduanas': aduanas}
    return _maestros_cache

@app.route('/api/maestros')
@login_required
def api_maestros():
    return jsonify(_load_maestros())

# ── Importar CSV existente ─────────────────────────────────────────────────
@app.route('/api/importar-csv', methods=['POST'])
@login_required
@roles('admin')
def api_importar_csv():
    """Migra datos del DB_Comex.csv al nuevo portal (solo admin, una vez)"""
    csv_path = None
    for c in [os.path.join(BASE_DIR, 'DB_Comex.csv'),
              os.path.join(BASE_DIR, '..', 'DB_Comex.csv')]:
        if os.path.isfile(c): csv_path = os.path.abspath(c); break
    if not csv_path:
        return jsonify({'error': 'DB_Comex.csv no encontrado'}), 404

    db = get_db()
    section = None
    bls = []; mercs = []; conts = []; cvs = []

    with open(csv_path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line.startswith('#SECTION:'):
                section = line.split(':')[1]; continue
            if not line or line.startswith('estado') or \
               line.startswith('numeroBL') or line.startswith('numero,'):
                continue
            parts = _parse_csv(line)
            if   section == 'DETALLE_BL':    bls.append(parts)
            elif section == 'MERCADERIA':    mercs.append(parts)
            elif section == 'CONTENEDORES':  conts.append(parts)
            elif section == 'CONT_VACIOS':   cvs.append(parts)

    importados = 0
    for p in bls:
        g = lambda i, d='': p[i] if i < len(p) else d
        try:
            db.execute("""
                INSERT OR IGNORE INTO bl(
                    estado,aduana,numero_bl,id_comex,tlat,cuit,pais,manifiesto,
                    fecha,hora,cond_merc,ubicacion,cond_imo,num_imo,
                    impedimento,tipo_impedimento,desc_impedimento,observaciones)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (g(0,'PRECARGA'),g(1,'0001'),g(2).upper(),g(3),g(4).upper(),
                  g(5),g(6),g(7),g(8),g(9),g(10),g(11),
                  1 if g(12)=='VERDADERO' else 0, g(13),
                  1 if g(14)=='VERDADERO' else 0, g(15), g(16), g(17)))
            importados += 1
        except Exception:
            pass

    bl_map = {r['numero_bl']: r['id']
              for r in db.execute("SELECT id,numero_bl FROM bl").fetchall()}

    for p in mercs:
        g = lambda i, d='': p[i] if i < len(p) else d
        bid = bl_map.get(g(0).upper())
        if bid:
            db.execute("""INSERT INTO mercaderia(bl_id,embalaje,condicion,cantidad,peso,
                          oc,pos_oc,ee,pos_ee,despacho) VALUES(?,?,?,?,?,?,?,?,?,?)""",
                       (bid,g(1),g(2) or 'Buena',g(3) or None,g(4) or None,
                        g(5),g(6),g(7),g(8),g(9)))

    for p in conts:
        g = lambda i, d='': p[i] if i < len(p) else d
        bid = bl_map.get(g(0).upper())
        if bid:
            db.execute("INSERT INTO contenedor(bl_id,numero,longitud,tipo,bultos) VALUES(?,?,?,?,?)",
                       (bid, g(1), g(2), g(3), g(4) or None))

    for p in cvs:
        g = lambda i, d='': p[i] if i < len(p) else d
        db.execute("INSERT INTO contenedor_vacio(numero,longitud,fecha,hora) VALUES(?,?,?,?)",
                   (g(0), g(1), g(2), g(3)))

    db.commit()
    return jsonify({'ok': True, 'importados': importados, 'total_csv': len(bls)})

def _parse_csv(line):
    result = []; field = ''; in_q = False
    for ch in line:
        if ch == '"':   in_q = not in_q
        elif ch == ',' and not in_q: result.append(field.strip()); field = ''
        else: field += ch
    result.append(field.strip())
    return result

# ── Exportación Excel para ARCA/SAP ───────────────────────────────────────
def _parse_fecha_excel(s):
    """Convierte 'YYYY-MM-DD' (texto en la BD) a date real para que Excel
    pueda mostrarlo con el formato dd/mm/aaaa."""
    if not s:
        return None
    try:
        return datetime.strptime(str(s)[:10], '%Y-%m-%d').date()
    except Exception:
        return s

def _parse_hora_excel(s):
    """Convierte 'HH:MM' o 'HH:MM:SS' (texto en la BD) a time real para que
    Excel pueda mostrarlo con el formato hh:mm:ss."""
    if not s:
        return None
    s = str(s).strip()
    try:
        if len(s) <= 5:
            return datetime.strptime(s, '%H:%M').time()
        return datetime.strptime(s[:8], '%H:%M:%S').time()
    except Exception:
        return s

def _generar_excel(con, ruta):
    """Genera el Excel con las 5 hojas en el formato original requerido por ARCA
    (nombres de hoja, títulos y orden de columnas tal como exportaba el portal HTML v5)
    y lo escribe directamente en 'ruta' (xlsxwriter exige el archivo de destino
    desde el momento en que se crea el Workbook, no admite "generar y guardar después").
    con: conexión SQLite existente o None (abre y cierra la propia).
    Devuelve la cantidad de BLs exportados."""
    if not XLSX_OK:
        raise RuntimeError("Instalar xlsxwriter: pip install xlsxwriter")
    cierra = con is None
    if cierra:
        con = sqlite3.connect(DB_FILE); con.row_factory = sqlite3.Row

    bls   = con.execute("SELECT * FROM bl ORDER BY numero_bl").fetchall()
    mercs = con.execute("SELECT * FROM mercaderia").fetchall()
    conts = con.execute("SELECT * FROM contenedor").fetchall()
    cvs   = con.execute("SELECT * FROM contenedor_vacio ORDER BY numero").fetchall()

    wb    = xlsxwriter.Workbook(ruta)
    TFONT = wb.add_format({'bold': True, 'font_size': 13})
    SFONT = wb.add_format({'italic': True, 'font_size': 10, 'font_color': '#555555'})
    HFMT  = wb.add_format({'bold': True, 'font_size': 10, 'font_color': '#FFFFFF',
                            'bg_color': '#1C2B3F', 'align': 'center'})
    _fmt_cache = {}
    def fmt_for(fmt_str, alt):
        """Devuelve (y cachea) el Format de xlsxwriter para una combinación
        de número-formato + sombreado de fila alternada."""
        key = (fmt_str, alt)
        if key not in _fmt_cache:
            props = {}
            if fmt_str: props['num_format'] = fmt_str
            if alt: props['bg_color'] = '#EEF2F7'
            _fmt_cache[key] = wb.add_format(props)
        return _fmt_cache[key]

    def hoja(name, titulo, subtitulo, headers, rows, formats=None):
        ws = wb.add_worksheet(name)
        ncols = len(headers)
        ws.merge_range(0, 0, 0, ncols - 1, titulo, TFONT)
        ws.merge_range(1, 0, 1, ncols - 1, subtitulo, SFONT)
        for ci, h in enumerate(headers):
            ws.write(2, ci, h, HFMT)
        for ri, r in enumerate(rows):
            row_idx = ri + 3                 # fila 4 (1-based) = primera fila de datos
            alt = (row_idx % 2 == 1)         # mismo criterio de sombreado que antes
            for ci, val in enumerate(r):
                fmt_str  = (formats or {}).get(ci + 1)
                cell_fmt = fmt_for(fmt_str, alt)
                ws.write(row_idx, ci, val, cell_fmt)
        # Ancho de columna en base a encabezado + datos
        for ci in range(ncols):
            w = len(str(headers[ci] or ''))
            for r in rows:
                if ci < len(r):
                    w = max(w, len(str(r[ci] if r[ci] is not None else '')))
            ws.set_column(ci, ci, min(w + 2, 40))
        return ws

    bl_id_map = {b['id']: b['numero_bl'] for b in bls}

    # Hoja 1 — Detalle del BL
    # (igual al formato original; se quita "Condición Mercadería" porque ese dato
    #  ahora se carga por línea en la hoja de Mercadería, no a nivel de BL)
    hoja('1. Detalle del BL',
         'REGISTRO DE BILL OF LADING - DETALLE DEL BL',
         'Área de Comercio Exterior - Importación — Gestión Stock Fiscal | GS',
         ['Estado', 'Código de Aduana', 'Número de BL', 'ID Comex', 'Número de TLAT',
          'CUIT del Consignatario', 'Procedencia (País)', 'Número de Manifiesto',
          'Fecha de Ingreso', 'Hora Ingreso', 'Ubicación', 'Condición IMO', 'Número IMO',
          'Impedimento Legal', 'Tipo Impedimento Legal', 'Descripción Impedimento Legal',
          'Observaciones'],
         [[b['estado'], b['aduana'], b['numero_bl'], b['id_comex'] or '', b['tlat'],
           b['cuit'], b['pais'], b['manifiesto'], _parse_fecha_excel(b['fecha']),
           _parse_hora_excel(b['hora']),
           b['ubicacion'] or '', bool(b['cond_imo']), b['num_imo'] or '',
           bool(b['impedimento']), b['tipo_impedimento'] or '', b['desc_impedimento'] or '',
           b['observaciones'] or ''] for b in bls],
         formats={9: 'DD/MM/YYYY', 10: 'HH:MM:SS'})

    # (con xlsxwriter no hay hoja "Sheet" vacía por defecto: solo se crean
    # las hojas que se piden explícitamente con add_worksheet)

    # Hoja 2 — Mercadería y Embalajes
    # (se agrega "Condición Mercadería" porque ahora es un dato por línea de mercadería)
    hoja('2. Mercaderia y Embalajes',
         'REGISTRO DE BL - LÍNEAS DE MERCADERÍA Y EMBALAJES',
         'Gestión Stock Fiscal | GS',
         ['Número de BL', 'Tipo de Embalaje', 'Condición Mercadería', 'Cantidad de Embalaje',
          'Peso Embalaje (kg)', 'Número de OC', 'Posición de la OC',
          'Número de Entrega Entrante', 'Posición de la EE', 'Número de Despacho'],
         [[bl_id_map.get(m['bl_id'], ''), m['embalaje'], m['condicion'] or 'Buena',
           m['cantidad'], m['peso'], m['oc'], m['pos_oc'], m['ee'], m['pos_ee'],
           m['despacho']] for m in mercs])

    # Hoja 3 — Contenedores
    hoja('3. Contenedores',
         'REGISTRO DE BL - GESTIÓN DE CONTENEDORES',
         'Gestión Stock Fiscal | GS',
         ['Número de BL', 'Número del Contenedor', 'Longitud del Contenedor',
          'Tipo de Contenedor', 'Cantidad de Bultos'],
         [[bl_id_map.get(c['bl_id'], ''), c['numero'], c['longitud'], c['tipo'],
           c['bultos']] for c in conts])

    # Hoja 4 — Contenedores Vacíos
    hoja('4. Contenedores Vacios',
         'GESTIÓN DE CONTENEDORES VACÍOS',
         'Gestión Stock Fiscal | GS',
         ['Número del Contenedor', 'Longitud del Contenedor', 'Fecha de Ingreso',
          'Hora Ingreso'],
         [[cv['numero'], cv['longitud'], _parse_fecha_excel(cv['fecha']),
           _parse_hora_excel(cv['hora'])] for cv in cvs],
         formats={3: 'DD/MM/YYYY', 4: 'HH:MM:SS'})

    # Hoja 5 — Control de Completitud
    merc_by_bl = {}
    for m in mercs: merc_by_bl.setdefault(m['bl_id'], []).append(m)
    cont_by_bl = {}
    for c in conts: cont_by_bl.setdefault(c['bl_id'], []).append(c)

    ctrl_rows = []
    for b in bls:
        nm = len(merc_by_bl.get(b['id'], []))
        nc = len(cont_by_bl.get(b['id'], []))
        falta_clave = not (b['numero_bl'] and b['tlat'] and b['cuit'] and b['pais'] and b['manifiesto'])
        if falta_clave or nm == 0:
            resultado = 'Error'
        elif not b['ubicacion']:
            resultado = 'Advertencia'
        else:
            resultado = 'Completo'
        ctrl_rows.append([b['estado'], b['numero_bl'], 'SI' if nm else 'NO',
                           'SI' if nc else 'N/A', nm, nc, resultado])

    ws5 = hoja('5. Control de Completitud',
               'CONTROL DE COMPLETITUD - BILL OF LADING',
               'Verde=Completo | Amarillo=Advertencia | Rojo=Error',
               ['Estado', 'Número de BL', 'Mercadería', 'Contenedor', 'Lín.Merc.',
                'Lín.Cont.', 'Resultado'],
               ctrl_rows)

    # Semáforo de la columna "Resultado": se reescribe la celda con un formato
    # de color (xlsxwriter no permite "restilar" una celda ya escrita, hay que
    # volver a escribirla — la última escritura sobre una celda es la que queda).
    COLOR = {'Completo': '#2ECC71', 'Advertencia': '#F1C40F', 'Error': '#E74C3C'}
    for ri, row in enumerate(ctrl_rows):
        col = COLOR.get(row[6])
        if col:
            sem_fmt = wb.add_format({
                'bold': True, 'bg_color': col,
                'font_color': '#000000' if row[6] == 'Advertencia' else '#FFFFFF'
            })
            ws5.write(ri + 3, 6, row[6], sem_fmt)

    if cierra: con.close()
    wb.close()                 # graba el archivo en 'ruta'
    return len(bls)

@app.route('/api/export', methods=['POST'])
@login_required
@roles('admin', 'comex', 'deposito')
def api_export():
    """Genera el Excel y lo deja ÚNICAMENTE en la carpeta de destino configurada
    (carpeta SAP/ARCA). No se descarga al browser."""
    db = get_db(); cfg = load_cfg()
    fname   = f"Stock_BL_ARCA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    carpeta = cfg.get('carpeta_arca', '').strip()

    if not carpeta:
        return jsonify({
            'error': 'No hay una carpeta de destino configurada. '
                     'Pedile a un administrador que la configure en Configuración → '
                     + cfg.get('label_arca', 'Carpeta SAP')
        }), 400

    try:
        os.makedirs(carpeta, exist_ok=True)
        ruta = os.path.join(carpeta, fname)
        n = _generar_excel(db, ruta)

        db.execute(
            "INSERT INTO export_log(archivo,ruta,registros,user_id,estado,mensaje) VALUES(?,?,?,?,?,?)",
            (fname, ruta, n, session['uid'], 'OK', None)
        )
        db.commit()
        return jsonify({'ok': True, 'archivo': fname, 'ruta': ruta, 'registros': n})

    except Exception as e:
        db.execute(
            "INSERT INTO export_log(archivo,ruta,registros,user_id,estado,mensaje) VALUES(?,?,?,?,?,?)",
            (fname, '', 0, session.get('uid'), 'ERROR', str(e))
        )
        db.commit()
        return jsonify({'error': str(e)}), 500

@app.route('/api/reporte')
@login_required
def api_reporte():
    """Devuelve todos los BLs con mercaderías y contenedores anidados + contenedores vacíos.
    Usado para generar el reporte PDF desde el browser."""
    db = get_db()
    bls_raw = db.execute("""
        SELECT b.*, u.nombre updated_by_nombre
        FROM bl b LEFT JOIN usuarios u ON b.updated_by = u.id
        ORDER BY b.fecha DESC NULLS LAST, b.numero_bl
    """).fetchall()
    result = []
    for b in bls_raw:
        bd = dict(b)
        bd['mercaderias']  = [dict(m) for m in
            db.execute("SELECT * FROM mercaderia WHERE bl_id=? ORDER BY id", (b['id'],)).fetchall()]
        bd['contenedores'] = [dict(c) for c in
            db.execute("SELECT * FROM contenedor WHERE bl_id=? ORDER BY id",  (b['id'],)).fetchall()]
        result.append(bd)
    cvs = [dict(c) for c in
        db.execute("SELECT * FROM contenedor_vacio ORDER BY fecha DESC, id DESC").fetchall()]
    return jsonify({'bls': result, 'contenedores_vacios': cvs})

@app.route('/api/export/log')
@login_required
def api_export_log():
    rows = get_db().execute("""
        SELECT e.*, u.nombre usuario FROM export_log e
        LEFT JOIN usuarios u ON e.user_id = u.id
        ORDER BY e.ts DESC LIMIT 50
    """).fetchall()
    return jsonify([dict(r) for r in rows])

# ── Exportación automática (ejecutada por scheduler) ──────────────────────
def _export_auto():
    """Sin contexto Flask — no usa 'session' ni 'g'"""
    cfg = load_cfg()
    carpeta = cfg.get('carpeta_arca', '').strip()
    if not carpeta:
        print(f"[{datetime.now():%H:%M}] ⚠ Carpeta ARCA no configurada"); return
    if not os.path.isdir(carpeta):
        print(f"[{datetime.now():%H:%M}] ⚠ Carpeta inaccesible: {carpeta}"); return

    fname = f"Stock_BL_ARCA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    tmp   = os.path.join(BASE_DIR, fname)
    con   = sqlite3.connect(DB_FILE); con.row_factory = sqlite3.Row
    try:
        n = _generar_excel(con, tmp)
        dest = os.path.join(carpeta, fname)
        shutil.copy2(tmp, dest)
        con.execute(
            "INSERT INTO export_log(archivo,ruta,registros,estado,mensaje) VALUES(?,?,?,?,?)",
            (fname, dest, n, 'OK', 'Automático')
        )
        con.commit()
        print(f"[{datetime.now():%H:%M}] ✓ Export automático: {n} BLs → {dest}")
    except Exception as e:
        print(f"[{datetime.now():%H:%M}] ✗ Error export automático: {e}")
    finally:
        con.close()
        try: os.remove(tmp)
        except: pass

# ── Scheduler ─────────────────────────────────────────────────────────────
_scheduler = None

def _restart_scheduler():
    global _scheduler
    if not SCHED_OK: return
    cfg = load_cfg()
    if _scheduler and _scheduler.running:
        try: _scheduler.shutdown(wait=False)
        except: pass
    if not cfg.get('envio_auto', True):
        print("  ℹ Envío automático deshabilitado"); return
    h, m = 17, 0
    try: h, m = map(int, cfg.get('hora_envio', '17:00').split(':'))
    except: pass
    _scheduler = BackgroundScheduler(daemon=True)
    _scheduler.add_job(_export_auto, 'cron', hour=h, minute=m, misfire_grace_time=600)
    _scheduler.start()
    print(f"  ✓ Envío automático programado todos los días a las {h:02d}:{m:02d}")

# ── Arranque ─────────────────────────────────────────────────────
if __name__ == '__main__':
    if '--init-only' in sys.argv:
        print('Inicializando base de datos...')
        init_db(); sys.exit(0)
    print('=' * 54)
    print('  Gestión Stock Fiscal | GS  -  v1.0')
    print('=' * 54)
    init_db()
    _load_maestros()
    _restart_scheduler()
    try:
        import socket
        ip = socket.gethostbyname(socket.gethostname())
    except Exception:
        ip = '127.0.0.1'
    print(f'  Acceso local:  http://localhost:5000')
    print(f'  Acceso red:    http://{ip}:5000')
    print('=' * 54)
    try:
        from waitress import serve
        print('  Usando waitress (modo produccion)')
        serve(app, host='0.0.0.0', port=5000, threads=8)
    except ImportError:
        print('  Usando Flask dev server')
        app.run(host='0.0.0.0', port=5000, debug=False)
