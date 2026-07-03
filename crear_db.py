#!/usr/bin/env python3
"""
crear_db.py — Inicializa la base de datos de Gestión Stock Fiscal | GS
────────────────────────────────────────────────────────────────────────
Uso:
    python crear_db.py                 → crea gestion_stock_bl.db en la misma carpeta
    python crear_db.py --ruta otra.db  → crea el archivo en la ruta indicada

Si el archivo ya existe NO lo pisa: solo agrega tablas/columnas faltantes
(el mismo comportamiento que tiene el portal al arrancar).

Usuario inicial creado: admin / admin123
¡Cambiar la contraseña en Configuración → Usuarios antes de usar en producción!
"""

import sqlite3
import sys
import os
from werkzeug.security import generate_password_hash

# ── Ruta del archivo de base de datos ──────────────────────────────────────
DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'gestion_stock_bl.db')

for i, arg in enumerate(sys.argv[1:]):
    if arg == '--ruta' and i + 2 <= len(sys.argv) - 1:
        DB_FILE = sys.argv[i + 2]

# ── Esquema completo ────────────────────────────────────────────────────────
DDL = """
PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS usuarios (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    username      TEXT    UNIQUE NOT NULL COLLATE NOCASE,
    nombre        TEXT    NOT NULL,
    password_hash TEXT    NOT NULL,
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

# ── Migraciones incrementales ───────────────────────────────────────────────
# Si la DB ya existía y le faltan columnas nuevas, se agregan sin perder datos.
MIGRACIONES = [
    "ALTER TABLE mercaderia ADD COLUMN condicion TEXT DEFAULT 'Buena'",
]

# ── Main ────────────────────────────────────────────────────────────────────
def main():
    nueva = not os.path.exists(DB_FILE)

    print(f"Base de datos: {DB_FILE}")
    print(f"{'Creando nueva...' if nueva else 'Ya existe — verificando estructura...'}")

    con = sqlite3.connect(DB_FILE)
    con.executescript(DDL)

    for sql in MIGRACIONES:
        try:
            con.execute(sql)
            con.commit()
        except Exception:
            pass  # columna ya existe → OK

    # Crear usuario admin solo si no hay ningún usuario todavía
    if con.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0] == 0:
        con.execute(
            "INSERT INTO usuarios(username, nombre, password_hash, rol) VALUES (?,?,?,?)",
            ('admin', 'Administrador', generate_password_hash('admin123'), 'admin')
        )
        con.commit()
        print("Usuario inicial creado: admin / admin123")
        print("¡Recordá cambiar la contraseña en Configuración → Usuarios!")
    else:
        print("Usuarios existentes conservados.")

    # Mostrar resumen de tablas
    tablas = con.execute(
        "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
    ).fetchall()
    print(f"\nTablas: {', '.join(t[0] for t in tablas)}")
    print("\n[OK] Base de datos lista.")
    con.close()

if __name__ == '__main__':
    main()
