#!/usr/bin/env python3
"""
reset_admin.py — Resetea la contraseña del usuario admin
─────────────────────────────────────────────────────────
Uso:
    python reset_admin.py                        → pide la nueva contraseña
    python reset_admin.py --pwd NuevaClave123    → la pasa directo
    python reset_admin.py --ruta /ruta/a/db.db   → DB en otra ruta
"""

import sqlite3, sys, os, getpass
from werkzeug.security import generate_password_hash

# ── Ruta de la base de datos ────────────────────────────────────────────────
DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'gestion_stock_bl.db')
nueva_pwd = None

for i, arg in enumerate(sys.argv[1:]):
    if arg == '--ruta' and i + 2 <= len(sys.argv) - 1:
        DB_FILE = sys.argv[i + 2]
    if arg == '--pwd' and i + 2 <= len(sys.argv) - 1:
        nueva_pwd = sys.argv[i + 2]

# ── Verificar que la DB existe ──────────────────────────────────────────────
if not os.path.exists(DB_FILE):
    print(f'[ERROR] No se encontró la base de datos: {DB_FILE}')
    sys.exit(1)

print(f'Base de datos: {DB_FILE}')

# ── Pedir contraseña si no viene por parámetro ──────────────────────────────
if not nueva_pwd:
    nueva_pwd = getpass.getpass('Nueva contraseña para admin: ')
    confirma  = getpass.getpass('Confirmar contraseña: ')
    if nueva_pwd != confirma:
        print('[ERROR] Las contraseñas no coinciden.')
        sys.exit(1)

if len(nueva_pwd) < 6:
    print('[ERROR] La contraseña debe tener al menos 6 caracteres.')
    sys.exit(1)

# ── Actualizar ──────────────────────────────────────────────────────────────
con = sqlite3.connect(DB_FILE)
cur = con.execute("SELECT id FROM usuarios WHERE username = 'admin'")
row = cur.fetchone()

if not row:
    print('[ERROR] No existe el usuario admin en la base de datos.')
    con.close()
    sys.exit(1)

con.execute(
    "UPDATE usuarios SET password_hash = ? WHERE username = 'admin'",
    (generate_password_hash(nueva_pwd),)
)
con.commit()
con.close()

print('[OK] Contraseña de admin actualizada correctamente.')
