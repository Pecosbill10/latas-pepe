"""
App satélite para el celular — "buzón" de latas pendientes.

Pensada para hosting gratuito 24/7 (PythonAnywhere free tier). A propósito
NO es una copia completa de la colección: no tiene fotos, ni estadísticas,
ni las 1000+ latas ya cargadas. Solo permite anotar rápido una lata nueva
desde el celular estando afuera de casa. Cuando volvés a la PC, el programa
principal (app.py) se sincroniza con esta app y pasa lo pendiente a la
colección real — después de eso, esta base queda vacía otra vez.

Esto evita el problema de tener 2 copias completas de la colección
desincronizadas: la única fuente de verdad sigue siendo la PC de casa.
"""
from flask import Flask, render_template, request, jsonify, session, redirect, url_for
import sqlite3
import os
import json
import secrets
from datetime import datetime
from functools import wraps

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'pendientes.db')
CONFIG_PATH = os.path.join(BASE_DIR, 'cloud_config.json')

# Mismos nombres de país que la app principal (app.py -> PAIS_ISO), para que
# al sincronizar coincidan exactamente y no se dupliquen países "distintos"
# por una diferencia de tipeo.
PAISES = [
    'Alemania', 'Angola', 'Argentina', 'Austria', 'Bélgica', 'Bolivia', 'Brasil',
    'Cabo Verde', 'Canadá', 'Chile', 'China', 'Colombia', 'Corea del Sur',
    'Costa Rica', 'Croacia', 'Dinamarca', 'E.E.U.U.', 'Escocia', 'Eslovaquia',
    'Eslovenia', 'España', 'Filipinas', 'Finlandia', 'Francia', 'Grecia',
    'Holanda', 'Hungría', 'India', 'Inglaterra', 'Irlanda', 'Italia',
    'Italia (Cerdeña)', 'Japón', 'Lituania', 'México', 'Nicaragua',
    'Nueva Zelanda', 'Panamá', 'Paraguay', 'Perú', 'Polonia', 'Portugal',
    'Rep. Checa', 'Rep. Dominicana', 'Rumania', 'Serbia', 'Sudáfrica', 'Suiza',
    'Tailandia', 'Turquía', 'Uruguay', 'Venezuela',
]

app = Flask(__name__)


def load_config():
    if not os.path.exists(CONFIG_PATH):
        raise RuntimeError(
            f'Falta {CONFIG_PATH}. Copiá cloud_config.example.json a cloud_config.json '
            'y completá tu PIN y una clave secreta antes de arrancar.'
        )
    with open(CONFIG_PATH, encoding='utf-8') as f:
        return json.load(f)


config = load_config()
app.secret_key = config['secret_key']
PIN = str(config['pin'])


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    conn.execute('''CREATE TABLE IF NOT EXISTS pendientes (
        id                INTEGER PRIMARY KEY AUTOINCREMENT,
        marca             TEXT NOT NULL,
        modelo            TEXT,
        tipo_lata         TEXT,
        pais              TEXT,
        cantidad          INTEGER DEFAULT 1,
        notas             TEXT,
        fecha_adquisicion TEXT,
        sabor             INTEGER,
        creado            TEXT
    )''')
    conn.commit()
    conn.close()


# ── Auth (PIN compartido, para uso personal — no es multiusuario) ──────────

def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get('ok'):
            return redirect(url_for('login'))
        return fn(*args, **kwargs)
    return wrapper


def token_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if request.args.get('token') != PIN and request.headers.get('X-Token') != PIN:
            return jsonify({'error': 'no autorizado'}), 401
        return fn(*args, **kwargs)
    return wrapper


@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        if request.form.get('pin', '') == PIN:
            session['ok'] = True
            return redirect(url_for('index'))
        error = 'PIN incorrecto'
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ── Páginas ──────────────────────────────────────────────────────────────

@app.route('/')
@login_required
def index():
    conn = get_conn()
    rows = conn.execute('SELECT * FROM pendientes ORDER BY id DESC').fetchall()
    conn.close()
    return render_template('index.html', pendientes=rows, paises=PAISES)


@app.route('/agregar', methods=['POST'])
@login_required
def agregar():
    marca = request.form.get('marca', '').strip()
    if not marca:
        return redirect(url_for('index'))
    conn = get_conn()
    conn.execute(
        '''INSERT INTO pendientes (marca, modelo, tipo_lata, pais, cantidad, notas, fecha_adquisicion, sabor, creado)
           VALUES (?,?,?,?,?,?,?,?,?)''',
        (
            marca,
            request.form.get('modelo', '').strip() or None,
            request.form.get('tipo_lata') or None,
            request.form.get('pais', '').strip() or None,
            max(int(request.form.get('cantidad') or 1), 1),
            request.form.get('notas', '').strip() or None,
            request.form.get('fecha_adquisicion') or None,
            int(request.form['sabor']) if request.form.get('sabor') else None,
            datetime.now().strftime('%Y-%m-%d %H:%M'),
        )
    )
    conn.commit()
    conn.close()
    return redirect(url_for('index'))


@app.route('/borrar/<int:pid>', methods=['POST'])
@login_required
def borrar(pid):
    conn = get_conn()
    conn.execute('DELETE FROM pendientes WHERE id=?', (pid,))
    conn.commit()
    conn.close()
    return redirect(url_for('index'))


# ── API para que la PC de casa sincronice ──────────────────────────────────

@app.route('/api/pendientes')
@token_required
def api_pendientes():
    conn = get_conn()
    rows = [dict(r) for r in conn.execute('SELECT * FROM pendientes ORDER BY id').fetchall()]
    conn.close()
    return jsonify({'data': rows})


@app.route('/api/pendientes/limpiar', methods=['POST'])
@token_required
def api_limpiar():
    ids = (request.get_json(silent=True) or {}).get('ids', [])
    if not ids:
        return jsonify({'error': 'no ids'}), 400
    conn = get_conn()
    ph = ','.join('?' * len(ids))
    conn.execute(f'DELETE FROM pendientes WHERE id IN ({ph})', ids)
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


init_db()

if __name__ == '__main__':
    app.run(debug=False, port=5050)
