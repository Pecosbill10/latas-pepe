import os
import sys
import tempfile

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as app_module


@pytest.fixture
def client():
    fd, path = tempfile.mkstemp(suffix='.db')
    os.close(fd)
    tmpdir = tempfile.mkdtemp()
    app_module.DB_PATH = path
    app_module.EXCEL_PATH = '__no_existe__.xlsx'  # evita importar el Excel real en cada test
    app_module.FOTOS_DIR = os.path.join(tmpdir, 'fotos')
    app_module.BACKUP_DIR = os.path.join(tmpdir, 'backups')
    os.makedirs(app_module.FOTOS_DIR, exist_ok=True)
    app_module.init_db()
    app_module.app.config['TESTING'] = True
    with app_module.app.test_client() as c:
        yield c
    os.remove(path)


def test_crear_lata_valida(client):
    r = client.post('/api/latas', json={'marca': 'Quilmes', 'pais': 'Argentina', 'cantidad': 2})
    assert r.status_code == 200
    assert r.get_json()['ok'] is True


def test_crear_lata_sin_marca_falla(client):
    r = client.post('/api/latas', json={'pais': 'Argentina'})
    assert r.status_code == 400


def test_crear_lata_sin_pais_falla(client):
    r = client.post('/api/latas', json={'marca': 'Quilmes'})
    assert r.status_code == 400


def test_sabor_fuera_de_rango_falla(client):
    r = client.post('/api/latas', json={'marca': 'Corona', 'pais': 'México', 'sabor': 9})
    assert r.status_code == 400


def test_deteccion_de_duplicados(client):
    client.post('/api/latas', json={'marca': 'Brahma', 'modelo': 'Chopp', 'pais': 'Brasil', 'cantidad': 1})
    r = client.post('/api/latas', json={'marca': 'brahma', 'modelo': 'chopp', 'pais': 'brasil', 'cantidad': 3})
    assert r.status_code == 409
    body = r.get_json()
    assert body['duplicate'] is True
    assert body['existing_cantidad'] == 1


def test_duplicado_se_puede_forzar(client):
    client.post('/api/latas', json={'marca': 'Stella', 'pais': 'Bélgica'})
    r = client.post('/api/latas', json={'marca': 'Stella', 'pais': 'Bélgica', 'force': True})
    assert r.status_code == 200


def test_patch_sabor_inline(client):
    r = client.post('/api/latas', json={'marca': 'Guinness', 'pais': 'Irlanda'})
    lid = r.get_json()['id']

    r2 = client.patch(f'/api/latas/{lid}/sabor', json={'sabor': 4})
    assert r2.status_code == 200
    assert r2.get_json()['sabor'] == 4

    # Volver a poner null desmarca la calificación
    r3 = client.patch(f'/api/latas/{lid}/sabor', json={'sabor': None})
    assert r3.status_code == 200
    assert r3.get_json()['sabor'] is None


def test_patch_sabor_invalido_falla(client):
    r = client.post('/api/latas', json={'marca': 'Guinness', 'pais': 'Irlanda'})
    lid = r.get_json()['id']
    r2 = client.patch(f'/api/latas/{lid}/sabor', json={'sabor': 7})
    assert r2.status_code == 400


def test_filtro_con_foto(client):
    client.post('/api/latas', json={'marca': 'SinFoto', 'pais': 'Chile'})
    r = client.get('/api/latas?con_foto=1')
    assert r.status_code == 200
    assert r.get_json()['total'] == 0


def test_actualizar_lata_get_then_put_con_nulos(client):
    """Regresión: GET devuelve modelo/notas en null cuando no están seteados;
    reenviar ese mismo objeto por PUT (patrón típico de un merge de duplicados)
    no debe romper con 500."""
    r = client.post('/api/latas', json={'marca': 'Heineken', 'pais': 'Holanda', 'cantidad': 1})
    lid = r.get_json()['id']

    lata = client.get(f'/api/latas/{lid}').get_json()
    assert lata['modelo'] is None
    assert lata['notas'] is None

    lata['cantidad'] = 5
    r2 = client.put(f'/api/latas/{lid}', json=lata)
    assert r2.status_code == 200


def test_wishlist_a_coleccion(client):
    r = client.post('/api/wishlist', json={'marca': 'Erdinger', 'pais': 'Alemania'})
    wid = r.get_json()['id']
    r2 = client.post(f'/api/wishlist/{wid}/conseguir')
    assert r2.status_code == 200
    r3 = client.get('/api/wishlist')
    assert all(w['id'] != wid for w in r3.get_json()['data'])


def test_wishlist_exportar(client):
    client.post('/api/wishlist', json={'marca': 'Erdinger', 'pais': 'Alemania'})
    r = client.get('/api/wishlist/exportar')
    assert r.status_code == 200
    assert r.mimetype == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


def test_foto_rechaza_archivo_que_no_es_imagen(client):
    r = client.post('/api/latas', json={'marca': 'Test', 'pais': 'Argentina'})
    lid = r.get_json()['id']
    data = {'foto': (tempfile_bytes(b'esto no es una imagen'), 'falsa.jpg')}
    r2 = client.post(f'/api/latas/{lid}/foto', data=data, content_type='multipart/form-data')
    assert r2.status_code == 400


def tempfile_bytes(content):
    import io
    return io.BytesIO(content)


def test_restore_rechaza_archivo_invalido(client):
    data = {'backup': (tempfile_bytes(b'no es una base de datos sqlite'), 'falso.db')}
    r = client.post('/api/restore', data=data, content_type='multipart/form-data')
    assert r.status_code == 400
