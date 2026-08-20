"""Base maestra en SQLite. Dedup por place_id via INSERT OR IGNORE."""
import sqlite3
import re
from datetime import datetime

RUTA = "output/negocios.db"

CAMPOS = ['place_id', 'nombre', 'telefono', 'categoria', 'direccion', 'web',
          'rating', 'resenas', 'ciudad', 'rubro', 'nicho', 'busqueda', 'url', 'fecha']


def conectar(ruta=RUTA):
    import os
    carpeta = os.path.dirname(ruta)       # vacio con ':memory:' o un nombre pelado
    if carpeta:
        os.makedirs(carpeta, exist_ok=True)
    con = sqlite3.connect(ruta)
    con.execute("""CREATE TABLE IF NOT EXISTS negocios (
        place_id TEXT PRIMARY KEY, nombre TEXT, telefono TEXT, categoria TEXT,
        direccion TEXT, web TEXT, rating TEXT, resenas TEXT, ciudad TEXT,
        rubro TEXT, nicho TEXT, busqueda TEXT, url TEXT, fecha TEXT)""")
    # Columna agregada despues: las bases viejas no la tienen.
    if 'rubro' not in [c[1] for c in con.execute('PRAGMA table_info(negocios)')]:
        con.execute("ALTER TABLE negocios ADD COLUMN rubro TEXT DEFAULT ''")
    con.execute("CREATE TABLE IF NOT EXISTS busquedas_hechas (clave TEXT PRIMARY KEY, fecha TEXT)")
    con.commit()
    return con


def place_id(url, nombre="", direccion=""):
    """Identidad estable del negocio. La url de Maps trae !19s<place_id>."""
    if url:
        m = re.search(r'!19s([^!?&]+)', url)
        if m:
            return m.group(1)
        m = re.search(r'!1s(0x[0-9a-f]+:0x[0-9a-f]+)', url)
        if m:
            return m.group(1)
    return f"{nombre}|{direccion}".lower().strip()


def guardar(con, negocios, nicho, ciudad, busqueda, rubro=''):
    """Devuelve cuantos eran nuevos (los repetidos se ignoran)."""
    ahora = datetime.now().isoformat(timespec='seconds')
    filas = [(place_id(n.get('url', ''), n.get('nombre', ''), n.get('direccion', '')),
              n.get('nombre', ''), n.get('telefono', ''), n.get('categoria', ''),
              n.get('direccion', ''), n.get('web', ''), n.get('rating', ''),
              n.get('resenas', ''), ciudad, rubro, nicho, busqueda, n.get('url', ''), ahora)
             for n in negocios]
    antes = con.execute("SELECT COUNT(*) FROM negocios").fetchone()[0]
    # Si ya existe, no lo pisa: solo rellena los campos que estaban vacios. Asi una
    # segunda pasada recupera el telefono que la primera no llego a leer.
    rellenar = ', '.join(
        f"{c} = COALESCE(NULLIF(excluded.{c}, ''), negocios.{c})"
        for c in CAMPOS if c not in ('place_id', 'fecha'))
    con.executemany(
        f"INSERT INTO negocios ({','.join(CAMPOS)}) VALUES ({','.join('?' * len(CAMPOS))}) "
        f"ON CONFLICT(place_id) DO UPDATE SET {rellenar}", filas)
    con.commit()
    return con.execute("SELECT COUNT(*) FROM negocios").fetchone()[0] - antes


def ya_hecha(con, clave):
    return con.execute("SELECT 1 FROM busquedas_hechas WHERE clave=?", (clave,)).fetchone() is not None


def marcar_hecha(con, clave):
    con.execute("INSERT OR REPLACE INTO busquedas_hechas VALUES (?,?)",
                (clave, datetime.now().isoformat(timespec='seconds')))
    con.commit()


def total(con):
    return con.execute("SELECT COUNT(*) FROM negocios").fetchone()[0]


def por_nicho(con):
    """{nicho: cuantas filas tiene en la planilla}. Decide si un nicho ya llego
    al tope y hay que dejar de scrapearlo.

    Cuenta por el nicho que decide el clasificador, NO por la columna nicho de
    la base. Esa columna guarda que busqueda encontro al negocio, que es otra
    cosa: "Cejas y pestañas" tiene 3794 filas por busqueda y 193 en la planilla,
    porque casi todo lo que devuelve esa busqueda Maps lo categoriza como centro
    de estetica. Contando crudo, el tope frenaba justo los nichos mas flacos.

    Mismo criterio que excel.exportar (solo con telefono, un telefono una fila),
    asi el tope se mide contra lo que el vendedor realmente ve. Queda ~0.2% por
    debajo del Excel porque no completa codigos de area; para un umbral sobra.
    """
    from modules.clasificador import clasificar
    cuenta, vistos = {}, set()
    # Por reseñas desc: si dos fichas comparten telefono, exportar se queda con
    # la de mas reseñas. Mismo desempate aca para contar la misma que ella.
    for categoria, nicho, nombre, telefono in con.execute(
            "SELECT categoria, nicho, nombre, telefono FROM negocios "
            "WHERE telefono != '' ORDER BY CAST(resenas AS INTEGER) DESC"):
        if telefono in vistos:
            continue
        vistos.add(telefono)
        decidido = clasificar(categoria, nicho, nombre)[0]
        if decidido:                      # None = descartado, no va a la planilla
            cuenta[decidido] = cuenta.get(decidido, 0) + 1
    return cuenta


def completos(con):
    """place_ids que ya tienen telefono: no hace falta volver a abrir su ficha.
    Los que quedaron sin telefono se reintentan, por si fue un fallo puntual."""
    return {r[0] for r in con.execute("SELECT place_id FROM negocios WHERE telefono != ''")}


def demo():
    """El tope se apoya en por_nicho: si vuelve a contar la columna cruda, frena
    los nichos flacos y sigue cavando en los gordos."""
    con = conectar(':memory:')
    filas = [
        # Encontrados buscando "cejas", pero Maps los llama centro de estetica:
        # a la planilla van como Centros de estetica, no como Cejas.
        ('p1', 'Bella', '341111', 'Centro de estética', 'Cejas y pestañas', '50'),
        ('p2', 'Glam', '341222', 'Centro de estética', 'Cejas y pestañas', '30'),
        ('p3', 'Mirada', '341333', 'Salón de cejas', 'Cejas y pestañas', '10'),
        # Mismo telefono que p3: una sola llamada, gana el de mas reseñas.
        ('p4', 'Mirada bis', '341333', 'Salón de cejas', 'Cejas y pestañas', '90'),
        # Sin telefono: no llega a la planilla, no cuenta para el tope.
        ('p5', 'Fantasma', '', 'Salón de cejas', 'Cejas y pestañas', '0'),
        # Un hotel no vende turnos: el clasificador lo descarta.
        ('p6', 'Hotel Spa', '341444', 'Hotel', 'Spas', '20'),
    ]
    con.executemany(
        "INSERT INTO negocios (place_id, nombre, telefono, categoria, nicho, resenas) "
        "VALUES (?,?,?,?,?,?)", filas)
    con.commit()

    c = por_nicho(con)
    assert c.get('Centros de estética') == 2, f'la categoria real no mando: {c}'
    assert c.get('Cejas y pestañas') == 1, f'el telefono repetido conto dos veces: {c}'
    assert 'Spas' not in c, f'conto un descartado: {c}'
    assert sum(c.values()) == 3, f'total mal: {c}'
    con.close()
    print('OK  por_nicho: cuenta lo que ve el vendedor, no lo que encontro la busqueda')


if __name__ == '__main__':
    demo()
