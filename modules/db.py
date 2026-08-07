"""Base maestra en SQLite. Dedup por place_id via INSERT OR IGNORE."""
import sqlite3
import re
from datetime import datetime

RUTA = "output/negocios.db"

CAMPOS = ['place_id', 'nombre', 'telefono', 'categoria', 'direccion', 'web',
          'rating', 'resenas', 'ciudad', 'nicho', 'busqueda', 'url', 'fecha']


def conectar(ruta=RUTA):
    import os
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    con = sqlite3.connect(ruta)
    con.execute("""CREATE TABLE IF NOT EXISTS negocios (
        place_id TEXT PRIMARY KEY, nombre TEXT, telefono TEXT, categoria TEXT,
        direccion TEXT, web TEXT, rating TEXT, resenas TEXT, ciudad TEXT,
        nicho TEXT, busqueda TEXT, url TEXT, fecha TEXT)""")
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


def guardar(con, negocios, nicho, ciudad, busqueda):
    """Devuelve cuantos eran nuevos (los repetidos se ignoran)."""
    ahora = datetime.now().isoformat(timespec='seconds')
    filas = [(place_id(n.get('url', ''), n.get('nombre', ''), n.get('direccion', '')),
              n.get('nombre', ''), n.get('telefono', ''), n.get('categoria', ''),
              n.get('direccion', ''), n.get('web', ''), n.get('rating', ''),
              n.get('resenas', ''), ciudad, nicho, busqueda, n.get('url', ''), ahora)
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
