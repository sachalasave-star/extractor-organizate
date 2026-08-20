"""Scraper de Google Maps: nombre, telefono y tipo de negocio a SQLite + Excel.

Corre las busquedas activas de config/busquedas.xlsx (Nicho / Busqueda / Ciudad).
Guarda despues de cada busqueda, deduplica por place_id y se puede reanudar.
"""
import json
import random
import time
import pandas as pd
from playwright.sync_api import sync_playwright

from modules import db
from modules.buscador import buscar_en_maps
from modules.scroll import hacer_scroll
from modules.extractor import extraer_negocios
from modules.excel import exportar

AFIRMATIVOS = {'si', 'sí', 's', 'x', '1', 'true', 'yes', 'y'}


def _config():
    try:
        with open('config/configuracion.json', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _col(df, *nombres):
    """Devuelve la columna sin importar tildes/mayusculas (Busqueda vs Búsqueda)."""
    for c in df.columns:
        if str(c).strip().lower().replace('ú', 'u').replace('í', 'i') in nombres:
            return c
    return None


def cargar_busquedas(ruta='config/busquedas.xlsx', ya_tengo=None, tope=0, sin_tope=()):
    """ya_tengo = {nicho: cuantos hay} y tope > 0 sacan de la cola los nichos
    que ya estan servidos.

    Es el freno de mano. Sin esto el scraper sigue cavando en Barberias (6829)
    mientras Opticas tiene 199, porque el orden del archivo manda y Barberias
    tiene mas busquedas. Con el tope, el nicho lleno desaparece de la cola y el
    tiempo se va solo a los que faltan.

    sin_tope = nichos que siguen igual aunque pasen el tope. El tope nunca tira
    un negocio ya scrapeado: solo decide que busquedas NUEVAS se lanzan. Pero
    una busqueda ya encolada de un nicho lleno es trabajo que igual sirve, asi
    que esta la lista para dejarla pasar."""
    df = pd.read_excel(ruta)
    c_nicho, c_busq = _col(df, 'nicho'), _col(df, 'busqueda')
    c_ciudad, c_activo = _col(df, 'ciudad', 'zona'), _col(df, 'activo')
    c_rubro = _col(df, 'rubro', 'categoria')
    if not c_nicho or not c_busq:
        raise ValueError(f"Faltan columnas Nicho/Busqueda en {ruta}. Hay: {df.columns.tolist()}")

    if c_activo:
        df = df[df[c_activo].astype(str).str.strip().str.lower().isin(AFIRMATIVOS)]

    if tope and ya_tengo:
        libres = {str(n).strip() for n in sin_tope}
        llenos = {n for n, c in ya_tengo.items() if c >= tope and n not in libres}
        frenados = df[c_nicho].astype(str).str.strip().isin(llenos)
        if frenados.any():
            print(f"{len(llenos)} nichos llegaron al tope de {tope} y quedan "
                  f"afuera: {', '.join(sorted(llenos))}")
            df = df[~frenados]

    filas = []
    for _, f in df.iterrows():
        ciudad = str(f[c_ciudad]).strip() if c_ciudad and pd.notna(f[c_ciudad]) else ""
        busqueda = str(f[c_busq]).strip()
        filas.append({
            'nicho': str(f[c_nicho]).strip(),
            'rubro': str(f[c_rubro]).strip() if c_rubro and pd.notna(f[c_rubro]) else "",
            'busqueda': busqueda,
            'ciudad': ciudad,
            # ", Argentina" ancla el pais: sin eso "en Córdoba" devolvia negocios
            # de Cordoba, España (91 casos con telefono +34).
            'termino': f"{busqueda} en {ciudad}, Argentina" if ciudad else busqueda,
        })
    return filas


def _nueva_pagina(p, headless):
    browser = p.chromium.launch(
        headless=headless, args=['--disable-blink-features=AutomationControlled'])
    page = browser.new_page(viewport={"width": 1400, "height": 900})
    return browser, page


def ejecutar_scraper(solo=None, reanudar=True, minutos_max=0):
    """minutos_max > 0 corta limpio al llegar al limite. En GitHub Actions es
    obligatorio: si el job se mata por timeout, el commit final nunca ocurre y
    se pierde el trabajo de toda la corrida."""
    cfg = _config()
    max_res = int(cfg.get('max_resultados_por_busqueda', 10000))
    headless = bool(cfg.get('headless', False))
    limite = time.monotonic() + minutos_max * 60 if minutos_max else None
    tope = int(cfg.get('tope_por_nicho', 0))
    sin_tope = cfg.get('nichos_sin_tope', [])

    con = db.conectar()
    busquedas = cargar_busquedas(ya_tengo=db.por_nicho(con), tope=tope,
                                 sin_tope=sin_tope)
    if solo:
        busquedas = busquedas[:solo]
    if not busquedas:
        print(f"Nada que hacer: todos los nichos activos llegaron al tope de "
              f"{tope}. Apaga el workflow." if tope else
              "No hay busquedas activas en config/busquedas.xlsx")
        con.close()
        return
    print(f"{len(busquedas)} busquedas activas | {db.total(con)} negocios ya en la base\n")

    with sync_playwright() as p:
        browser, page = _nueva_pagina(p, headless)
        try:
            for i, b in enumerate(busquedas, 1):
                if limite and time.monotonic() > limite:
                    print(f"\nLimite de {minutos_max} min alcanzado. Corte limpio "
                          f"en {i - 1}/{len(busquedas)}; el resto queda para la proxima.")
                    break

                clave = f"{b['nicho']}|{b['termino']}"
                if reanudar and db.ya_hecha(con, clave):
                    print(f"[{i}/{len(busquedas)}] {b['termino']} - ya hecha, salteando")
                    continue

                print(f"\n[{i}/{len(busquedas)}] {b['nicho']} -> {b['termino']}")
                try:
                    buscar_en_maps(page, b['termino'])
                    hacer_scroll(page, max_res)
                    negocios = extraer_negocios(page, max_negocios=max_res,
                                                ya_tengo=db.completos(con))

                    if negocios:
                        nuevos = db.guardar(con, negocios, b['nicho'], b['ciudad'],
                                            b['termino'], b['rubro'])
                        db.marcar_hecha(con, clave)
                        print(f"   +{nuevos} nuevos ({len(negocios) - nuevos} repetidos) | total {db.total(con)}")
                    else:
                        print("   Sin resultados.")

                    # ponytail: pausa fija; rotar user-agent/proxy solo si Maps empieza a cortar
                    time.sleep(random.uniform(3, 8))

                except KeyboardInterrupt:
                    raise
                except Exception as e:
                    print(f"   Error: {e}")
                    if page.is_closed() or not browser.is_connected():
                        print("   Reiniciando navegador...")
                        try:
                            browser.close()
                        except Exception:
                            pass
                        browser, page = _nueva_pagina(p, headless)
        except KeyboardInterrupt:
            print("\nInterrumpido. Lo guardado se conserva.")
        finally:
            browser.close()

    exportar(con)
    print(f"Listo. {db.total(con)} negocios en la base.")
    con.close()


def demo():
    """El tope es la unica logica nueva: que saque al nicho lleno y no toque al
    flaco. Si esto se rompe, el scraper vuelve a cavar en Barberias para siempre."""
    import os
    ruta = 'output/_tope_test.xlsx'
    pd.DataFrame([
        {'Rubro': 'Estética y belleza', 'Nicho': 'Barberías', 'Busqueda': 'barbería',
         'Ciudad': 'Rosario', 'Activo': 'Si'},
        {'Rubro': 'Otros servicios', 'Nicho': 'Ópticas', 'Busqueda': 'óptica',
         'Ciudad': 'Rosario', 'Activo': 'Si'},
        {'Rubro': 'Deporte y movimiento', 'Nicho': 'Gimnasios', 'Busqueda': 'gym',
         'Ciudad': 'Rosario', 'Activo': 'No'},
    ]).to_excel(ruta, index=False)
    tengo = {'Barberías': 6829, 'Ópticas': 199}

    nichos = lambda r: {b['nicho'] for b in r}
    assert nichos(cargar_busquedas(ruta)) == {'Barberías', 'Ópticas'}, 'Activo=No no filtro'
    assert nichos(cargar_busquedas(ruta, tengo, tope=2000)) == {'Ópticas'},         'el tope no saco a Barberías'
    # tope=0 es "sin freno": tiene que dejar todo como estaba.
    assert nichos(cargar_busquedas(ruta, tengo, tope=0)) == {'Barberías', 'Ópticas'}
    # Justo en el numero ya cuenta como lleno, si no nunca frena.
    assert nichos(cargar_busquedas(ruta, {'Ópticas': 2000}, tope=2000)) == {'Barberías'}
    # El termino tiene que seguir armandose igual para no rehacer lo ya buscado.
    assert cargar_busquedas(ruta, tengo, tope=2000)[0]['termino'] == 'óptica en Rosario, Argentina'

    # La lista de exentos deja pasar al lleno sin aflojar el tope para el resto.
    assert nichos(cargar_busquedas(ruta, tengo, 2000, ['Barberías'])) == {'Barberías', 'Ópticas'}
    assert nichos(cargar_busquedas(ruta, tengo, 2000, [' Barberías '])) == {'Barberías', 'Ópticas'},         'no aguanta espacios de mas en la config'

    os.remove(ruta)
    print('OK  tope: frena al lleno, deja pasar al flaco y respeta los exentos')


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Scraper de Google Maps")
    ap.add_argument('--minutos', type=int, default=0,
                    help="cortar limpio despues de N minutos (0 = sin limite)")
    ap.add_argument('--solo', type=int, default=None,
                    help="procesar solo las primeras N busquedas")
    ap.add_argument('--test', action='store_true', help="correr el check del tope")
    args = ap.parse_args()
    if args.test:
        demo()
    else:
        ejecutar_scraper(solo=args.solo, minutos_max=args.minutos)
