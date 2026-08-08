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


def cargar_busquedas(ruta='config/busquedas.xlsx'):
    df = pd.read_excel(ruta)
    c_nicho, c_busq = _col(df, 'nicho'), _col(df, 'busqueda')
    c_ciudad, c_activo = _col(df, 'ciudad', 'zona'), _col(df, 'activo')
    c_rubro = _col(df, 'rubro', 'categoria')
    if not c_nicho or not c_busq:
        raise ValueError(f"Faltan columnas Nicho/Busqueda en {ruta}. Hay: {df.columns.tolist()}")

    if c_activo:
        df = df[df[c_activo].astype(str).str.strip().str.lower().isin(AFIRMATIVOS)]

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

    busquedas = cargar_busquedas()
    if solo:
        busquedas = busquedas[:solo]
    if not busquedas:
        print("No hay busquedas activas en config/busquedas.xlsx")
        return

    con = db.conectar()
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


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Scraper de Google Maps")
    ap.add_argument('--minutos', type=int, default=0,
                    help="cortar limpio despues de N minutos (0 = sin limite)")
    ap.add_argument('--solo', type=int, default=None,
                    help="procesar solo las primeras N busquedas")
    args = ap.parse_args()
    ejecutar_scraper(solo=args.solo, minutos_max=args.minutos)
