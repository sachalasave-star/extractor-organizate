"""Check rapido: una busqueda, 5 negocios, y falla ruidosamente si los
selectores de Google Maps cambiaron. Correr esto antes de una tanda larga.

    python verificar.py
"""
import os
import re
import sys
from playwright.sync_api import sync_playwright

from modules import db
from modules.buscador import buscar_en_maps
from modules.scroll import hacer_scroll
from modules.extractor import extraer_negocios

# Psicologos a proposito: sus nombres traen tildes ("Psicóloga ..."), que es
# donde fallaba el match del panel. Con barberias (todo ASCII) el bug no salta.
TERMINO = "Psicologo en Rosario"
N = 6


def check_db():
    if os.path.exists("output/_test.db"):
        os.remove("output/_test.db")
    con = db.conectar("output/_test.db")
    muestra = [{'nombre': 'A', 'telefono': '341 1', 'url': 'x/data=!19sABC?hl=es'},
               {'nombre': 'A bis', 'telefono': '', 'url': 'y/data=!19sABC?hl=es'}]
    db.guardar(con, muestra[:1], 'Test', 'Rosario', TERMINO)
    db.guardar(con, muestra, 'Test', 'Rosario', TERMINO)
    assert db.total(con) == 1, f"dedup rota: {db.total(con)} filas, esperaba 1"
    # el segundo llega sin telefono y no debe borrar el que ya estaba
    tel = con.execute("SELECT telefono FROM negocios").fetchone()[0]
    assert tel == '341 1', f"upsert piso un dato bueno con vacio: {tel!r}"
    assert db.place_id("z/data=!19sChIJxyz?authuser=0") == "ChIJxyz"
    con.close()
    print("OK  db: dedup, upsert y place_id")


def check_scraper():
    en_ci = os.environ.get("CI") == "true"      # GitHub Actions lo setea solo
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=en_ci,
                                    args=['--disable-blink-features=AutomationControlled'])
        page = browser.new_page(viewport={"width": 1400, "height": 900})
        try:
            buscar_en_maps(page, TERMINO)
            hacer_scroll(page, max_resultados=N)
            negocios = extraer_negocios(page, max_negocios=N)
        except Exception as e:
            page.screenshot(path="output/fallo.png")
            raise AssertionError(
                f"no se pudo leer el listado ({type(e).__name__}). "
                "Si corre en la nube, lo mas probable es que Google haya bloqueado la IP: "
                "mirar output/fallo.png") from e
        finally:
            try:                       # que un screenshot fallido no tape el error real
                page.screenshot(path="output/verificacion.png")
            except Exception:
                pass
            browser.close()

    assert negocios, "no se extrajo ningun negocio"
    for n in negocios:
        print(f"  {n['nombre'][:32]:34} tel={n['telefono'] or '-':16} cat={n['categoria'] or '-'}")

    con_tel = [n for n in negocios if n['telefono']]
    assert len(con_tel) >= 3, (
        f"solo {len(con_tel)}/{len(negocios)} con telefono. "
        "Revisar button[data-item-id^='phone:tel:'] en la ficha.")

    # La ficha no abrio => categoria vacia. Con tildes en el nombre esto fallaba.
    con_tildes = [n for n in negocios if any(c in n['nombre'] for c in 'áéíóúñÁÉÍÓÚÑ')]
    if con_tildes:
        abiertas = [n for n in con_tildes if n['categoria']]
        assert len(abiertas) >= len(con_tildes) * 0.7, (
            f"solo {len(abiertas)}/{len(con_tildes)} fichas abiertas en nombres con tilde. "
            "El match del panel no soporta acentos.")

    # regresiones del bug viejo: categoria traia el rating, direccion traia "(12)"
    for n in negocios:
        assert not re.fullmatch(r'[\d.,]+', n['categoria'] or 'ok'), \
            f"categoria es un numero ({n['categoria']!r}): volvio el bug de los spans"
        assert not re.match(r'^\(\d', n['direccion'] or 'ok'), \
            f"direccion trae el conteo de resenas ({n['direccion']!r})"

    print(f"OK  scraper: {len(con_tel)}/{len(negocios)} con telefono, categoria y direccion sanas")


if __name__ == "__main__":
    try:
        check_db()
        check_scraper()
        print("\nTodo OK. Podes correr: python scraper.py")
    except AssertionError as e:
        print(f"\nFALLO: {e}")
        sys.exit(1)
