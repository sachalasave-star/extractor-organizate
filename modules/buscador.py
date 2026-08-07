from urllib.parse import quote_plus
from playwright.sync_api import Page


def buscar_en_maps(page: Page, termino: str):
    """Navega directo a la url de busqueda.

    Tipear en el input y apretar Enter no sirve: Maps toma la sugerencia
    resaltada del autocompletado, no lo tipeado. Con "Barberia en Rosario"
    terminaba en el mapa de la ciudad Rosario, sin resultados.
    """
    # gl=AR fuerza region Argentina: sin eso, desde un servidor fuera del pais
    # Google devuelve los telefonos en formato internacional (+54341...) y la
    # base termina con dos formatos mezclados.
    page.goto(f"https://www.google.com/maps/search/{quote_plus(termino)}?hl=es&gl=AR",
              wait_until="domcontentloaded")
    page.wait_for_selector('div[role="feed"]', timeout=20000)
    page.wait_for_timeout(2000)
