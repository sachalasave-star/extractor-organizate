"""Extrae negocios del listado de Google Maps abriendo la ficha de cada uno.

El telefono, la direccion y el tipo de negocio SOLO existen en la ficha, no en
la tarjeta del listado. Se leen por atributo `data-item-id`, que no depende del
idioma ni de las clases ofuscadas de Google.
"""
from playwright.sync_api import Page

from modules.db import place_id

TARJETA = 'div[role="article"]'


def _attr(loc, nombre_attr, timeout=1500):
    try:
        if loc.count() > 0:
            return (loc.first.get_attribute(nombre_attr, timeout=timeout) or "").strip()
    except Exception:
        pass
    return ""


def _texto(loc, timeout=1500):
    try:
        if loc.count() > 0:
            return (loc.first.text_content(timeout=timeout) or "").strip()
    except Exception:
        pass
    return ""


def _leer_ficha(panel):
    """Lee los datos del panel abierto. Scopeado al panel del negocio correcto
    para no arrastrar el telefono de la ficha anterior si esta no cerro."""
    telefono = _attr(panel.locator('button[data-item-id^="phone:tel:"]'), 'data-item-id')
    telefono = telefono.replace('phone:tel:', '') if telefono else ""

    direccion = _attr(panel.locator('button[data-item-id="address"]'), 'aria-label')
    if ':' in direccion:
        direccion = direccion.split(':', 1)[1].strip()

    web = _attr(panel.locator('a[data-item-id="authority"]'), 'href')

    categoria = _texto(panel.locator('button[jsaction*="category"]'))
    if not categoria:
        categoria = _texto(panel.locator('.DkEaL'))
    # Cuando el negocio no declara rubro, Maps pone en su lugar botones de
    # edicion ("Añadir categoría", "Añadir una foto") y se colaban como categoria.
    if categoria.lower().startswith(('añadir', 'agregar', 'reclamar', 'sugerir')):
        categoria = ""

    return telefono, direccion, web, categoria


def extraer_negocios(page: Page, max_negocios=0, ya_tengo=frozenset()) -> list:
    """`ya_tengo`: place_ids que ya estan completos en la base. Se saltean sin
    abrir la ficha, que es lo caro (~3s c/u). Con varias keywords por nicho el
    solapamiento es alto, asi que esto es la diferencia entre horas y horas."""
    negocios = []
    sin_telefono = salteados = 0

    try:
        page.wait_for_selector(TARJETA, timeout=10000)
    except Exception:
        print("   No se encontraron tarjetas.")
        return []

    total = page.locator(TARJETA).count()
    if max_negocios > 0:
        total = min(total, max_negocios)
    print(f"   {total} tarjetas a procesar.")

    for i in range(total):
        try:
            # Re-resolver por indice: al abrir/cerrar fichas el feed se re-renderiza
            # y un handle guardado al inicio queda stale.
            enlace = page.locator(TARJETA).nth(i).locator('a.hfpxzc')
            if enlace.count() == 0:
                continue

            nombre = _attr(enlace, 'aria-label')
            url = _attr(enlace, 'href')
            if not nombre:
                continue

            # Ya lo tenemos completo: no vale la pena abrir la ficha de nuevo
            if place_id(url, nombre) in ya_tengo:
                salteados += 1
                continue

            tarjeta = page.locator(TARJETA).nth(i)
            rating = _texto(tarjeta.locator('span.MW4etd'))
            resenas = _texto(tarjeta.locator('span.UY7F9')).strip('()').replace('.', '').replace(',', '')

            telefono = direccion = web = categoria = ""
            try:
                enlace.first.click(timeout=5000)
                # El panel de la ficha es un role=main con el nombre del negocio como
                # aria-label (el listado es OTRO role=main, con aria-label vacio).
                # get_by_role y no un selector CSS: los nombres con tilde no se pueden
                # escapar a mano para CSS ("ó" no es escape valido, "\f3 " si).
                panel = page.get_by_role("main", name=nombre, exact=True)
                panel.first.wait_for(state="visible", timeout=6000)
                page.wait_for_timeout(400)
                telefono, direccion, web, categoria = _leer_ficha(panel.first)
            except Exception:
                pass
            finally:
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(300)
                except Exception:
                    pass

            if not telefono:
                sin_telefono += 1

            negocios.append({
                'nombre': nombre, 'telefono': telefono, 'categoria': categoria,
                'direccion': direccion, 'web': web, 'rating': rating,
                'resenas': resenas, 'url': url,
            })

            if (i + 1) % 10 == 0:
                print(f"   {i + 1}/{total}...")

        except Exception:
            continue

    if negocios and sin_telefono == len(negocios):
        print(f"   AVISO: 0 telefonos en {len(negocios)} negocios. Revisar selectores de la ficha.")
    elif sin_telefono:
        print(f"   {sin_telefono}/{len(negocios)} sin telefono.")
    if salteados:
        print(f"   {salteados} ya estaban en la base (no se abrieron).")
    print(f"   Extraidos {len(negocios)} negocios.")
    return negocios
