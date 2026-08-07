from playwright.sync_api import Page


def hacer_scroll(page: Page, max_resultados=10000):
    """Baja el listado hasta el final real: cuenta tarjetas en vez de mirar
    scrollTop, que se estanca mientras Maps todavia carga async."""
    try:
        panel = page.locator('div[role="feed"]')
        panel.wait_for(timeout=20000)
    except Exception:
        print("   No se encontro el contenedor de resultados.")
        return 0

    previo, quietas = 0, 0
    while quietas < 3:
        panel.evaluate("(e) => e.scrollBy(0, 5000)")
        page.wait_for_timeout(1200)
        actual = page.locator('div[role="article"]').count()

        if actual >= max_resultados:
            break
        # Maps corta cerca de 120 resultados y muestra el cartel de fin de lista
        if page.get_by_text("Llegaste al final de la lista").count() > 0:
            break

        quietas = quietas + 1 if actual == previo else 0
        previo = actual

    total = page.locator('div[role="article"]').count()
    print(f"   Listado cargado: {total} resultados.")
    return total
