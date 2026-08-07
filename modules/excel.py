"""Regenera el xlsx completo desde SQLite. Una hoja por rubro + una con todo."""
import os
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

COLUMNAS = ['nicho', 'nombre', 'telefono', 'categoria', 'ciudad', 'direccion',
            'web', 'rating', 'resenas', 'busqueda', 'url']
TITULOS = {'rubro': 'Rubro', 'nicho': 'Nicho', 'nombre': 'Negocio', 'telefono': 'Teléfono',
           'categoria': 'Rubro en Maps', 'ciudad': 'Ciudad', 'direccion': 'Dirección',
           'web': 'Web', 'rating': 'Puntaje', 'resenas': 'Reseñas',
           'busqueda': 'Búsqueda', 'url': 'Link en Maps'}
ANCHOS = {'nicho': 22, 'nombre': 42, 'telefono': 16, 'categoria': 24, 'ciudad': 18,
          'direccion': 44, 'web': 34, 'rating': 8, 'resenas': 9, 'busqueda': 26, 'url': 16}


def _nombre_hoja(texto, usados):
    hoja = re.sub(r'[\[\]:*?/\\]', '-', str(texto)).strip()[:31] or "Sin rubro"
    base, n = hoja, 2
    while hoja in usados:
        hoja = f"{base[:28]}_{n}"
        n += 1
    usados.add(hoja)
    return hoja


def exportar(con, archivo_salida="output/Organizate.xlsx"):
    """Solo exporta negocios CON telefono. Los que no tienen quedan igual en la
    base: si en una pasada posterior aparece el numero, el upsert lo rellena y
    entran solos al Excel. Descartarlos de la base seria irreversible."""
    cols = ','.join(COLUMNAS)
    todo = pd.read_sql_query(
        f"SELECT rubro,{cols} FROM negocios WHERE telefono != '' "
        "ORDER BY rubro, nicho, ciudad, nombre", con)
    if todo.empty:
        print("No hay negocios con telefono para exportar.")
        return

    os.makedirs(os.path.dirname(archivo_salida), exist_ok=True)
    usados, hojas = set(), []
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
        # Primero "Todos": es la que sirve para filtrar y cruzar.
        hoja = _nombre_hoja("Todos", usados)
        todo.rename(columns=TITULOS).to_excel(writer, sheet_name=hoja, index=False)
        hojas.append((hoja, len(todo), True))

        for rubro, grupo in todo.groupby('rubro', sort=True):
            hoja = _nombre_hoja(rubro or "Sin rubro", usados)
            grupo[COLUMNAS].rename(columns=TITULOS).to_excel(
                writer, sheet_name=hoja, index=False)
            hojas.append((hoja, len(grupo), False))

    _formatear(archivo_salida, hojas)
    sin_tel = con.execute("SELECT COUNT(*) FROM negocios WHERE telefono = ''").fetchone()[0]
    print(f"Excel: {len(todo)} negocios con telefono -> {archivo_salida}")
    for h, n, _ in hojas[1:]:
        print(f"       {h}: {n}")
    if sin_tel:
        print(f"       ({sin_tel} sin telefono quedaron en la base, fuera del Excel)")


def _formatear(archivo, hojas):
    try:
        wb = load_workbook(archivo)
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill("solid", start_color="2F5597")
        fino = Side(style='thin', color="D0D0D0")
        borde = Border(left=fino, right=fino, top=fino, bottom=fino)

        for hoja, _, con_rubro in hojas:
            ws = wb[hoja]
            ws.freeze_panes = "B2"                # fija encabezado y primera columna
            ws.auto_filter.ref = ws.dimensions    # filtros en todas las columnas

            claves = (['rubro'] if con_rubro else []) + COLUMNAS
            for i, clave in enumerate(claves, start=1):
                letra = get_column_letter(i)
                ws.column_dimensions[letra].width = 20 if clave == 'rubro' else ANCHOS[clave]
                if clave == 'telefono':
                    # Formato texto: sin esto Excel lee 03416359476 como numero y
                    # se come el 0 inicial, que hace inservible el numero.
                    for fila in range(2, ws.max_row + 1):
                        ws.cell(row=fila, column=i).number_format = '@'
                elif clave == 'url':
                    for fila in range(2, ws.max_row + 1):
                        celda = ws.cell(row=fila, column=i)
                        if celda.value:
                            celda.hyperlink, celda.value = celda.value, "ver en Maps"
                            celda.font = Font(color="0563C1", underline="single")

            for celda in ws[1]:
                celda.font, celda.fill, celda.border = header_font, header_fill, borde
                celda.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 22
        wb.save(archivo)
    except Exception as e:
        print(f"Aviso: no se pudo formatear el Excel: {e}")
