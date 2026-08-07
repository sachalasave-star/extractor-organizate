"""Regenera el xlsx completo desde SQLite. Una hoja por nicho."""
import os
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

COLUMNAS = ['nombre', 'telefono', 'categoria', 'direccion', 'web',
            'rating', 'resenas', 'ciudad', 'busqueda', 'url']


def _nombre_hoja(nicho, usados):
    hoja = re.sub(r'[\[\]:*?/\\]', '-', str(nicho)).strip()[:31] or "Sin nicho"
    base, n = hoja, 2
    while hoja in usados:                      # Excel no admite hojas repetidas
        hoja = f"{base[:28]}_{n}"
        n += 1
    usados.add(hoja)
    return hoja


def exportar(con, archivo_salida="output/Organizate.xlsx"):
    """Solo exporta negocios CON telefono. Los que no tienen quedan igual en la
    base: si en una pasada posterior aparece el numero, el upsert lo rellena y
    entran solos al Excel. Descartarlos de la base seria irreversible."""
    nichos = [r[0] for r in con.execute(
        "SELECT DISTINCT nicho FROM negocios WHERE telefono != '' ORDER BY nicho")]
    if not nichos:
        print("No hay negocios con telefono para exportar.")
        return

    os.makedirs(os.path.dirname(archivo_salida), exist_ok=True)
    usados, hojas = set(), []
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
        for nicho in nichos:
            df = pd.read_sql_query(
                f"SELECT {','.join(COLUMNAS)} FROM negocios "
                "WHERE nicho=? AND telefono != '' ORDER BY nombre",
                con, params=(nicho,))
            hoja = _nombre_hoja(nicho, usados)
            df.to_excel(writer, sheet_name=hoja, index=False)
            hojas.append((hoja, len(df)))

    _formatear(archivo_salida, [h for h, _ in hojas])
    total = sum(n for _, n in hojas)
    sin_tel = con.execute("SELECT COUNT(*) FROM negocios WHERE telefono = ''").fetchone()[0]
    print(f"Excel: {total} negocios con telefono en {len(hojas)} hojas -> {archivo_salida}")
    if sin_tel:
        print(f"       ({sin_tel} sin telefono quedaron en la base, fuera del Excel)")


def _formatear(archivo, hojas):
    try:
        wb = load_workbook(archivo)
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        fino = Side(style='thin')
        borde = Border(left=fino, right=fino, top=fino, bottom=fino)
        col_tel = COLUMNAS.index('telefono') + 1
        for hoja in hojas:
            ws = wb[hoja]
            ws.freeze_panes = "A2"
            # Formato texto: sin esto Excel puede leer 03416359476 como numero
            # y comerse el 0 inicial, que en Argentina hace inservible el numero.
            for fila in range(2, ws.max_row + 1):
                ws.cell(row=fila, column=col_tel).number_format = '@'
            for cell in ws[1]:
                cell.font, cell.fill, cell.border = header_font, header_fill, borde
                cell.alignment = Alignment(horizontal='center', vertical='center')
            for column in ws.columns:
                largo = max((len(str(c.value)) for c in column if c.value is not None), default=0)
                ws.column_dimensions[column[0].column_letter].width = min(largo + 2, 50)
        wb.save(archivo)
    except Exception as e:
        print(f"Aviso: no se pudo formatear el Excel: {e}")
