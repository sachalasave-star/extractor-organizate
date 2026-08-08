"""Arma la planilla de ventas en Google Sheets: una hoja por nicho, con
desplegables de vendedor y estado, colores por estado y un ranking en vivo.

    python armar_planilla.py            # arma o rehace la estructura
    python armar_planilla.py --test     # solo la logica, sin tocar Google

CUIDADO: rehace las hojas de nicho desde cero. Se corre una vez al principio;
para el dia a dia esta sincronizar_sheets.py, que solo agrega filas.
"""
import os
import sys
import time

import pandas as pd

from modules.planilla import (COLUMNAS, ANCHOS, CLAVE, VENDEDORES, ESTADOS,
                              NOMBRES_ESTADO, CONFIG, RANKING, FILA_VENDEDORES,
                              FILA_ESTADOS, COL_VENDEDOR, COL_ESTADO, COL_TELEFONO,
                              IDX, fila_desde)
from modules.estilo import (FUENTE, FUENTE_DATOS, TAM_ENCABEZADO, TAM_DATOS,
                            color_rubro, hex_a_rgb, TINTA, TINTA_SUAVE, LINEA, BLANCO)
from sincronizar_sheets import _credenciales, GENERADO

MAX_VENDEDORES = 40      # margen del desplegable, no cuesta nada
# Filas del ranking. Cada una son 6 formulas x 43 hojas: con 40 la planilla se
# arrastra, con 12 hay margen de sobra sobre los 4 vendedores actuales.
FILAS_RANKING = 12


def _abrir_libro():
    import gspread
    from google.oauth2.service_account import Credentials
    sheet_id = os.environ.get("SHEET_ID")
    if not sheet_id:
        sys.exit("Falta SHEET_ID")
    cred = Credentials.from_service_account_info(
        _credenciales(), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(cred).open_by_key(sheet_id)


def _rgb(c):
    return {"red": c[0], "green": c[1], "blue": c[2]}


def reintentar(fn, *a, **kw):
    """Sheets corta a las 60 escrituras por minuto. Espera y sigue en vez de
    dejar la planilla a medio armar."""
    for intento in range(6):
        try:
            return fn(*a, **kw)
        except Exception as e:
            if '429' not in str(e) or intento == 5:
                raise
            espera = 20 * (intento + 1)
            print(f"   (limite de Google, esperando {espera}s)")
            time.sleep(espera)


def _formato_hoja(sid, filas, rubro=''):
    """Requests de formato para una hoja de nicho."""
    ultima = max(filas + 1, 2000)        # deja margen para que crezca sola
    oscuro, claro = color_rubro(rubro)
    reqs = [
        # Encabezado, con el color del rubro
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb(hex_a_rgb(oscuro)),
                "textFormat": {"bold": True, "foregroundColor": _rgb(hex_a_rgb(BLANCO)),
                               "fontSize": TAM_ENCABEZADO, "fontFamily": FUENTE},
                "verticalAlignment": "MIDDLE",
                "horizontalAlignment": "LEFT",
                "padding": {"left": 10, "right": 10}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,"
                      "horizontalAlignment,padding)"}},
        # Cuerpo: una sola tipografia, texto tranquilo
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": ultima},
            "cell": {"userEnteredFormat": {
                "textFormat": {"fontFamily": FUENTE_DATOS, "fontSize": TAM_DATOS,
                               "foregroundColor": _rgb(hex_a_rgb(TINTA))},
                "verticalAlignment": "MIDDLE",
                "padding": {"left": 10, "right": 10}}},
            "fields": "userEnteredFormat(textFormat,verticalAlignment,padding)"}},
        # Filas alternadas en el tono claro del rubro: guia el ojo en listas largas
        {"addBanding": {"bandedRange": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": ultima,
                      "startColumnIndex": 0, "endColumnIndex": len(COLUMNAS)},
            "rowProperties": {
                "headerColor": _rgb(hex_a_rgb(oscuro)),
                "firstBandColor": _rgb(hex_a_rgb(BLANCO)),
                "secondBandColor": _rgb(hex_a_rgb(claro))}}}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid, "tabColor": _rgb(hex_a_rgb(oscuro)),
                           "gridProperties": {"frozenRowCount": 1, "frozenColumnCount": 1}},
            "fields": "tabColor,gridProperties(frozenRowCount,frozenColumnCount)"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 34}, "fields": "pixelSize"}},
        # El nombre del negocio en seminegrita: es lo primero que se lee
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": ultima,
                      "startColumnIndex": 0, "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
            "fields": "userEnteredFormat.textFormat.bold"}},
        # Telefono monoespaciado: se leen en columna, uno abajo del otro
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": ultima,
                      "startColumnIndex": IDX[CLAVE], "endColumnIndex": IDX[CLAVE] + 1},
            "cell": {"userEnteredFormat": {
                "textFormat": {"fontFamily": "Roboto Mono", "fontSize": TAM_DATOS}}},
            "fields": "userEnteredFormat.textFormat"}},
        # Categoria y ciudad son contexto, no protagonistas
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": ultima,
                      "startColumnIndex": IDX['Categoría'], "endColumnIndex": IDX['Categoría'] + 1},
            "cell": {"userEnteredFormat": {
                "textFormat": {"foregroundColor": _rgb(hex_a_rgb(TINTA_SUAVE))}}},
            "fields": "userEnteredFormat.textFormat.foregroundColor"}},
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": ultima,
                      "startColumnIndex": IDX['Ciudad'], "endColumnIndex": IDX['Ciudad'] + 1},
            "cell": {"userEnteredFormat": {
                "textFormat": {"foregroundColor": _rgb(hex_a_rgb(TINTA_SUAVE))}}},
            "fields": "userEnteredFormat.textFormat.foregroundColor"}},
        # Estado y vendedor centrados: son etiquetas, no texto
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": ultima,
                      "startColumnIndex": IDX['Vendedor'], "endColumnIndex": IDX['Estado'] + 1},
            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat.horizontalAlignment"}},
        # El telefono como texto: sin esto Sheets se come el 0 inicial
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 1,
                      "startColumnIndex": IDX[CLAVE], "endColumnIndex": IDX[CLAVE] + 1},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "TEXT"}}},
            "fields": "userEnteredFormat.numberFormat"}},
        # Observaciones: recorta el texto largo en vez de estirar la fila
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 1,
                      "startColumnIndex": IDX['Observaciones'],
                      "endColumnIndex": IDX['Observaciones'] + 1},
            "cell": {"userEnteredFormat": {"wrapStrategy": "CLIP"}},
            "fields": "userEnteredFormat.wrapStrategy"}},
        {"setBasicFilter": {"filter": {"range": {
            "sheetId": sid, "startRowIndex": 0, "endColumnIndex": len(COLUMNAS)}}}},
    ]
    for i, ancho in enumerate(ANCHOS):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": ancho}, "fields": "pixelSize"}})

    # Desplegables. Apuntan a Config: agregar un vendedor ahi lo habilita en
    # las 43 hojas de una, sin rehacer ninguna validacion.
    for col, rango in ((IDX['Vendedor'], f"='{CONFIG}'!$A${FILA_VENDEDORES}:$A${FILA_VENDEDORES + MAX_VENDEDORES}"),
                       (IDX['Estado'], f"='{CONFIG}'!$C${FILA_ESTADOS}:$C${FILA_ESTADOS + len(ESTADOS) - 1}")):
        reqs.append({"setDataValidation": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": ultima,
                      "startColumnIndex": col, "endColumnIndex": col + 1},
            "rule": {"condition": {"type": "ONE_OF_RANGE",
                                   "values": [{"userEnteredValue": rango}]},
                     "showCustomUi": True, "strict": False}}})

    # Un color por estado, sobre toda la fila del negocio
    for i, (estado, fondo, texto) in enumerate(ESTADOS):
        reqs.append({"addConditionalFormatRule": {"index": i, "rule": {
            "ranges": [{"sheetId": sid, "startRowIndex": 1, "endRowIndex": ultima,
                        "startColumnIndex": IDX['Estado'], "endColumnIndex": IDX['Estado'] + 1}],
            "booleanRule": {
                "condition": {"type": "TEXT_EQ",
                              "values": [{"userEnteredValue": estado}]},
                "format": {"backgroundColor": _rgb(fondo),
                           "textFormat": {"foregroundColor": _rgb(texto), "bold": True}}}}}})
    return reqs


def _hoja_config(libro):
    try:
        h = libro.worksheet(CONFIG)
        libro.del_worksheet(h)
    except Exception:
        pass
    h = libro.add_worksheet(title=CONFIG, rows=60, cols=6, index=0)
    filas = [["VENDEDORES", "", "ESTADOS", ""],
             ["Agregá o borrá acá: se actualiza en las 43 hojas", "",
              "No cambies el texto: los colores dependen de él", ""]]
    for i in range(max(len(VENDEDORES), len(ESTADOS))):
        filas.append([VENDEDORES[i] if i < len(VENDEDORES) else "", "",
                      NOMBRES_ESTADO[i] if i < len(ESTADOS) else "", ""])
    h.update(values=filas, range_name="A1")

    azul = hex_a_rgb('#1D4E89')
    reqs = [
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb(azul),
                "textFormat": {"bold": True, "fontSize": 11, "fontFamily": FUENTE,
                               "foregroundColor": _rgb(hex_a_rgb(BLANCO))},
                "verticalAlignment": "MIDDLE", "padding": {"left": 10}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,padding)"}},
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 1, "endRowIndex": 2},
            "cell": {"userEnteredFormat": {"textFormat": {
                "italic": True, "fontSize": 9, "fontFamily": FUENTE,
                "foregroundColor": _rgb(hex_a_rgb(TINTA_SUAVE))}}},
            "fields": "userEnteredFormat.textFormat"}},
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 2, "endRowIndex": 50},
            "cell": {"userEnteredFormat": {
                "textFormat": {"fontFamily": FUENTE, "fontSize": TAM_DATOS},
                "verticalAlignment": "MIDDLE", "padding": {"left": 10}}},
            "fields": "userEnteredFormat(textFormat,verticalAlignment,padding)"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": h.id, "tabColor": _rgb(azul),
                           "gridProperties": {"frozenRowCount": 2}},
            "fields": "tabColor,gridProperties.frozenRowCount"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 300}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2},
            "properties": {"pixelSize": 30}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3},
            "properties": {"pixelSize": 300}, "fields": "pixelSize"}},
    ]
    # Cada estado con su color, para que se vea la escala de una
    for i, (estado, fondo, texto) in enumerate(ESTADOS):
        reqs.append({"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": FILA_ESTADOS - 1 + i,
                      "endRowIndex": FILA_ESTADOS + i,
                      "startColumnIndex": 2, "endColumnIndex": 3},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb(fondo),
                "textFormat": {"bold": True, "fontFamily": FUENTE,
                               "foregroundColor": _rgb(texto)}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"}})
    reintentar(libro.batch_update, {"requests": reqs})
    return h


def _hoja_ranking(libro, nichos):
    try:
        libro.del_worksheet(libro.worksheet(RANKING))
    except Exception:
        pass
    h = libro.add_worksheet(title=RANKING, rows=60, cols=10, index=1)

    # Una suma explicita por hoja. Da largo, pero INDIRECT sobre un array no
    # sirve: Sheets lo evalua solo con el primer elemento y el total daba el
    # de la primera hoja (818 = Academias) en vez de los 7391.
    def contar(col, condicion):
        return "=" + "+".join(
            f"COUNTIF('{n}'!{col}2:{col},{condicion})" for n in nichos)

    def contar_por(fila, estado=None):
        """Cuenta lo del vendedor de esa fila, opcionalmente filtrando estado."""
        partes = []
        for n in nichos:
            if estado:
                partes.append(f"COUNTIFS('{n}'!{COL_VENDEDOR}2:{COL_VENDEDOR},$A{fila},"
                              f"'{n}'!{COL_ESTADO}2:{COL_ESTADO},\"{estado}\")")
            else:
                partes.append(f"COUNTIF('{n}'!{COL_VENDEDOR}2:{COL_VENDEDOR},$A{fila})")
        return f'=IF($A{fila}="","",' + "+".join(partes) + ")"

    filas = [["RANKING DE VENDEDORES", "", "", "", "", ""],
             ["Se actualiza solo. Contá desde acá quién movió el embudo.", "", "", "", "", ""],
             ["Vendedor", "Clientes activos", "Demos", "Interesados", "Asignados", "Sin contactar"]]
    for i in range(FILAS_RANKING):
        f = 4 + i
        filas.append([
            f"=IFERROR('{CONFIG}'!A{FILA_VENDEDORES + i},\"\")",
            contar_por(f, "Cliente activo"),
            contar_por(f, "Demo iniciada"),
            contar_por(f, "Le interesó"),
            contar_por(f),
            contar_por(f, "Sin contactar"),
        ])

    f0 = 3 + FILAS_RANKING + 1
    filas.append([""] * 6)
    filas.append(["ESTADO GENERAL DE LA BASE", "", "", "", "", ""])
    filas.append(["Total negocios", contar(COL_TELEFONO, '"<>"'), "", "", "", ""])
    for estado in NOMBRES_ESTADO:
        filas.append([estado, contar(COL_ESTADO, f'"{estado}"'), "", "", "", ""])

    h.update(values=filas, range_name="A1", value_input_option="USER_ENTERED")

    azul, azul_claro = hex_a_rgb('#1D4E89'), hex_a_rgb('#EAF0F8')
    fin_vend = 3 + len(VENDEDORES)
    reintentar(libro.batch_update, {"requests": [
        # Todo con la misma tipografia de base
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 0, "endRowIndex": 60},
            "cell": {"userEnteredFormat": {
                "textFormat": {"fontFamily": FUENTE, "fontSize": TAM_DATOS,
                               "foregroundColor": _rgb(hex_a_rgb(TINTA))},
                "verticalAlignment": "MIDDLE", "padding": {"left": 10}}},
            "fields": "userEnteredFormat(textFormat,verticalAlignment,padding)"}},
        # Titulo
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {
                "bold": True, "fontSize": 16, "fontFamily": FUENTE,
                "foregroundColor": _rgb(azul)}}},
            "fields": "userEnteredFormat.textFormat"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 42}, "fields": "pixelSize"}},
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 1, "endRowIndex": 2},
            "cell": {"userEnteredFormat": {"textFormat": {
                "italic": True, "fontSize": 9, "fontFamily": FUENTE,
                "foregroundColor": _rgb(hex_a_rgb(TINTA_SUAVE))}}},
            "fields": "userEnteredFormat.textFormat"}},
        # Encabezado de la tabla de vendedores
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 2, "endRowIndex": 3},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb(azul),
                "textFormat": {"bold": True, "fontFamily": FUENTE, "fontSize": TAM_ENCABEZADO,
                               "foregroundColor": _rgb(hex_a_rgb(BLANCO))},
                "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,"
                      "verticalAlignment)"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "ROWS", "startIndex": 2, "endIndex": 3},
            "properties": {"pixelSize": 32}, "fields": "pixelSize"}},
        # Numeros centrados y en seminegrita: son la metrica, que se lean rapido
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 3, "endRowIndex": fin_vend,
                      "startColumnIndex": 1, "endColumnIndex": 6},
            "cell": {"userEnteredFormat": {
                "horizontalAlignment": "CENTER",
                "textFormat": {"bold": True, "fontSize": 11, "fontFamily": FUENTE}}},
            "fields": "userEnteredFormat(horizontalAlignment,textFormat)"}},
        {"addBanding": {"bandedRange": {
            "range": {"sheetId": h.id, "startRowIndex": 2, "endRowIndex": fin_vend,
                      "startColumnIndex": 0, "endColumnIndex": 6},
            "rowProperties": {"headerColor": _rgb(azul),
                              "firstBandColor": _rgb(hex_a_rgb(BLANCO)),
                              "secondBandColor": _rgb(azul_claro)}}}},
        # Bloque de estado general
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": f0, "endRowIndex": f0 + 1},
            "cell": {"userEnteredFormat": {"textFormat": {
                "bold": True, "fontSize": 12, "fontFamily": FUENTE,
                "foregroundColor": _rgb(azul)}}},
            "fields": "userEnteredFormat.textFormat"}},
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": f0 + 1, "endRowIndex": f0 + 9,
                      "startColumnIndex": 1, "endColumnIndex": 2},
            "cell": {"userEnteredFormat": {
                "horizontalAlignment": "CENTER",
                "textFormat": {"bold": True, "fontFamily": FUENTE}}},
            "fields": "userEnteredFormat(horizontalAlignment,textFormat)"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": h.id, "tabColor": _rgb(azul),
                           "gridProperties": {"frozenRowCount": 3}},
            "fields": "tabColor,gridProperties.frozenRowCount"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 230}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 6},
            "properties": {"pixelSize": 135}, "fields": "pixelSize"}},
    ]})
    # Los estados del resumen, cada uno con su color
    reintentar(libro.batch_update, {"requests": [
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": f0 + 2 + i, "endRowIndex": f0 + 3 + i,
                      "startColumnIndex": 0, "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb(fondo),
                "textFormat": {"bold": True, "fontFamily": FUENTE,
                               "foregroundColor": _rgb(texto)}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"}}
        for i, (estado, fondo, texto) in enumerate(ESTADOS)]})


def armar(generado=GENERADO):
    if not os.path.exists(generado):
        sys.exit(f"No existe {generado}. Corre el scraper primero.")
    df = pd.read_excel(generado, sheet_name="Todos", dtype=str).fillna("")
    df = df[df[CLAVE].str.strip() != ""]

    libro = _abrir_libro()
    nichos = sorted(df['Nicho'].unique())
    print(f"{len(df)} negocios en {len(nichos)} nichos")

    _hoja_config(libro)
    print(f"   {CONFIG} lista ({len(VENDEDORES)} vendedores)")

    previas = {h.title: h for h in libro.worksheets()}
    datos = {nicho[:99]: [COLUMNAS] + [fila_desde(r) for r in
                                       df[df['Nicho'] == nicho].to_dict('records')]
             for nicho in nichos}

    # Crear todas las hojas que falten en UN request: una por una se come el
    # limite de 60 escrituras por minuto antes de llegar a la mitad.
    faltantes = [(t, f) for t, f in datos.items() if t not in previas]
    if faltantes:
        reintentar(libro.batch_update, {"requests": [
            {"addSheet": {"properties": {
                "title": t,
                "gridProperties": {"rowCount": max(len(f) + 200, 2000),
                                   "columnCount": len(COLUMNAS)}}}}
            for t, f in faltantes]})
        print(f"   {len(faltantes)} hojas creadas")

    hojas = {h.title: h for h in libro.worksheets()}
    for titulo in datos:
        if titulo in previas:
            reintentar(hojas[titulo].clear)

    # Todos los valores de una sola vez.
    reintentar(libro.values_batch_update, {
        "valueInputOption": "RAW",
        "data": [{"range": f"'{t}'!A1", "values": f} for t, f in datos.items()]})
    print(f"   {sum(len(f) - 1 for f in datos.values())} filas escritas en {len(datos)} hojas")

    # Banding y reglas condicionales se acumulan al rehacer: hay que sacar las
    # viejas o Google rechaza las nuevas por superponerse.
    meta = reintentar(libro.fetch_sheet_metadata)
    limpieza = []
    for h in meta.get('sheets', []):
        for b in h.get('bandedRanges', []):
            limpieza.append({"deleteBanding": {"bandedRangeId": b['bandedRangeId']}})
        sid = h['properties']['sheetId']
        for i in range(len(h.get('conditionalFormats', [])) - 1, -1, -1):
            limpieza.append({"deleteConditionalFormatRule": {"sheetId": sid, "index": i}})
    for i in range(0, len(limpieza), 150):
        reintentar(libro.batch_update, {"requests": limpieza[i:i + 150]})

    rubro_de = dict(zip(df['Nicho'], df['Rubro']))
    reqs_formato = []
    for titulo, filas in datos.items():
        reqs_formato += _formato_hoja(hojas[titulo].id, len(filas), rubro_de.get(titulo, ''))
    for i in range(0, len(reqs_formato), 150):
        reintentar(libro.batch_update, {"requests": reqs_formato[i:i + 150]})
    print("   formato, desplegables y colores aplicados")

    _hoja_ranking(libro, nichos)
    print(f"   {RANKING} lista")

    for titulo, hoja in previas.items():
        if titulo not in nichos and titulo not in (CONFIG, RANKING):
            libro.del_worksheet(hoja)
            print(f"   borrada hoja vieja: {titulo}")

    print(f"\nListo: {libro.url}")


def demo():
    from modules.planilla import SIN_CONTACTAR
    n = {'Negocio': 'Barber X', 'Teléfono': '03411234567', 'Categoría': 'Barbería',
         'Ciudad': 'Rosario', 'Link en Maps': 'http://x', 'Sitio web': 'no va'}
    fila = fila_desde(n)
    assert len(fila) == len(COLUMNAS), f'{len(fila)} valores para {len(COLUMNAS)} columnas'
    assert fila[IDX['Negocio']] == 'Barber X'
    assert fila[IDX['Teléfono']] == '03411234567'
    assert fila[IDX['Categoría']] == 'Barbería'
    assert fila[IDX['Estado']] == SIN_CONTACTAR, 'el estado inicial no quedo seteado'
    assert fila[IDX['Vendedor']] == '', 'el vendedor lo asigna una persona'
    assert fila[IDX['Observaciones']] == ''
    assert 'Sitio web' not in COLUMNAS, 'la web no va en la planilla'
    print(f'OK  planilla: {len(COLUMNAS)} columnas, fila armada correctamente')


if __name__ == "__main__":
    if '--test' in sys.argv:
        demo()
    else:
        armar()
