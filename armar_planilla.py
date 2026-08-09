"""Arma la planilla de ventas en Google Sheets: una hoja por nicho, con
desplegables de vendedor y estado, colores por estado y un ranking en vivo.

    python armar_planilla.py            # arma o rehace la estructura
    python armar_planilla.py --test     # solo la logica, sin tocar Google

CUIDADO: rehace las hojas de nicho desde cero. Se corre una vez al principio;
para el dia a dia esta sincronizar_sheets.py, que solo agrega filas.
"""
import os
import sys

import pandas as pd

from modules.planilla import (COLUMNAS, ANCHOS, CLAVE, VENDEDORES, ESTADOS,
                              NOMBRES_ESTADO, MOTIVOS, CONFIG, PANEL, RESUMEN,
                              FILA_VENDEDORES, FILA_ESTADOS, FILA_MOTIVOS,
                              COL_VENDEDOR, COL_ESTADO, COL_MOTIVO, COL_FECHA,
                              COL_TELEFONO, IDX, col_letra, fila_desde,
                              reintentar)
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
                       (IDX['Estado'], f"='{CONFIG}'!$C${FILA_ESTADOS}:$C${FILA_ESTADOS + len(ESTADOS) - 1}"),
                       (IDX['Motivo'], f"='{CONFIG}'!$E${FILA_MOTIVOS}:$E${FILA_MOTIVOS + len(MOTIVOS) - 1}")):
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
        reintentar(libro.del_worksheet, h)
    except Exception:
        pass
    h = reintentar(libro.add_worksheet, title=CONFIG, rows=60, cols=6, index=0)
    filas = [["VENDEDORES", "", "ESTADOS", "", "MOTIVOS DE NO AVANCE"],
             ["Agregá o borrá acá: se actualiza en todas las hojas", "",
              "No cambies el texto: los colores dependen de él", "",
              "Por qué no avanzó el lead"]]
    for i in range(max(len(VENDEDORES), len(ESTADOS), len(MOTIVOS))):
        filas.append([VENDEDORES[i] if i < len(VENDEDORES) else "", "",
                      NOMBRES_ESTADO[i] if i < len(ESTADOS) else "", "",
                      MOTIVOS[i] if i < len(MOTIVOS) else ""])
    reintentar(h.update, values=filas, range_name="A1")

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
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 3, "endIndex": 4},
            "properties": {"pixelSize": 30}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 4, "endIndex": 5},
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

def _formulas_panel(nichos):
    """Devuelve helpers que suman una hoja por vez.

    No se puede usar INDIRECT sobre un array: Sheets lo evalua solo con el
    primer elemento y el total daba el de la primera hoja (818 en vez de 7391).
    """
    def por_estado(estado):
        return "=" + "+".join(
            f"COUNTIF('{n}'!{COL_ESTADO}2:{COL_ESTADO},\"{estado}\")" for n in nichos)

    def total_leads():
        # "?*" = al menos un caracter. COUNTA contaria tambien las celdas con
        # cadena vacia que quedan al crear la fila, y daba 7391 asignados.
        return "=" + "+".join(
            f"COUNTIF('{n}'!{COL_TELEFONO}2:{COL_TELEFONO},\"?*\")" for n in nichos)

    def asignados():
        return "=" + "+".join(
            f"COUNTIF('{n}'!{COL_VENDEDOR}2:{COL_VENDEDOR},\"?*\")" for n in nichos)

    def por_motivo(motivo):
        return "=" + "+".join(
            f"COUNTIF('{n}'!{COL_MOTIVO}2:{COL_MOTIVO},\"{motivo}\")" for n in nichos)

    def del_vendedor(fila, estado=None, hoy=False):
        partes = []
        for n in nichos:
            v = f"'{n}'!{COL_VENDEDOR}2:{COL_VENDEDOR},$A{fila}"
            if estado:
                partes.append(f"COUNTIFS({v},'{n}'!{COL_ESTADO}2:{COL_ESTADO},\"{estado}\")")
            elif hoy:
                partes.append(f"COUNTIFS({v},'{n}'!{COL_FECHA}2:{COL_FECHA},TODAY())")
            else:
                partes.append(f"COUNTIF({v})")
        return f'=IF($A{fila}="","",' + "+".join(partes) + ")"

    return por_estado, total_leads, asignados, por_motivo, del_vendedor


def _hoja_panel(libro, nichos):
    """Panel de ventas: embudo, ranking de vendedores y motivos de no avance."""
    try:
        reintentar(libro.del_worksheet, libro.worksheet(PANEL))
    except Exception:
        pass
    h = reintentar(libro.add_worksheet, title=PANEL, rows=80, cols=10, index=1)
    por_estado, total_leads, asignados, por_motivo, del_vendedor = _formulas_panel(nichos)

    f = []                                   # filas, 0-based mientras se arma
    f.append(["PANEL DE VENTAS"] + [""] * 7)
    f.append(['=CONCATENATE("Actualizado: ",TEXT(TODAY(),"dd/mm/yyyy"))'] + [""] * 7)
    f.append([""] * 8)

    F_EMBUDO = len(f)
    f.append(["TOTALES DEL EMBUDO"] + [""] * 7)
    f.append(["Leads totales", "Asignados"] + NOMBRES_ESTADO)
    f.append([total_leads(), asignados()] + [por_estado(e) for e in NOMBRES_ESTADO])
    f.append([""] * 8)

    F_RANKING = len(f)
    f.append(["RANKING DE VENDEDORES"] + [""] * 7)
    f.append(["Vendedor", "Leads asignados", "Demos iniciadas", "Clientes activos",
              "% conversión", "Trabajados hoy", "", ""])
    for i in range(FILAS_RANKING):
        fila = len(f) + 1                    # 1-based, como la ve Sheets
        f.append([
            f"=IFERROR('{CONFIG}'!A{FILA_VENDEDORES + i},\"\")",
            del_vendedor(fila),
            del_vendedor(fila, "Demo iniciada"),
            del_vendedor(fila, "Cliente activo"),
            f'=IF(N(${col_letra(1)}{fila})=0,"",${col_letra(3)}{fila}/${col_letra(1)}{fila})',
            del_vendedor(fila, hoy=True), "", ""])
    f.append([""] * 8)

    F_MOTIVOS = len(f)
    f.append(["MOTIVOS DE NO AVANCE"] + [""] * 7)
    f.append(["Motivo", "Cantidad", "", "", "", "", "", ""])
    for m in MOTIVOS:
        f.append([m, por_motivo(m), "", "", "", "", "", ""])
    f.append([""] * 8)

    F_AYUDA = len(f)
    f.append(["CÓMO USARLO"] + [""] * 7)
    for t in ["Los desplegables (Vendedor / Estado / Motivo) salen de la hoja Config.",
              f"Agregá o quitá un vendedor en Config columna A y se actualiza en las {len(nichos)} hojas.",
              "El teléfono está en formato texto: conserva los ceros iniciales.",
              'Cargá la fecha en "Última gestión" para que cuente en Trabajados hoy.',
              "El resto se calcula solo. Sin macros: anda igual en Sheets y en Excel.",
              "El detalle por nicho está en la hoja Resumen por nicho."]:
        f.append(["• " + t] + [""] * 7)

    reintentar(h.update, values=f, range_name="A1", value_input_option="USER_ENTERED")

    azul, azul_claro = hex_a_rgb('#1D4E89'), hex_a_rgb('#EAF0F8')
    gris_claro = hex_a_rgb('#F1F3F4')

    def titulo(fila, tam=12):
        return {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": fila, "endRowIndex": fila + 1},
            "cell": {"userEnteredFormat": {"textFormat": {
                "bold": True, "fontSize": tam, "fontFamily": FUENTE,
                "foregroundColor": _rgb(azul)}}},
            "fields": "userEnteredFormat.textFormat"}}

    def encabezado(fila, hasta):
        return [{"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": fila, "endRowIndex": fila + 1,
                      "startColumnIndex": 0, "endColumnIndex": hasta},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb(azul),
                "textFormat": {"bold": True, "fontSize": TAM_ENCABEZADO, "fontFamily": FUENTE,
                               "foregroundColor": _rgb(hex_a_rgb(BLANCO))},
                "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                "wrapStrategy": "WRAP"}},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,"
                      "verticalAlignment,wrapStrategy)"}},
            {"updateDimensionProperties": {
                "range": {"sheetId": h.id, "dimension": "ROWS",
                          "startIndex": fila, "endIndex": fila + 1},
                "properties": {"pixelSize": 34}, "fields": "pixelSize"}}]

    def numeros(desde, hasta, col_ini, col_fin, tam=11):
        return {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": desde, "endRowIndex": hasta,
                      "startColumnIndex": col_ini, "endColumnIndex": col_fin},
            "cell": {"userEnteredFormat": {
                "horizontalAlignment": "CENTER",
                "textFormat": {"bold": True, "fontSize": tam, "fontFamily": FUENTE}}},
            "fields": "userEnteredFormat(horizontalAlignment,textFormat)"}}

    fin_rank = F_RANKING + 2 + FILAS_RANKING
    reqs = [
        # Base tipografica de toda la hoja
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 0, "endRowIndex": 80},
            "cell": {"userEnteredFormat": {
                "textFormat": {"fontFamily": FUENTE, "fontSize": TAM_DATOS,
                               "foregroundColor": _rgb(hex_a_rgb(TINTA))},
                "verticalAlignment": "MIDDLE", "padding": {"left": 10, "right": 10}}},
            "fields": "userEnteredFormat(textFormat,verticalAlignment,padding)"}},
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {
                "bold": True, "fontSize": 18, "fontFamily": FUENTE,
                "foregroundColor": _rgb(azul)}}},
            "fields": "userEnteredFormat.textFormat"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 46}, "fields": "pixelSize"}},
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 1, "endRowIndex": 2},
            "cell": {"userEnteredFormat": {"textFormat": {
                "italic": True, "fontSize": 9, "fontFamily": FUENTE,
                "foregroundColor": _rgb(hex_a_rgb(TINTA_SUAVE))}}},
            "fields": "userEnteredFormat.textFormat"}},
        titulo(F_EMBUDO), titulo(F_RANKING), titulo(F_MOTIVOS), titulo(F_AYUDA),
        # Numeros grandes del embudo
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": F_EMBUDO + 2,
                      "endRowIndex": F_EMBUDO + 3, "startColumnIndex": 0, "endColumnIndex": 8},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb(gris_claro),
                "horizontalAlignment": "CENTER",
                "textFormat": {"bold": True, "fontSize": 15, "fontFamily": FUENTE}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "ROWS",
                      "startIndex": F_EMBUDO + 2, "endIndex": F_EMBUDO + 3},
            "properties": {"pixelSize": 40}, "fields": "pixelSize"}},
        numeros(F_RANKING + 2, fin_rank, 1, 6),
        numeros(F_MOTIVOS + 2, F_MOTIVOS + 2 + len(MOTIVOS), 1, 2),
        # % conversion como porcentaje de verdad
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": F_RANKING + 2, "endRowIndex": fin_rank,
                      "startColumnIndex": 4, "endColumnIndex": 5},
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "PERCENT", "pattern": "0.0%"}}},
            "fields": "userEnteredFormat.numberFormat"}},
        # Nombres de vendedor a la izquierda y en negrita
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": F_RANKING + 2, "endRowIndex": fin_rank,
                      "startColumnIndex": 0, "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontFamily": FUENTE}}},
            "fields": "userEnteredFormat.textFormat"}},
        # Barra proporcional en clientes activos: el ranking se lee de un vistazo
        {"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [{"sheetId": h.id, "startRowIndex": F_RANKING + 2, "endRowIndex": fin_rank,
                        "startColumnIndex": 3, "endColumnIndex": 4}],
            "gradientRule": {
                "minpoint": {"color": _rgb(hex_a_rgb('#FFFFFF')), "type": "NUMBER", "value": "0"},
                "maxpoint": {"color": _rgb(hex_a_rgb('#34A853')), "type": "MAX"}}}}},
        # Ayuda en gris chico: esta para consultarla, no para leerla siempre
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": F_AYUDA + 1, "endRowIndex": F_AYUDA + 8},
            "cell": {"userEnteredFormat": {"textFormat": {
                "fontSize": 9, "fontFamily": FUENTE,
                "foregroundColor": _rgb(hex_a_rgb(TINTA_SUAVE))}}},
            "fields": "userEnteredFormat.textFormat"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": h.id, "tabColor": _rgb(azul),
                           "gridProperties": {"hideGridlines": True}},
            "fields": "tabColor,gridProperties.hideGridlines"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 210}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 8},
            "properties": {"pixelSize": 145}, "fields": "pixelSize"}},
    ]
    reqs += encabezado(F_EMBUDO + 1, 8)
    reqs += encabezado(F_RANKING + 1, 6)
    reqs += encabezado(F_MOTIVOS + 1, 2)
    # Cada motivo con el color del estado que lo genera
    for i, m in enumerate(MOTIVOS):
        reqs.append({"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": F_MOTIVOS + 2 + i,
                      "endRowIndex": F_MOTIVOS + 3 + i, "startColumnIndex": 0, "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {
                "fontFamily": FUENTE,
                "foregroundColor": _rgb(hex_a_rgb('#8A5A1B'))}}},
            "fields": "userEnteredFormat.textFormat"}})
    for i in range(0, len(reqs), 150):
        reintentar(libro.batch_update, {"requests": reqs[i:i + 150]})


def _hoja_resumen(libro, nichos):
    """Una fila por nicho: cuantos hay, cuantos se trabajaron y cuantos cerraron."""
    try:
        reintentar(libro.del_worksheet, libro.worksheet(RESUMEN))
    except Exception:
        pass
    h = reintentar(libro.add_worksheet, title=RESUMEN, rows=len(nichos) + 12, cols=7, index=2)

    filas = [["RESUMEN POR NICHO"] + [""] * 6,
             ["Para ver qué rubro rinde y cuál conviene dejar de llamar."] + [""] * 6,
             ["Nicho", "Leads", "Sin contactar", "Le interesó", "Demos",
              "Clientes activos", "% conversión"]]
    for n in nichos:
        f = len(filas) + 1
        filas.append([
            n,
            f"=COUNTIF('{n}'!{COL_TELEFONO}2:{COL_TELEFONO},\"?*\")",
            f"=COUNTIF('{n}'!{COL_ESTADO}2:{COL_ESTADO},\"Sin contactar\")",
            f"=COUNTIF('{n}'!{COL_ESTADO}2:{COL_ESTADO},\"Le interesó\")",
            f"=COUNTIF('{n}'!{COL_ESTADO}2:{COL_ESTADO},\"Demo iniciada\")",
            f"=COUNTIF('{n}'!{COL_ESTADO}2:{COL_ESTADO},\"Cliente activo\")",
            f'=IF($B{f}=0,"",$F{f}/$B{f})'])
    total = len(filas) + 1
    filas.append(["TOTAL"] + [f"=SUM({c}4:{c}{total - 1})" for c in "BCDEF"] +
                 [f'=IF($B{total}=0,"",$F{total}/$B{total})'])

    reintentar(h.update, values=filas, range_name="A1", value_input_option="USER_ENTERED")

    azul, azul_claro = hex_a_rgb('#1D4E89'), hex_a_rgb('#EAF0F8')
    reintentar(libro.batch_update, {"requests": [
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 0, "endRowIndex": len(filas) + 2},
            "cell": {"userEnteredFormat": {
                "textFormat": {"fontFamily": FUENTE, "fontSize": TAM_DATOS,
                               "foregroundColor": _rgb(hex_a_rgb(TINTA))},
                "verticalAlignment": "MIDDLE", "padding": {"left": 10, "right": 10}}},
            "fields": "userEnteredFormat(textFormat,verticalAlignment,padding)"}},
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
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 2, "endRowIndex": 3},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb(azul),
                "textFormat": {"bold": True, "fontSize": TAM_ENCABEZADO, "fontFamily": FUENTE,
                               "foregroundColor": _rgb(hex_a_rgb(BLANCO))},
                "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                "wrapStrategy": "WRAP"}},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,"
                      "verticalAlignment,wrapStrategy)"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "ROWS", "startIndex": 2, "endIndex": 3},
            "properties": {"pixelSize": 34}, "fields": "pixelSize"}},
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 3, "endRowIndex": len(filas),
                      "startColumnIndex": 1, "endColumnIndex": 7},
            "cell": {"userEnteredFormat": {
                "horizontalAlignment": "CENTER",
                "textFormat": {"fontFamily": FUENTE, "fontSize": TAM_DATOS}}},
            "fields": "userEnteredFormat(horizontalAlignment,textFormat)"}},
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 3, "endRowIndex": len(filas),
                      "startColumnIndex": 6, "endColumnIndex": 7},
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "PERCENT", "pattern": "0.0%"}}},
            "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": len(filas) - 1, "endRowIndex": len(filas)},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb(azul_claro),
                "textFormat": {"bold": True, "fontFamily": FUENTE}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"}},
        {"addBanding": {"bandedRange": {
            "range": {"sheetId": h.id, "startRowIndex": 2, "endRowIndex": len(filas) - 1,
                      "startColumnIndex": 0, "endColumnIndex": 7},
            "rowProperties": {"headerColor": _rgb(azul),
                              "firstBandColor": _rgb(hex_a_rgb(BLANCO)),
                              "secondBandColor": _rgb(azul_claro)}}}},
        {"updateSheetProperties": {
            "properties": {"sheetId": h.id, "tabColor": _rgb(azul),
                           "gridProperties": {"frozenRowCount": 3}},
            "fields": "tabColor,gridProperties.frozenRowCount"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 220}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 7},
            "properties": {"pixelSize": 125}, "fields": "pixelSize"}},
        {"setBasicFilter": {"filter": {"range": {
            "sheetId": h.id, "startRowIndex": 2, "endRowIndex": len(filas) - 1,
            "endColumnIndex": 7}}}},
    ]})


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

    previas = {h.title: h for h in reintentar(libro.worksheets)}
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

    hojas = {h.title: h for h in reintentar(libro.worksheets)}
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

    _hoja_panel(libro, nichos)
    print(f"   {PANEL} listo")
    _hoja_resumen(libro, nichos)
    print(f"   {RESUMEN} listo")

    for titulo, hoja in previas.items():
        if titulo not in nichos and titulo not in (CONFIG, PANEL, RESUMEN):
            reintentar(libro.del_worksheet, hoja)
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

    # reintentar: un 429 tiene que esperar y volver, cualquier otro error tiene
    # que salir de una. Si esto se rompe, el sync se saltea nichos en silencio.
    import modules.planilla as P
    dormido = []
    real_sleep = P.time.sleep
    P.time.sleep = dormido.append          # no esperar 30s de verdad en el test
    try:
        _probar_reintentar(dormido)
    finally:
        P.time.sleep = real_sleep
    assert dormido == [30, 60], f'espera mal escalonada: {dormido}'
    print('OK  reintentar: espera el 429 y deja pasar el resto')


def _probar_reintentar(dormido):
    llamadas = []
    def falla_dos_veces():
        llamadas.append(1)
        if len(llamadas) <= 2:
            raise Exception("APIError: [429]: Quota exceeded")
        return "listo"
    assert reintentar(falla_dos_veces) == "listo", 'no reintento tras el 429'
    assert len(llamadas) == 3, f'reintento {len(llamadas)} veces'

    def falla_404():
        raise Exception("APIError: [404]: not found")
    try:
        reintentar(falla_404)
        assert False, 'se comio un error que no era de cuota'
    except Exception as e:
        assert '404' in str(e), 'tapo un error que no era de cuota'


if __name__ == "__main__":
    if '--test' in sys.argv:
        demo()
    else:
        armar()
