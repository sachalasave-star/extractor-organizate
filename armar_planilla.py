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
                              COL_TELEFONO, IDX, fila_desde,
                              reintentar)
from modules.estilo import (FUENTE, FUENTE_DATOS, TAM_ENCABEZADO, TAM_DATOS,
                            color_rubro, hex_a_rgb, TINTA, TINTA_SUAVE, LINEA, BLANCO)
from sincronizar_sheets import _credenciales, _sheet_id, GENERADO

MAX_VENDEDORES = 40      # margen del desplegable, no cuesta nada
# Filas del ranking. Cada una son 6 formulas x 43 hojas: con 40 la planilla se
# arrastra, con 12 hay margen de sobra sobre los 4 vendedores actuales.
FILAS_RANKING = 12


def _abrir_libro():
    import gspread
    from google.oauth2.service_account import Credentials
    sheet_id = _sheet_id()
    if not sheet_id:
        sys.exit("Falta SHEET_ID")
    cred = Credentials.from_service_account_info(
        _credenciales(), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    cliente = gspread.authorize(cred)
    return reintentar(cliente.open_by_key, sheet_id)


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
    # cols=16: A-H son la hoja visible, J-O son un staging area escondida (ver
    # mas abajo por que hace falta).
    h = reintentar(libro.add_worksheet, title=PANEL, rows=80, cols=16, index=1)
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

    # RANKING: no alcanza con listar a los vendedores en el orden de Config,
    # eso es una tabla, no un ranking. Las metricas de cada uno se calculan en
    # J:O (escondida) y A:F se arma con UN solo SORT(FILTER(...)) que ordena
    # por Leads asignados de mayor a menor y descarta los slots sin vendedor.
    # Asi arriba queda el mas activo y abajo el menos activo, sin filas
    # vacias de relleno.
    F_RANKING = len(f)
    f.append(["RANKING DE VENDEDORES"] + [""] * 7)
    f.append(["Vendedor", "Leads asignados", "Demos iniciadas", "Clientes activos",
              "% conversión", "Trabajados hoy", "", ""])
    fila_r1 = len(f) + 1                     # primera fila de datos, 1-based
    fila_r2 = fila_r1 + FILAS_RANKING - 1
    f.append([f'=IFERROR(SORT(FILTER(J{fila_r1}:O{fila_r2},J{fila_r1}:J{fila_r2}<>""),2,FALSE),'
              f'"Cargá vendedores en Config")'] + [""] * 7)
    for _ in range(FILAS_RANKING - 1):       # el resto lo llena el spill del SORT
        f.append([""] * 8)
    f.append([""] * 8)

    F_MOTIVOS = len(f)
    f.append(["MOTIVOS DE NO AVANCE"] + [""] * 7)
    f.append(["El que más frena, no la lista completa."] + [""] * 7)
    fila_m1 = len(f) + 1
    fila_m2 = fila_m1 + len(MOTIVOS) - 1
    rango_motivos = f"J{fila_m1}:J{fila_m2}"
    nombres = "{" + ";".join(f'"{m}"' for m in MOTIVOS) + "}"
    f.append([f'=IFERROR(INDEX({nombres},MATCH(MAX({rango_motivos}),{rango_motivos},0)),'
              f'"Sin datos aún")'] + [""] * 7)
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

    # Staging escondido: las mismas cuentas por vendedor que antes iban directo
    # en A:F, ahora en J:O para que SORT(FILTER(...)) las ordene sin que nadie
    # vea la tabla cruda. Mismo criterio para el motivo dominante: MAX/MATCH
    # necesitan un rango de numeros, no se puede armar sobre formulas sueltas.
    staging_ranking = []
    for i in range(FILAS_RANKING):
        fila = fila_r1 + i
        staging_ranking.append([
            f"=IFERROR('{CONFIG}'!A{FILA_VENDEDORES + i},\"\")",
            del_vendedor(fila),
            del_vendedor(fila, "Demo iniciada"),
            del_vendedor(fila, "Cliente activo"),
            f'=IF(N(K{fila})=0,"",M{fila}/K{fila})',
            del_vendedor(fila, hoy=True)])
    reintentar(h.update, values=staging_ranking, range_name=f"J{fila_r1}",
              value_input_option="USER_ENTERED")

    staging_motivos = [[por_motivo(m)] for m in MOTIVOS]
    reintentar(h.update, values=staging_motivos, range_name=f"J{fila_m1}",
              value_input_option="USER_ENTERED")

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
        # Subtitulo de motivos, mismo estilo italic/chico que el de Resumen
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": F_MOTIVOS + 1, "endRowIndex": F_MOTIVOS + 2},
            "cell": {"userEnteredFormat": {"textFormat": {
                "italic": True, "fontSize": 9, "fontFamily": FUENTE,
                "foregroundColor": _rgb(hex_a_rgb(TINTA_SUAVE))}}},
            "fields": "userEnteredFormat.textFormat"}},
        # El motivo dominante, destacado como los numeros grandes del embudo
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": F_MOTIVOS + 2,
                      "endRowIndex": F_MOTIVOS + 3, "startColumnIndex": 0, "endColumnIndex": 8},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb(gris_claro),
                "textFormat": {"bold": True, "fontSize": 14, "fontFamily": FUENTE}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "ROWS",
                      "startIndex": F_MOTIVOS + 2, "endIndex": F_MOTIVOS + 3},
            "properties": {"pixelSize": 40}, "fields": "pixelSize"}},
        # J:O son el staging del ranking y del motivo dominante: no es para
        # que lo vea nadie, solo para que SORT/MATCH tengan un rango de donde leer.
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 9, "endIndex": 15},
            "properties": {"hiddenByUser": True}, "fields": "hiddenByUser"}},
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
    # Sin fila TOTAL a proposito: sumar las 43 hojas da los leads totales del
    # negocio contra los interesados totales en una sola linea, que es
    # exactamente el numero que un vendedor nuevo no necesita ver el primer
    # dia. El detalle por nicho de arriba sigue completo.

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
        # Sin resaltado de TOTAL: esa fila ya no existe.
        {"addBanding": {"bandedRange": {
            "range": {"sheetId": h.id, "startRowIndex": 2, "endRowIndex": len(filas),
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
            "sheetId": h.id, "startRowIndex": 2, "endRowIndex": len(filas),
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


def _crudas(meta):
    """Hojas de nicho sin formato, sacadas del metadata del libro.

    El sello es frozenRowCount: _formato_hoja congela la primera fila, y la
    hoja que crea sincronizar_sheets con add_worksheet viene en 0. Es lo unico
    que distingue una hoja cruda de una lista sin leer una sola celda.
    """
    salida = []
    for h in meta.get('sheets', []):
        p = h['properties']
        if p['title'] in (CONFIG, PANEL, RESUMEN):
            continue
        if not p.get('gridProperties', {}).get('frozenRowCount'):
            salida.append(p)
    return salida


def poner_al_dia(generado=GENERADO):
    """Formatea las hojas que creo el sync y rehace Panel y Resumen.

    armar() no se puede correr con los vendedores adentro: rehace Config (y se
    lleva puestos los vendedores cargados a mano), hace clear() de cada hoja y
    borra las hojas de los nichos que ya no estan en el Excel, con el trabajo
    hecho adentro. Esto toca solo lo que no tiene nada de nadie: el formato de
    las hojas nuevas, que salen crudas y sin desplegables, y las dos hojas de
    formulas, que hay que rehacer para que cuenten los nichos nuevos.
    """
    if not os.path.exists(generado):
        sys.exit(f"No existe {generado}. Corre el scraper primero.")
    df = pd.read_excel(generado, sheet_name="Todos", dtype=str).fillna("")
    df = df[df[CLAVE].str.strip() != ""]

    libro = _abrir_libro()
    meta = reintentar(libro.fetch_sheet_metadata)
    crudas = _crudas(meta)

    if crudas:
        rubro_de = dict(zip(df['Nicho'], df['Rubro']))
        filas_de = df['Nicho'].value_counts().to_dict()
        reqs = []
        for p in crudas:
            titulo = p['title']
            reqs += _formato_hoja(p['sheetId'], filas_de.get(titulo, 0) + 1,
                                  rubro_de.get(titulo, ''))
            print(f"   formateada: {titulo}")
        for i in range(0, len(reqs), 150):
            reintentar(libro.batch_update, {"requests": reqs[i:i + 150]})
    else:
        print("   no hay hojas nuevas sin formato")

    # Panel y Resumen se rehacen siempre: sus totales son un COUNTIF explicito
    # por hoja (INDIRECT sobre un array no recorre nada en Sheets), asi que un
    # nicho nuevo no se cuenta solo. Son formulas, no tienen datos de nadie.
    # La lista sale del libro y no del Excel: los nichos apagados siguen
    # teniendo su hoja con trabajo adentro y tienen que seguir sumando.
    nichos = sorted(p['title'] for h in meta.get('sheets', [])
                    for p in [h['properties']]
                    if p['title'] not in (CONFIG, PANEL, RESUMEN))
    _hoja_panel(libro, nichos)
    print(f"   {PANEL} rehecho sobre {len(nichos)} nichos")
    _hoja_resumen(libro, nichos)
    print(f"   {RESUMEN} rehecho")

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
    assert dormido == [30, 60, 30], f'espera mal escalonada: {dormido}'
    print('OK  reintentar: espera el 429 y deja pasar el resto')

    # _sheet_id: el id sucio da 404 y el 404 no dice por que. Las dos corridas
    # del 9/8 scrapearon 20203 negocios y no subieron una fila por esto.
    limpio = '1MICJFEZXlGLJPTdBSAtmkUEMzQdjAzxsxJglFupQp_c'
    previo = os.environ.get('SHEET_ID')
    try:
        for sucio in (limpio, limpio + '\n', '﻿' + limpio, ' ' + limpio + ' \r\n'):
            os.environ['SHEET_ID'] = sucio
            assert _sheet_id() == limpio, f'no limpio {sucio!r}'
        os.environ.pop('SHEET_ID')
        assert _sheet_id() == '', 'sin secret tiene que dar vacio, no reventar'
    finally:
        if previo is not None:
            os.environ['SHEET_ID'] = previo
    print('OK  _sheet_id: le saca el BOM y el \\n que mete PowerShell')

    _probar_hoja_panel()
    print('OK  _hoja_panel: el ranking ordena por la columna que hay que ordenar '
          'y el motivo dominante mira el rango correcto')

    # _crudas: si se equivoca para el lado de "esta lista", la hoja nueva queda
    # sin desplegables para siempre; si se equivoca para el otro, le manda
    # addBanding a una hoja que ya tiene y Google rechaza el batch entero.
    def _hoja(titulo, frozen):
        props = {"title": titulo, "sheetId": abs(hash(titulo)) % 1000}
        if frozen is not None:
            props["gridProperties"] = {"frozenRowCount": frozen}
        return {"properties": props}
    meta = {"sheets": [
        _hoja("Peluquerías", 1),        # formateada por armar()
        _hoja("Flebología", 0),         # recien creada por el sync
        _hoja("Ópticas", None),         # sin gridProperties en el metadata
        _hoja(CONFIG, 0),               # no es hoja de nicho
        _hoja(PANEL, 2),
        _hoja(RESUMEN, 3),
    ]}
    crudas = {p['title'] for p in _crudas(meta)}
    assert crudas == {"Flebología", "Ópticas"}, f'detecto mal: {crudas}'
    assert not _crudas({}), 'metadata vacio tiene que dar vacio'
    print('OK  _crudas: agarra las del sync y no toca Config, Panel ni Resumen')


class _HojaFalsa:
    """Imita lo justo de un Worksheet de gspread para que _hoja_panel corra sin
    tocar Google. Guarda cada .update() para poder revisar las formulas."""
    def __init__(self, sheet_id):
        self.id = sheet_id
        self.escrituras = []           # [(range_name, values)]

    def update(self, values, range_name, value_input_option=None):
        self.escrituras.append((range_name, values))

    def celda(self, range_name):
        """El valor guardado en esa celda exacta, entre todas las escrituras.
        Busca la escritura que la cubre aunque haya sido una fila con varias
        columnas (ej. el staging J:O escribe 6 columnas de una)."""
        col, fila = ord(range_name[0]) - ord('A'), int(range_name[1:])
        for rango, values in self.escrituras:
            c0, f0 = ord(rango[0]) - ord('A'), int(rango[1:])
            j = col - c0
            i = fila - f0
            if j < 0 or not (0 <= i < len(values)) or j >= len(values[i] or []):
                continue
            return values[i][j]
        return None


class _LibroFalso:
    """Imita lo justo de un Spreadsheet: crear una hoja y guardar los
    batch_update, sin llamar a Google."""
    def __init__(self):
        self.hojas = {}
        self.batches = []

    def worksheet(self, titulo):
        raise Exception("no existe (planilla nueva)")

    def del_worksheet(self, hoja):
        pass

    def add_worksheet(self, title, rows, cols, index):
        h = _HojaFalsa(sheet_id=len(self.hojas) + 1)
        self.hojas[title] = h
        return h

    def batch_update(self, body):
        self.batches.append(body)


def _probar_hoja_panel():
    """El ranking se arma con SORT/FILTER sobre una zona escondida (J:O) que
    tiene que apuntar exactamente a las mismas filas y columnas que llena el
    staging. Un desfasaje ahi no tira error en Sheets: silenciosamente ordena
    o calcula mal y nadie lo nota hasta que alguien cuenta a mano.

    Las filas se ubican por contenido, no por numero calculado a mano: si se
    agrega o saca una linea del layout, el test la sigue encontrando sola en
    vez de quedar apuntando a la fila vieja sin avisar."""
    libro = _LibroFalso()
    nichos = ['Barberías', 'Odontología']
    _hoja_panel(libro, nichos)
    h = libro.hojas[PANEL]
    valores = next(v for r, v in h.escrituras if r == "A1")   # el bloque `f` completo

    def fila_1based(texto):
        return next(i for i, fila in enumerate(valores) if fila and fila[0] == texto) + 1

    fila_r1 = fila_1based("Vendedor") + 1    # header del ranking + 1 = primera fila de datos
    visible = h.celda(f"A{fila_r1}")
    assert visible.startswith("=IFERROR(SORT(FILTER("), f'no arranca con SORT/FILTER: {visible}'
    assert f"J{fila_r1}:O{fila_r1 + FILAS_RANKING - 1}" in visible, \
        f'el SORT no lee el mismo rango que llena el staging: {visible}'
    assert ",2,FALSE)" in visible, f'no ordena por Leads asignados desc: {visible}'

    # % conversion en el staging: numerador Clientes activos (M), denominador
    # Leads asignados (K). Mezclarlas (paso una vez con Demos en vez de
    # Clientes) da un % que no es el que dice el encabezado.
    pct = h.celda(f"N{fila_r1}")
    assert pct == f'=IF(N(K{fila_r1})=0,"",M{fila_r1}/K{fila_r1})', f'% mal armado: {pct}'

    # Motivo dominante: el rango de MAX/MATCH tiene que ser el mismo que
    # llena el staging de motivos, ni una fila mas ni una menos.
    fila_m1 = fila_1based("El que más frena, no la lista completa.") + 1
    rango_esperado = f"J{fila_m1}:J{fila_m1 + len(MOTIVOS) - 1}"
    dominante = h.celda(f"A{fila_m1}")
    assert dominante.count(rango_esperado) == 2, \
        f'MAX y MATCH no miran el mismo rango que el staging: {dominante}'
    assert all(f'"{m}"' in dominante for m in MOTIVOS), f'falta un motivo: {dominante}'
    assert "Cantidad" not in str(h.escrituras), 'quedo la columna Cantidad a la vista'

    # El staging tiene que estar escondido, si no todo esto era para nada.
    ocultos = [r for b in libro.batches for r in b['requests']
              if 'updateDimensionProperties' in r
              and r['updateDimensionProperties']['range'].get('dimension') == 'COLUMNS'
              and r['updateDimensionProperties']['properties'].get('hiddenByUser')]
    assert ocultos, 'las columnas J:O quedaron a la vista'
    rango_oculto = ocultos[0]['updateDimensionProperties']['range']
    assert (rango_oculto['startIndex'], rango_oculto['endIndex']) == (9, 15), \
        f'esconde las columnas que no son: {rango_oculto}'


def _probar_reintentar(dormido):
    llamadas = []
    def falla_dos_veces():
        llamadas.append(1)
        if len(llamadas) <= 2:
            raise Exception("APIError: [429]: Quota exceeded")
        return "listo"
    assert reintentar(falla_dos_veces) == "listo", 'no reintento tras el 429'
    assert len(llamadas) == 3, f'reintento {len(llamadas)} veces'

    # El 20/8 un 503 al abrir el libro tiro abajo la sync entera y se perdieron
    # 4290 negocios de esa corrida. Tiene que reintentar igual que el 429.
    llamadas_503 = []
    def falla_una_vez_503():
        llamadas_503.append(1)
        if len(llamadas_503) == 1:
            raise Exception("APIError: [503]: The service is currently unavailable.")
        return "listo"
    assert reintentar(falla_una_vez_503) == "listo", 'no reintento tras el 503'

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
    elif '--nuevas' in sys.argv:
        poner_al_dia()
    else:
        # armar() borra y reescribe TODO. Con vendedores cargando datos va
        # --nuevas, que solo formatea lo que creo el sync.
        if input("armar() rehace la planilla de cero y se lleva puesto lo que "
                 "hayan cargado los vendedores.\nSi solo queres formatear las "
                 "hojas nuevas es --nuevas.\nEscribi REHACER para seguir: ") != "REHACER":
            sys.exit("Cancelado.")
        armar()
