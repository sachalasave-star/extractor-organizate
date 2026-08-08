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
from sincronizar_sheets import _credenciales, GENERADO

MAX_VENDEDORES = 40          # margen para sumar vendedores sin rehacer nada


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


def _formato_hoja(sid, filas):
    """Requests de formato para una hoja de nicho."""
    ultima = max(filas + 1, 2000)        # deja margen para que crezca sola
    reqs = [
        # Encabezado
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb((0.12, 0.24, 0.42)),
                "textFormat": {"bold": True, "foregroundColor": _rgb((1, 1, 1)), "fontSize": 11},
                "verticalAlignment": "MIDDLE"}},
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment)"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1}},
            "fields": "gridProperties.frozenRowCount"}},
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
             ["(agregá o borrá acá: se actualiza en todas las hojas)", "", "(no tocar el texto)", ""]]
    for i in range(max(len(VENDEDORES), len(ESTADOS))):
        filas.append([VENDEDORES[i] if i < len(VENDEDORES) else "", "",
                      NOMBRES_ESTADO[i] if i < len(ESTADOS) else "", ""])
    h.update(values=filas, range_name="A1")
    return h


def _hoja_ranking(libro, nichos):
    try:
        libro.del_worksheet(libro.worksheet(RANKING))
    except Exception:
        pass
    h = libro.add_worksheet(title=RANKING, rows=60, cols=10, index=1)

    # SUMPRODUCT sobre INDIRECT recorre las 43 hojas sin escribir 43 formulas.
    hojas = "{" + ";".join(f'"{n}"' for n in nichos) + "}"

    def contar(col, condicion):
        return (f'=SUMPRODUCT(COUNTIF(INDIRECT("\'"&{hojas}&"\'!{col}2:{col}"),{condicion}))')

    filas = [["RANKING DE VENDEDORES", "", "", "", "", ""],
             ["Se actualiza solo. Contá desde acá quién movió el embudo.", "", "", "", "", ""],
             ["Vendedor", "Clientes activos", "Demos", "Interesados", "Asignados", "Sin contactar"]]
    for i in range(len(VENDEDORES)):
        f = 4 + i
        v = f"'{CONFIG}'!A{FILA_VENDEDORES + i}"
        filas.append([
            f"={v}",
            f'=IF($A{f}="","",SUMPRODUCT(COUNTIFS(INDIRECT("\'"&{hojas}&"\'!{COL_VENDEDOR}2:{COL_VENDEDOR}"),$A{f},INDIRECT("\'"&{hojas}&"\'!{COL_ESTADO}2:{COL_ESTADO}"),"Cliente activo")))',
            f'=IF($A{f}="","",SUMPRODUCT(COUNTIFS(INDIRECT("\'"&{hojas}&"\'!{COL_VENDEDOR}2:{COL_VENDEDOR}"),$A{f},INDIRECT("\'"&{hojas}&"\'!{COL_ESTADO}2:{COL_ESTADO}"),"Demo iniciada")))',
            f'=IF($A{f}="","",SUMPRODUCT(COUNTIFS(INDIRECT("\'"&{hojas}&"\'!{COL_VENDEDOR}2:{COL_VENDEDOR}"),$A{f},INDIRECT("\'"&{hojas}&"\'!{COL_ESTADO}2:{COL_ESTADO}"),"Le interesó")))',
            f'=IF($A{f}="","",SUMPRODUCT(COUNTIF(INDIRECT("\'"&{hojas}&"\'!{COL_VENDEDOR}2:{COL_VENDEDOR}"),$A{f})))',
            f'=IF($A{f}="","",SUMPRODUCT(COUNTIFS(INDIRECT("\'"&{hojas}&"\'!{COL_VENDEDOR}2:{COL_VENDEDOR}"),$A{f},INDIRECT("\'"&{hojas}&"\'!{COL_ESTADO}2:{COL_ESTADO}"),"Sin contactar")))',
        ])

    f0 = 4 + len(VENDEDORES) + 1
    filas.append([""] * 6)
    filas.append(["ESTADO GENERAL DE LA BASE", "", "", "", "", ""])
    filas.append(["Total negocios", contar(COL_TELEFONO, '"<>"'), "", "", "", ""])
    for estado in NOMBRES_ESTADO:
        filas.append([estado, contar(COL_ESTADO, f'"{estado}"'), "", "", "", ""])

    h.update(values=filas, range_name="A1", value_input_option="USER_ENTERED")
    libro.batch_update({"requests": [
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 14}}},
            "fields": "userEnteredFormat.textFormat"}},
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": 2, "endRowIndex": 3},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb((0.12, 0.24, 0.42)),
                "textFormat": {"bold": True, "foregroundColor": _rgb((1, 1, 1))}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"}},
        {"repeatCell": {
            "range": {"sheetId": h.id, "startRowIndex": f0, "endRowIndex": f0 + 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 12}}},
            "fields": "userEnteredFormat.textFormat"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 220}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": h.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 6},
            "properties": {"pixelSize": 130}, "fields": "pixelSize"}},
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

    reqs_formato = []
    for titulo, filas in datos.items():
        reqs_formato += _formato_hoja(hojas[titulo].id, len(filas))
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
