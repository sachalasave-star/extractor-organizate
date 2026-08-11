"""Manda los negocios nuevos a la planilla de Google Sheets, SIN tocar lo que
anotaron los vendedores.

    python sincronizar_sheets.py

Solo AGREGA filas al final de la hoja de cada nicho. Las que ya estan no se
leen para reescribirlas: se dejan donde estan. Asi Vendedor, Estado y
Observaciones sobreviven aunque alguien este escribiendo en ese momento.

La estructura la arma armar_planilla.py; esto es el dia a dia.

Necesita:
  GOOGLE_CREDENTIALS  json de la cuenta de servicio (o archivo credenciales.json)
  SHEET_ID            el id de la planilla, sale de su URL
"""
import json
import os
import sys

import pandas as pd

from modules.planilla import (COLUMNAS, CLAVE, CONFIG, PANEL, RESUMEN, IDX,
                              col_letra, fila_desde, reintentar)
from modules.telefono import canonico

GENERADO = "output/Organizate.xlsx"


def _credenciales():
    crudo = os.environ.get("GOOGLE_CREDENTIALS")
    if crudo:
        # lstrip del BOM: guardar el secret con PowerShell (Get-Content -Raw)
        # le mete un ﻿ adelante y json.loads no lo tolera. Fallaba en todas
        # las corridas de la nube con "Unexpected UTF-8 BOM".
        return json.loads(crudo.lstrip('﻿').strip())
    if os.path.exists("credenciales.json"):
        with open("credenciales.json", encoding="utf-8-sig") as f:
            return json.load(f)
    sys.exit("Falta GOOGLE_CREDENTIALS (o el archivo credenciales.json). "
             "Ver las instrucciones en el README.")


def _sheet_id():
    # Mismo cuento que el BOM de las credenciales: el secret cargado desde
    # PowerShell viene con un ﻿ adelante o un \n atras, y Sheets contesta
    # 404 "Requested entity was not found" (no dice que el id venga sucio).
    # Dos dias de corridas scrapearon bien y no subieron una sola fila.
    return os.environ.get("SHEET_ID", "").strip('﻿ \t\r\n')


def _abrir_libro():
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        sys.exit("Falta gspread: pip install gspread google-auth")
    cred = Credentials.from_service_account_info(
        _credenciales(), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(cred).open_by_key(_sheet_id())


def nuevos_por_nicho(df, ya_cargados):
    """{nicho: [filas]} con lo que todavia no esta en la planilla."""
    salida = {}
    for nicho, grupo in df.groupby('Nicho'):
        faltan = grupo[~grupo[CLAVE].isin(ya_cargados.get(nicho, set()))]
        if not faltan.empty:
            salida[nicho] = [fila_desde(r) for r in faltan.to_dict('records')]
    return salida


def sincronizar(generado=GENERADO):
    if not os.path.exists(generado):
        sys.exit(f"No existe {generado}. Corre el scraper primero.")
    if not _sheet_id() or not (
            os.environ.get("GOOGLE_CREDENTIALS") or os.path.exists("credenciales.json")):
        print("Google Sheets sin configurar (falta SHEET_ID o las credenciales), salteando.")
        return

    df = pd.read_excel(generado, sheet_name="Todos", dtype=str).fillna("")
    df = df[df[CLAVE].str.strip() != ""]

    libro = _abrir_libro()
    hojas = {h.title: h for h in reintentar(libro.worksheets)
             if h.title not in (CONFIG, PANEL, RESUMEN)}

    # Solo la columna del telefono: traer las hojas enteras seria lentisimo y
    # ademas arriesga leer datos a medio escribir por un vendedor.
    col = chr(ord('A') + IDX[CLAVE])
    ya_cargados = {}
    for titulo, hoja in hojas.items():
        try:
            ya_cargados[titulo] = set(reintentar(hoja.col_values, IDX[CLAVE] + 1)[1:])
        except Exception as e:
            print(f"   no pude leer '{titulo}' ({e}), la salteo")
            ya_cargados[titulo] = None

    total = 0
    for nicho, filas in nuevos_por_nicho(df, {k: v for k, v in ya_cargados.items() if v}).items():
        if nicho not in hojas:
            hoja = reintentar(libro.add_worksheet, title=nicho[:99],
                              rows=len(filas) + 500, cols=len(COLUMNAS))
            reintentar(hoja.update, values=[COLUMNAS], range_name="A1")
            print(f"   hoja nueva: {nicho} (correr armar_planilla.py para darle formato)")
        elif ya_cargados.get(nicho) is None:
            continue                      # no se pudo leer: mejor no duplicar
        else:
            hoja = hojas[nicho]
        # append_rows escribe solo al final: no toca una celda de lo que ya hay.
        reintentar(hoja.append_rows, filas, value_input_option="RAW")
        print(f"   {nicho}: +{len(filas)}")
        total += len(filas)

    print(f"Agregados {total} negocios nuevos." if total else "Sin novedades.")


def calificar_existentes(generado=GENERADO):
    """Rellena Prioridad y Por que en las filas que ya estaban en la planilla.

    Las filas nuevas las trae el sync solo, porque Prioridad entro en COLUMNAS.
    Las 18291 que ya estaban quedaron con las dos columnas vacias, que son
    justo las que el equipo esta llamando.

    Escribe SOLO esas dos columnas, cruzando por telefono normalizado. No toca
    Vendedor, Estado, Motivo, Observaciones ni Ultima gestion: se calcula el
    rango exacto de las dos ultimas columnas y se escribe ahi.
    """
    df = pd.read_excel(generado, sheet_name="Todos", dtype=str).fillna("")
    calif = {canonico(t): (p, q) for t, p, q in
             zip(df[CLAVE], df['Prioridad'], df['Por qué']) if canonico(t)}
    print(f"{len(calif)} negocios calificados en el Excel")

    libro = _abrir_libro()
    meta = reintentar(libro.fetch_sheet_metadata)
    hojas = {h['properties']['title']: h['properties'] for h in meta.get('sheets', [])
             if h['properties']['title'] not in (CONFIG, PANEL, RESUMEN)}

    # Las hojas se crearon con exactamente 10 columnas. La API de valores no
    # agranda la grilla sola: escribir en la 11 da "exceeds grid limits".
    angostas = [p for p in hojas.values()
                if p.get('gridProperties', {}).get('columnCount', 0) < len(COLUMNAS)]
    if angostas:
        reintentar(libro.batch_update, {"requests": [
            {"updateSheetProperties": {
                "properties": {"sheetId": p['sheetId'],
                               "gridProperties": {"columnCount": len(COLUMNAS)}},
                "fields": "gridProperties.columnCount"}} for p in angostas]})
        print(f"   {len(angostas)} hojas ensanchadas a {len(COLUMNAS)} columnas")

    desde = col_letra(IDX['Prioridad'])
    hasta = col_letra(IDX['Por qué'])
    col_tel = col_letra(IDX[CLAVE])

    # Las 45 columnas de telefono en UN request. Una hoja por vez son 45
    # lecturas, y libro.worksheet(titulo) se trae el metadata entero cada vez,
    # asi que eran 90: el limite de Google son 60 por minuto y reventaba.
    titulos = list(hojas)
    leidas = reintentar(libro.values_batch_get,
                        [f"'{t}'!{col_tel}2:{col_tel}" for t in titulos])

    datos, sin_datos = [], 0
    for titulo, rango in zip(titulos, leidas.get('valueRanges', [])):
        tels = [f[0] if f else '' for f in rango.get('values', [])]
        if not tels:
            continue
        valores = [[COLUMNAS[IDX['Prioridad']], COLUMNAS[IDX['Por qué']]]]
        for t in tels:
            p, q = calif.get(canonico(t), ('', ''))
            if not p:
                sin_datos += 1
            valores.append([p, q])
        datos.append({"range": f"'{titulo}'!{desde}1:{hasta}{len(valores)}",
                      "values": valores})
        print(f"   {titulo}: {len(tels)} filas")

    if datos:
        reintentar(libro.values_batch_update,
                   {"valueInputOption": "RAW", "data": datos})
    filas = sum(len(d['values']) - 1 for d in datos)
    print(f"\nCalificadas {filas - sin_datos}/{filas} filas en {len(datos)} hojas.")
    if sin_datos:
        print(f"   {sin_datos} sin calificar: el telefono no esta en el Excel "
              f"(negocio borrado de la base o telefono editado a mano).")


def demo():
    from modules.planilla import SIN_CONTACTAR
    df = pd.DataFrame([
        {'Nicho': 'Barberías', 'Negocio': 'A', CLAVE: '0341111', 'Categoría': 'Barbería',
         'Ciudad': 'Rosario', 'Link en Maps': 'u1'},
        {'Nicho': 'Barberías', 'Negocio': 'B', CLAVE: '0341222', 'Categoría': 'Barbería',
         'Ciudad': 'Rosario', 'Link en Maps': 'u2'},
        {'Nicho': 'Spas', 'Negocio': 'C', CLAVE: '0341333', 'Categoría': 'Spa',
         'Ciudad': 'Rosario', 'Link en Maps': 'u3'},
    ])
    # En Barberías ya esta cargado el 0341111
    r = nuevos_por_nicho(df, {'Barberías': {'0341111'}})

    assert set(r) == {'Barberías', 'Spas'}, f'nichos mal separados: {set(r)}'
    assert len(r['Barberías']) == 1, 'remando un negocio que ya estaba'
    assert r['Barberías'][0][IDX['Negocio']] == 'B'
    assert len(r['Spas']) == 1
    assert r['Spas'][0][IDX['Estado']] == SIN_CONTACTAR
    assert all(len(f) == len(COLUMNAS) for fs in r.values() for f in fs), 'ancho de fila mal'
    print('OK  sync: separa por nicho y no repite lo ya cargado')


if __name__ == "__main__":
    if '--test' in sys.argv:
        demo()
    elif '--calificar' in sys.argv:
        calificar_existentes()
    else:
        sincronizar()
