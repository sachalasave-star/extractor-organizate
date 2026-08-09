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
                              fila_desde, reintentar)

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


def _abrir_libro():
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        sys.exit("Falta gspread: pip install gspread google-auth")
    cred = Credentials.from_service_account_info(
        _credenciales(), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(cred).open_by_key(os.environ["SHEET_ID"])


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
    if not os.environ.get("SHEET_ID") or not (
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
    else:
        sincronizar()
