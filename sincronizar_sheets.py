"""Manda los negocios nuevos a una planilla de Google Sheets, SIN tocar lo que
anotaron los vendedores.

    python sincronizar_sheets.py

Igual que actualizar.py pero contra Sheets: solo AGREGA filas al final. Las que
ya estan no se leen para reescribirlas, se dejan donde estan. Asi Estado,
Vendedor y Notas sobreviven aunque un vendedor este escribiendo en ese momento.

Se adapta a las columnas de la planilla: si le agregas o renombras una desde
Google Sheets, las filas nuevas se acomodan a lo que encuentre.

Necesita:
  GOOGLE_CREDENTIALS  json de la cuenta de servicio (o archivo credenciales.json)
  SHEET_ID            el id de la planilla, sale de su URL
"""
import json
import os
import sys

import pandas as pd

GENERADO = "output/Organizate.xlsx"
CLAVE = "Teléfono"
HOJA = "Negocios"
GESTION = ["Estado", "Vendedor", "Fecha de llamada", "Notas"]
SIN_LLAMAR = "Sin llamar"


def _credenciales():
    crudo = os.environ.get("GOOGLE_CREDENTIALS")
    if crudo:
        return json.loads(crudo)
    if os.path.exists("credenciales.json"):
        with open("credenciales.json", encoding="utf-8") as f:
            return json.load(f)
    sys.exit("Falta GOOGLE_CREDENTIALS (o el archivo credenciales.json). "
             "Ver las instrucciones en el README.")


def _abrir_hoja():
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        sys.exit("Falta gspread: pip install gspread google-auth")

    sheet_id = os.environ.get("SHEET_ID")
    if not sheet_id:
        sys.exit("Falta SHEET_ID: es el codigo largo que aparece en la url de la "
                 "planilla, entre /d/ y /edit")

    cred = Credentials.from_service_account_info(
        _credenciales(),
        scopes=["https://www.googleapis.com/auth/spreadsheets"])
    libro = gspread.authorize(cred).open_by_key(sheet_id)
    try:
        return libro.worksheet(HOJA)
    except Exception:
        return libro.add_worksheet(title=HOJA, rows=1000, cols=30)


def filas_nuevas(generado, encabezados, telefonos_existentes):
    """Las filas del Excel generado que todavia no estan en la planilla,
    acomodadas al orden de columnas que ya tiene la planilla."""
    hoja_datos = "Todos" if "Todos" in pd.ExcelFile(generado).sheet_names else 0
    df = pd.read_excel(generado, sheet_name=hoja_datos, dtype=str).fillna("")
    df = df[df[CLAVE].astype(str).str.strip() != ""]
    df = df[~df[CLAVE].isin(telefonos_existentes)]
    if df.empty:
        return []

    for col in encabezados:
        if col not in df.columns:
            df[col] = SIN_LLAMAR if col == "Estado" else ""
    return df[encabezados].astype(str).values.tolist()


def sincronizar(generado=GENERADO):
    if not os.path.exists(generado):
        sys.exit(f"No existe {generado}. Corre el scraper primero.")

    # Sin configurar todavia no es un error: el scraping igual sirve y se guarda.
    if not os.environ.get("SHEET_ID") or not (
            os.environ.get("GOOGLE_CREDENTIALS") or os.path.exists("credenciales.json")):
        print("Google Sheets sin configurar (falta SHEET_ID o las credenciales), salteando.")
        return

    hoja = _abrir_hoja()
    existente = hoja.get_all_values()

    if not existente or not any(existente[0]):
        # Planilla vacia: la inicializa con las columnas del export + gestion.
        cols = pd.read_excel(generado, sheet_name="Todos", nrows=0).columns.tolist()
        encabezados = cols + GESTION
        hoja.update(values=[encabezados], range_name="A1")
        hoja.freeze(rows=1)
        existente = [encabezados]

    encabezados = existente[0]
    if CLAVE not in encabezados:
        sys.exit(f"La planilla no tiene columna '{CLAVE}'. Sin ella no puedo "
                 f"saber que negocio ya esta cargado.")

    i_tel = encabezados.index(CLAVE)
    existentes = {f[i_tel] for f in existente[1:] if len(f) > i_tel and f[i_tel]}

    nuevas = filas_nuevas(generado, encabezados, existentes)
    if not nuevas:
        print(f"Sin novedades: los {len(existentes)} negocios ya estaban en la planilla.")
        return

    # append_rows escribe solo al final: no toca ni una celda de lo que ya hay.
    hoja.append_rows(nuevas, value_input_option="RAW")
    print(f"Agregados {len(nuevas)} negocios ({len(existentes)} ya estaban, intactos).")


def demo():
    """Lo que importa: acomodar las filas nuevas a las columnas de la planilla y
    no volver a mandar las que ya estan."""
    import tempfile
    d = tempfile.mkdtemp()
    gen = os.path.join(d, 'g.xlsx')
    pd.DataFrame({'Negocio': ['A', 'B', 'C'],
                  CLAVE: ['0341111', '0341222', '0341333'],
                  'Ciudad': ['Rosario'] * 3}).to_excel(gen, index=False, sheet_name='Todos')

    # La planilla tiene otro orden de columnas y una propia que el export no conoce
    encabezados = [CLAVE, 'Negocio', 'Estado', 'Vendedor', 'Prioridad']
    filas = filas_nuevas(gen, encabezados, {'0341111'})

    assert len(filas) == 2, f'esperaba 2 filas nuevas, hay {len(filas)}'
    assert filas[0][0] == '0341222', f'no respeto el orden de columnas: {filas[0]}'
    assert filas[0][1] == 'B'
    assert filas[0][2] == SIN_LLAMAR, 'no inicializo el Estado'
    assert filas[0][4] == '', 'una columna propia de la planilla deberia quedar vacia'
    assert all(f[0] != '0341111' for f in filas), 'remando un negocio que ya estaba'
    print('OK  sheets: respeta las columnas de la planilla y no repite negocios')


if __name__ == "__main__":
    if '--test' in sys.argv:
        demo()
    else:
        sincronizar()
