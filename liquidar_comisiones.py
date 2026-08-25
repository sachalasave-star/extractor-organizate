"""Liquida las comisiones del mes: cruza los clientes que pagan en la web de
Organizate contra la planilla de ventas, y escribe la hoja Liquidacion.

    python liquidar_comisiones.py

Necesita, ademas de GOOGLE_CREDENTIALS y SHEET_ID (ver sincronizar_sheets.py):
  ORGANIZATE_TOKEN    token fijo del endpoint de clientes de Organizate

La hoja Referidos (vendedor -> quien lo trajo) la completa el equipo a mano;
si no existe se crea vacia la primera vez. La hoja Liquidacion se rehace
entera en cada corrida: sale toda de datos de origen (el endpoint + la
planilla + Referidos), nunca de algo que alguien tipeo ahi, asi que no hay
nada que perder al reescribirla.
"""
import os
import sys

from modules.comisiones import (COLUMNAS_LIQUIDACION, COLUMNAS_REFERIDOS,
                                LIQUIDACION, REFERIDOS, liquidar, obtener_clientes)
from modules.planilla import CLAVE, CONFIG, IDX, PANEL, RANKING, RESUMEN, VENDEDORES, col_letra, reintentar
from modules.telefono import canonico
from sincronizar_sheets import _abrir_libro, _sheet_id

FUERA_DE_NICHOS = (CONFIG, PANEL, RESUMEN, RANKING, LIQUIDACION, REFERIDOS)


def _telefono_a_vendedor(libro):
    """{telefono canonico: vendedor} cruzando Telefono y Vendedor de las 45
    hojas de nicho. Una fila sin vendedor cargado no aporta nada: nadie cerro
    ese negocio todavia."""
    meta = reintentar(libro.fetch_sheet_metadata)
    titulos = [h['properties']['title'] for h in meta.get('sheets', [])
               if h['properties']['title'] not in FUERA_DE_NICHOS]
    col_tel = col_letra(IDX[CLAVE])
    col_vend = col_letra(IDX['Vendedor'])
    leidas = reintentar(libro.values_batch_get,
                        [f"'{t}'!{col_tel}2:{col_vend}" for t in titulos])

    salida = {}
    for rango in leidas.get('valueRanges', []):
        for fila in rango.get('values', []):
            tel = fila[0] if fila else ''
            vend = fila[-1] if len(fila) > 1 else ''
            c = canonico(tel)
            if c and vend:
                salida[c] = vend
    return salida


def _hoja_referidos(libro):
    """Trae la hoja Referidos, creandola vacia (un vendedor por fila) si es
    la primera corrida. Nunca se reescribe si ya existe: la completa el
    equipo a mano."""
    try:
        return libro.worksheet(REFERIDOS)
    except Exception:
        h = reintentar(libro.add_worksheet, title=REFERIDOS,
                       rows=len(VENDEDORES) + 10, cols=len(COLUMNAS_REFERIDOS))
        filas = [COLUMNAS_REFERIDOS] + [[v, ''] for v in VENDEDORES]
        reintentar(h.update, values=filas, range_name="A1")
        print(f"   {REFERIDOS} creada vacia: completar 'Referido por' a mano")
        return h


def _referido_por(hoja):
    valores = reintentar(hoja.get_all_values)[1:]
    return {fila[0]: fila[1] for fila in valores if len(fila) > 1 and fila[1].strip()}


def _hoja_liquidacion(libro, filas):
    try:
        reintentar(libro.del_worksheet, libro.worksheet(LIQUIDACION))
    except Exception:
        pass
    h = reintentar(libro.add_worksheet, title=LIQUIDACION,
                   rows=len(filas) + 10, cols=len(COLUMNAS_LIQUIDACION))
    reintentar(h.update, values=[COLUMNAS_LIQUIDACION] + filas, range_name="A1")
    return h


def main():
    if not _sheet_id() or not (
            os.environ.get("GOOGLE_CREDENTIALS") or os.path.exists("credenciales.json")):
        print("Google Sheets sin configurar (falta SHEET_ID o las credenciales), salteando.")
        return
    if not os.environ.get("ORGANIZATE_TOKEN", "").strip():
        print("Falta ORGANIZATE_TOKEN, salteando la liquidacion de comisiones.")
        return

    libro = _abrir_libro()
    tel_a_vend = _telefono_a_vendedor(libro)
    referidos_h = _hoja_referidos(libro)
    referido_por = _referido_por(referidos_h)

    clientes = obtener_clientes()
    filas = liquidar(clientes, tel_a_vend, referido_por)
    _hoja_liquidacion(libro, filas)

    sin_match = sum(1 for c in clientes if c.get('pagos')
                    and not tel_a_vend.get(canonico(c.get('telefono', ''))))
    print(f"{LIQUIDACION}: {len(filas)} pagos liquidados, {len(clientes)} clientes en Organizate"
          f"{f', {sin_match} sin match en la planilla' if sin_match else ''}.")


if __name__ == "__main__":
    if '--test' in sys.argv:
        import modules.comisiones as demo_mod
        demo_mod.demo()
    else:
        main()
