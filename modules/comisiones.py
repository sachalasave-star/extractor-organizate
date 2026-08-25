"""Liquidacion de comisiones: cruza los clientes que pagan en la web de
Organizate contra la planilla de ventas, por telefono.

50% el vendedor que cerro, 20% el que lo trajo (Referidos), 30% Organizate.
Una fila por cliente por mes por cada pago que aparezca en el endpoint: el
disparador es plata que entro de verdad, no un estado marcado a mano.

El telefono es la clave de cruce, igual que en el resto del sistema (ver
modules/telefono.py). El endpoint lo manda crudo y sin validar a proposito:
sacar el 15 sin poner el 9 deja un celular identico a un fijo, asi que
`canonico()` es lo unico que puede desambiguarlo.
"""
import os
import sys

from modules.telefono import canonico

ENDPOINT = "https://www.organizate.click/api/organizate/clientes"

PCT_VENDEDOR = 0.50
PCT_REFERIDO = 0.20
PCT_ORGANIZATE = 0.30

LIQUIDACION = '💰 Liquidación'
REFERIDOS = '🤝 Referidos'

COLUMNAS_LIQUIDACION = ['Cliente', 'Negocio', 'Teléfono', 'Período', 'Monto',
                        'Vendedor', 'Comisión vendedor', 'Referido por',
                        'Comisión referido', 'Organizate', 'Corte', 'Alta']
COLUMNAS_REFERIDOS = ['Vendedor', 'Referido por']


def _token():
    # Mismo problema de BOM que GOOGLE_CREDENTIALS y SHEET_ID en
    # sincronizar_sheets.py: un secret cargado desde PowerShell puede traer
    # basura invisible adelante o atras.
    t = os.environ.get("ORGANIZATE_TOKEN", "").strip('﻿ \t\r\n')
    if not t:
        sys.exit("Falta ORGANIZATE_TOKEN.")
    return t


def obtener_clientes():
    """Trae la lista de clientes que pagan, desde el endpoint de Organizate."""
    import requests
    r = requests.get(ENDPOINT, headers={"Authorization": f"Bearer {_token()}"}, timeout=30)
    r.raise_for_status()
    return r.json()


def corte(fecha_alta):
    """15 o 30: que dia del mes cae la liquidacion, segun el dia de alta."""
    try:
        dia = int(fecha_alta.split('-')[2])
    except (IndexError, ValueError):
        dia = 1
    return 15 if dia <= 15 else 30


def liquidar(clientes, telefono_a_vendedor, referido_por):
    """[filas] para la hoja Liquidacion: una por cliente por pago cobrado.

    telefono_a_vendedor: {telefono canonico: vendedor}, sale de cruzar las
    45 hojas de la planilla. referido_por: {vendedor: quien lo trajo}, sale
    de la hoja Referidos (la completa el equipo a mano).

    Un cliente cuyo telefono no matchea ningun negocio de la planilla no
    genera fila: no hay a quien pagarle, y una comision al vendedor
    equivocado es peor que ninguna.
    """
    filas = []
    for cliente in clientes:
        vendedor = telefono_a_vendedor.get(canonico(cliente.get('telefono', '')), '')
        if not vendedor:
            continue
        referido = referido_por.get(vendedor, '')
        alta = cliente.get('alta', '')
        for pago in cliente.get('pagos', []):
            monto = float(pago.get('monto', 0))
            filas.append([
                cliente.get('id', ''), cliente.get('negocio', ''),
                cliente.get('telefono', ''), pago.get('periodo', ''), monto,
                vendedor, round(monto * PCT_VENDEDOR, 2),
                referido, round(monto * PCT_REFERIDO, 2) if referido else 0,
                round(monto * PCT_ORGANIZATE, 2), corte(alta), alta,
            ])
    return filas


def demo():
    clientes = [
        {'id': 'cli_1', 'negocio': 'Barberia A', 'telefono': '0341 353-9510',
         'alta': '2026-08-10', 'pagos': [{'periodo': '2026-08', 'monto': 15000}]},
        {'id': 'cli_2', 'negocio': 'Spa B', 'telefono': '011 15-2620-5229',
         'alta': '2026-08-20', 'pagos': [{'periodo': '2026-08', 'monto': 10000},
                                         {'periodo': '2026-09', 'monto': 10000}]},
        # Sin match en la planilla: no genera fila.
        {'id': 'cli_3', 'negocio': 'Nadie', 'telefono': '0341 000-0000',
         'alta': '2026-08-01', 'pagos': [{'periodo': '2026-08', 'monto': 5000}]},
    ]
    tel_a_vend = {canonico('3413539510'): 'Augusto', canonico('1126205229'): 'Valentino'}
    referido_por = {'Valentino': 'Augusto'}   # Augusto trajo a Valentino

    filas = liquidar(clientes, tel_a_vend, referido_por)
    assert len(filas) == 3, f'esperaba 3 pagos liquidados (2 de cli_2), dio {len(filas)}'

    f1 = next(f for f in filas if f[0] == 'cli_1')
    assert f1[5] == 'Augusto' and f1[6] == 7500, f1          # 50% de 15000
    assert f1[7] == '' and f1[8] == 0, 'Augusto no tiene referido, no se inventa comision'
    assert f1[9] == 4500, f1                                  # 30% de 15000
    assert f1[10] == 15, 'alta el 10 cae en el corte del 15'

    f2 = next(f for f in filas if f[0] == 'cli_2' and f[3] == '2026-09')
    assert f2[5] == 'Valentino' and f2[6] == 5000, f2         # 50% de 10000
    assert f2[7] == 'Augusto' and f2[8] == 2000, f2           # 20% de 10000
    assert f2[10] == 30, 'alta el 20 cae en el corte del 30'

    assert not any(f[0] == 'cli_3' for f in filas), 'un telefono sin match no debe cobrar'

    print('OK  comisiones: 50/20/30 por pago, corte por dia de alta, sin match no liquida')


if __name__ == '__main__':
    demo()
