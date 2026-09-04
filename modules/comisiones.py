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
from datetime import datetime

from modules.telefono import canonico

ENDPOINT = "https://www.organizate.click/api/organizate/clientes"

PCT_VENDEDOR = 0.50
PCT_REFERIDO = 0.20
PCT_ORGANIZATE = 0.30

LIQUIDACION = '💰 Liquidación'
REFERIDOS = '🤝 Referidos'
RESUMEN_COMISIONES = '💵 Comisiones por vendedor'
CLIENTES = '💳 Clientes de Organizate'

COLUMNAS_LIQUIDACION = ['Cliente', 'Negocio', 'Teléfono', 'Período', 'Monto',
                        'Vendedor', 'Comisión vendedor', 'Referido por',
                        'Comisión referido', 'Organizate', 'Corte', 'Alta']
COLUMNAS_REFERIDOS = ['Vendedor', 'Referido por']

ASIGNAR = '🔗 Asignar a mano'
# El ID va ultimo y angosto: es la clave con la que se guarda la asignacion
# entre corridas, no algo para leer. El nombre del negocio no sirve de clave
# porque el dueño lo puede cambiar cuando quiera.
COLUMNAS_ASIGNAR = ['Negocio en Organizate', 'Alta', 'Estado', 'Vendedor', 'ID']

# Como llega 'estado' del endpoint -> como se muestra.
#   activo      paga hoy
#   demo        en prueba, todavia no pago nunca
#   cancelado   pagaba y se dio de baja (baja trae la fecha)
#   abandonado  se registro, nunca pago, tampoco esta en demo
# 'baja' es el nombre viejo de cancelado; se deja mapeado por si vuelve.
ESTADO_LABEL = {'activo': 'Activo', 'demo': 'Demo', 'cancelado': 'Cancelado',
                'baja': 'Cancelado', 'abandonado': 'Abandonado'}

# Mismos colores que los estados del embudo en planilla.ESTADOS (Cliente
# activo, Demo iniciada, No interesado), para que la planilla se lea igual
# de una hoja a la otra. Abandonado va gris: no es una mala noticia como una
# baja, es alguien que nunca arranco.
COLOR_ESTADO = {
    'Activo':     ((0.20, 0.66, 0.33), (1, 1, 1)),
    'Demo':       ((0.55, 0.30, 0.80), (1, 1, 1)),
    'Cancelado':  ((0.85, 0.24, 0.24), (1, 1, 1)),
    'Abandonado': ((0.42, 0.42, 0.46), (1, 1, 1)),
}


def estado_label(estado):
    """Un estado que no conocemos se muestra TAL CUAL, nunca como 'Activo'.

    Antes el default era 'Activo', con el argumento de que era el unico que
    cobra y mejor mostrar de mas. Salio mal apenas Organizate sumo estados:
    los 8 clientes 'abandonado' se mostraban como Activo, o sea justo al
    reves de la realidad. Un estado desconocido tiene que cantar que es
    desconocido, no disfrazarse del mas optimista.
    """
    return ESTADO_LABEL.get(estado, (estado or 'sin estado').capitalize())


def par_telefono_vendedor(fila):
    """Una fila leida del rango Telefono..Vendedor -> {clave: vendedor} o {}.

    El vendedor se toma por su POSICION exacta, nunca con fila[-1]. Sheets
    recorta las celdas vacias del final, asi que en una fila sin vendedor
    cargado -que son casi todas- el ultimo valor es la Categoria. Con fila[-1]
    la planilla devolvia 86734 "vendedores" que en realidad eran nichos
    ("Peluquería", "Centro de estética"), y una comision se le hubiera
    liquidado a una categoria en vez de a una persona.
    """
    from modules.planilla import CLAVE as _CLAVE, IDX as _IDX
    salto = _IDX['Vendedor'] - _IDX[_CLAVE]
    if len(fila) <= salto:
        return {}
    clave = canonico(fila[0] if fila else '')
    vendedor = (fila[salto] or '').strip()
    return {clave: vendedor} if clave and vendedor else {}


def sin_cuentas_internas(clientes):
    """Saca las cuentas internas (demos de vendedores, pruebas de desarrollo).

    Son negocios gratis que existen a proposito. Si entran al calculo, un
    vendedor puede terminar cobrando comision por una cuenta de mentira.
    """
    return [c for c in clientes if not c.get('es_cuenta_interna')]


def _token():
    # Mismo problema de BOM que GOOGLE_CREDENTIALS y SHEET_ID en
    # sincronizar_sheets.py: un secret cargado desde PowerShell puede traer
    # basura invisible adelante o atras.
    t = os.environ.get("ORGANIZATE_TOKEN", "").strip('﻿ \t\r\n')
    if not t:
        sys.exit("Falta ORGANIZATE_TOKEN.")
    return t


def obtener_clientes():
    """Trae los clientes desde el endpoint de Organizate, sin cuentas internas.

    El filtro va aca, en la puerta de entrada, y no en cada consumidor: si
    alguna vez se agrega otro calculo sobre los clientes, arranca limpio sin
    tener que acordarse de filtrar.
    """
    import requests
    r = requests.get(ENDPOINT, headers={"Authorization": f"Bearer {_token()}"}, timeout=30)
    r.raise_for_status()
    return sin_cuentas_internas(r.json())


def corte(fecha_alta):
    """15 o 30: que dia del mes cae la liquidacion, segun el dia de alta.

    El endpoint manda datetime completo con zona ("2026-08-25T21:18:13.55+00:00"),
    no solo la fecha, asi que hace falta parsearlo entero en vez de partir por '-'
    (el dia no es el tercer campo separado por guiones)."""
    try:
        dia = datetime.fromisoformat(fecha_alta).day
    except (TypeError, ValueError):
        dia = 1
    return 15 if dia <= 15 else 30


def vendedor_por_cliente(clientes, telefono_a_vendedor, a_mano=None,
                         validos=None):
    """{id de cliente: vendedor}. Tres escalones, en este orden.

      1. `vendedor` que devuelve el endpoint. Es el unico dato de primera mano:
         lo carga la web en el momento de la venta, no lo deduce nadie. Si
         viene, gana.
      2. El cruce por telefono contra el master.
      3. La hoja de asignacion a mano.

    El telefono nunca va a cruzar el 100%: un negocio se registra con el
    celular del dueño y en Maps figura el fijo del local, o se registro antes
    de que la web pidiera telefono (los primeros clientes de Organizate estan
    en ese caso y ya no tiene arreglo).

    `validos` son los vendedores que existen en Config. Se pide porque el
    endpoint lo escribe gente: un nombre mal tipeado o el de alguien que ya no
    esta pagaria una comision a nadie, y en silencio. Un vendedor que no esta
    en la lista se ignora y el cliente baja al escalon siguiente, que termina
    en la hoja de asignacion a mano.

    La asignacion a mano se guarda contra el ID del cliente, no contra el
    nombre: el nombre lo puede editar el dueño del negocio cuando quiera y ahi
    se perderia la asignacion.
    """
    a_mano = a_mano or {}
    conocidos = {v.strip().lower(): v.strip() for v in (validos or []) if v.strip()}
    salida = {}
    for c in clientes:
        cid = c.get('id', '')
        del_endpoint = (c.get('vendedor') or '').strip()
        if del_endpoint and conocidos:
            del_endpoint = conocidos.get(del_endpoint.lower(), '')
        vendedor = (del_endpoint
                    or telefono_a_vendedor.get(canonico(c.get('telefono', '')))
                    or a_mano.get(cid, ''))
        if cid and vendedor:
            salida[cid] = vendedor
    return salida


def sin_vendedor(clientes, vendedor_de):
    """Los clientes que no se pudieron atribuir a nadie. Son los que hay que
    resolver a mano en la planilla."""
    return [c for c in clientes if not vendedor_de.get(c.get('id', ''))]


def liquidar(clientes, vendedor_de, referido_por):
    """[filas] para la hoja Liquidacion: una por cliente por pago cobrado.

    vendedor_de: {id de cliente: vendedor}, lo que devuelve
    vendedor_por_cliente. referido_por: {vendedor: quien lo trajo}, sale de la
    hoja Referidos (la completa el equipo a mano).

    Un cliente que no se pudo atribuir a nadie no genera fila: no hay a quien
    pagarle, y una comision al vendedor equivocado es peor que ninguna.
    """
    filas = []
    for cliente in clientes:
        vendedor = vendedor_de.get(cliente.get('id', ''), '')
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


def _negocios_de(vendedor, clientes, vendedor_de):
    """Negocios de ESTE vendedor (los que el cerro), con estado y lo pagado
    en total hasta ahora. Incluye demo y cancelados: el pedido es distinguir
    pipeline de plata real, no esconder lo que todavia no cobra."""
    salida = []
    for c in clientes:
        if vendedor_de.get(c.get('id', '')) != vendedor:
            continue
        total = sum(float(p.get('monto', 0)) for p in c.get('pagos', []))
        salida.append({'negocio': c.get('negocio', ''),
                       'estado': estado_label(c.get('estado', '')), 'total': total})
    return salida


def resumen_por_vendedor(clientes, vendedor_de, referido_por, vendedores):
    """{vendedor: {propios, comision_propia, afiliados, comision_afiliados, total}}

    afiliados es {nombre_afiliado: {negocios, comision}}: quien tiene a quien
    de afiliado sale de invertir referido_por (si Valentino fue 'referido
    por' Augusto, Valentino es afiliado de Augusto). Un vendedor puede tener
    varios afiliados; cada uno le suma su 20%, sin tope.
    """
    resultado = {}
    for v in vendedores:
        propios = _negocios_de(v, clientes, vendedor_de)
        comision_propia = round(sum(n['total'] for n in propios) * PCT_VENDEDOR, 2)

        afiliados = {}
        for a in vendedores:
            if referido_por.get(a) != v:
                continue
            negocios_a = _negocios_de(a, clientes, vendedor_de)
            afiliados[a] = {'negocios': negocios_a,
                            'comision': round(sum(n['total'] for n in negocios_a) * PCT_REFERIDO, 2)}
        comision_afiliados = round(sum(info['comision'] for info in afiliados.values()), 2)

        resultado[v] = {'propios': propios, 'comision_propia': comision_propia,
                        'afiliados': afiliados, 'comision_afiliados': comision_afiliados,
                        'total': round(comision_propia + comision_afiliados, 2)}
    return resultado


def filas_resumen(resumen):
    """(filas, marcas) para la hoja de comisiones.

    Arranca con una tabla de totales -una linea por vendedor- y recien
    despues el detalle negocio por negocio. Es al reves de como salia antes:
    la pregunta que se hace todo el mundo al abrir la hoja es "cuanto le toca
    a cada uno", y eso no puede estar al final de 200 filas de detalle.

    marcas es {rol: [indices de fila]}, 0-based, para que el formato sepa que
    pintar sin tener que adivinar leyendo el texto de cada fila.
    """
    filas, marcas = [], {}

    def poner(fila, rol):
        marcas.setdefault(rol, []).append(len(filas))
        filas.append(fila)

    poner(['💵 COMISIONES POR VENDEDOR', '', '', ''], 'titulo')
    poner(['Se calcula solo con los pagos reales de Organizate. No editar a mano.',
           '', '', ''], 'ayuda')
    poner(['', '', '', ''], 'vacia')

    poner(['Vendedor', 'Ventas propias (50%)', 'Afiliados (20%)', 'TOTAL A COBRAR'],
          'enc_resumen')
    for v, info in resumen.items():
        poner([v, info['comision_propia'], info['comision_afiliados'], info['total']],
              'resumen')
    poner(['TOTAL EQUIPO',
           round(sum(i['comision_propia'] for i in resumen.values()), 2),
           round(sum(i['comision_afiliados'] for i in resumen.values()), 2),
           round(sum(i['total'] for i in resumen.values()), 2)], 'total_equipo')

    poner(['', '', '', ''], 'vacia')
    poner(['', '', '', ''], 'vacia')
    poner(['DETALLE POR VENDEDOR', '', '', ''], 'titulo_detalle')

    for v, info in resumen.items():
        poner(['', '', '', ''], 'vacia')
        poner([v.upper(), '', '', info['total']], 'vendedor')

        poner(['Ventas propias (50%)', '', '', ''], 'subseccion')
        poner(['Negocio', 'Estado', 'Total pagado', 'Comisión'], 'encabezado')
        if info['propios']:
            for n in info['propios']:
                poner([n['negocio'], n['estado'], n['total'],
                       round(n['total'] * PCT_VENDEDOR, 2)], 'dato')
        else:
            poner(['Todavía no tiene negocios asignados', '', '', ''], 'sin_datos')
        poner(['Subtotal propio', '', '', info['comision_propia']], 'subtotal')

        if info['afiliados']:
            poner(['Afiliados (20% de lo que venden ellos)', '', '', ''], 'subseccion')
            for a, ainfo in info['afiliados'].items():
                poner([f'↳ {a}', '', '', ainfo['comision']], 'afiliado')
                poner(['Negocio', 'Estado', 'Total pagado', 'Comisión (20%)'], 'encabezado')
                if ainfo['negocios']:
                    for n in ainfo['negocios']:
                        poner([n['negocio'], n['estado'], n['total'],
                               round(n['total'] * PCT_REFERIDO, 2)], 'dato')
                else:
                    poner(['Todavía no vendió nada', '', '', ''], 'sin_datos')
            poner(['Subtotal afiliados', '', '', info['comision_afiliados']], 'subtotal')

        poner([f'TOTAL {v.upper()}', '', '', info['total']], 'total')

    return filas, marcas


COLUMNAS_CLIENTES = ['Negocio', 'Teléfono', 'Estado', 'Alta', 'Vendedor',
                     'Cuotas pagadas', 'Último pago', 'Total cobrado']


def filas_clientes(clientes, vendedor_de):
    """Una fila por cliente de Organizate, del que mas paga al que menos.

    Logica pura: se testea sin tocar Google ni el endpoint.

    Las cuentas internas quedan afuera. Son las de prueba de los vendedores,
    para grabar los videos: no son negocios que pagan y contarlas como clientes
    inflaria el numero justo donde se mira si el negocio funciona.
    """
    filas, marcas = [], []
    datos = []
    for c in sin_cuentas_internas(clientes):
        pagos = c.get('pagos') or []
        montos = [_monto(p) for p in pagos]
        fechas = sorted(f for f in (_fecha_pago(p) for p in pagos) if f)
        datos.append({
            'negocio': (c.get('negocio') or '').strip() or '(sin nombre)',
            'telefono': (c.get('telefono') or '').strip() or '—',
            'estado': estado_label(c.get('estado')),
            'alta': _dia(c.get('alta')),
            'vendedor': vendedor_de.get(c.get('id', ''), '—'),
            'cuotas': len(pagos),
            'ultimo': max(fechas).strftime('%d/%m/%Y') if fechas else '—',
            'total': sum(montos),
        })

    # Primero el que mas plata trajo; entre los que no pagaron nada, el activo
    # antes que el que se fue.
    orden = {'Activo': 0, 'Demo': 1, 'Cancelado': 2, 'Abandonado': 3}
    datos.sort(key=lambda d: (-d['total'], -d['cuotas'],
                              orden.get(d['estado'], 9), d['negocio']))

    for d in datos:
        filas.append([d['negocio'], d['telefono'], d['estado'], d['alta'],
                      d['vendedor'], d['cuotas'], d['ultimo'],
                      d['total'] or ''])
        marcas.append(d['estado'])
    return filas, marcas


def _monto(pago):
    try:
        return float(pago.get('monto') or pago.get('amount') or 0)
    except (TypeError, ValueError):
        return 0.0


def _fecha_pago(pago):
    for campo in ('fecha', 'fecha_pago', 'date', 'created_at'):
        crudo = pago.get(campo)
        if crudo:
            try:
                return datetime.fromisoformat(str(crudo).replace('Z', '+00:00')).date()
            except ValueError:
                pass
    return None


def _dia(iso):
    try:
        return datetime.fromisoformat(str(iso).replace('Z', '+00:00')).strftime('%d/%m/%Y')
    except (TypeError, ValueError):
        return '—'


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

    filas = liquidar(clientes, vendedor_por_cliente(clientes, tel_a_vend), referido_por)
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

    # Fecha completa con hora y zona, como la manda de verdad el endpoint.
    assert corte('2026-08-25T21:18:13.55173+00:00') == 30
    assert corte('2026-08-10T00:00:00+00:00') == 15
    print('OK  corte: parsea el datetime completo del endpoint, no solo la fecha')

    # Marto va a proposito sin una sola venta: el caso normal al principio.
    vendedores = ['Augusto', 'Valentino', 'Joaquin', 'Marto']
    referido_por = {'Valentino': 'Augusto', 'Joaquin': 'Augusto'}
    clientes_r = [
        {'id': 'r1', 'telefono': '3413539510', 'negocio': 'Barberia A', 'estado': 'activo',
         'pagos': [{'monto': 15000}]},                                  # Augusto, propio
        {'id': 'r2', 'telefono': '1126205229', 'negocio': 'Spa B', 'estado': 'activo',
         'pagos': [{'monto': 10000}]},                                  # Valentino, afiliado de Augusto
        {'id': 'r3', 'telefono': '3411112222', 'negocio': 'Nails C', 'estado': 'activo',
         'pagos': [{'monto': 8000}]},                                   # Joaquin, afiliado de Augusto
        {'id': 'r4', 'telefono': '3413539510', 'negocio': 'Barberia A demo',
         'estado': 'demo', 'pagos': []},
    ]
    tel_a_vend_r = {canonico('3413539510'): 'Augusto', canonico('1126205229'): 'Valentino',
                    canonico('3411112222'): 'Joaquin'}
    r = resumen_por_vendedor(clientes_r, vendedor_por_cliente(clientes_r, tel_a_vend_r),
                             referido_por, vendedores)

    assert r['Augusto']['comision_propia'] == 7500, r['Augusto']       # 50% de 15000
    assert set(r['Augusto']['afiliados']) == {'Valentino', 'Joaquin'}
    assert r['Augusto']['afiliados']['Valentino']['comision'] == 2000  # 20% de 10000
    assert r['Augusto']['afiliados']['Joaquin']['comision'] == 1600    # 20% de 8000
    assert r['Augusto']['comision_afiliados'] == 3600
    assert r['Augusto']['total'] == 11100, r['Augusto']                # 7500 + 3600
    # El demo de Augusto aparece en propios con total 0, no suma comision.
    demo_propio = next(n for n in r['Augusto']['propios'] if n['estado'] == 'Demo')
    assert demo_propio['total'] == 0

    assert r['Valentino']['comision_propia'] == 5000 and not r['Valentino']['afiliados']
    assert r['Joaquin']['comision_propia'] == 4000 and not r['Joaquin']['afiliados']

    print('OK  resumen_por_vendedor: 50% propio + 20% de cada afiliado, demo no cobra')

    filas, marcas = filas_resumen(r)
    assert all(len(f) == 4 for f in filas), 'todas las filas van de 4 columnas'
    # La tabla de totales va ARRIBA: la primera fila de resumen tiene que
    # aparecer antes que cualquier fila de detalle.
    assert max(marcas['resumen']) < min(marcas['vendedor']), \
        'el resumen quedo despues del detalle'
    assert len(marcas['resumen']) == len(vendedores), 'falta un vendedor en el resumen'
    # El total del equipo es la suma de los totales, no algo recalculado aparte.
    fila_equipo = filas[marcas['total_equipo'][0]]
    assert fila_equipo[3] == round(sum(i['total'] for i in r.values()), 2), fila_equipo
    # Marto no vendio nada: tiene que salir igual, con su fila y en cero.
    assert filas[marcas['resumen'][-1]][0] == 'Marto'
    assert any('no tiene negocios' in f[0] for f in filas), \
        'el vendedor sin ventas tiene que decirlo, no quedar en blanco'
    # Cada indice marcado apunta a una fila que existe.
    for rol, idxs in marcas.items():
        assert all(0 <= i < len(filas) for i in idxs), rol
    print('OK  filas_resumen: totales arriba, detalle abajo, sin vendedor perdido')

    # Estados del endpoint. 'abandonado' NO puede caer en Activo: cuando
    # Organizate sumo ese estado, 8 de 9 clientes eran abandonados y con el
    # default viejo se mostraban todos como Activo.
    assert estado_label('activo') == 'Activo'
    assert estado_label('abandonado') == 'Abandonado'
    assert estado_label('cancelado') == 'Cancelado'
    assert estado_label('baja') == 'Cancelado', 'el nombre viejo tiene que seguir mapeado'
    assert estado_label('demo') == 'Demo'
    assert estado_label('jubilado') == 'Jubilado', 'un estado nuevo se muestra, no se disfraza'
    assert estado_label(None) == 'Sin estado'
    assert set(ESTADO_LABEL.values()) <= set(COLOR_ESTADO), 'falta un color de estado'
    print('OK  estado_label: ningun estado desconocido se hace pasar por Activo')

    internas = [{'negocio': 'Real', 'es_cuenta_interna': False},
                {'negocio': 'Demo del vendedor', 'es_cuenta_interna': True},
                {'negocio': 'Viejo sin el campo'}]
    quedan = [c['negocio'] for c in sin_cuentas_internas(internas)]
    assert quedan == ['Real', 'Viejo sin el campo'], quedan
    print('OK  sin_cuentas_internas: la cuenta interna no llega al calculo')

    # El rango leido es Telefono, Categoria, Vendedor. Sheets recorta las
    # celdas vacias del final, asi que la fila SIN vendedor llega con 2
    # elementos y la ultima es la categoria: no puede colarse como vendedor.
    assert par_telefono_vendedor(['0341 353-9510', 'Barbería', 'Augusto']) == \
        {'3413539510': 'Augusto'}
    assert par_telefono_vendedor(['0341 353-9510', 'Barbería']) == {}, \
        'la categoria se colo como vendedor'
    assert par_telefono_vendedor(['0341 353-9510']) == {}
    assert par_telefono_vendedor(['0341 353-9510', 'Barbería', '   ']) == {}
    assert par_telefono_vendedor(['sin telefono', 'Barbería', 'Augusto']) == {}, \
        'sin telefono valido no hay clave para cruzar'
    assert par_telefono_vendedor([]) == {}
    print('OK  par_telefono_vendedor: la categoria nunca se hace pasar por vendedor')

    # Fecha con Z, como figura en el ejemplo del endpoint (los datos reales
    # llegan con +00:00; los dos formatos tienen que andar).
    assert corte('2026-01-15T12:00:00.000Z') == 15
    assert corte('2026-01-20T12:00:00.000Z') == 30

    # De punta a punta con la forma NUEVA del endpoint, para tener probado el
    # dia que empiecen a llegar telefono y pagos de verdad.
    reales = sin_cuentas_internas([
        {'id': 'b1a2', 'telefono': '0341 353-9510', 'negocio': 'Barber Valen',
         'alta': '2026-01-15T12:00:00.000Z', 'estado': 'activo', 'baja': None,
         'es_cuenta_interna': False,
         'pagos': [{'fecha': '2026-02-15', 'monto': 24990, 'periodo': '2026-02'}]},
        {'id': 'c3d4', 'telefono': '+54 9 11 2620-5229', 'negocio': 'Spa Norte',
         'alta': '2026-02-20T09:00:00.000Z', 'estado': 'cancelado',
         'baja': '2026-03-10', 'es_cuenta_interna': False,
         'pagos': [{'fecha': '2026-02-20', 'monto': 24990, 'periodo': '2026-02'}]},
        {'id': 'demo', 'telefono': '0341 353-9510', 'negocio': 'Demo de Augusto',
         'alta': '2026-01-01T00:00:00.000Z', 'estado': 'activo', 'baja': None,
         'es_cuenta_interna': True,
         'pagos': [{'fecha': '2026-02-01', 'monto': 99999, 'periodo': '2026-02'}]},
    ])
    tv = {canonico('3413539510'): 'Augusto', canonico('1126205229'): 'Valentino'}
    vd = vendedor_por_cliente(reales, tv)
    liq = liquidar(reales, vd, {'Valentino': 'Augusto'})
    assert len(liq) == 2, f'la cuenta interna se colo en la liquidacion: {liq}'
    assert not any(f[4] == 99999 for f in liq), 'cobro una comision por la cuenta demo'

    res = resumen_por_vendedor(reales, vd, {'Valentino': 'Augusto'}, ['Augusto', 'Valentino'])
    # Augusto: 50% de 24990 propio + 20% de los 24990 de Valentino.
    assert res['Augusto']['comision_propia'] == 12495, res['Augusto']
    assert res['Augusto']['afiliados']['Valentino']['comision'] == 4998
    assert res['Augusto']['total'] == 17493, res['Augusto']
    # El cancelado igual se muestra, con su plata ya cobrada y su etiqueta.
    assert [n['estado'] for n in res['Valentino']['propios']] == ['Cancelado']
    print('OK  end-to-end: forma nueva del endpoint, cuenta interna afuera, 50/20 correcto')

    # Asignacion a mano: el caso de los 9 primeros clientes de Organizate, que
    # se registraron antes de que la web pidiera telefono y por eso lo tienen
    # en null para siempre.
    viejos = [
        {'id': 'uuid-1', 'negocio': 'Pestañas ROSARIO', 'telefono': None,
         'alta': '2026-07-09T13:16:45+00:00', 'estado': 'activo',
         'pagos': [{'periodo': '2026-08', 'monto': 24990}]},
        {'id': 'uuid-2', 'negocio': 'Estetica Peralta', 'telefono': None,
         'alta': '2026-07-16T14:09:10+00:00', 'estado': 'abandonado', 'pagos': []},
    ]
    # Sin asignar nada: nadie cobra y los dos figuran como pendientes.
    vacio = vendedor_por_cliente(viejos, {})
    assert vacio == {}, vacio
    assert len(sin_vendedor(viejos, vacio)) == 2
    assert liquidar(viejos, vacio, {}) == [], 'sin vendedor no se puede liquidar'

    # Con la asignacion a mano el cliente activo ya cobra.
    vd2 = vendedor_por_cliente(viejos, {}, {'uuid-1': 'Marto'})
    assert vd2 == {'uuid-1': 'Marto'}, vd2
    assert [c['negocio'] for c in sin_vendedor(viejos, vd2)] == ['Estetica Peralta']
    liq2 = liquidar(viejos, vd2, {'Marto': 'Gige'})
    assert len(liq2) == 1 and liq2[0][5] == 'Marto'
    assert liq2[0][6] == 12495, liq2[0]                    # 50% de 24990
    assert liq2[0][7] == 'Gige' and liq2[0][8] == 4998     # 20% para quien lo trajo

    # El telefono, cuando existe, le gana a lo cargado a mano: el dato de la
    # planilla es el que mantiene el equipo, la asignacion manual es el parche.
    con_tel = [{'id': 'uuid-3', 'telefono': '0341 353-9510', 'negocio': 'X', 'pagos': []}]
    gana = vendedor_por_cliente(con_tel, {canonico('3413539510'): 'Augusto'},
                                {'uuid-3': 'Marto'})
    assert gana == {'uuid-3': 'Augusto'}, gana
    # --- los tres escalones -------------------------------------------------
    # 1. Lo que dice el endpoint gana sobre el cruce por telefono: lo carga la
    #    web cuando se hace la venta, no lo deduce nadie.
    cs = [{'id': 'c1', 'negocio': 'Barberia A', 'telefono': '0341 353-9510',
           'vendedor': 'Marto', 'estado': 'activo', 'pagos': []}]
    tel = {canonico('03413539510'): 'Gige'}
    assert vendedor_por_cliente(cs, tel, validos=['Marto', 'Gige']) == {'c1': 'Marto'}

    # 2. Sin vendedor en el endpoint, manda el telefono.
    cs[0]['vendedor'] = None
    assert vendedor_por_cliente(cs, tel, validos=['Marto', 'Gige']) == {'c1': 'Gige'}

    # 3. Sin telefono que cruce, la hoja de asignacion a mano.
    assert vendedor_por_cliente([{'id': 'c2', 'telefono': None}], tel,
                                a_mano={'c2': 'Tomas'}, validos=['Tomas']) == {'c2': 'Tomas'}

    # Un vendedor que no esta en Config se ignora y el cliente baja de escalon.
    # Si no, un nombre mal tipeado en la web le pagaria comision a nadie y en
    # silencio, que es la peor forma de perder plata.
    cs[0]['vendedor'] = 'Martoo'
    assert vendedor_por_cliente(cs, tel, validos=['Marto', 'Gige']) == {'c1': 'Gige'}

    # Sin lista de validos no se puede chequear nada, asi que se confia.
    assert vendedor_por_cliente(cs, {}) == {'c1': 'Martoo'}

    # Mayusculas y espacios no tienen que romper el cruce.
    cs[0]['vendedor'] = '  marto '
    assert vendedor_por_cliente(cs, tel, validos=['Marto']) == {'c1': 'Marto'}
    # --- la tabla de quien esta pagando -------------------------------------
    cartera = [
        {'id': 'p1', 'negocio': 'Pestañas ROSARIO', 'telefono': '0341 111-1111',
         'estado': 'activo', 'alta': '2026-07-01T10:00:00+00:00', 'es_cuenta_interna': False,
         'pagos': [{'monto': 25000, 'fecha': '2026-07-05T10:00:00+00:00'},
                   {'monto': 25000, 'fecha': '2026-08-05T10:00:00+00:00'}]},
        {'id': 'p2', 'negocio': 'Spa Sin Pagos', 'telefono': None, 'estado': 'activo',
         'alta': '2026-08-20T10:00:00+00:00', 'es_cuenta_interna': False, 'pagos': []},
        {'id': 'p3', 'negocio': 'Se Fue', 'telefono': '011 2222-2222',
         'estado': 'abandonado', 'alta': '2026-06-01T10:00:00+00:00',
         'es_cuenta_interna': False, 'pagos': []},
        {'id': 'p4', 'negocio': 'Prueba de Tomas', 'telefono': '', 'estado': 'activo',
         'alta': '2026-08-01T10:00:00+00:00', 'es_cuenta_interna': True, 'pagos': []},
    ]
    fs, ms = filas_clientes(cartera, {'p1': 'Marto'})

    assert len(fs) == 3, 'la cuenta interna no es un cliente que paga'
    assert [f[0] for f in fs] == ['Pestañas ROSARIO', 'Spa Sin Pagos', 'Se Fue'], fs
    assert fs[0][1] == '0341 111-1111' and fs[0][2] == 'Activo'
    assert fs[0][4] == 'Marto' and fs[0][5] == 2, fs[0]
    assert fs[0][6] == '05/08/2026', fs[0][6]
    assert fs[0][7] == 50000, fs[0][7]
    # El que nunca pago no muestra un cero disfrazado de cobro.
    assert fs[1][5] == 0 and fs[1][6] == '—' and fs[1][7] == ''
    # Sin telefono se ve que falta, no una celda vacia que parece un bug.
    assert fs[1][1] == '—'
    # Entre los que no pagaron nada, el activo antes que el que se fue.
    assert ms == ['Activo', 'Activo', 'Abandonado'], ms
    print('OK  filas_clientes: quien paga, cuantas cuotas y quien lo vendio')

    print('OK  vendedor_por_cliente: endpoint, telefono y a mano, en ese orden')

    print('OK  asignacion a mano: rescata al cliente sin telefono, y el telefono manda')


if __name__ == '__main__':
    demo()
