"""Reparte leads a los vendedores y trae de vuelta lo que trabajaron.

    python repartir_leads.py              # reparte y sincroniza
    python repartir_leads.py --simular    # dice que haria, sin escribir nada

Cada vendedor tiene su propio archivo con su lote. No ve la base ni el trabajo
de los demas. El archivo se lo crea el Apps Script (apps_script/Vendedores.gs);
aca se le carga el lote, se le lee lo que gestiono y se le arma el panel.

En cada corrida, por cada vendedor:

  1. Se lee su archivo y lo que cambio se copia al master (Estado, Motivo,
     Observaciones). Si algo cambio respecto del master, se le pone la fecha
     de hoy en Ultima gestion: el master hace de foto anterior, asi que no
     hace falta ni un trigger ni guardar un snapshot aparte.
  2. Se mira que nicho eligio en su panel. Si pidio uno distinto del que esta
     trabajando, se le cambia (ver `puede_cambiar_nicho` para cuando si y
     cuando no).
  3. Se decide si le toca lote nuevo (modules/asignacion.puede_reponer).
  4. Si le toca: salen de su archivo los leads cerrados, quedan los que
     siguen vivos, y se completa hasta 30 con leads nuevos del nicho elegido.
  5. Se le reescribe el panel con como viene y que le falta para el proximo.

Necesita GOOGLE_CREDENTIALS y SHEET_ID, igual que el resto del pipeline.
"""
import os
import sys
from datetime import date

sys.stdout.reconfigure(encoding='utf-8')

import pandas as pd

from collections import Counter

from modules.asignacion import (CALIFICADOS, DIAS_SIN_VENDER, FORMATO_FECHA, LOTE,
                                actividad,
                                contactado, elegir_lote, fecha, minimo_para,
                                puede_cambiar_nicho, puede_reponer, resumen_lote,
                                _clave)
from modules.estilo import (BLANCO, DINERO, FUENTE, TINTA, TINTA_SUAVE,
                            hex_a_rgb, rgb)
from modules.planilla import (CLAVE, COLUMNAS, CONFIG, FILA_VENDEDORES, IDX,
                              MOTIVOS, NOMBRES_ESTADO, SEGUIMIENTO, SIN_CONTACTAR,
                              col_letra, reintentar)
from modules.telefono import canonico
from sincronizar_sheets import GENERADO, _abrir_libro, _cliente_gspread, _sheet_id

# Las dos hojas del archivo personal. El panel lo crea Python (agregar una
# hoja a un archivo que ya existe no consume cuota de Drive, crearlo si);
# la de leads la crea el Apps Script con este nombre exacto.
HOJA_VENDEDOR = 'Mis clientes'
PANEL_VENDEDOR = '📊 Mi panel'
HOJA_CALIFICADOS = '⭐ Interesados'

# Columnas del archivo personal, en el orden que las crea el Apps Script.
# Si se tocan alla, hay que tocarlas aca: es el unico acoplamiento entre los dos.
COLUMNAS_VENDEDOR = ['Negocio', 'Teléfono', 'Estado', 'Motivo', 'Observaciones',
                     'Última gestión', 'Ciudad', 'Categoría', 'Link en Maps',
                     'Prioridad']
V = {c: i for i, c in enumerate(COLUMNAS_VENDEDOR)}

# Un lead cerrado ya no tiene vuelta y sale del archivo cuando entra una tanda
# nueva. "Cliente activo" salia por aca antes: ahora va a la carpeta de
# interesados, que es donde tiene que quedar guardado.
ESTADOS_CERRADOS = {'No interesado'}

# Columnas del master que se leen para saber que esta tomado y con que estado.
# Va de Telefono a Ultima gestion en un solo rango.
COL_DESDE, COL_HASTA = IDX[CLAVE], IDX['Última gestión']


def _hoy():
    return date.today().strftime('%d/%m/%Y')


def vendedores_con_archivo(libro):
    """[{nombre, email, archivo}] de Config, solo los que ya tienen archivo."""
    hoja = reintentar(libro.worksheet, CONFIG)
    filas = reintentar(hoja.get_all_values)[FILA_VENDEDORES - 1:]
    salida = []
    for f in filas:
        f = (list(f) + [''] * 10)[:10]
        nombre, email, archivo = f[0].strip(), f[6].strip(), f[7].strip()
        if nombre and email and archivo:
            salida.append({'nombre': nombre, 'email': email, 'archivo': archivo})
    return salida


def estado_del_master(libro, titulos):
    """{clave: {hoja, fila, vendedor, estado, motivo, observaciones}}.

    Es la foto de lo que el master cree hoy. Sirve para dos cosas: saber que
    lead esta tomado y por quien, y comparar contra el archivo del vendedor
    para detectar que cambio.
    """
    d, h = col_letra(COL_DESDE), col_letra(COL_HASTA)
    leidas = reintentar(libro.values_batch_get, [f"'{t}'!{d}2:{h}" for t in titulos])
    salida = {}
    for titulo, rango in zip(titulos, leidas.get('valueRanges', [])):
        for i, fila in enumerate(rango.get('values', [])):
            fila = (list(fila) + [''] * (COL_HASTA - COL_DESDE + 1))
            clave = canonico(fila[0])
            if not clave or clave in salida:
                continue
            salida[clave] = {
                'hoja': titulo, 'fila': i + 2,          # +2: la 1 es el encabezado
                'vendedor': fila[IDX['Vendedor'] - COL_DESDE].strip(),
                'estado': fila[IDX['Estado'] - COL_DESDE].strip(),
                'motivo': fila[IDX['Motivo'] - COL_DESDE].strip(),
                'observaciones': fila[IDX['Observaciones'] - COL_DESDE].strip(),
                'ultima': fila[IDX['Última gestión'] - COL_DESDE].strip(),
            }
    return salida


def _hoja_leads(libro):
    """La hoja de leads del vendedor.

    Por nombre y no por `.sheet1`: desde que existe el panel, la primera hoja
    del archivo es el panel, y leer esa seria leer cualquier cosa.
    """
    hojas = reintentar(libro.worksheets)
    for h in hojas:
        if h.title == HOJA_VENDEDOR:
            return h
    for h in hojas:                        # archivo armado a mano, sin el nombre
        if h.title != PANEL_VENDEDOR:
            return h
    raise RuntimeError('el archivo no tiene hoja de leads')


def leer_archivo(hoja):
    """Las filas del archivo del vendedor, con TODAS sus columnas.

    Se leen todas y no solo las tres que el vendedor edita, porque cuando le
    toca lote nuevo el archivo se rearma: si aca se perdiera Ciudad o Link en
    Maps, los seguimientos que se conservan quedarian sin esos datos.
    """
    filas = reintentar(hoja.get_all_values)[1:]
    salida = []
    for f in filas:
        f = (list(f) + [''] * len(COLUMNAS_VENDEDOR))[:len(COLUMNAS_VENDEDOR)]
        clave = canonico(f[V['Teléfono']])
        if not clave:
            continue
        salida.append({'clave': clave,
                       'negocio': f[V['Negocio']], 'telefono': f[V['Teléfono']],
                       'estado': f[V['Estado']].strip(),
                       'motivo': f[V['Motivo']].strip(),
                       'observaciones': f[V['Observaciones']].strip(),
                       'ultima': f[V['Última gestión']],
                       'ciudad': f[V['Ciudad']], 'categoria': f[V['Categoría']],
                       'link': f[V['Link en Maps']], 'prioridad': f[V['Prioridad']]})
    return salida


def rescatar_del_master(nombre, master, por_clave, ya_en_archivo):
    """Los leads que el master le tiene asignados y no estan en su archivo.

    Es el trabajo viejo, de cuando todo el equipo compartia una sola planilla.
    Gige tiene 37 asignados asi y Fran Majul 7, con seguimientos abiertos
    adentro. Si no se rescatan, al pasar a los archivos personales quedan
    asignados a su nombre en el master pero invisibles para ellos, o sea que se
    pierde la gestion hecha.

    Los datos del negocio salen del Excel (por_clave) y el trabajo del master:
    asi no hay que bajarse las 12 columnas de las 89407 filas.
    """
    salida = []
    for clave, m in master.items():
        if m['vendedor'] != nombre or clave in ya_en_archivo:
            continue
        lead = por_clave.get(clave, {})
        salida.append({'clave': clave,
                       'negocio': lead.get('negocio', ''),
                       'telefono': lead.get('telefono', ''),
                       'estado': m['estado'] or SIN_CONTACTAR,
                       'motivo': m['motivo'], 'observaciones': m['observaciones'],
                       'ultima': '',
                       'ciudad': lead.get('ciudad', ''),
                       'categoria': lead.get('categoria', ''),
                       'link': lead.get('link', ''),
                       'prioridad': lead.get('prioridad', '')})
    return salida


def cambios_al_master(filas_vendedor, master):
    """[(clave, campos)] de lo que el vendedor toco y el master todavia no sabe.

    Compara contra el master en vez de guardar un snapshot: el master ES el
    snapshot, porque la corrida anterior lo dejo al dia.
    """
    cambios = []
    for f in filas_vendedor:
        m = master.get(f['clave'])
        if not m:
            continue
        if (f['estado'], f['motivo'], f['observaciones']) != \
           (m['estado'], m['motivo'], m['observaciones']):
            cambios.append((f['clave'], f))
    return cambios


def _fila_para_vendedor(lead):
    """Un lead del Excel -> la fila que ve el vendedor en su archivo."""
    fila = [''] * len(COLUMNAS_VENDEDOR)
    fila[V['Negocio']] = lead.get('negocio', '')
    fila[V['Teléfono']] = lead.get('telefono', '')
    fila[V['Estado']] = SIN_CONTACTAR
    fila[V['Ciudad']] = lead.get('ciudad', '')
    fila[V['Categoría']] = lead.get('categoria', '')
    fila[V['Link en Maps']] = lead.get('link', '')
    fila[V['Prioridad']] = lead.get('prioridad', '')
    return fila


def _fila_conservada(f):
    """Una fila que ya estaba en el archivo y se queda (seguimiento abierto)."""
    fila = [''] * len(COLUMNAS_VENDEDOR)
    fila[V['Negocio']] = f.get('negocio', '')
    fila[V['Teléfono']] = f.get('telefono', '')
    fila[V['Estado']] = f.get('estado', '')
    fila[V['Motivo']] = f.get('motivo', '')
    fila[V['Observaciones']] = f.get('observaciones', '')
    fila[V['Última gestión']] = f.get('ultima', '')
    fila[V['Ciudad']] = f.get('ciudad', '')
    fila[V['Categoría']] = f.get('categoria', '')
    fila[V['Link en Maps']] = f.get('link', '')
    fila[V['Prioridad']] = f.get('prioridad', '')
    return fila


# --------------------------------------------------------------------------
# El panel del vendedor
#
# Una hoja de dos columnas al frente del archivo: como viene el lote, que le
# falta para el proximo, y el unico control que tiene, el desplegable de nicho.
# La fila 5 (Nicho) es lo unico que se deja editable.
# --------------------------------------------------------------------------

FILA_NICHO = 5                # B5: desplegable de rubro
FILA_CIUDAD = 6               # B6: desplegable de ciudad, opcional
TODA_ARGENTINA = 'Toda Argentina'   # en B6 significa "no me filtres por ciudad"
CIUDADES_OFRECIDAS = 20       # las que mas negocios libres tienen
ALTO_PANEL = 16               # filas que ocupa
TITULOS = (1, 4, 9, 15)       # filas que son encabezado de seccion
MENSAJES = (2, 7, 16)         # filas de texto largo, van con wrap


def _hoja_panel(libro, crear=True):
    """(hoja, recien_creada). Sin crear devuelve (None, False) si no existe."""
    for h in reintentar(libro.worksheets):
        if h.title == PANEL_VENDEDOR:
            return h, False
    if not crear:
        return None, False
    hoja = reintentar(libro.add_worksheet, title=PANEL_VENDEDOR,
                      rows=ALTO_PANEL + 4, cols=2, index=0)
    return hoja, True


def leer_filtro(hoja):
    """(rubro, ciudad) que el vendedor dejo elegidos. Ciudad vacia = toda Argentina."""
    b5, b6 = reintentar(hoja.batch_get, [f'B{FILA_NICHO}', f'B{FILA_CIUDAD}'])
    saca = lambda r: (r[0][0] if r and r[0] else '').strip()
    ciudad = saca(b6)
    return saca(b5), ('' if ciudad == TODA_ARGENTINA else ciudad)


def _frase(filtro):
    """(nicho, ciudad) -> 'Peluquerias en Rosario' / 'Peluquerias' / 'negocios en Rosario'."""
    nicho, ciudad = filtro
    if nicho and ciudad:
        return f'{nicho} en {ciudad}'
    return nicho or (f'negocios en {ciudad}' if ciudad else 'nada')


def mensaje_filtro(pedido, logrado, conseguidos, puede, motivo, repone):
    """Que decirle sobre el rubro y la ciudad, en castellano.

    `logrado` es lo que EFECTIVAMENTE quedo en su tanda, y `conseguidos` cuantos
    de los nuevos son de lo que pidio. Si pidio una cosa y quedo otra hay tres
    razones y se dicen distinto: se le nego el cambio por estar a mitad de tanda,
    lo que pidio se agoto del todo, o alcanzaba para unos pocos y se completo con
    otra cosa. La ultima es la que mas confunde si no se explica.
    """
    ln, lc = logrado
    ahora = ('Ahora estas llamando ' + (f'{ln} en {lc}.' if ln and lc else f'{ln or "nada"}.'))
    pn, pc = pedido
    if (not pn or pn == ln) and (not pc or pc == lc):
        return (ahora + ' Podes cambiar el rubro y la ciudad hasta la llamada 10, y otra '
                'vez cada vez que te entra una tanda nueva. Tarda hasta una hora en '
                'aplicarse, no es al instante.')
    if conseguidos:
        return (ahora + f' Solo quedaban {conseguidos} de {_frase(pedido)}, asi que te '
                'completamos la tanda con lo mas parecido que habia.')
    if puede or repone:
        return (ahora + f' No quedan negocios de {_frase(pedido)} sin repartir. Proba con '
                'otro rubro, o con toda Argentina.')
    return (ahora + f' Pediste {_frase(pedido)} pero {motivo}: cuando termines la tanda '
            'te lo cambiamos.')


def mensaje_proximo(res, minimo):
    """Que le falta para que le entre otro lote."""
    if not res['asignados']:
        return 'Todavia no tenes negocios. Te entran en la proxima actualizacion.'
    if res['sin_contactar']:
        return (f"Te faltan {res['sin_contactar']} por contactar. Cuando llames a los "
                f"{res['asignados']} y hayas hablado con {minimo}, te entra "
                f'una tanda nueva.')
    if res['hablados'] < minimo:
        return (f"Ya contactaste a los {res['asignados']}. Te faltan "
                f"{minimo - res['hablados']} conversaciones para que entre "
                f'la tanda nueva.')
    return 'Listo: en la proxima actualizacion te entra una tanda nueva.'


def filas_panel(nombre, pedido, res, guardados, aviso_filtro, aviso_proximo):
    """Las 16 filas del panel. Logica pura: se testea sin tocar Google.

    Las filas 5 y 6 son lo que el vendedor PIDIO, no lo que le toco: son sus dos
    controles y no se le pisan. Lo que efectivamente esta llamando se lo dice el
    mensaje de la fila 7, que si no la ciudad quedaria fijada sin que la haya
    elegido nunca.
    """
    return [
        [f'Panel de {nombre}', ''],
        ['Se actualiza solo cada hora. Vos elegis el rubro y la ciudad.' +
         (f' Tenes {guardados} guardados en la pestaña {HOJA_CALIFICADOS}.'
          if guardados else ''), ''],
        ['', ''],
        ['LO QUE QUIERO LLAMAR', ''],
        ['Rubro', pedido[0] or '(sin elegir)'],
        ['Ciudad', pedido[1] or TODA_ARGENTINA],
        [aviso_filtro, ''],
        ['', ''],
        ['MI TANDA', ''],
        ['Negocios asignados', res['asignados']],
        ['Ya contactados', res['contactados']],
        ['Con los que hablaste', res['hablados']],
        ['Sin contactar', res['sin_contactar']],
        ['', ''],
        ['PROXIMA TANDA', ''],
        [aviso_proximo, ''],
    ]


def _formato_panel(libro, hoja):
    """El maquillaje, una sola vez cuando se crea la hoja."""
    sid = hoja.id
    oscuro, claro = (rgb(hex_a_rgb(c)) for c in DINERO)

    def rango(f1, f2, c1=0, c2=2):
        return {'sheetId': sid, 'startRowIndex': f1 - 1, 'endRowIndex': f2,
                'startColumnIndex': c1, 'endColumnIndex': c2}

    def celdas(f1, f2, fmt, campos, c1=0, c2=2):
        return {'repeatCell': {'range': rango(f1, f2, c1, c2),
                               'cell': {'userEnteredFormat': fmt}, 'fields': campos}}

    reqs = [
        {'updateSheetProperties': {
            'properties': {'sheetId': sid, 'gridProperties': {'hideGridlines': True}},
            'fields': 'gridProperties.hideGridlines'}},
        {'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'COLUMNS',
                      'startIndex': 0, 'endIndex': 1},
            'properties': {'pixelSize': 215}, 'fields': 'pixelSize'}},
        {'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'COLUMNS',
                      'startIndex': 1, 'endIndex': 2},
            'properties': {'pixelSize': 300}, 'fields': 'pixelSize'}},
        # base: misma tipografia y aire vertical en todo el panel
        celdas(1, ALTO_PANEL,
               {'verticalAlignment': 'MIDDLE',
                'textFormat': {'fontFamily': FUENTE, 'fontSize': 10,
                               'foregroundColor': rgb(hex_a_rgb(TINTA))}},
               'userEnteredFormat(verticalAlignment,textFormat)'),
    ]

    # El nombre, arriba de todo.
    reqs += [
        {'mergeCells': {'range': rango(1, 1), 'mergeType': 'MERGE_ALL'}},
        {'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'ROWS',
                      'startIndex': 0, 'endIndex': 1},
            'properties': {'pixelSize': 52}, 'fields': 'pixelSize'}},
        celdas(1, 1, {'backgroundColor': oscuro, 'horizontalAlignment': 'CENTER',
                      'textFormat': {'bold': True, 'fontSize': 16,
                                     'foregroundColor': rgb(hex_a_rgb(BLANCO))}},
               'userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)'),
    ]

    # Encabezados de seccion.
    for f in TITULOS[1:]:
        reqs += [
            {'mergeCells': {'range': rango(f, f), 'mergeType': 'MERGE_ALL'}},
            celdas(f, f, {'backgroundColor': claro,
                          'textFormat': {'bold': True, 'fontSize': 9,
                                         'foregroundColor': oscuro}},
                   'userEnteredFormat(backgroundColor,textFormat)'),
        ]

    # Mensajes: una celda ancha, chica y que ajuste sola.
    for f in MENSAJES:
        reqs += [
            {'mergeCells': {'range': rango(f, f), 'mergeType': 'MERGE_ALL'}},
            celdas(f, f, {'wrapStrategy': 'WRAP',
                          'textFormat': {'italic': True, 'fontSize': 9,
                                         'foregroundColor': rgb(hex_a_rgb(TINTA_SUAVE))}},
                   'userEnteredFormat(wrapStrategy,textFormat)'),
        ]

    # Los numeros del lote, grandes y pegados a su etiqueta.
    reqs.append(celdas(10, 13, {'horizontalAlignment': 'RIGHT',
                                'textFormat': {'bold': True, 'fontSize': 13}},
                       'userEnteredFormat(horizontalAlignment,textFormat)', 1, 2))

    # Las dos celdas editables, resaltadas para que se note cuales son.
    reqs += [
        celdas(FILA_NICHO, FILA_CIUDAD,
               {'backgroundColor': rgb(hex_a_rgb('#FFF8E1')),
                'textFormat': {'bold': True, 'fontSize': 11}},
               'userEnteredFormat(backgroundColor,textFormat)', 1, 2),
        {'updateBorders': dict(
            {'range': rango(FILA_NICHO, FILA_CIUDAD, 1, 2)},
            **{lado: {'style': 'SOLID', 'color': oscuro}
               for lado in ('top', 'bottom', 'left', 'right')})},
    ]

    # Todo protegido menos las dos celdas que elige, en modo aviso: el archivo es
    # del vendedor, no hace falta trabarlo. Alcanza con que Sheets le avise antes
    # de pisar algo que igual se reescribe en la proxima corrida.
    for f1, f2 in ((1, FILA_NICHO - 1), (FILA_CIUDAD + 1, ALTO_PANEL)):
        reqs.append({'addProtectedRange': {'protectedRange': {
            'range': rango(f1, f2), 'warningOnly': True,
            'description': 'Lo escribe el sistema'}}})

    reqs.append({'updateSheetProperties': {
        'properties': {'sheetId': sid, 'tabColor': oscuro}, 'fields': 'tabColor'}})

    reintentar(libro.batch_update, {'requests': reqs})


def _desplegables(libro, hoja, nichos, ciudades):
    """Refresca las dos listas del panel.

    Va en cada corrida y no solo al crear la hoja: los rubros y las ciudades que
    todavia tienen negocios libres cambian a medida que el equipo los consume, y
    ofrecerle uno vacio seria ofrecerle quedarse sin tanda.
    """
    def regla(fila, valores, ayuda):
        return {'setDataValidation': {
            'range': {'sheetId': hoja.id, 'startRowIndex': fila - 1, 'endRowIndex': fila,
                      'startColumnIndex': 1, 'endColumnIndex': 2},
            'rule': {'condition': {'type': 'ONE_OF_LIST',
                                   'values': [{'userEnteredValue': v} for v in valores]},
                     'showCustomUi': True, 'strict': False, 'inputMessage': ayuda}}}

    reintentar(libro.batch_update, {'requests': [
        regla(FILA_NICHO, nichos, 'El rubro que queres trabajar'),
        regla(FILA_CIUDAD, [TODA_ARGENTINA] + ciudades,
              'Una sola ciudad si queres ir a los locales, o toda Argentina')]})


def _mas_comun(valores):
    valores = [v for v in valores if v]
    return Counter(valores).most_common(1)[0][0] if valores else ''


def nicho_y_ciudad(filas, por_clave):
    """Que esta trabajando hoy el vendedor, deducido de sus propios leads.

    El nicho no se guarda en ningun lado: sale de cruzar sus leads contra el
    Excel. Un estado menos que mantener sincronizado.
    """
    nicho = _mas_comun([por_clave.get(f['clave'], {}).get('nicho') for f in filas])
    return nicho, _mas_comun([f.get('ciudad') for f in filas])


def _hoja_calificados(libro, crear=True):
    """(hoja, recien_creada) de la carpeta de interesados.

    Sin crear devuelve (None, False) si no existe. Igual que el panel, la crea
    Python: agregar una hoja a un archivo que ya existe no consume cuota de
    Drive, crear el archivo si.
    """
    for h in reintentar(libro.worksheets):
        if h.title == HOJA_CALIFICADOS:
            return h, False
    if not crear:
        return None, False
    hoja = reintentar(libro.add_worksheet, title=HOJA_CALIFICADOS, rows=300,
                      cols=len(COLUMNAS_VENDEDOR), index=2)
    return hoja, True


def _formato_calificados(libro, hoja):
    """Encabezado, anchos y los mismos desplegables que la hoja de la tanda."""
    sid = hoja.id
    oscuro = rgb(hex_a_rgb(DINERO[0]))
    reqs = [
        {'updateSheetProperties': {
            'properties': {'sheetId': sid, 'tabColor': oscuro,
                           'gridProperties': {'frozenRowCount': 1}},
            'fields': 'tabColor,gridProperties.frozenRowCount'}},
        {'repeatCell': {
            'range': {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': 1},
            'cell': {'userEnteredFormat': {
                'backgroundColor': oscuro,
                'textFormat': {'bold': True, 'fontFamily': FUENTE,
                               'foregroundColor': rgb(hex_a_rgb(BLANCO))}}},
            'fields': 'userEnteredFormat(backgroundColor,textFormat)'}},
    ]
    for col, ancho in ((V['Negocio'], 260), (V['Teléfono'], 130), (V['Estado'], 140),
                       (V['Motivo'], 170), (V['Observaciones'], 320)):
        reqs.append({'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'COLUMNS',
                      'startIndex': col, 'endIndex': col + 1},
            'properties': {'pixelSize': ancho}, 'fields': 'pixelSize'}})
    # El vendedor sigue trabajando estos negocios, asi que necesita las mismas
    # listas para moverlos de "Le interesó" a "Demo iniciada" y a "Cliente activo".
    for col, lista in ((V['Estado'], NOMBRES_ESTADO), (V['Motivo'], MOTIVOS)):
        reqs.append({'setDataValidation': {
            'range': {'sheetId': sid, 'startRowIndex': 1,
                      'startColumnIndex': col, 'endColumnIndex': col + 1},
            'rule': {'condition': {'type': 'ONE_OF_LIST',
                                   'values': [{'userEnteredValue': x} for x in lista]},
                     'showCustomUi': True, 'strict': False}}})
    reintentar(libro.batch_update, {'requests': reqs})


def repartir_por_hoja(filas, carpeta):
    """Adonde va cada negocio cuando entra una tanda nueva.

    Devuelve (siguen, nueva_carpeta). Logica pura, se testea sin tocar Google.

      - Le interesó / Demo iniciada / Cliente activo -> a la carpeta. Son los
        que valen y no se pueden perder entre 30 llamados en frio nuevos.
      - No interesado -> afuera, de las dos hojas.
      - el resto (Volver a llamar, y el que pidio que lo llamen despues) se
        queda en la tanda: son seguimientos con fecha.

    La carpeta se acumula entre tandas, asi que lo que ya estaba adentro entra
    de nuevo salvo que el vendedor lo haya bajado a No interesado.
    """
    nueva, vistos = [], set()
    for f in carpeta + [f for f in filas if f['estado'] in CALIFICADOS]:
        if f['estado'] in ESTADOS_CERRADOS or f['clave'] in vistos:
            continue
        vistos.add(f['clave'])
        nueva.append(f)
    siguen = [f for f in filas
              if f['estado'] not in CALIFICADOS and f['estado'] not in ESTADOS_CERRADOS]
    return siguen, nueva


# --------------------------------------------------------------------------
# Seguimiento del equipo
#
# Una hoja en el master, para el dueño, con una fila por vendedor: cuanto
# trabajo y hace cuanto que no vende. Nadie se da de baja solo: el que lleva un
# mes sin vender aparece arriba con sus numeros al lado, y decide una persona.
# --------------------------------------------------------------------------

COL_ALTA = 10                 # Config!K, "en el equipo desde"
COLUMNAS_SEGUIMIENTO = ['Vendedor', 'En el equipo desde', 'Asignados', 'Llamados',
                        'Conversaciones', 'En el embudo', 'Ventas', 'Última venta',
                        'Días sin vender', 'Situación']


def altas(libro, vendedores, master, hoy=None, escribir=True):
    """{vendedor: fecha en que entro al equipo}, guardada en Config!K.

    La primera vez que se ve a alguien se le escribe la fecha, para que en la
    corrida siguiente ya se sepa si es nuevo o lleva meses. Al que ya tiene
    gestiones hechas se le pone la mas vieja: entro al menos ese dia, y arrancar
    el reloj hoy lo haria pasar por recien llegado un mes entero.
    """
    hoy = hoy or date.today()
    hoja = reintentar(libro.worksheet, CONFIG)
    if hoja.col_count <= COL_ALTA:
        reintentar(libro.batch_update, {'requests': [{'updateSheetProperties': {
            'properties': {'sheetId': hoja.id,
                           'gridProperties': {'columnCount': COL_ALTA + 1}},
            'fields': 'gridProperties.columnCount'}}]})
    col = col_letra(COL_ALTA)
    guardadas = reintentar(hoja.col_values, COL_ALTA + 1)

    primera = {}
    for m in master.values():
        f = fecha(m.get('ultima'))
        if m['vendedor'] and f:
            primera[m['vendedor']] = min(primera.get(m['vendedor'], f), f)

    salida, faltantes = {}, []
    for v in vendedores:
        i = _fila_en_config(libro, v['nombre'])
        texto = guardadas[i - 1] if i - 1 < len(guardadas) else ''
        d = fecha(texto)
        if not d:
            d = primera.get(v['nombre'], hoy)
            faltantes.append({'range': f"'{CONFIG}'!{col}{i}",
                              'values': [[d.strftime(FORMATO_FECHA)]]})
        salida[v['nombre']] = d

    if faltantes and escribir:
        faltantes.append({'range': f"'{CONFIG}'!{col}1",
                          'values': [['EN EL EQUIPO DESDE']]})
        reintentar(libro.values_batch_update,
                   {'valueInputOption': 'RAW', 'data': faltantes})
    return salida


def _fila_en_config(libro, nombre):
    """La fila de Config donde esta ese vendedor. Se cachea: son 17 llamadas
    iguales si no."""
    if not hasattr(_fila_en_config, 'cache'):
        col = reintentar(libro.worksheet(CONFIG).col_values, 1)
        _fila_en_config.cache = {n.strip(): i for i, n in enumerate(col, 1) if n.strip()}
    return _fila_en_config.cache[nombre]


def filas_seguimiento(master, vendedores, desde, hoy=None):
    """Las filas de la hoja, con los que hay que mirar arriba de todo.

    Logica pura sobre el master: se testea sin tocar Google.
    """
    hoy = hoy or date.today()
    por_vendedor = {}
    for m in master.values():
        if m['vendedor']:
            por_vendedor.setdefault(m['vendedor'], []).append(m)

    datos = []
    for v in vendedores:
        n = v['nombre']
        a = actividad(por_vendedor.get(n, []), desde.get(n), hoy)
        datos.append((n, a))

    # Primero los que hay que mirar, y dentro de esos el que hace mas que no
    # vende. Los que venden quedan abajo ordenados por ventas.
    datos.sort(key=lambda x: (not x[1]['alerta'],
                              -(x[1]['dias_sin_vender'] or 9999) if x[1]['alerta']
                              else -x[1]['ventas']))

    marcados = [n for n, a in datos if a['alerta']]
    if marcados:
        aviso = f'{len(marcados)} para mirar: ' + ', '.join(marcados)
    else:
        # Sin esto decia "todos vendieron en el ultimo mes" con CERO ventas
        # registradas en toda la planilla, que es lo contrario de la verdad.
        vendieron = sum(1 for _, a in datos if a['ventas'])
        nuevos = sum(1 for _, a in datos
                     if a['antiguedad'] is not None and a['antiguedad'] < DIAS_SIN_VENDER)
        aviso = (f'Ninguno para mirar: {vendieron} de {len(datos)} tienen ventas '
                 f'registradas y {nuevos} estan dentro de su primer mes.')

    filas = [['SEGUIMIENTO DEL EQUIPO'] + [''] * 9,
             [aviso] + [''] * 9,
             ['Una venta se cuenta cuando el vendedor marca "Cliente activo" en su '
              'archivo: si no la marca, no aparece aca. La fecha de alta la pone '
              'el sistema con la gestion mas vieja que encuentra, y se corrige a mano '
              'en la columna K de Config.'] + [''] * 9,
             COLUMNAS_SEGUIMIENTO]
    for n, a in datos:
        filas.append([
            n,
            desde[n].strftime(FORMATO_FECHA) if desde.get(n) else '—',
            a['asignados'], a['llamados'], a['conversaciones'], a['en_el_embudo'],
            a['ventas'],
            a['ultima_venta'].strftime(FORMATO_FECHA) if a['ultima_venta'] else '—',
            a['dias_sin_vender'] if a['dias_sin_vender'] is not None else '—',
            a['situacion'],
        ])
    return filas, [a['alerta'] for _, a in datos]


def _hoja_seguimiento(libro, filas, alertas):
    """Rehace la hoja y la deja de solo lectura: son numeros calculados."""
    from liquidar_comisiones import _proteger, _rehacer
    hoja = _rehacer(libro, SEGUIMIENTO, filas, len(COLUMNAS_SEGUIMIENTO))
    sid = hoja.id
    oscuro, claro = (rgb(hex_a_rgb(c)) for c in DINERO)
    n = len(COLUMNAS_SEGUIMIENTO)
    encabezado = len(filas) - len(alertas)      # la fila de titulos de columna

    def rango(f1, f2, c1=0, c2=n):
        return {'sheetId': sid, 'startRowIndex': f1 - 1, 'endRowIndex': f2,
                'startColumnIndex': c1, 'endColumnIndex': c2}

    reqs = [
        {'updateSheetProperties': {
            'properties': {'sheetId': sid, 'tabColor': oscuro,
                           'gridProperties': {'frozenRowCount': encabezado}},
            'fields': 'tabColor,gridProperties.frozenRowCount'}},
        {'mergeCells': {'range': rango(1, 1), 'mergeType': 'MERGE_ALL'}},
        {'mergeCells': {'range': rango(2, 2), 'mergeType': 'MERGE_ALL'}},
        {'mergeCells': {'range': rango(3, 3), 'mergeType': 'MERGE_ALL'}},
        {'repeatCell': {
            'range': rango(1, 1),
            'cell': {'userEnteredFormat': {
                'backgroundColor': oscuro,
                'textFormat': {'bold': True, 'fontSize': 14, 'fontFamily': FUENTE,
                               'foregroundColor': rgb(hex_a_rgb(BLANCO))}}},
            'fields': 'userEnteredFormat(backgroundColor,textFormat)'}},
        {'repeatCell': {
            'range': rango(2, 2),
            'cell': {'userEnteredFormat': {
                'backgroundColor': claro, 'wrapStrategy': 'WRAP',
                'textFormat': {'bold': True, 'fontSize': 10, 'foregroundColor': oscuro}}},
            'fields': 'userEnteredFormat(backgroundColor,wrapStrategy,textFormat)'}},
        {'repeatCell': {
            'range': rango(3, 3),
            'cell': {'userEnteredFormat': {
                'wrapStrategy': 'WRAP',
                'textFormat': {'italic': True, 'fontSize': 9,
                               'foregroundColor': rgb(hex_a_rgb(TINTA_SUAVE))}}},
            'fields': 'userEnteredFormat(wrapStrategy,textFormat)'}},
        {'repeatCell': {
            'range': rango(encabezado, encabezado),
            'cell': {'userEnteredFormat': {
                'backgroundColor': oscuro, 'verticalAlignment': 'MIDDLE',
                'wrapStrategy': 'WRAP',
                'textFormat': {'bold': True, 'fontSize': 9, 'fontFamily': FUENTE,
                               'foregroundColor': rgb(hex_a_rgb(BLANCO))}}},
            'fields': 'userEnteredFormat(backgroundColor,verticalAlignment,'
                      'wrapStrategy,textFormat)'}},
        {'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': 1},
            'properties': {'pixelSize': 170}, 'fields': 'pixelSize'}},
        {'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': 1, 'endIndex': n - 1},
            'properties': {'pixelSize': 105}, 'fields': 'pixelSize'}},
        {'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'COLUMNS',
                      'startIndex': n - 1, 'endIndex': n},
            'properties': {'pixelSize': 260}, 'fields': 'pixelSize'}},
    ]

    # Los numeros centrados, y el nombre en negrita.
    if alertas:
        reqs += [
            {'repeatCell': {
                'range': rango(encabezado + 1, len(filas), 1, n - 1),
                'cell': {'userEnteredFormat': {'horizontalAlignment': 'CENTER'}},
                'fields': 'userEnteredFormat.horizontalAlignment'}},
            {'repeatCell': {
                'range': rango(encabezado + 1, len(filas), 0, 1),
                'cell': {'userEnteredFormat': {'textFormat': {'bold': True}}},
                'fields': 'userEnteredFormat.textFormat.bold'}},
        ]
    # Fondo ambar en la fila del que hay que mirar. Se pinta por fila y no con
    # formato condicional porque la condicion ya la calculo Python.
    for i, marcado in enumerate(alertas):
        if marcado:
            reqs.append({'repeatCell': {
                'range': rango(encabezado + 1 + i, encabezado + 1 + i),
                'cell': {'userEnteredFormat': {
                    'backgroundColor': rgb(hex_a_rgb('#FBF3E0'))}},
                'fields': 'userEnteredFormat.backgroundColor'}})

    reintentar(libro.batch_update, {'requests': reqs})
    return _proteger(libro, sid, 'Seguimiento del equipo: lo calcula el sistema')


def leads_del_excel(generado=GENERADO):
    """Todos los negocios del Excel, en el formato que espera elegir_lote."""
    df = pd.read_excel(generado, sheet_name='Todos', dtype=str).fillna('')
    df = df[df[CLAVE].str.strip() != '']
    return [{'negocio': r['Negocio'], 'telefono': r[CLAVE], 'nicho': r['Nicho'],
             'ciudad': r['Ciudad'], 'prioridad': r.get('Prioridad', ''),
             'categoria': r.get('Categoría', ''), 'link': r.get('Link en Maps', '')}
            for r in df.to_dict('records')]


def main(simular=False):
    if not _sheet_id() or not (
            os.environ.get('GOOGLE_CREDENTIALS') or os.path.exists('credenciales.json')):
        print('Google Sheets sin configurar, salteando.')
        return
    if not os.path.exists(GENERADO):
        sys.exit(f'No existe {GENERADO}. Corre el scraper primero.')

    cliente = _cliente_gspread()
    libro = reintentar(cliente.open_by_key, _sheet_id())

    vendedores = vendedores_con_archivo(libro)
    if not vendedores:
        print('Ningun vendedor tiene archivo todavia. Corre "Ventas > Reparar '
              'accesos de todos" en la planilla.')
        return
    print(f'{len(vendedores)} vendedores con archivo')

    meta = reintentar(libro.fetch_sheet_metadata)
    from liquidar_comisiones import FUERA_DE_NICHOS
    titulos = [s['properties']['title'] for s in meta.get('sheets', [])
               if s['properties']['title'] not in FUERA_DE_NICHOS
               and s['properties']['title'] != 'Sheet2']
    master = estado_del_master(libro, titulos)
    print(f'{len(master)} leads en el master, '
          f'{sum(1 for m in master.values() if m["vendedor"])} ya asignados')

    todos = leads_del_excel()
    por_clave = {_clave(l): l for l in todos if _clave(l)}
    disponibles = [l for l in todos
                   if not master.get(_clave(l), {}).get('vendedor')]
    tomados = {k for k, m in master.items() if m['vendedor']}

    # Los desplegables ofrecen solo lo que todavia tiene negocios sin repartir:
    # elegir algo vacio seria elegir quedarse sin tanda. Las ciudades ademas van
    # por cantidad y cortadas: son 36 y la cola larga son partidos del conurbano
    # con dos negocios, que no le sirven a nadie.
    nichos = sorted({l['nicho'] for l in disponibles if l.get('nicho')})
    ciudades = [c for c, _ in Counter(
        l['ciudad'].strip() for l in disponibles if l.get('ciudad', '').strip()
    ).most_common(CIUDADES_OFRECIDAS)]

    escrituras_master, nuevas_asignaciones, paneles = [], [], []

    for v in vendedores:
        try:
            libro_v = reintentar(cliente.open_by_key, v['archivo'])
            filas = leer_archivo(_hoja_leads(libro_v))
            hoja_panel, panel_nuevo = _hoja_panel(libro_v, crear=not simular)
            hoja_carpeta, carpeta_nueva = _hoja_calificados(libro_v, crear=not simular)
            carpeta = leer_archivo(hoja_carpeta) if hoja_carpeta else []
        except Exception as e:
            print(f"   {v['nombre']}: no pude abrir su archivo ({str(e)[:60]}), lo salteo")
            continue

        if panel_nuevo:
            try:
                _formato_panel(libro_v, hoja_panel)
            except Exception:
                # Sin formato la hoja queda inservible, y ninguna corrida la
                # volveria a formatear porque ya existe. Se borra y se reintenta
                # de cero la proxima vez.
                reintentar(libro_v.del_worksheet, hoja_panel)
                raise
            print(f"   {v['nombre']}: panel creado")
        if hoja_panel:
            _desplegables(libro_v, hoja_panel, nichos, ciudades)
        if carpeta_nueva:
            _formato_calificados(libro_v, hoja_carpeta)

        def tomar(nicho, ciudad, faltan):
            """Saca `faltan` leads del pozo y se los anota en el master.

            Si lo pedido se quedo sin negocios se va aflojando: primero se suelta
            el rubro y no la ciudad, porque el que eligio ciudad lo hizo para ir
            a los locales y no le sirve el mismo rubro a 500 km.
            """
            nonlocal disponibles
            if faltan <= 0:
                return []
            intentos = [(nicho, ciudad)]
            if ciudad:
                intentos.append(('', ciudad))
            if nicho:
                intentos.append((nicho, ''))
            intentos.append(('', ''))
            elegidos = []
            for n, c in intentos:
                if len(elegidos) >= faltan:
                    break
                pozo = [l for l in disponibles
                        if (not n or l.get('nicho') == n)
                        and (not c or l.get('ciudad', '').strip() == c)]
                ya = tomados | {_clave(x) for x in elegidos}
                elegidos += elegir_lote(pozo, faltan - len(elegidos), ya)
            for lead in elegidos:
                clave = _clave(lead)
                tomados.add(clave)
                m = master.get(clave)
                if m:
                    escrituras_master.append({
                        'range': f"'{m['hoja']}'!{col_letra(IDX['Vendedor'])}{m['fila']}",
                        'values': [[v['nombre']]]})
                    m['vendedor'] = v['nombre']
            disponibles = [l for l in disponibles if _clave(l) not in tomados]
            return elegidos

        # 0. Lo que el master le tiene asignado de antes y no esta en su
        # archivo se suma ahora: es gestion ya hecha y no se puede perder.
        rescatados = rescatar_del_master(
            v['nombre'], master, por_clave,
            {f['clave'] for f in filas} | {f['clave'] for f in carpeta})
        if rescatados:
            print(f"   {v['nombre']}: {len(rescatados)} leads que ya tenia asignados "
                  f"en el master y no estaban en su archivo, se los llevo")
            filas = filas + rescatados

        # 1. Lo que trabajo vuelve al master.
        # Las dos hojas: el vendedor sigue moviendo los de la carpeta de
        # "Le interesó" a "Demo iniciada" y a "Cliente activo", y si eso no
        # llegara al master no habria venta que liquidar.
        cambios = cambios_al_master(filas + carpeta, master)
        for clave, f in cambios:
            m = master[clave]
            escrituras_master.append({
                'range': f"'{m['hoja']}'!{col_letra(IDX['Estado'])}{m['fila']}:"
                         f"{col_letra(IDX['Última gestión'])}{m['fila']}",
                'values': [[f['estado'], f['motivo'], f['observaciones'], _hoy()]]})
            m.update(estado=f['estado'], motivo=f['motivo'],
                     observaciones=f['observaciones'])

        # 2. Que esta trabajando y que pidio.
        actual = nicho_y_ciudad(filas, por_clave)
        pedido = leer_filtro(hoja_panel) if hoja_panel and not panel_nuevo else ('', '')
        cambiar = ((pedido[0] and pedido[0] != actual[0]) or
                   (pedido[1] and pedido[1] != actual[1]))
        libre, motivo_cambio = puede_cambiar_nicho(filas)

        # 3. Le toca lote nuevo?
        r = resumen_lote(filas)
        repone, por_que = puede_reponer(filas)
        print(f"   {v['nombre']}: {r['asignados']} leads, {r['contactados']} contactados, "
              f"{r['hablados']} hablados, {len(cambios)} cambios -> "
              f"{'REPONE' if repone else 'sigue'} ({por_que})")

        siguen, lote, sueltos = filas, [], []
        carpeta_final, movidos = carpeta, 0
        # El rubro pedido manda; si nunca eligio, sigue con el que tiene. La
        # ciudad solo filtra si la eligio a proposito: vacia es toda Argentina.
        nicho, ciudad = pedido[0] or actual[0], pedido[1]

        if repone:
            # Lote nuevo: los cerrados salen, los seguimientos se quedan. El
            # cambio de nicho aca siempre vale: el lote arranca de cero igual.
            siguen, carpeta_final = repartir_por_hoja(filas, carpeta)
            movidos = len(carpeta_final) - len(carpeta)
            lote = tomar(nicho, ciudad, LOTE - len(siguen))
            if cambiar:
                print(f'      cambia: {_frase(actual)} -> {_frase(pedido)}')
            if not lote:
                print(f'      {len(siguen)} seguimientos abiertos, no entra ninguno nuevo'
                      if len(siguen) >= LOTE else '      no quedan leads para repartir')
            else:
                print(f"      lote nuevo: {len(lote)} leads ({lote[0].get('nicho')}, "
                      f"{lote[0].get('ciudad')}), mas {len(siguen)} seguimientos")
        elif cambiar and libre:
            # Cambio a mitad de lote: lo que ya trabajo se queda, lo que ni
            # miro vuelve a la bolsa para que lo agarre otro.
            trabajados = [f for f in filas if contactado(f['estado'])]
            siguen, carpeta_final = repartir_por_hoja(trabajados, carpeta)
            movidos = len(carpeta_final) - len(carpeta)
            sueltos = [f for f in filas if not contactado(f['estado'])]
            for f in sueltos:
                m = master.get(f['clave'])
                if m and m['vendedor'] == v['nombre']:
                    escrituras_master.append({
                        'range': f"'{m['hoja']}'!{col_letra(IDX['Vendedor'])}{m['fila']}",
                        'values': [['']]})
                    m['vendedor'] = ''
            lote = tomar(nicho, ciudad, LOTE - len(siguen))
            print(f'      cambia: {_frase(actual)} -> {_frase(pedido)} '
                  f'({len(sueltos)} devueltos, {len(lote)} nuevos)')
        elif cambiar:
            print(f'      pidio {_frase(pedido)} pero {motivo_cambio}')

        # Se rearma el archivo si cambio algo de lo que el vendedor ve. Los
        # rescatados cuentan: si no, quedan asignados a su nombre en el master
        # y el nunca los ve.
        if movidos:
            print(f'      {movidos} pasan a {HOJA_CALIFICADOS} '
                  f'({len(carpeta_final)} guardados en total)')

        if lote or rescatados or sueltos or movidos:
            nuevas_asignaciones.append((v, siguen, lote, carpeta_final))

        # 4. El panel, siempre: aunque no cambie nada, los numeros se mueven.
        if hoja_panel:
            final = siguen + [{'estado': SIN_CONTACTAR, 'motivo': ''} for _ in lote]
            res = resumen_lote(final)
            logrado = ((_mas_comun([l.get('nicho') for l in lote]),
                        _mas_comun([l.get('ciudad', '').strip() for l in lote]))
                       if lote else actual)
            conseguidos = sum(1 for l in lote
                              if (not pedido[0] or l.get('nicho') == pedido[0])
                              and (not pedido[1] or l.get('ciudad', '').strip() == pedido[1]))
            paneles.append((v, hoja_panel, filas_panel(
                v['nombre'], (pedido[0] or logrado[0], pedido[1]), res,
                len(carpeta_final),
                mensaje_filtro(pedido, logrado, conseguidos if conseguidos < len(lote) else 0,
                               libre, motivo_cambio, repone),
                mensaje_proximo(res, minimo_para(final)))))

    if simular:
        print(f'\n[SIMULACION] {len(escrituras_master)} escrituras al master, '
              f'{len(nuevas_asignaciones)} archivos a rearmar y {len(paneles)} '
              f'paneles a refrescar. No se escribio nada.')

    if escrituras_master and not simular:
        reintentar(libro.values_batch_update,
                   {'valueInputOption': 'RAW', 'data': escrituras_master})
        print(f'\nMaster actualizado: {len(escrituras_master)} celdas')

    for v, siguen, lote, carpeta_final in ([] if simular else nuevas_asignaciones):
        libro_v = reintentar(cliente.open_by_key, v['archivo'])
        hoja = _hoja_leads(libro_v)
        filas = ([COLUMNAS_VENDEDOR] +
                 [_fila_conservada(f) for f in siguen] +
                 [_fila_para_vendedor(l) for l in lote])
        reintentar(hoja.clear)
        reintentar(hoja.update, values=filas, range_name='A1')

        carpeta_hoja, _ = _hoja_calificados(libro_v)
        reintentar(carpeta_hoja.clear)
        reintentar(carpeta_hoja.update,
                   values=[COLUMNAS_VENDEDOR] + [_fila_conservada(f) for f in carpeta_final],
                   range_name='A1')
        print(f"   {v['nombre']}: {len(filas) - 1} en la tanda, "
              f"{len(carpeta_final)} guardados")

    for v, hoja, valores in ([] if simular else paneles):
        reintentar(hoja.update, values=valores, range_name=f'A1:B{ALTO_PANEL}')
    if paneles and not simular:
        print(f'   {len(paneles)} paneles actualizados')

    # El reporte para el dueño, al final: necesita el master ya actualizado.
    desde = altas(libro, vendedores, master, escribir=not simular)
    filas_seg, alertas = filas_seguimiento(master, vendedores, desde)
    if simular:
        print(f'\n[SIMULACION] {SEGUIMIENTO}: {filas_seg[1][0]}')
    else:
        duro = _hoja_seguimiento(libro, filas_seg, alertas)
        print(f'\n{SEGUIMIENTO}: {filas_seg[1][0]}'
              + ('' if duro else '  (proteccion en modo aviso: falta SHEET_OWNER)'))

    print('\nListo.')


def demo():
    """Los pedazos que no hablan con Google."""
    master = {'aaa': {'hoja': 'Barberías', 'fila': 5, 'vendedor': 'Marto',
                      'estado': 'Sin contactar', 'motivo': '', 'observaciones': ''},
              'bbb': {'hoja': 'Spas', 'fila': 9, 'vendedor': 'Marto',
                      'estado': 'No interesado', 'motivo': 'Precio alto', 'observaciones': ''}}
    filas = [{'clave': 'aaa', 'negocio': 'A', 'estado': 'Le interesó', 'motivo': '',
              'observaciones': 'llamar el lunes'},
             {'clave': 'bbb', 'negocio': 'B', 'estado': 'No interesado',
              'motivo': 'Precio alto', 'observaciones': ''},
             {'clave': 'zzz', 'negocio': 'Z', 'estado': 'Le interesó', 'motivo': '',
              'observaciones': ''}]
    cambios = cambios_al_master(filas, master)
    assert [c[0] for c in cambios] == ['aaa'], cambios
    assert cambios[0][1]['observaciones'] == 'llamar el lunes'
    print('OK  cambios_al_master: solo lo que cambio, y lo que no esta en el master se ignora')

    fila = _fila_para_vendedor({'negocio': 'Barberia X', 'telefono': '0341 353-9510',
                                'ciudad': 'Rosario', 'prioridad': 'A'})
    assert len(fila) == len(COLUMNAS_VENDEDOR)
    assert fila[V['Estado']] == SIN_CONTACTAR
    assert fila[V['Negocio']] == 'Barberia X' and fila[V['Prioridad']] == 'A'
    assert fila[V['Motivo']] == '' and fila[V['Última gestión']] == ''
    print('OK  _fila_para_vendedor: entra Sin contactar y sin gestion previa')

    # --- adonde va cada negocio cuando entra una tanda nueva ----------------
    tanda = [{'clave': 'a', 'estado': 'Volver a llamar'},
             {'clave': 'b', 'estado': 'Le interesó'},
             {'clave': 'c', 'estado': 'Demo iniciada'},
             {'clave': 'd', 'estado': 'Cliente activo'},
             {'clave': 'e', 'estado': 'No interesado'}]
    siguen, carpeta = repartir_por_hoja(tanda, [])
    assert [f['clave'] for f in siguen] == ['a'], siguen
    assert [f['clave'] for f in carpeta] == ['b', 'c', 'd'], carpeta
    print('OK  repartir_por_hoja: el seguimiento se queda, el interesado se guarda')

    # La carpeta se acumula entre tandas y no se duplica.
    vieja = [{'clave': 'x', 'estado': 'Le interesó'}, {'clave': 'b', 'estado': 'Demo iniciada'}]
    _, carpeta = repartir_por_hoja(tanda, vieja)
    assert [f['clave'] for f in carpeta] == ['x', 'b', 'c', 'd'], carpeta
    assert carpeta[1]['estado'] == 'Demo iniciada', 'gana lo que ya estaba en la carpeta'

    # Un interesado que despues dijo que no, sale de la carpeta.
    _, carpeta = repartir_por_hoja([], [{'clave': 'x', 'estado': 'No interesado'}])
    assert carpeta == [], carpeta
    print('OK  repartir_por_hoja: la carpeta acumula, no duplica y suelta a los que dijeron que no')

    # --- el panel -----------------------------------------------------------
    por_clave = {'aaa': {'nicho': 'Peluquerías'}, 'bbb': {'nicho': 'Peluquerías'},
                 'ccc': {'nicho': 'Spas'}}
    filas_v = [{'clave': 'aaa', 'ciudad': 'Rosario'}, {'clave': 'bbb', 'ciudad': 'Rosario'},
               {'clave': 'ccc', 'ciudad': 'Córdoba'}]
    assert nicho_y_ciudad(filas_v, por_clave) == ('Peluquerías', 'Rosario')
    assert nicho_y_ciudad([], por_clave) == ('', ''), 'el vendedor nuevo no tiene nicho'
    print('OK  nicho_y_ciudad: sale de los propios leads, sin guardar estado')

    assert _frase(('Peluquerías', 'Rosario')) == 'Peluquerías en Rosario'
    assert _frase(('Peluquerías', '')) == 'Peluquerías'
    assert _frase(('', 'Rosario')) == 'negocios en Rosario'

    # Le dieron lo que pidio: rubro y ciudad.
    ok = mensaje_filtro(('Spas', 'Rosario'), ('Spas', 'Rosario'), 0, True, '', False)
    assert 'Ahora estas llamando Spas en Rosario.' in ok and 'Podes cambiar' in ok, ok
    # Nunca eligio ciudad: que le haya tocado Córdoba no es un pedido incumplido.
    assert 'Podes cambiar' in mensaje_filtro(('Spas', ''), ('Spas', 'Córdoba'), 0, True, '', False)
    # Pidio a mitad de tanda y no le toca: se le guarda.
    bloq = mensaje_filtro(('Spas', 'Rosario'), ('Peluquerías', 'Córdoba'), 0, False,
                          'ya llamaste a 12', False)
    assert 'Pediste Spas en Rosario' in bloq and 'cuando termines la tanda' in bloq, bloq
    # Se lo permitieron pero no quedaban negocios de eso.
    vacio = mensaje_filtro(('Spas', 'Rosario'), ('Peluquerías', 'Rosario'), 0, True, '', False)
    assert 'No quedan negocios de Spas en Rosario' in vacio, vacio
    # Alcanzaban para 11 de 30: eso no es "no quedan", y confunde si se dice mal.
    poco = mensaje_filtro(('Spas', 'Rosario'), ('Peluquerías', 'Rosario'), 11, True, '', False)
    assert 'Solo quedaban 11 de Spas en Rosario' in poco, poco
    print('OK  mensaje_filtro: distingue cumplido, negado y agotado, con ciudad')

    assert 'faltan 18 por contactar' in mensaje_proximo(
        {'asignados': 30, 'contactados': 12, 'hablados': 8, 'sin_contactar': 18}, 20)
    assert 'faltan 5 conversaciones' in mensaje_proximo(
        {'asignados': 30, 'contactados': 30, 'hablados': 15, 'sin_contactar': 0}, 20)
    assert 'tanda nueva' in mensaje_proximo(
        {'asignados': 30, 'contactados': 30, 'hablados': 10, 'sin_contactar': 0}, 10)
    assert 'Todavia no tenes negocios' in mensaje_proximo(
        {'asignados': 0, 'contactados': 0, 'hablados': 0, 'sin_contactar': 0}, 20)
    # Al de 7 leads se le pide 5, no 20: la vara del panel es la misma que la
    # que despues decide si repone.
    assert 'hablado con 5' in mensaje_proximo(
        {'asignados': 7, 'contactados': 6, 'hablados': 4, 'sin_contactar': 1}, 5)
    print('OK  mensaje_proximo: dice exactamente que le falta')

    res = {'asignados': 30, 'contactados': 12, 'hablados': 8, 'sin_contactar': 18}
    p = filas_panel('Marto', ('Peluquerías', 'Rosario'), res, 4, 'aviso', 'proximo')
    assert '4 guardados' in p[1][0], p[1][0]
    assert 'guardados' not in filas_panel('Marto', ('P', ''), res, 0, 'a', 'b')[1][0]
    assert len(p) == ALTO_PANEL and all(len(f) == 2 for f in p)
    assert p[0][0] == 'Panel de Marto', 'el nombre va arriba de todo'
    assert p[FILA_NICHO - 1] == ['Rubro', 'Peluquerías']
    assert p[FILA_CIUDAD - 1] == ['Ciudad', 'Rosario']
    # Sin ciudad elegida la celda dice "Toda Argentina", no la ciudad que le
    # toco: si no, quedaria filtrado por una ciudad que nunca pidio.
    sin = filas_panel('Marto', ('Peluquerías', ''), res, 0, 'aviso', 'proximo')
    assert sin[FILA_CIUDAD - 1] == ['Ciudad', TODA_ARGENTINA], sin[FILA_CIUDAD - 1]
    for fila in TITULOS[1:] + MENSAJES:
        assert p[fila - 1][1] == '', f'la fila {fila} se combina, la B tiene que ir vacia'
    print('OK  filas_panel: rubro en B5, ciudad en B6 y las combinadas sin B')


if __name__ == '__main__':
    if '--test' in sys.argv:
        demo()
    else:
        main(simular='--simular' in sys.argv)
