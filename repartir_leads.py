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

from modules.asignacion import (LOTE, contactado, elegir_lote, minimo_para,
                                puede_cambiar_nicho, puede_reponer, resumen_lote,
                                _clave)
from modules.estilo import (BLANCO, DINERO, FUENTE, TINTA, TINTA_SUAVE,
                            hex_a_rgb, rgb)
from modules.planilla import (CLAVE, COLUMNAS, CONFIG, FILA_VENDEDORES, IDX,
                              SIN_CONTACTAR, col_letra, reintentar)
from modules.telefono import canonico
from sincronizar_sheets import GENERADO, _abrir_libro, _cliente_gspread, _sheet_id

# Las dos hojas del archivo personal. El panel lo crea Python (agregar una
# hoja a un archivo que ya existe no consume cuota de Drive, crearlo si);
# la de leads la crea el Apps Script con este nombre exacto.
HOJA_VENDEDOR = 'Mis clientes'
PANEL_VENDEDOR = '📊 Mi panel'

# Columnas del archivo personal, en el orden que las crea el Apps Script.
# Si se tocan alla, hay que tocarlas aca: es el unico acoplamiento entre los dos.
COLUMNAS_VENDEDOR = ['Negocio', 'Teléfono', 'Estado', 'Motivo', 'Observaciones',
                     'Última gestión', 'Ciudad', 'Categoría', 'Link en Maps',
                     'Prioridad']
V = {c: i for i, c in enumerate(COLUMNAS_VENDEDOR)}

# Un lead cerrado ya no tiene vuelta y sale del archivo del vendedor cuando le
# toca lote nuevo. Lo demas (incluido "Volver a llamar") se queda: son los
# seguimientos, y borrarlos seria tirar el trabajo hecho.
ESTADOS_CERRADOS = {'No interesado', 'Cliente activo'}

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

FILA_NICHO = 5                # B5: el desplegable
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


def leer_nicho_elegido(hoja):
    """Lo que el vendedor dejo en el desplegable de nicho."""
    return (reintentar(hoja.acell, f'B{FILA_NICHO}').value or '').strip()


def mensaje_nicho(elegido, actual, puede, motivo, repone):
    """Que decirle sobre el cambio de nicho, en castellano.

    `actual` es el nicho que EFECTIVAMENTE quedo en su lote despues de repartir,
    asi que si pidio uno y quedo otro, o se le nego el cambio o el nicho que
    pidio se quedo sin negocios. Las dos cosas hay que decirselas distinto.
    """
    if not elegido or elegido == actual:
        return ('Podes cambiar de nicho hasta la llamada 10, y otra vez cada vez que te '
                'entra un lote nuevo. Elegi del desplegable y se cambia solo.')
    if puede or repone:
        return (f'No quedan negocios de {elegido} sin repartir, asi que seguis en '
                f'{actual}. Proba con otro rubro.')
    return (f'Pediste {elegido} pero {motivo}. Tu lote sigue siendo {actual}: '
            f'cuando lo termines te lo cambiamos.')


def mensaje_proximo(res, minimo):
    """Que le falta para que le entre otro lote."""
    if not res['asignados']:
        return 'Todavia no tenes leads. Te entran en la proxima actualizacion.'
    if res['sin_contactar']:
        return (f"Te faltan {res['sin_contactar']} por contactar. Cuando llames a los "
                f"{res['asignados']} y hayas hablado con {minimo}, te entra "
                f'un lote nuevo.')
    if res['hablados'] < minimo:
        return (f"Ya contactaste a los {res['asignados']}. Te faltan "
                f"{minimo - res['hablados']} conversaciones para que entre "
                f'el lote nuevo.')
    return 'Listo: en la proxima actualizacion te entra un lote nuevo.'


def filas_panel(nombre, nicho, ciudad, res, aviso_nicho, aviso_proximo):
    """Las 16 filas del panel. Logica pura: se testea sin tocar Google."""
    return [
        [f'Panel de {nombre}', ''],
        ['Se actualiza solo dos veces por dia. Vos solo tocas el nicho.', ''],
        ['', ''],
        ['MI NICHO', ''],
        ['Nicho', nicho or '(sin asignar)'],
        ['Ciudad', ciudad or '-'],
        [aviso_nicho, ''],
        ['', ''],
        ['MI LOTE', ''],
        ['Negocios asignados', res['asignados']],
        ['Ya contactados', res['contactados']],
        ['Con los que hablaste', res['hablados']],
        ['Sin contactar', res['sin_contactar']],
        ['', ''],
        ['PROXIMO LOTE', ''],
        [aviso_proximo, ''],
    ]


def _formato_panel(libro, hoja, nichos):
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

    # El desplegable de nicho: lo unico editable, con marco para que se note.
    reqs += [
        {'setDataValidation': {
            'range': rango(FILA_NICHO, FILA_NICHO, 1, 2),
            'rule': {'condition': {'type': 'ONE_OF_LIST',
                                   'values': [{'userEnteredValue': n} for n in nichos]},
                     'showCustomUi': True, 'strict': False,
                     'inputMessage': 'Elegi el rubro que queres trabajar'}}},
        celdas(FILA_NICHO, FILA_NICHO,
               {'backgroundColor': rgb(hex_a_rgb('#FFF8E1')),
                'textFormat': {'bold': True, 'fontSize': 11}},
               'userEnteredFormat(backgroundColor,textFormat)', 1, 2),
        {'updateBorders': dict(
            {'range': rango(FILA_NICHO, FILA_NICHO, 1, 2)},
            **{lado: {'style': 'SOLID', 'color': oscuro}
               for lado in ('top', 'bottom', 'left', 'right')})},
    ]

    # Todo protegido menos la fila del nicho, en modo aviso: el archivo es del
    # vendedor, no hace falta trabarlo. Alcanza con que Sheets le avise antes
    # de pisar algo que igual se reescribe en la proxima corrida.
    for f1, f2 in ((1, FILA_NICHO - 1), (FILA_NICHO + 1, ALTO_PANEL)):
        reqs.append({'addProtectedRange': {'protectedRange': {
            'range': rango(f1, f2), 'warningOnly': True,
            'description': 'Lo escribe el sistema'}}})

    reqs.append({'updateSheetProperties': {
        'properties': {'sheetId': sid, 'tabColor': oscuro}, 'fields': 'tabColor'}})

    reintentar(libro.batch_update, {'requests': reqs})


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

    # El desplegable ofrece solo nichos que todavia tienen leads sin repartir:
    # elegir uno vacio seria elegir quedarse sin lote.
    nichos = sorted({l['nicho'] for l in disponibles if l.get('nicho')})

    escrituras_master, nuevas_asignaciones, paneles = [], [], []

    for v in vendedores:
        try:
            libro_v = reintentar(cliente.open_by_key, v['archivo'])
            hoja_leads = _hoja_leads(libro_v)
            filas = leer_archivo(hoja_leads)
            hoja_panel, panel_nuevo = _hoja_panel(libro_v, crear=not simular)
        except Exception as e:
            print(f"   {v['nombre']}: no pude abrir su archivo ({str(e)[:60]}), lo salteo")
            continue

        if panel_nuevo:
            try:
                _formato_panel(libro_v, hoja_panel, nichos)
            except Exception:
                # Sin formato la hoja queda inservible, y ninguna corrida la
                # volveria a formatear porque ya existe. Se borra y se reintenta
                # de cero la proxima vez.
                reintentar(libro_v.del_worksheet, hoja_panel)
                raise
            print(f"   {v['nombre']}: panel creado")

        def tomar(nicho, faltan):
            """Saca `faltan` leads del pozo y se los anota en el master."""
            nonlocal disponibles
            if faltan <= 0:
                return []
            pozo = [l for l in disponibles if l.get('nicho') == nicho] if nicho else disponibles
            elegidos = elegir_lote(pozo, faltan, tomados)
            if not elegidos and nicho:      # el nicho pedido se quedo sin leads
                elegidos = elegir_lote(disponibles, faltan, tomados)
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
        rescatados = rescatar_del_master(v['nombre'], master, por_clave,
                                         {f['clave'] for f in filas})
        if rescatados:
            print(f"   {v['nombre']}: {len(rescatados)} leads que ya tenia asignados "
                  f"en el master y no estaban en su archivo, se los llevo")
            filas = filas + rescatados

        # 1. Lo que trabajo vuelve al master.
        cambios = cambios_al_master(filas, master)
        for clave, f in cambios:
            m = master[clave]
            escrituras_master.append({
                'range': f"'{m['hoja']}'!{col_letra(IDX['Estado'])}{m['fila']}:"
                         f"{col_letra(IDX['Última gestión'])}{m['fila']}",
                'values': [[f['estado'], f['motivo'], f['observaciones'], _hoy()]]})
            m.update(estado=f['estado'], motivo=f['motivo'],
                     observaciones=f['observaciones'])

        # 2. Que nicho esta trabajando y cual pidio.
        nicho_actual, ciudad = nicho_y_ciudad(filas, por_clave)
        elegido = leer_nicho_elegido(hoja_panel) if hoja_panel and not panel_nuevo else ''
        cambiar = bool(elegido) and elegido != nicho_actual
        libre, motivo_cambio = puede_cambiar_nicho(filas)

        # 3. Le toca lote nuevo?
        r = resumen_lote(filas)
        repone, por_que = puede_reponer(filas)
        print(f"   {v['nombre']}: {r['asignados']} leads, {r['contactados']} contactados, "
              f"{r['hablados']} hablados, {len(cambios)} cambios -> "
              f"{'REPONE' if repone else 'sigue'} ({por_que})")

        siguen, lote, sueltos = filas, [], []
        nicho = elegido if (cambiar and (repone or libre)) else nicho_actual

        if repone:
            # Lote nuevo: los cerrados salen, los seguimientos se quedan. El
            # cambio de nicho aca siempre vale: el lote arranca de cero igual.
            siguen = [f for f in filas if f['estado'] not in ESTADOS_CERRADOS]
            lote = tomar(nicho, LOTE - len(siguen))
            if cambiar:
                print(f'      cambia de nicho: {nicho_actual or "-"} -> {nicho}')
            if not lote:
                print(f'      {len(siguen)} seguimientos abiertos, no entra ninguno nuevo'
                      if len(siguen) >= LOTE else '      no quedan leads para repartir')
            else:
                print(f"      lote nuevo: {len(lote)} leads ({lote[0].get('nicho')}, "
                      f"{lote[0].get('ciudad')}), mas {len(siguen)} seguimientos")
        elif cambiar and libre:
            # Cambio a mitad de lote: lo que ya trabajo se queda, lo que ni
            # miro vuelve a la bolsa para que lo agarre otro.
            siguen = [f for f in filas if contactado(f['estado'])]
            sueltos = [f for f in filas if not contactado(f['estado'])]
            for f in sueltos:
                m = master.get(f['clave'])
                if m and m['vendedor'] == v['nombre']:
                    escrituras_master.append({
                        'range': f"'{m['hoja']}'!{col_letra(IDX['Vendedor'])}{m['fila']}",
                        'values': [['']]})
                    m['vendedor'] = ''
            lote = tomar(nicho, LOTE - len(siguen))
            print(f'      cambia de nicho: {nicho_actual or "-"} -> {nicho} '
                  f'({len(sueltos)} devueltos, {len(lote)} nuevos)')
        elif cambiar:
            print(f'      pidio {elegido} pero {motivo_cambio}')

        # Se rearma el archivo si cambio algo de lo que el vendedor ve. Los
        # rescatados cuentan: si no, quedan asignados a su nombre en el master
        # y el nunca los ve.
        if lote or rescatados or sueltos:
            nuevas_asignaciones.append((v, siguen, lote))

        # 4. El panel, siempre: aunque no cambie nada, los numeros se mueven.
        if hoja_panel:
            final = siguen + [{'estado': SIN_CONTACTAR, 'motivo': ''} for _ in lote]
            res = resumen_lote(final)
            nicho_final = (lote[0].get('nicho') if lote else nicho) or nicho_actual
            ciudad_final = lote[0].get('ciudad') if lote else ciudad
            paneles.append((v, hoja_panel, filas_panel(
                v['nombre'], elegido or nicho_final, ciudad_final, res,
                mensaje_nicho(elegido, nicho_final, libre, motivo_cambio, repone),
                mensaje_proximo(res, minimo_para(final)))))

    if simular:
        print(f'\n[SIMULACION] {len(escrituras_master)} escrituras al master, '
              f'{len(nuevas_asignaciones)} archivos a rearmar y {len(paneles)} '
              f'paneles a refrescar. No se escribio nada.')
        return

    if escrituras_master:
        reintentar(libro.values_batch_update,
                   {'valueInputOption': 'RAW', 'data': escrituras_master})
        print(f'\nMaster actualizado: {len(escrituras_master)} celdas')

    for v, siguen, lote in nuevas_asignaciones:
        hoja = _hoja_leads(reintentar(cliente.open_by_key, v['archivo']))
        filas = ([COLUMNAS_VENDEDOR] +
                 [_fila_conservada(f) for f in siguen] +
                 [_fila_para_vendedor(l) for l in lote])
        reintentar(hoja.clear)
        reintentar(hoja.update, values=filas, range_name='A1')
        print(f"   {v['nombre']}: archivo rearmado con {len(filas) - 1} leads")

    for v, hoja, valores in paneles:
        reintentar(hoja.update, values=valores, range_name=f'A1:B{ALTO_PANEL}')
    if paneles:
        print(f'   {len(paneles)} paneles actualizados')

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

    abiertos = [{'estado': 'Volver a llamar'}, {'estado': 'Le interesó'},
                {'estado': 'Demo iniciada'}, {'estado': SIN_CONTACTAR}]
    cerrados = [{'estado': 'No interesado'}, {'estado': 'Cliente activo'}]
    assert all(f['estado'] not in ESTADOS_CERRADOS for f in abiertos)
    assert all(f['estado'] in ESTADOS_CERRADOS for f in cerrados)
    print('OK  ESTADOS_CERRADOS: los seguimientos abiertos no se tiran')

    # --- el panel -----------------------------------------------------------
    por_clave = {'aaa': {'nicho': 'Peluquerías'}, 'bbb': {'nicho': 'Peluquerías'},
                 'ccc': {'nicho': 'Spas'}}
    filas_v = [{'clave': 'aaa', 'ciudad': 'Rosario'}, {'clave': 'bbb', 'ciudad': 'Rosario'},
               {'clave': 'ccc', 'ciudad': 'Córdoba'}]
    assert nicho_y_ciudad(filas_v, por_clave) == ('Peluquerías', 'Rosario')
    assert nicho_y_ciudad([], por_clave) == ('', ''), 'el vendedor nuevo no tiene nicho'
    print('OK  nicho_y_ciudad: sale de los propios leads, sin guardar estado')

    # Pidio otro nicho a mitad de lote y no le toca: se le dice que se le
    # guarda, y el desplegable NO se le pisa.
    bloqueado = mensaje_nicho('Spas', 'Peluquerías', False, 'ya llamaste a 12', False)
    assert 'Spas' in bloqueado and 'sigue siendo Peluquerías' in bloqueado, bloqueado
    # Se lo permitieron pero el nicho no tenia negocios libres.
    vacio = mensaje_nicho('Spas', 'Peluquerías', True, '', False)
    assert 'No quedan negocios de Spas' in vacio, vacio
    # Le dieron lo que pidio.
    assert 'Podes cambiar' in mensaje_nicho('Spas', 'Spas', True, '', False)
    assert 'Podes cambiar' in mensaje_nicho('', 'Spas', True, '', False)
    print('OK  mensaje_nicho: distingue negado, agotado y concedido')

    assert 'faltan 18 por contactar' in mensaje_proximo(
        {'asignados': 30, 'contactados': 12, 'hablados': 8, 'sin_contactar': 18}, 20)
    assert 'faltan 5 conversaciones' in mensaje_proximo(
        {'asignados': 30, 'contactados': 30, 'hablados': 15, 'sin_contactar': 0}, 20)
    assert 'lote nuevo' in mensaje_proximo(
        {'asignados': 30, 'contactados': 30, 'hablados': 20, 'sin_contactar': 0}, 20)
    assert 'Todavia no tenes leads' in mensaje_proximo(
        {'asignados': 0, 'contactados': 0, 'hablados': 0, 'sin_contactar': 0}, 20)
    # Al de 7 leads se le pide 5, no 20: la vara del panel es la misma que la
    # que despues decide si repone.
    assert 'hablado con 5' in mensaje_proximo(
        {'asignados': 7, 'contactados': 6, 'hablados': 4, 'sin_contactar': 1}, 5)
    print('OK  mensaje_proximo: dice exactamente que le falta')

    p = filas_panel('Marto', 'Peluquerías', 'Rosario',
                    {'asignados': 30, 'contactados': 12, 'hablados': 8, 'sin_contactar': 18},
                    'aviso', 'proximo')
    assert len(p) == ALTO_PANEL and all(len(f) == 2 for f in p)
    assert p[0][0] == 'Panel de Marto', 'el nombre va arriba de todo'
    assert p[FILA_NICHO - 1] == ['Nicho', 'Peluquerías'], 'el desplegable esta en B5'
    assert [f[0] for f in p if f[0] in ('MI NICHO', 'MI LOTE', 'PROXIMO LOTE')] == \
           ['MI NICHO', 'MI LOTE', 'PROXIMO LOTE']
    for fila in TITULOS[1:] + MENSAJES:
        assert p[fila - 1][1] == '', f'la fila {fila} se combina, la B tiene que ir vacia'
    print('OK  filas_panel: 16 filas, el nicho en B5 y las combinadas sin B')



if __name__ == '__main__':
    if '--test' in sys.argv:
        demo()
    else:
        main(simular='--simular' in sys.argv)
