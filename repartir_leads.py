"""Reparte leads a los vendedores y trae de vuelta lo que trabajaron.

    python repartir_leads.py              # reparte y sincroniza
    python repartir_leads.py --simular    # dice que haria, sin escribir nada

Cada vendedor tiene su propio archivo con su lote. No ve la base ni el trabajo
de los demas. El archivo se lo crea el Apps Script (apps_script/Vendedores.gs);
aca se le carga el lote y se le lee lo que gestiono.

En cada corrida, por cada vendedor:

  1. Se lee su archivo y lo que cambio se copia al master (Estado, Motivo,
     Observaciones). Si algo cambio respecto del master, se le pone la fecha
     de hoy en Ultima gestion: el master hace de foto anterior, asi que no
     hace falta ni un trigger ni guardar un snapshot aparte.
  2. Se decide si le toca lote nuevo (modules/asignacion.puede_reponer).
  3. Si le toca: salen de su archivo los leads cerrados, quedan los que
     siguen vivos, y se completa hasta 30 con leads nuevos.

Necesita GOOGLE_CREDENTIALS y SHEET_ID, igual que el resto del pipeline.
"""
import os
import sys
from datetime import date

sys.stdout.reconfigure(encoding='utf-8')

import pandas as pd

from modules.asignacion import LOTE, elegir_lote, puede_reponer, resumen_lote, _clave
from modules.planilla import (CLAVE, COLUMNAS, CONFIG, FILA_VENDEDORES, IDX,
                              SIN_CONTACTAR, col_letra, reintentar)
from modules.telefono import canonico
from sincronizar_sheets import GENERADO, _abrir_libro, _cliente_gspread, _sheet_id

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


def leer_archivo(cliente, archivo_id):
    """Las filas del archivo del vendedor, con TODAS sus columnas.

    Se leen todas y no solo las tres que el vendedor edita, porque cuando le
    toca lote nuevo el archivo se rearma: si aca se perdiera Ciudad o Link en
    Maps, los seguimientos que se conservan quedarian sin esos datos.
    """
    hoja = reintentar(cliente.open_by_key, archivo_id).sheet1
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

    escrituras_master, nuevas_asignaciones = [], []

    for v in vendedores:
        try:
            filas = leer_archivo(cliente, v['archivo'])
        except Exception as e:
            print(f"   {v['nombre']}: no pude abrir su archivo ({str(e)[:60]}), lo salteo")
            continue

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

        # 2. Le toca lote nuevo?
        r = resumen_lote(filas)
        ok, por_que = puede_reponer(filas)
        print(f"   {v['nombre']}: {r['asignados']} leads, {r['contactados']} contactados, "
              f"{r['hablados']} hablados, {len(cambios)} cambios -> "
              f"{'REPONE' if ok else 'sigue'} ({por_que})")
        # 3. Si repone: salen los cerrados, quedan los vivos y se completa
        # hasta 30. Si no repone, se queda con lo mismo que tenia.
        siguen = [f for f in filas if f['estado'] not in ESTADOS_CERRADOS] if ok else filas
        lote = []

        if ok:
            faltan = LOTE - len(siguen)
            if faltan <= 0:
                print(f'      {len(siguen)} seguimientos abiertos, no entra ninguno nuevo')
            else:
                lote = elegir_lote(disponibles, faltan, tomados)
                if not lote:
                    print('      no quedan leads disponibles para repartir')
                else:
                    for lead in lote:
                        clave = _clave(lead)
                        tomados.add(clave)
                        m = master.get(clave)
                        if m:
                            escrituras_master.append({
                                'range': f"'{m['hoja']}'!"
                                         f"{col_letra(IDX['Vendedor'])}{m['fila']}",
                                'values': [[v['nombre']]]})
                            m['vendedor'] = v['nombre']
                    disponibles = [l for l in disponibles if _clave(l) not in tomados]
                    print(f"      lote nuevo: {len(lote)} leads "
                          f"({lote[0].get('nicho')}, {lote[0].get('ciudad')}), "
                          f"mas {len(siguen)} seguimientos que se quedan")

        # Se rearma el archivo si hay lote nuevo, o si hay que meterle los
        # rescatados aunque no le toque reponer: si no, esos leads se quedarian
        # asignados a su nombre en el master y el nunca los veria.
        if lote or rescatados:
            nuevas_asignaciones.append((v, siguen, lote))

    if simular:
        print(f'\n[SIMULACION] {len(escrituras_master)} escrituras al master y '
              f'{len(nuevas_asignaciones)} archivos a rearmar. No se escribio nada.')
        return

    if escrituras_master:
        reintentar(libro.values_batch_update,
                   {'valueInputOption': 'RAW', 'data': escrituras_master})
        print(f'\nMaster actualizado: {len(escrituras_master)} celdas')

    for v, siguen, lote in nuevas_asignaciones:
        hoja = reintentar(cliente.open_by_key, v['archivo']).sheet1
        filas = ([COLUMNAS_VENDEDOR] +
                 [_fila_conservada(f) for f in siguen] +
                 [_fila_para_vendedor(l) for l in lote])
        reintentar(hoja.clear)
        reintentar(hoja.update, values=filas, range_name='A1')
        print(f"   {v['nombre']}: archivo rearmado con {len(filas) - 1} leads")

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


if __name__ == '__main__':
    if '--test' in sys.argv:
        demo()
    else:
        main(simular='--simular' in sys.argv)
