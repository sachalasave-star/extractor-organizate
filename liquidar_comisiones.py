"""Liquida las comisiones: cruza los clientes que pagan en la web de Organizate
contra la planilla de ventas y arma las hojas de plata.

    python liquidar_comisiones.py

Necesita, ademas de GOOGLE_CREDENTIALS y SHEET_ID (ver sincronizar_sheets.py):
  ORGANIZATE_TOKEN    token fijo del endpoint de clientes de Organizate
  SHEET_OWNER         (opcional) mail del dueño de la planilla. Sin esto las
                      hojas se protegen en modo aviso en vez de bloqueo real.

Escribe tres hojas:
  Referidos              quien trajo a cada vendedor. Es LO UNICO que se carga
                         a mano, y ni siquiera del todo: la lista de vendedores
                         se rellena sola desde Config y solo hay que elegir de
                         un desplegable. Protegida: si cada vendedor pudiera
                         editarla, se adjudicaria haber referido a todo el
                         mundo y el 20% se pagaria mal.
  Liquidacion            una fila por pago cobrado.
  Comisiones por vendedor  cuanto le toca a cada uno: su 50% mas el 20% de cada
                         afiliado, con el detalle negocio por negocio.

Las tres se rehacen enteras en cada corrida (Referidos conserva lo elegido
cruzando por nombre). Ninguna guarda nada que no se pueda recalcular.
"""
import os
import sys

# La consola de Windows imprime en cp1252 por default, que no tiene los
# emojis de PANEL/REFERIDOS/etc. La nube (Linux) es UTF-8 y nunca lo sufre;
# esto es solo para poder correr el script a mano en Windows sin que un
# print() con emoji tire el proceso abajo a mitad de una escritura real.
sys.stdout.reconfigure(encoding='utf-8')

from armar_planilla import MAX_VENDEDORES
from modules.comisiones import (ASIGNAR, COLUMNAS_ASIGNAR, COLUMNAS_LIQUIDACION,
                                COLUMNAS_REFERIDOS, COLOR_ESTADO, LIQUIDACION,
                                REFERIDOS, RESUMEN_COMISIONES, estado_label,
                                filas_resumen, liquidar, obtener_clientes,
                                par_telefono_vendedor, resumen_por_vendedor,
                                sin_vendedor, vendedor_por_cliente)
from modules.estilo import (BLANCO, DINERO, FUENTE, FUENTE_DATOS, LINEA, TINTA,
                            TINTA_SUAVE, hex_a_rgb, rgb)
from modules.planilla import (CLAVE, CONFIG, FILA_VENDEDORES, IDX, PANEL, RANKING,
                              RESUMEN, col_letra, reintentar)
from sincronizar_sheets import _abrir_libro, _credenciales, _sheet_id

FUERA_DE_NICHOS = (CONFIG, PANEL, RESUMEN, RANKING, LIQUIDACION, REFERIDOS,
                   RESUMEN_COMISIONES, ASIGNAR)

# Referidos: 1 titulo, 2 ayuda, 3 encabezados, 4 en adelante los vendedores.
FILA_DATOS_REFERIDOS = 4

PLATA = {"numberFormat": {"type": "CURRENCY", "pattern": '"$"#,##0.00'}}

DUENO = os.environ.get("SHEET_OWNER", "").strip('﻿ \t\r\n')


# ---------------------------------------------------------------- utilidades

def _celda(sid, fila, formato, campos, col0=0, col1=4):
    """repeatCell sobre una fila entera. Es el ladrillo de todo el formato."""
    return {"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": fila, "endRowIndex": fila + 1,
                  "startColumnIndex": col0, "endColumnIndex": col1},
        "cell": {"userEnteredFormat": formato}, "fields": campos}}


def _texto(negrita=False, tam=10, color=TINTA, fondo=None, alineacion=None,
           italica=False):
    """Formato de texto + los 'fields' que hay que declarar para que aplique."""
    fmt = {"textFormat": {"bold": negrita, "italic": italica, "fontSize": tam,
                          "fontFamily": FUENTE, "foregroundColor": rgb(hex_a_rgb(color))},
           "verticalAlignment": "MIDDLE",
           "padding": {"left": 10, "right": 10}}
    campos = "userEnteredFormat(textFormat,verticalAlignment,padding"
    if fondo:
        fmt["backgroundColor"] = rgb(hex_a_rgb(fondo))
        campos += ",backgroundColor"
    if alineacion:
        fmt["horizontalAlignment"] = alineacion
        campos += ",horizontalAlignment"
    return fmt, campos + ")"


def _fusionar(sid, fila, col1=4):
    return {"mergeCells": {"range": {
        "sheetId": sid, "startRowIndex": fila, "endRowIndex": fila + 1,
        "startColumnIndex": 0, "endColumnIndex": col1}, "mergeType": "MERGE_ALL"}}


def _proteger(libro, sid, descripcion):
    """Deja la hoja de solo lectura para todos menos el dueño. -> True si el
    bloqueo es real.

    Sin SHEET_OWNER cae a warningOnly (avisa, pero deja pasar). Motivo: una
    proteccion dura sin lista de editores la termina administrando la cuenta
    de servicio, y ahi el dueño no puede tocar ni su propia planilla. Mejor
    avisar que dejar a alguien afuera de su hoja.
    """
    protegida = {"range": {"sheetId": sid}, "description": descripcion}
    if DUENO:
        # La cuenta de servicio tiene que figurar en la lista: es la que manda
        # el request, y Google rechaza con "You can't remove yourself as an
        # editor" si se protege dejandose afuera. Ademas la necesita para poder
        # reescribir la hoja en la corrida siguiente.
        protegida["editors"] = {"users": [DUENO, _credenciales()["client_email"]],
                                "domainUsersCanEdit": False}
    else:
        protegida["warningOnly"] = True
    reintentar(libro.batch_update, {"requests": [
        {"addProtectedRange": {"protectedRange": protegida}}]})
    return bool(DUENO)


def _rehacer(libro, titulo, filas, cols):
    """Borra la hoja si estaba y la crea de nuevo con estos valores."""
    try:
        reintentar(libro.del_worksheet, libro.worksheet(titulo))
    except Exception:
        pass
    h = reintentar(libro.add_worksheet, title=titulo, rows=len(filas) + 20, cols=cols)
    reintentar(h.update, values=filas, range_name="A1")
    return h


# ------------------------------------------------------------------- lectura

def _vendedores_de_config(libro):
    """Los vendedores que hay HOY en Config, en orden.

    Antes esto usaba planilla.VENDEDORES, una lista fija de 4 nombres escrita
    en el codigo. En Config ya hay 12 y uno estaba mal escrito (Gije/Gige), asi
    que las comisiones ignoraban a 8 personas. Config es la fuente de verdad:
    es donde el equipo agrega y saca gente.
    """
    col = reintentar(libro.worksheet(CONFIG).col_values, 1)
    return [v.strip() for v in col[FILA_VENDEDORES - 1:] if v.strip()]


def _telefono_a_vendedor(libro):
    """{telefono canonico: vendedor} cruzando Telefono y Vendedor de las hojas
    de nicho. Una fila sin vendedor cargado no aporta nada: nadie cerro ese
    negocio todavia."""
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
            salida.update(par_telefono_vendedor(fila))
    return salida


# ----------------------------------------------------------------- Referidos

def _formato_referidos(sid, n_vend):
    oscuro, claro = DINERO
    ultima = FILA_DATOS_REFERIDOS - 1 + n_vend        # 0-based, exclusivo
    fmt_tit, campos_tit = _texto(negrita=True, tam=13, color=BLANCO, fondo=oscuro)
    fmt_ayu, campos_ayu = _texto(tam=9, color=TINTA_SUAVE, fondo=claro, italica=True)
    fmt_enc, campos_enc = _texto(negrita=True, tam=10, color=TINTA, fondo=claro)
    fmt_dat, campos_dat = _texto(tam=11, color=TINTA)
    reqs = [
        _fusionar(sid, 0, 2), _celda(sid, 0, fmt_tit, campos_tit, col1=2),
        _fusionar(sid, 1, 2), _celda(sid, 1, fmt_ayu, campos_ayu, col1=2),
        _celda(sid, 2, fmt_enc, campos_enc, col1=2),
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 3, "endRowIndex": ultima,
                      "startColumnIndex": 0, "endColumnIndex": 2},
            "cell": {"userEnteredFormat": fmt_dat}, "fields": campos_dat}},
        # El nombre del vendedor en negrita: es la etiqueta de la fila.
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 3, "endRowIndex": ultima,
                      "startColumnIndex": 0, "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
            "fields": "userEnteredFormat.textFormat.bold"}},
        # La celda que se elige, con fondo claro para que se vea que es la editable.
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 3, "endRowIndex": ultima,
                      "startColumnIndex": 1, "endColumnIndex": 2},
            "cell": {"userEnteredFormat": {
                "backgroundColor": rgb(hex_a_rgb(BLANCO)),
                "borders": {"left": {"style": "SOLID", "color": rgb(hex_a_rgb(LINEA))},
                            "right": {"style": "SOLID", "color": rgb(hex_a_rgb(LINEA))},
                            "top": {"style": "SOLID", "color": rgb(hex_a_rgb(LINEA))},
                            "bottom": {"style": "SOLID", "color": rgb(hex_a_rgb(LINEA))}}}},
            "fields": "userEnteredFormat(backgroundColor,borders)"}},
        # El desplegable apunta a Config, igual que la columna Vendedor de las
        # hojas de nicho: agregar gente ahi la habilita aca sin tocar nada.
        {"setDataValidation": {
            "range": {"sheetId": sid, "startRowIndex": 3, "endRowIndex": ultima,
                      "startColumnIndex": 1, "endColumnIndex": 2},
            "rule": {"condition": {"type": "ONE_OF_RANGE", "values": [{"userEnteredValue":
                     f"='{CONFIG}'!$A${FILA_VENDEDORES}:$A${FILA_VENDEDORES + MAX_VENDEDORES}"}]},
                     "showCustomUi": True, "strict": False}}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid, "tabColor": rgb(hex_a_rgb(oscuro)),
                           "gridProperties": {"frozenRowCount": 3}},
            "fields": "tabColor,gridProperties.frozenRowCount"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 42}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 3,
                      "endIndex": ultima},
            "properties": {"pixelSize": 30}, "fields": "pixelSize"}},
    ]
    for i, ancho in enumerate((200, 260)):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": ancho}, "fields": "pixelSize"}})
    return reqs


def _hoja_referidos(libro, vendedores):
    """Rehace Referidos: un renglon por vendedor de Config y un desplegable
    para elegir quien lo trajo. -> (hoja, {vendedor: referido}, bloqueo_real)

    Conserva lo ya elegido cruzando por NOMBRE, no por posicion. Si se
    conservara por fila, agregar un vendedor en el medio de Config le correria
    el referido a todos los de abajo, y una asignacion corrida le paga el 20%
    al que no vendio.
    """
    previo = {}
    try:
        vieja = libro.worksheet(REFERIDOS)
    except Exception:
        vieja = None
    if vieja is not None:
        # Leer ANTES de borrar. Si esta lectura falla, la excepcion sube y la
        # hoja queda intacta: es la unica configuracion cargada a mano de todo
        # el sistema y no se puede regenerar sola.
        for fila in reintentar(vieja.get_all_values)[FILA_DATOS_REFERIDOS - 1:]:
            if len(fila) > 1 and fila[0].strip() and fila[1].strip():
                previo[fila[0].strip()] = fila[1].strip()
        reintentar(libro.del_worksheet, vieja)

    filas = [['🤝 REFERIDOS — quién trajo a cada vendedor', ''],
             ['Elegí de la lista. El que refiere cobra 20% de todo lo que venda '
              'su afiliado, mes a mes.', ''],
             COLUMNAS_REFERIDOS]
    filas += [[v, previo.get(v, '')] for v in vendedores]

    h = _rehacer(libro, REFERIDOS, filas, cols=2)
    reintentar(libro.batch_update, {"requests": _formato_referidos(h.id, len(vendedores))})
    duro = _proteger(libro, h.id,
                     "Referidos: define el 20% de comision. Solo lo edita el dueño.")
    vigente = {v: previo[v] for v in vendedores if v in previo}
    return h, vigente, duro


# ----------------------------------------------------------- Asignar a mano

def _hoja_asignar(libro, pendientes, vendedores):
    """Los clientes que el telefono no pudo atribuir, para resolver a mano.
    -> {id de cliente: vendedor}

    Lista TODOS los que no cruzan por telefono, incluidos los que ya tienen
    vendedor cargado. Si se listaran solo los que faltan, el cliente
    desapareceria de la hoja apenas se lo asigna y en la corrida siguiente se
    perderia lo cargado.

    La clave es el ID de Organizate, no el nombre: el dueño del negocio puede
    renombrarlo cuando quiera.
    """
    previo = {}
    try:
        vieja = libro.worksheet(ASIGNAR)
    except Exception:
        vieja = None
    if vieja is not None:
        # Leer antes de borrar, igual que en Referidos: es trabajo manual que
        # no se puede regenerar solo.
        col_v, col_id = COLUMNAS_ASIGNAR.index('Vendedor'), COLUMNAS_ASIGNAR.index('ID')
        for fila in reintentar(vieja.get_all_values)[FILA_DATOS_REFERIDOS - 1:]:
            if len(fila) > col_id and fila[col_id].strip() and fila[col_v].strip():
                previo[fila[col_id].strip()] = fila[col_v].strip()
        reintentar(libro.del_worksheet, vieja)

    filas = [['🔗 ASIGNAR A MANO — clientes que no cruzaron por teléfono', '', '', '', ''],
             ['Se registraron con otro número (o antes de que la web lo pidiera). '
              'Elegí quién lo vendió y cobra la comisión igual.', '', '', '', ''],
             COLUMNAS_ASIGNAR]
    for c in pendientes:
        cid = c.get('id', '')
        filas.append([c.get('negocio', ''), (c.get('alta') or '')[:10],
                      estado_label(c.get('estado', '')), previo.get(cid, ''), cid])

    h = _rehacer(libro, ASIGNAR, filas, cols=len(COLUMNAS_ASIGNAR))
    reintentar(libro.batch_update,
               {"requests": _formato_asignar(h.id, len(pendientes))})
    _proteger(libro, h.id,
              "Asignacion manual de ventas. Solo la edita el dueño.")
    return {c.get('id', ''): previo[c['id']] for c in pendientes if c.get('id') in previo}


def _formato_asignar(sid, n):
    oscuro, claro = DINERO
    ultima = FILA_DATOS_REFERIDOS - 1 + n
    ncols = len(COLUMNAS_ASIGNAR)
    col_v = COLUMNAS_ASIGNAR.index('Vendedor')
    fmt_tit, campos_tit = _texto(negrita=True, tam=13, color=BLANCO, fondo=oscuro)
    fmt_ayu, campos_ayu = _texto(tam=9, color=TINTA_SUAVE, fondo=claro, italica=True)
    fmt_enc, campos_enc = _texto(negrita=True, tam=10, color=TINTA, fondo=claro)
    reqs = [
        _fusionar(sid, 0, ncols), _celda(sid, 0, fmt_tit, campos_tit, col1=ncols),
        _fusionar(sid, 1, ncols), _celda(sid, 1, fmt_ayu, campos_ayu, col1=ncols),
        _celda(sid, 2, fmt_enc, campos_enc, col1=ncols),
        # El ID es la clave tecnica: se muestra en gris y chico para que no
        # compita con el nombre del negocio, que es lo que se lee.
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 3, "endRowIndex": max(ultima, 4),
                      "startColumnIndex": ncols - 1, "endColumnIndex": ncols},
            "cell": {"userEnteredFormat": {"textFormat": {
                "fontSize": 8, "foregroundColor": rgb(hex_a_rgb(TINTA_SUAVE))}}},
            "fields": "userEnteredFormat.textFormat"}},
        {"setDataValidation": {
            "range": {"sheetId": sid, "startRowIndex": 3, "endRowIndex": max(ultima, 4),
                      "startColumnIndex": col_v, "endColumnIndex": col_v + 1},
            "rule": {"condition": {"type": "ONE_OF_RANGE", "values": [{"userEnteredValue":
                     f"='{CONFIG}'!$A${FILA_VENDEDORES}:$A${FILA_VENDEDORES + MAX_VENDEDORES}"}]},
                     "showCustomUi": True, "strict": False}}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid, "tabColor": rgb(hex_a_rgb(oscuro)),
                           "gridProperties": {"frozenRowCount": 3}},
            "fields": "tabColor,gridProperties.frozenRowCount"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 42}, "fields": "pixelSize"}},
    ]
    for i, ancho in enumerate((300, 100, 120, 170, 230)):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": ancho}, "fields": "pixelSize"}})
    for i, (etiqueta, (fondo, texto)) in enumerate(COLOR_ESTADO.items()):
        reqs.append({"addConditionalFormatRule": {"index": i, "rule": {
            "ranges": [{"sheetId": sid, "startRowIndex": 3, "endRowIndex": max(ultima, 4),
                        "startColumnIndex": 2, "endColumnIndex": 3}],
            "booleanRule": {
                "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": etiqueta}]},
                "format": {"backgroundColor": rgb(fondo),
                           "textFormat": {"foregroundColor": rgb(texto), "bold": True}}}}}})
    return reqs


# --------------------------------------------------- Comisiones por vendedor

def _formato_resumen(sid, marcas, n_filas):
    oscuro, claro = DINERO
    gris = '#F1F3F4'
    reqs = [
        # Plata en las tres columnas de numeros, de una. El formato de numero
        # no toca las celdas de texto, asi que no hace falta esquivar titulos.
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": n_filas,
                      "startColumnIndex": 1, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": PLATA},
            "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": n_filas},
            "cell": {"userEnteredFormat": {
                "textFormat": {"fontFamily": FUENTE_DATOS, "fontSize": 10,
                               "foregroundColor": rgb(hex_a_rgb(TINTA))},
                "verticalAlignment": "MIDDLE", "padding": {"left": 10, "right": 10}}},
            "fields": "userEnteredFormat(textFormat,verticalAlignment,padding)"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid, "tabColor": rgb(hex_a_rgb(oscuro)),
                           "gridProperties": {"frozenRowCount": 1}},
            "fields": "tabColor,gridProperties.frozenRowCount"}},
    ]

    # (rol, formato, fusionar, alto de fila)
    estilos = [
        ('titulo',         _texto(True, 14, BLANCO, oscuro),        True,  44),
        ('ayuda',          _texto(False, 9, TINTA_SUAVE, claro, italica=True), True, 26),
        ('titulo_detalle', _texto(True, 12, BLANCO, oscuro),        True,  36),
        ('enc_resumen',    _texto(True, 10, TINTA, claro),          False, 32),
        ('resumen',        _texto(False, 11, TINTA),                False, 30),
        ('total_equipo',   _texto(True, 12, BLANCO, oscuro),        False, 36),
        ('vendedor',       _texto(True, 12, BLANCO, oscuro),        False, 36),
        ('subseccion',     _texto(True, 10, TINTA, claro),          True,  28),
        ('encabezado',     _texto(True, 9, TINTA_SUAVE, gris),      False, 24),
        ('afiliado',       _texto(True, 10, TINTA, claro),          False, 28),
        ('sin_datos',      _texto(False, 9, TINTA_SUAVE, italica=True), True, 24),
        ('subtotal',       _texto(True, 10, TINTA, gris),           False, 28),
        ('total',          _texto(True, 12, TINTA, claro),          False, 34),
    ]
    for rol, (fmt, campos), fusiona, alto in estilos:
        for fila in marcas.get(rol, []):
            reqs.append(_celda(sid, fila, fmt, campos))
            if fusiona:
                reqs.append(_fusionar(sid, fila))
            reqs.append({"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "ROWS",
                          "startIndex": fila, "endIndex": fila + 1},
                "properties": {"pixelSize": alto}, "fields": "pixelSize"}})

    # El nombre del vendedor de la tabla de arriba, en negrita.
    for fila in marcas.get('resumen', []):
        reqs.append(_celda(sid, fila, {"textFormat": {"bold": True}},
                           "userEnteredFormat.textFormat.bold", col0=0, col1=1))
    # Los numeros a la derecha: se leen en columna y se comparan de un vistazo.
    for rol in ('enc_resumen', 'resumen', 'total_equipo', 'encabezado', 'dato',
                'subtotal', 'total', 'vendedor', 'afiliado'):
        for fila in marcas.get(rol, []):
            reqs.append(_celda(sid, fila, {"horizontalAlignment": "RIGHT"},
                               "userEnteredFormat.horizontalAlignment", col0=1, col1=4))

    # El estado va centrado: es una etiqueta, no un texto que se lee de corrido.
    # Va aparte de la regla condicional de abajo a proposito: la API solo acepta
    # negrita, cursiva, tachado y colores adentro de un formato condicional, y
    # meterle alineacion devuelve 400.
    for fila in marcas.get('dato', []):
        reqs.append(_celda(sid, fila, {"horizontalAlignment": "CENTER"},
                           "userEnteredFormat.horizontalAlignment", col0=1, col1=2))

    # Un color por estado del cliente, igual que el embudo de las hojas de
    # nicho, para no tener que aprender dos codigos de color distintos.
    for i, (etiqueta, (fondo, texto)) in enumerate(COLOR_ESTADO.items()):
        reqs.append({"addConditionalFormatRule": {"index": i, "rule": {
            "ranges": [{"sheetId": sid, "startRowIndex": 0, "endRowIndex": n_filas,
                        "startColumnIndex": 1, "endColumnIndex": 2}],
            "booleanRule": {
                "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": etiqueta}]},
                "format": {"backgroundColor": rgb(fondo),
                           "textFormat": {"foregroundColor": rgb(texto), "bold": True}}}}}})

    for i, ancho in enumerate((320, 150, 150, 170)):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": ancho}, "fields": "pixelSize"}})
    return reqs


def _hoja_resumen_comisiones(libro, filas, marcas):
    h = _rehacer(libro, RESUMEN_COMISIONES, filas, cols=4)
    reintentar(libro.batch_update,
               {"requests": _formato_resumen(h.id, marcas, len(filas))})
    duro = _proteger(libro, h.id,
                     "Calculado por liquidar_comisiones.py. No editar a mano.")
    return h, duro


def _hoja_liquidacion(libro, filas):
    h = _rehacer(libro, LIQUIDACION, [COLUMNAS_LIQUIDACION] + filas,
                 cols=len(COLUMNAS_LIQUIDACION))
    fmt, campos = _texto(negrita=True, color=BLANCO, fondo=DINERO[0])
    reintentar(libro.batch_update, {"requests": [
        _celda(h.id, 0, fmt, campos, col1=len(COLUMNAS_LIQUIDACION)),
        {"updateSheetProperties": {
            "properties": {"sheetId": h.id, "tabColor": rgb(hex_a_rgb(DINERO[0])),
                           "gridProperties": {"frozenRowCount": 1}},
            "fields": "tabColor,gridProperties.frozenRowCount"}}]})
    _proteger(libro, h.id, "Calculado por liquidar_comisiones.py. No editar a mano.")
    return h


# ---------------------------------------------------------------------- main

def main():
    if not _sheet_id() or not (
            os.environ.get("GOOGLE_CREDENTIALS") or os.path.exists("credenciales.json")):
        print("Google Sheets sin configurar (falta SHEET_ID o las credenciales), salteando.")
        return
    if not os.environ.get("ORGANIZATE_TOKEN", "").strip():
        print("Falta ORGANIZATE_TOKEN, salteando la liquidacion de comisiones.")
        return

    libro = _abrir_libro()
    vendedores = _vendedores_de_config(libro)
    print(f"{len(vendedores)} vendedores en {CONFIG}")

    tel_a_vend = _telefono_a_vendedor(libro)
    _, referido_por, duro = _hoja_referidos(libro, vendedores)
    print(f"   {REFERIDOS}: {len(vendedores)} filas, {len(referido_por)} con referido cargado"
          f"{'' if duro else ' (proteccion en modo aviso: falta SHEET_OWNER)'}")

    clientes = obtener_clientes()

    # El telefono resuelve solo; lo que no cruza cae en la hoja de asignacion
    # manual, que es la unica forma de rescatar al cliente que se registro con
    # otro numero (o antes de que la web pidiera telefono).
    por_telefono = vendedor_por_cliente(clientes, tel_a_vend)
    pendientes = sin_vendedor(clientes, por_telefono)
    a_mano = _hoja_asignar(libro, pendientes, vendedores)
    print(f"   {ASIGNAR}: {len(pendientes)} sin cruce por teléfono, "
          f"{len(a_mano)} ya asignados a mano")

    vendedor_de = vendedor_por_cliente(clientes, tel_a_vend, a_mano)
    _hoja_liquidacion(libro, liquidar(clientes, vendedor_de, referido_por))

    resumen = resumen_por_vendedor(clientes, vendedor_de, referido_por, vendedores)
    filas, marcas = filas_resumen(resumen)
    _hoja_resumen_comisiones(libro, filas, marcas)

    cobran = sum(1 for i in resumen.values() if i['total'])
    con_pagos = sum(1 for c in clientes if c.get('pagos'))
    print(f"   {RESUMEN_COMISIONES}: {cobran} de {len(vendedores)} vendedores cobran algo")
    print(f"{len(clientes)} clientes en Organizate, {len(vendedor_de)} atribuidos a un "
          f"vendedor, {con_pagos} con pagos registrados.")
    if clientes and not con_pagos:
        print("   OJO: ningun cliente tiene pagos registrados, asi que no hay "
              "comision que calcular todavia (falta el webhook de Mercado Pago).")


if __name__ == "__main__":
    if '--test' in sys.argv:
        import modules.comisiones as demo_mod
        demo_mod.demo()
    else:
        main()
