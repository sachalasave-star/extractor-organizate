"""Estructura de la planilla de ventas en Google Sheets.

Una hoja por nicho, mas Config (vendedores editables) y Ranking (metricas
vivas). Los desplegables de Vendedor apuntan a Config, asi sumar o sacar un
vendedor es escribir en una celda y no tocar 43 validaciones.
"""

# Columnas de la hoja de cada nicho, en orden. Las 5 primeras son las que se
# usan llamando; el resto es contexto que se mira solo cuando hace falta.
COLUMNAS = ['Negocio', 'Teléfono', 'Categoría', 'Vendedor', 'Estado', 'Motivo',
            'Observaciones', 'Última gestión', 'Ciudad', 'Link en Maps']
ANCHOS = [270, 120, 160, 110, 130, 170, 400, 115, 125, 90]

# De donde sale cada columna en el Excel generado. '' = la llena el vendedor.
ORIGEN = {'Negocio': 'Negocio', 'Teléfono': 'Teléfono', 'Categoría': 'Categoría',
          'Ciudad': 'Ciudad', 'Link en Maps': 'Link en Maps'}

CLAVE = 'Teléfono'

VENDEDORES = ['Augusto', 'Valentino', 'Gije', 'Marto']

# (estado, fondo, texto). El orden es el del embudo, de frio a cliente.
ESTADOS = [
    ('Sin contactar',  (0.26, 0.26, 0.26), (1, 1, 1)),
    ('No interesado',  (0.85, 0.24, 0.24), (1, 1, 1)),
    ('Volver a llamar', (0.99, 0.83, 0.31), (0.15, 0.15, 0.15)),
    ('Le interesó',    (0.20, 0.49, 0.90), (1, 1, 1)),
    ('Demo iniciada',  (0.55, 0.30, 0.80), (1, 1, 1)),
    ('Cliente activo', (0.20, 0.66, 0.33), (1, 1, 1)),
]
NOMBRES_ESTADO = [e[0] for e in ESTADOS]
SIN_CONTACTAR = NOMBRES_ESTADO[0]

# Por que un lead no avanzo. Sirve para saber si el problema es el pitch, el
# precio o la lista: "numero equivocado" es un dato del scraper, "precio alto"
# es un dato del negocio.
MOTIVOS = ['No atiende', 'Número equivocado', 'Precio alto', 'Ya tiene sistema',
           'Pidió que llame después', 'No es el rubro', 'Cerró']

CONFIG = '⚙️ Config'
PANEL = '📊 Panel de ventas'
RESUMEN = '📁 Resumen por nicho'

# Filas de Config donde viven las listas (1-based, como las ve el usuario)
FILA_VENDEDORES = 3          # A3 hacia abajo
FILA_ESTADOS = 3             # C3 hacia abajo
FILA_MOTIVOS = 3             # E3 hacia abajo


def col_letra(i):
    """0 -> A. Solo se usa hasta la Z, alcanza para 10 columnas."""
    return chr(ord('A') + i)


def reintentar(fn, *a, **kw):
    """Sheets corta a las 60 lecturas y 60 escrituras por minuto. Espera y
    sigue en vez de dejar la planilla a medio armar o a medio sincronizar."""
    import time
    for intento in range(6):
        try:
            return fn(*a, **kw)
        except Exception as e:
            if '429' not in str(e) or intento == 5:
                raise
            espera = 30 * (intento + 1)
            print(f"   (limite de Google, esperando {espera}s)")
            time.sleep(espera)


IDX = {c: i for i, c in enumerate(COLUMNAS)}
COL_VENDEDOR = col_letra(IDX['Vendedor'])
COL_ESTADO = col_letra(IDX['Estado'])
COL_MOTIVO = col_letra(IDX['Motivo'])
COL_FECHA = col_letra(IDX['Última gestión'])
COL_TELEFONO = col_letra(IDX[CLAVE])


def fila_desde(negocio):
    """Arma la fila de la planilla desde un registro del Excel generado."""
    return [str(negocio.get(ORIGEN[c], '')) if c in ORIGEN
            else (SIN_CONTACTAR if c == 'Estado' else '')
            for c in COLUMNAS]
