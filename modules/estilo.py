"""Paleta y tipografia de la planilla de ventas.

Un color por rubro, aplicado a la pestaña y al encabezado, para que se sepa
donde uno esta parado sin leer el nombre de la hoja. Los estados usan otra
escala aparte: son semaforo, no identidad.
"""

FUENTE = "Inter"            # Sheets cae a Arial solo si no la tiene
FUENTE_DATOS = "Inter"
TAM_ENCABEZADO = 10
TAM_DATOS = 10


def hex_a_rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


# Un tono por rubro: oscuro para el encabezado, claro para las filas alternadas.
RUBROS = {
    'Estética y belleza':   ('#8E3B6B', '#F7EDF3'),
    'Salud y bienestar':    ('#1F6F6B', '#E9F4F3'),
    'Deporte y movimiento': ('#1D4E89', '#EAF0F8'),
    'Formación':            ('#8A5A1B', '#FAF1E4'),
    'Otros servicios':      ('#4A4A55', '#F0F0F2'),
}
RUBRO_POR_DEFECTO = RUBROS['Otros servicios']

TINTA = '#1A1A1A'           # texto principal
TINTA_SUAVE = '#6B6B76'     # texto secundario
LINEA = '#DADCE0'           # bordes
BLANCO = '#FFFFFF'


def color_rubro(rubro):
    return RUBROS.get(rubro, RUBRO_POR_DEFECTO)
