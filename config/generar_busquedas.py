"""Regenera config/busquedas.xlsx desde esta lista.

Editar aca y correr `python config/generar_busquedas.py`. Las busquedas ya
hechas quedan marcadas en la base, asi que agregar keywords no rehace trabajo.

Criterio para incluir un nicho: tiene que vender TURNOS frecuentes y que el
cliente vuelva. Por eso quedan afuera chapa y pintura, gomerias o inmobiliarias:
poco volumen de turnos, se coordinan sobre la hora y el cliente no vuelve en
meses.

El ORDEN de este archivo es el orden en que el scraper trabaja, y como nunca
llega al final, lo de arriba es lo unico que existe. Antes se ordenaba por
Rubro alfabetico y "Deporte y movimiento" empezaba con D: 464 de las primeras
569 busquedas se fueron en gimnasios y academias, que son justo los que peor
encajan. Ahora manda PRIORIDAD.
"""
import pandas as pd

# nicho -> (rubro, [keywords]). Varias keywords por nicho porque Google devuelve
# resultados distintos segun el termino: "psicologo" y "psicoanalista" no traen
# la misma lista aunque sean el mismo negocio.
NICHOS = {
 # ---------- Estética y belleza ----------
 'Peluquerías':            ('Estética y belleza', ['peluquería', 'peluquería unisex', 'salón de belleza', 'peluquería para mujeres', 'estilista']),
 'Barberías':              ('Estética y belleza', ['barbería', 'barber shop', 'peluquería masculina', 'barbería para hombres', 'barbero']),
 'Centros de estética':    ('Estética y belleza', ['centro de estética', 'estética facial', 'clínica estética', 'estética corporal', 'cosmetología', 'esteticista']),
 'Manicura y pedicura':    ('Estética y belleza', ['manicura', 'pedicura', 'salón de uñas', 'nail salon', 'esculpidas de uñas']),
 'Cejas y pestañas':       ('Estética y belleza', ['extensiones de pestañas', 'diseño de cejas', 'lifting de pestañas', 'microblading', 'perfilado de cejas', 'lash extension', 'pestañas pelo a pelo', 'diseño de cejas con henna']),
 'Depilación definitiva':  ('Estética y belleza', ['depilación definitiva', 'depilación láser', 'centro de depilación', 'depilación con cera', 'depilación masculina']),
 'Salones de maquillaje':  ('Estética y belleza', ['maquillaje profesional', 'maquilladora', 'estudio de maquillaje', 'maquillaje para eventos', 'automaquillaje']),
 'Spas':                   ('Estética y belleza', ['spa', 'day spa', 'centro de relajación', 'spa y masajes', 'sauna']),
 'Centros de masajes':     ('Estética y belleza', ['centro de masajes', 'masajes terapéuticos', 'masajista', 'masajes descontracturantes', 'drenaje linfático']),
 'Centros de bronceado':   ('Estética y belleza', ['centro de bronceado', 'solarium', 'bronceado natural', 'cama solar']),
 'Peluquería canina':      ('Estética y belleza', ['peluquería canina', 'peluquería para mascotas', 'grooming canino', 'baño y corte de perros', 'estética canina']),

 # ---------- Salud y bienestar ----------
 'Kinesiología':           ('Salud y bienestar', ['kinesiología', 'kinesiólogo', 'centro de kinesiología', 'rehabilitación kinésica', 'kinesióloga']),
 'Fisioterapia':           ('Salud y bienestar', ['fisioterapia', 'fisioterapeuta', 'clínica de fisioterapia', 'rehabilitación física', 'centro de rehabilitación']),
 'Quiropraxia y osteopatía': ('Salud y bienestar', ['quiropraxia', 'quiropráctico', 'osteopatía', 'osteópata', 'terapia manual']),
 'Psicólogos':             ('Salud y bienestar', ['psicólogo', 'psicóloga', 'psicoanalista', 'psicoterapeuta', 'consultorio psicológico', 'terapia psicológica', 'psicoanálisis']),
 'Psicopedagogía':         ('Salud y bienestar', ['psicopedagogía', 'psicopedagoga', 'consultorio psicopedagógico', 'apoyo escolar profesional']),
 'Terapia ocupacional':    ('Salud y bienestar', ['terapia ocupacional', 'terapista ocupacional', 'terapeuta ocupacional']),
 'Fonoaudiología':         ('Salud y bienestar', ['fonoaudiología', 'fonoaudióloga', 'terapia del lenguaje', 'consultorio fonoaudiológico', 'logopedia']),
 'Nutricionistas':         ('Salud y bienestar', ['nutricionista', 'consultorio nutricional', 'licenciada en nutrición', 'asesoramiento nutricional', 'nutrición deportiva']),
 'Odontología':            ('Salud y bienestar', ['odontología', 'dentista', 'clínica dental', 'consultorio odontológico', 'ortodoncia', 'implantes dentales']),
 'Podología':              ('Salud y bienestar', ['podología', 'podólogo', 'consultorio podológico', 'tratamiento de pies', 'podóloga']),
 'Oftalmología':           ('Salud y bienestar', ['oftalmología', 'oftalmólogo', 'consultorio oftalmológico', 'control de la vista']),
 'Dermatología':           ('Salud y bienestar', ['dermatología', 'dermatólogo', 'consultorio dermatológico', 'dermatóloga']),
 'Laboratorios de análisis': ('Salud y bienestar', ['laboratorio de análisis clínicos', 'análisis clínicos', 'extracción de sangre', 'laboratorio bioquímico']),
 'Diagnóstico por imágenes': ('Salud y bienestar', ['diagnóstico por imágenes', 'ecografías', 'radiología', 'resonancia magnética', 'centro de imágenes']),
 'Medicina alternativa':   ('Salud y bienestar', ['medicina alternativa', 'acupuntura', 'terapias holísticas', 'medicina naturista', 'reiki', 'homeopatía']),
 'Clínicas':               ('Salud y bienestar', ['clínica médica', 'centro médico', 'clínica privada', 'policlínico', 'consultorios médicos']),
 # Nuevos: consultorio de un solo profesional, sesion de 30-45 min y control
 # periodico. Es exactamente la forma del calendario.
 'Psiquiatría':            ('Salud y bienestar', ['psiquiatra', 'psiquiatría', 'consultorio psiquiátrico', 'psiquiatra infantil']),
 'Ginecología y obstetricia': ('Salud y bienestar', ['ginecólogo', 'ginecóloga', 'consultorio ginecológico', 'obstetra', 'partera']),
 'Pediatría':              ('Salud y bienestar', ['pediatra', 'consultorio pediátrico', 'pediatría', 'médico de niños']),
 'Traumatología':          ('Salud y bienestar', ['traumatólogo', 'traumatología', 'consultorio traumatológico', 'ortopedia y traumatología']),
 'Flebología':             ('Salud y bienestar', ['flebología', 'flebólogo', 'tratamiento de várices', 'escleroterapia']),

 # ---------- Deporte y movimiento ----------
 'Gimnasios':              ('Deporte y movimiento', ['gimnasio', 'gym', 'centro de entrenamiento', 'gimnasio funcional', 'crossfit', 'entrenamiento funcional']),
 'Yoga y pilates':         ('Deporte y movimiento', ['yoga', 'pilates', 'estudio de yoga', 'centro de pilates', 'pilates reformer']),
 'Entrenadores personales': ('Deporte y movimiento', ['entrenador personal', 'personal trainer', 'preparador físico', 'coach fitness', 'entrenadora personal']),
 'Artes marciales':        ('Deporte y movimiento', ['escuela de artes marciales', 'karate', 'jiu jitsu', 'taekwondo', 'boxeo', 'muay thai']),
 'Natación':               ('Deporte y movimiento', ['natatorio', 'escuela de natación', 'clases de natación', 'pileta climatizada']),
 'Escuelas de danza':      ('Deporte y movimiento', ['academia de baile', 'escuela de danza', 'clases de baile', 'ballet', 'danza clásica']),
 'Academias':              ('Deporte y movimiento', ['academia deportiva', 'escuela de deportes', 'club deportivo']),

 # ---------- Formación ----------
 'Escuelas de música':     ('Formación', ['escuela de música', 'clases de guitarra', 'clases de piano', 'clases de canto', 'profesor de música']),
 'Escuelas de idiomas':    ('Formación', ['instituto de idiomas', 'clases de inglés', 'escuela de idiomas', 'profesor de inglés', 'clases de portugués']),
 'Autoescuelas':           ('Formación', ['autoescuela', 'escuela de manejo', 'escuela de conducir', 'curso de manejo', 'clases de manejo']),
 'Clases particulares':    ('Formación', ['profesor particular', 'clases particulares', 'apoyo escolar', 'profesora particular de matemática', 'clases de apoyo']),

 # ---------- Otros servicios ----------
 'Estudios de tatuajes':   ('Otros servicios', ['estudio de tatuajes', 'tattoo studio', 'tatuador', 'piercing y tatuajes', 'tatuadora']),
 'Estudios de fotografía': ('Otros servicios', ['estudio fotográfico', 'fotógrafo profesional', 'sesión de fotos', 'book fotográfico']),
 'Veterinarias':           ('Otros servicios', ['veterinaria', 'clínica veterinaria', 'consultorio veterinario', 'veterinario', 'urgencias veterinarias']),
 'Ópticas':                ('Otros servicios', ['óptica', 'óptica y lentes', 'óptica profesional', 'venta de lentes', 'optometrista']),
 'Talleres':               ('Otros servicios', ['taller mecánico', 'taller de autos', 'mecánica del automotor', 'taller multimarca', 'service de autos']),
 'Lavaderos y detailing':  ('Otros servicios', ['lavadero de autos', 'car wash', 'detailing automotor', 'lavadero y polarizado', 'pulido de autos']),
 'Adiestramiento canino':  ('Otros servicios', ['adiestramiento canino', 'adiestrador de perros', 'educación canina', 'entrenador de perros']),
}

# Que tan bien encaja el nicho con un calendario de UN turno por franja.
#
# 1 = el turno ES el negocio. Un profesional, un cliente, reserva previa, y el
#     cliente vuelve. Sin agenda no laburan. Un tatuador bloquea dos horas y
#     pide seña; un psicologo tiene el mismo horario todos los martes.
# 2 = encaja, pero el turno compite con algo. Varias cabinas o profesionales en
#     paralelo, obra social de por medio, o una recepcionista que ya lo resuelve
#     por telefono. Se vende, cuesta mas.
# 3 = no encaja. Clase grupal, cupo por capacidad o directamente por orden de
#     llegada. Un gimnasio no reserva: entra el que llega, y son 40 a la vez.
#     Se generan igual, para no perder el historial de lo ya buscado, pero
#     salen con Activo=No y el scraper no los toca.
PRIORIDAD = {
 # --- 1: el turno es el negocio ---
 'Estudios de tatuajes': 1, 'Peluquerías': 1, 'Barberías': 1,
 'Manicura y pedicura': 1, 'Cejas y pestañas': 1, 'Depilación definitiva': 1,
 'Centros de estética': 1, 'Salones de maquillaje': 1, 'Centros de masajes': 1,
 'Peluquería canina': 1, 'Psicólogos': 1, 'Psiquiatría': 1,
 'Nutricionistas': 1, 'Kinesiología': 1, 'Fisioterapia': 1,
 'Quiropraxia y osteopatía': 1, 'Odontología': 1, 'Podología': 1,
 'Fonoaudiología': 1, 'Psicopedagogía': 1, 'Terapia ocupacional': 1,
 'Dermatología': 1, 'Flebología': 1, 'Medicina alternativa': 1,
 'Veterinarias': 1, 'Estudios de fotografía': 1, 'Adiestramiento canino': 1,
 'Clases particulares': 1, 'Entrenadores personales': 1, 'Autoescuelas': 1,

 # --- 2: encaja con friccion ---
 'Spas': 2, 'Centros de bronceado': 2, 'Ópticas': 2,
 'Ginecología y obstetricia': 2, 'Pediatría': 2, 'Traumatología': 2,
 'Oftalmología': 2, 'Laboratorios de análisis': 2,
 'Diagnóstico por imágenes': 2, 'Clínicas': 2,
 'Talleres': 2, 'Lavaderos y detailing': 2, 'Escuelas de música': 2,

 # --- 3: no encaja, apagados ---
 'Gimnasios': 3, 'Academias': 3, 'Artes marciales': 3, 'Natación': 3,
 'Escuelas de danza': 3, 'Yoga y pilates': 3, 'Escuelas de idiomas': 3,
}

APAGADO = 3          # de esta prioridad en adelante, Activo=No

# Los 24 partidos del conurbano. "Buenos Aires" a secas devuelve CABA sola, y
# afuera de la General Paz vive mas gente que en Cordoba, Rosario y Mendoza
# juntas: era el agujero mas grande que teniamos.
#
# Todos llevan ", Buenos Aires" pegado y NO es decorativo. El termino que se
# busca es "{keyword} en {ciudad}, Argentina", asi que el nombre del partido va
# crudo a Maps, y varios existen en otra provincia: Merlo es San Luis, Pilar es
# Cordoba, Avellaneda es Santa Fe, Ituzaingo es Corrientes, San Fernando es
# Catamarca y San Miguel es Tucuman. Es el mismo bug que "Cordoba" trayendo
# España, que no se noto hasta que aparecieron 91 telefonos con +34.
CONURBANO = ['La Matanza', 'Lomas de Zamora', 'Quilmes', 'Almirante Brown',
             'Merlo', 'Moreno', 'Lanús', 'Florencio Varela',
             'General San Martín', 'Tigre', 'Avellaneda', 'Berazategui',
             'Malvinas Argentinas', 'Morón', 'Tres de Febrero', 'Pilar',
             'Esteban Echeverría', 'San Isidro', 'San Miguel', 'José C. Paz',
             'Escobar', 'Vicente López', 'Hurlingham', 'Ituzaingó']

# El conurbano va pegado a Buenos Aires porque es la misma plaza. No le saca el
# lugar a nadie que importe: prioridad 1 ya cerro Rosario, Buenos Aires,
# Cordoba, Mendoza, La Plata y Mar del Plata, asi que lo unico que retrasa es
# Tucuman para abajo, que son plazas chicas al lado de estos 24 partidos.
#
# NO renombrar una ciudad ya scrapeada. La busqueda queda marcada como hecha
# por su termino completo, asi que tocarle el nombre a "Buenos Aires" manda a
# rehacer de cero las 147 busquedas que ya cerro.
CIUDADES = (['Rosario', 'Buenos Aires']
            + [f'{p}, Buenos Aires' for p in CONURBANO]
            + ['Córdoba', 'Mendoza', 'La Plata', 'Mar del Plata',
               'San Miguel de Tucumán', 'Salta', 'Santa Fe', 'Neuquén',
               'Bahía Blanca', 'Paraná'])


def generar(destino='config/busquedas.xlsx'):
    faltan = set(NICHOS) - set(PRIORIDAD)
    if faltan:
        raise ValueError(f"Nichos sin prioridad: {sorted(faltan)}. Sin eso no se "
                         f"sabe donde ponerlos en la cola.")

    filas = [{'Rubro': rubro, 'Nicho': nicho, 'Prioridad': PRIORIDAD[nicho],
              'Busqueda': kw, 'Ciudad': ciudad,
              'Activo': 'No' if PRIORIDAD[nicho] >= APAGADO else 'Si'}
             for ciudad in CIUDADES
             for nicho, (rubro, kws) in NICHOS.items()
             for kw in kws]
    # Este sort ES la decision de que se scrapea: el scraper recorre el archivo
    # en orden y nunca llega al final.
    #
    # Ciudad antes que Nicho a proposito. Ordenando por nicho se termina con
    # tatuajes en las 12 ciudades y cero peluquerias, porque la A va antes; con
    # la ciudad adelante cada plaza queda COMPLETA en todos los nichos que
    # sirven antes de pasar a la siguiente. Y el orden de CIUDADES no es
    # alfabetico, es por donde conviene vender primero.
    orden_ciudad = {c: i for i, c in enumerate(CIUDADES)}
    df = pd.DataFrame(filas)
    df = (df.assign(_c=df['Ciudad'].map(orden_ciudad))
            .sort_values(['Prioridad', '_c', 'Nicho', 'Busqueda'])
            .drop(columns='_c'))
    df.to_excel(destino, index=False)

    kws = sum(len(k) for _, k in NICHOS.values())
    print(f"{len(df)} busquedas = {len(NICHOS)} nichos / {kws} keywords x {len(CIUDADES)} ciudades")
    for p, g in df.groupby('Prioridad'):
        estado = 'apagados' if p >= APAGADO else 'activos'
        print(f"   prioridad {p}: {g['Nicho'].nunique():2d} nichos, "
              f"{len(g):4d} busquedas ({estado})")
    return df


def demo():
    df = generar('output/_busquedas_test.xlsx')
    prio = df['Prioridad'].tolist()
    assert prio == sorted(prio), 'el archivo no quedo ordenado por prioridad'
    assert df.iloc[0]['Prioridad'] == 1, 'no arranca por los que mejor encajan'

    activos = df[df['Activo'] == 'Si']
    assert set(activos['Prioridad']) <= {1, 2}, 'quedo activo un nicho que no encaja'
    assert 'Gimnasios' not in set(activos['Nicho']), 'los gimnasios siguen prendidos'
    # Dentro de cada prioridad, la primera ciudad tiene que quedar entera antes
    # de pasar a la segunda: es lo que evita terminar con un nicho en 12
    # ciudades y el resto en cero. La prioridad manda sobre la ciudad, asi que
    # esto se cumple por bloque, no sobre el archivo entero.
    for p, g in activos.groupby('Prioridad'):
        primera = g[g['Ciudad'] == CIUDADES[0]]
        assert set(primera['Nicho']) == set(g['Nicho']), \
            f'prioridad {p}: {CIUDADES[0]} no cubre todos sus nichos'
        assert g.head(len(primera))['Ciudad'].eq(CIUDADES[0]).all(), \
            f'prioridad {p}: no arranca completando {CIUDADES[0]}'

    # Un partido sin provincia se va a la otra punta del pais y no se nota:
    # vuelve con negocios reales, telefono valido y todo. Solo se cae si alguien
    # mira los codigos de area.
    for p in CONURBANO:
        assert f'{p}, Buenos Aires' in CIUDADES, f'{p} quedo sin provincia'
    assert len(CIUDADES) == len(set(CIUDADES)), 'hay una ciudad repetida'
    # Y que sigan pegados a Buenos Aires: si el conurbano cae al final del
    # archivo, el scraper no llega nunca y agregarlo no sirvio de nada.
    i = CIUDADES.index('Buenos Aires')
    assert CIUDADES[i + 1:i + 1 + len(CONURBANO)] == [f'{p}, Buenos Aires' for p in CONURBANO], \
        'el conurbano se despego de Buenos Aires'

    import os
    os.remove('output/_busquedas_test.xlsx')
    print(f'OK  busquedas: {len(activos)} activas, ordenadas por prioridad')


if __name__ == '__main__':
    import sys
    demo() if '--test' in sys.argv else generar()
