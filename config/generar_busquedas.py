"""Regenera config/busquedas.xlsx desde esta lista.

Editar aca y correr `python config/generar_busquedas.py`. Las busquedas ya
hechas quedan marcadas en la base, asi que agregar keywords no rehace trabajo.

Criterio para incluir un nicho: tiene que vender TURNOS frecuentes y que el
cliente vuelva. Por eso quedan afuera chapa y pintura, gomerias o inmobiliarias:
poco volumen de turnos, se coordinan sobre la hora y el cliente no vuelve en
meses.
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
 'Cejas y pestañas':       ('Estética y belleza', ['extensiones de pestañas', 'diseño de cejas', 'lifting de pestañas', 'microblading', 'perfilado de cejas']),
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

 # ---------- Otros servicios ----------
 'Estudios de tatuajes':   ('Otros servicios', ['estudio de tatuajes', 'tattoo studio', 'tatuador', 'piercing y tatuajes', 'tatuadora']),
 'Estudios de fotografía': ('Otros servicios', ['estudio fotográfico', 'fotógrafo profesional', 'sesión de fotos', 'book fotográfico']),
 'Veterinarias':           ('Otros servicios', ['veterinaria', 'clínica veterinaria', 'consultorio veterinario', 'veterinario', 'urgencias veterinarias']),
 'Ópticas':                ('Otros servicios', ['óptica', 'óptica y lentes', 'óptica profesional', 'venta de lentes', 'optometrista']),
 'Talleres':               ('Otros servicios', ['taller mecánico', 'taller de autos', 'mecánica del automotor', 'taller multimarca', 'service de autos']),
 'Lavaderos y detailing':  ('Otros servicios', ['lavadero de autos', 'car wash', 'detailing automotor', 'lavadero y polarizado', 'pulido de autos']),
}

CIUDADES = ['Rosario', 'Buenos Aires', 'Córdoba', 'Mendoza', 'La Plata',
            'Mar del Plata', 'San Miguel de Tucumán', 'Salta', 'Santa Fe',
            'Neuquén', 'Bahía Blanca', 'Paraná']


def generar(destino='config/busquedas.xlsx'):
    filas = [{'Rubro': rubro, 'Nicho': nicho, 'Busqueda': kw, 'Ciudad': ciudad, 'Activo': 'Si'}
             for ciudad in CIUDADES
             for nicho, (rubro, kws) in NICHOS.items()
             for kw in kws]
    df = pd.DataFrame(filas).sort_values(['Rubro', 'Nicho', 'Ciudad', 'Busqueda'])
    df.to_excel(destino, index=False)
    kws = sum(len(k) for _, k in NICHOS.values())
    print(f"{len(df)} busquedas = {len(NICHOS)} nichos / {kws} keywords x {len(CIUDADES)} ciudades")
    for rubro, g in df.groupby('Rubro'):
        print(f"   {rubro}: {g['Nicho'].nunique()} nichos")
    return df


if __name__ == '__main__':
    generar()
