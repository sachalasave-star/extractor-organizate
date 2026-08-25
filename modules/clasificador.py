"""Decide el nicho de cada negocio por su categoria REAL de Google Maps, no por
la busqueda que lo encontro.

Buscar "veterinaria en Rosario" devuelve un club nocturno llamado "Animal", y
"gimnasio" devuelve escuelas de arte. La busqueda es una pista floja; la
categoria que Maps le asigna al negocio es el dato confiable.

Ademas descarta lo que no vende turnos: tiendas, hoteles, bares, laboratorios
industriales. Un local que vende articulos de kinesiologia no es un kinesiologo.
"""
import re
import unicodedata


def _norm(t):
    """minusculas y sin tildes: 'Kinesiólogo' -> 'kinesiologo'."""
    t = unicodedata.normalize('NFD', str(t or '').lower())
    return ''.join(c for c in t if unicodedata.category(c) != 'Mn')


# Categorias que NO venden turnos. Se evaluan primero que todo.
DESCARTAR = re.compile(_norm(
    r'tienda|venta al por menor|supermercado|kiosco|almacen|libreria|ferreteria|'
    r'hotel|hostel|apart|alojamiento|cabañas|'
    r'bar$|pub|club nocturno|discoteca|boliche|restaurante|pizzeria|cafeteria|'
    r'heladeria|panaderia|rotiseria|comida|'
    r'laboratorio farmaceutico|fabrica|mayorista|distribuidor|importador|'
    r'inmobiliaria|banco|aseguradora|iglesia|escuela primaria|escuela secundaria|'
    r'jardin de infantes|universidad|municipalidad|estacionamiento|'
    r'añadir|agregar|reclamar'          # botones de Maps que se colaron como categoria
))

# La primera regla que matchea gana: lo especifico va antes que lo generico
# ("peluqueria canina" antes que "peluqueria", "kinesiologo" antes que "gimnasio").
REGLAS = [
    ('Peluquería canina',       r'mascota|canin|grooming|peluquero de perros'),
    ('Veterinarias',            r'veterinari'),

    ('Barberías',               r'barber'),
    ('Manicura y pedicura',     r'manicura|pedicura|uñas|nail'),
    ('Cejas y pestañas',        r'pestañ|cejas|microblading'),
    ('Depilación definitiva',   r'depilacion|laser'),
    ('Salones de maquillaje',   r'maquilla'),
    ('Centros de masajes',      r'masaj|masajista'),
    ('Spas',                    r'\bspa\b|sauna|termas|banos termales'),
    ('Centros de bronceado',    r'broncead|solarium'),
    ('Centros de estética',     r'estetic|esteticista|cosmetolog|cosmiatr|'
                                r'belleza|depilacion con cera|tratamiento facial'),
    ('Peluquerías',             r'peluquer|salon de belleza|estilista'),

    ('Kinesiología',            r'kinesiolog|kinesiter'),
    ('Fisioterapia',            r'fisioterap|rehabilitacion|terapia fisica'),
    ('Quiropraxia y osteopatía', r'quiropract|osteopat'),
    ('Psicopedagogía',          r'psicopedagog'),
    ('Terapia ocupacional',     r'terapia ocupacional|terapeuta ocupacional'),
    ('Psicólogos',              r'psicolog|psicoanal|psicoterap|salud mental|psiquiatr'),
    ('Fonoaudiología',          r'fonoaudiolog|logoped|terapia del lenguaje'),
    ('Nutricionistas',          r'nutricion|dietist'),
    ('Odontología',             r'dentista|dental|odontolog|ortodon|implante'),
    ('Podología',               r'podolog|podiatr|pie'),
    ('Oftalmología',            r'oftalmolog|oculist'),
    ('Dermatología',            r'dermatolog'),
    ('Laboratorios de análisis', r'laboratorio de analisis|analisis clinico|'
                                r'laboratorio medico|extraccion de sangre'),
    ('Diagnóstico por imágenes', r'radiolog|ecograf|resonancia|tomograf|'
                                r'diagnostico por imagen|mamograf'),
    ('Medicina alternativa',    r'acupuntur|holistic|naturis|homeopat|reiki|ayurved|'
                                r'medicina alternativa|terapia alternativa'),
    ('Clínicas',                r'clinica|centro medico|hospital|sanatorio|consultorio medico|'
                                r'oficina medica|medico|obstetr|ginecolog|pediatr|traumatolog'),

    ('Yoga y pilates',          r'yoga|pilates'),
    ('Artes marciales',         r'artes marciales|karate|judo|taekwondo|jiu.?jitsu|'
                                r'boxeo|muay thai|kickboxing|mma'),
    ('Natación',                r'natacion|natatorio|pileta|piscina|nadar'),
    ('Escuelas de danza',       r'baile|danza|ballet|dance'),
    ('Escuelas de música',      r'musica|guitarra|piano|canto|conservatorio'),
    ('Escuelas de idiomas',     r'idioma|ingles|frances|italiano|portugues|traduccion'),
    ('Entrenadores personales', r'entrenador personal|personal trainer|preparador fisico'),
    ('Gimnasios',               r'gimnasio|\bgym\b|crossfit|centro deportivo|'
                                r'centro de fitness|entrenamiento funcional|musculacion'),
    ('Academias',               r'academia|escuela de|instituto'),

    ('Estudios de tatuajes',    r'tatuaj|tattoo|piercing'),
    ('Estudios de fotografía',  r'fotograf|foto estudio'),
    ('Ópticas',                 r'optic|lentes|anteojos|gafas'),
    ('Autoescuelas',            r'autoescuela|escuela de manejo|escuela de conducir|conduccion'),
    ('Lavaderos y detailing',   r'lavado de coche|lavadero|car wash|detailing|polarizado|encerado'),
    ('Talleres',                r'taller|mecanic|automotor|automovil|chapa y pintura|'
                                r'reparacion de auto|gomeria|neumatic'),
]

_COMPILADAS = [(nicho, re.compile(_norm(patron))) for nicho, patron in REGLAS]

# Categorias que media salud usa por igual: un psicologo, un nutricionista y un
# odontologo pueden figurar los tres como "Centro medico". No alcanzan para
# decidir el nicho, asi que se desempata por el nombre del negocio
# ("Juliana Vitale. Psicóloga" es psicologa aunque Maps diga Clinica ambulatoria).
GENERICAS = re.compile(_norm(
    r'^(clinica ambulatoria|centro medico|oficina medica|medico de familia|'
    r'centro de salud y bienestar|consultorio|profesional de|'
    r'centro de bienestar|especialista)'))

# "Centro de estetica" es el catch-all mas grande de Maps (6325 negocios, mas
# que ningun otro): una lashista, una manicura y una depiladora terminan las
# tres con esta categoria. A diferencia de las GENERICAS de arriba, ACA la
# categoria es un nicho valido en si mismo (a la mayoria no le sobra nada mas
# especifico), asi que el nombre solo la desplaza cuando dice algo puntual
# ("Lashes Rosario"); sin esa pista se queda en Centros de estetica, no cae a
# la busqueda que la encontro (esa SI es una pista floja, ver arriba).
AMPLIAS = {'Centros de estética'}


def _por_reglas(texto):
    for nicho, patron in _COMPILADAS:
        if patron.search(texto):
            return nicho
    return None


def _canonico(nicho):
    """Normaliza nombres viejos de nicho al actual, para que renombrar uno en la
    config no abra una pestaña duplicada con los registros ya scrapeados
    ('Manicura y pedicure' -> 'Manicura y pedicura')."""
    if not nicho:
        return None
    return _por_reglas(_norm(nicho)) or nicho


def clasificar(categoria, nicho_busqueda='', nombre=''):
    """Devuelve (nicho, motivo).

    nicho None = descartar. Si no se puede decidir se respeta el nicho de la
    busqueda en vez de tirar el negocio: mejor un dato mal ordenado que perdido.
    """
    cat = _norm(categoria)
    respaldo = _canonico(nicho_busqueda)
    if not cat:
        # Sin categoria, el nombre es lo unico que hay.
        return (_por_reglas(_norm(nombre)) or respaldo, 'sin categoria en Maps')
    if DESCARTAR.search(cat):
        return (None, f'no vende turnos ({categoria})')

    if GENERICAS.match(cat):
        por_nombre = _por_reglas(_norm(nombre))
        if por_nombre:
            return (por_nombre, f'categoria generica ({categoria}), resuelto por el nombre')
        return (respaldo, f'categoria generica ({categoria})')

    nicho = _por_reglas(cat)
    if nicho:
        if nicho in AMPLIAS:
            por_nombre = _por_reglas(_norm(nombre))
            if por_nombre and por_nombre != nicho:
                return (por_nombre, f'categoria amplia ({categoria}), resuelto por el nombre')
        return (nicho, 'ok' if nicho == nicho_busqueda else f'reclasificado desde {nicho_busqueda}')
    return (_por_reglas(_norm(nombre)) or respaldo, f'categoria sin regla ({categoria})')


def demo():
    casos = [
        # (categoria, nicho de la busqueda, nombre, nicho esperado)
        ('Club nocturno',        'Veterinarias',  'Animal Rosario',      None),
        ('Tienda de deportes',   'Gimnasios',     'Kineshop',            None),
        ('Hotel',                'Spas',          'Puerto Norte Hotel',  None),
        ('Añadir categoría',     'Peluquerías',   'Mirarte Estética',    None),
        ('Universidad',          'Psicólogos',    'Facultad de Psi',     None),
        ('Tienda de suministros para peluquería', 'Barberías', 'IDC',    None),

        ('Veterinario',          'Veterinarias',  'Vet Central',         'Veterinarias'),
        ('Farmacia veterinaria', 'Veterinarias',  'Vet Farma',           'Veterinarias'),
        ('Masajista',            'Spas',          'Manos Sanas',         'Centros de masajes'),
        ('Centro de pilates',    'Gimnasios',     'Studio P',            'Yoga y pilates'),
        ('Kinesiólogo',          'Spas',          'Centro Kine',         'Kinesiología'),
        ('Academia de baile',    'Gimnasios',     'Ritmo',               'Escuelas de danza'),
        ('Peluquero de mascotas', 'Peluquerías',  'Pet Style',           'Peluquería canina'),
        ('Escuela de boxeo',     'Gimnasios',     'Box Club',            'Artes marciales'),
        ('Psicoanalista',        'Psicólogos',    'Lic. Perez',          'Psicólogos'),

        # Categoria generica: manda el nombre, no la categoria
        ('Clínica ambulatoria',  'Psicólogos', 'Juliana Vitale. Psicóloga', 'Psicólogos'),
        ('Centro médico',        'Psicólogos', 'Centro John Bowlby',    'Psicólogos'),
        ('Oficina médica',       'Odontología', 'Consultorio Dental Sur', 'Odontología'),
        # Generica y el nombre no aporta: se respeta la busqueda
        ('Centro médico',        'Psicólogos',  'Arkhé Salud Integral',  'Psicólogos'),
        # Sin categoria: no se tira, se usa el nombre o la busqueda
        ('',                     'Barberías',   'Corte Fino',            'Barberías'),
        ('',                     'Peluquerías', 'Barbería Don Juan',     'Barberías'),

        # Centro de estetica es el catch-all mas grande de Maps (6325 negocios):
        # el nombre desplaza la categoria si es especifico...
        ('Centro de estética',   'Cejas y pestañas', 'Pestañas Bea',     'Cejas y pestañas'),
        # ...pero sin nombre especifico se queda en Centros de estetica, no cae
        # a la busqueda (a diferencia de una GENERICA como Centro medico).
        ('Centro de estética',   'Cejas y pestañas', 'Bella',            'Centros de estética'),
    ]
    for categoria, busqueda, nombre, esperado in casos:
        obtenido, motivo = clasificar(categoria, busqueda, nombre)
        assert obtenido == esperado, \
            f'{nombre!r} [{categoria}] buscado como {busqueda}: esperaba {esperado!r}, dio {obtenido!r} [{motivo}]'
    print(f'OK  clasificador: {len(casos)} casos')


if __name__ == '__main__':
    demo()
