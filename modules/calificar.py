"""Decide a quien conviene llamar primero, con lo que Google Maps de verdad da.

Devuelve una letra y el motivo en castellano. La letra ordena la lista; el
motivo es para que el vendedor pueda estar en desacuerdo. Un score 0-100 seria
falsa precision: no tenemos con que sostener el decimal.

Que senal sirve, medido sobre los 28744 negocios con telefono de la base:

  nicho apagado   29%   la mas fuerte y la mas barata. Gimnasios, Academias,
                        Artes marciales, Yoga, Natacion, Danza e Idiomas ya se
                        habian apagado en las busquedas por no vender turnos,
                        pero los scrapeados antes siguen en la planilla.
  web             52%   parte la base al medio. Es la unica senal de madurez
                        digital que tenemos: al que tiene web le podes mandar
                        el link de la demo y sabe que hacer con el.
  celular         12%   telefono de 13 digitos (area + 15 + numero) contra 11
                        del fijo. Un celular suele ser el profesional solo
                        atendiendo de su casa; el fijo es un local con puerta.
  resenas          5%   la mejor de todas y la que casi no tenemos. Es lo unico
                        que mide cuanta gente entra de verdad. Google la
                        devuelve solo en una de sus dos variantes de ficha, y
                        los 28744 ya scrapeados no se vuelven a abrir.
  rating          88%   inservible sola. 10184 de 21530 estan entre 4,9 y 5,0,
                        que es la firma de tener tres resenas, no de ser bueno.
                        Solo se usa para bajar a los que estan muy mal puntuados
                        CON resenas suficientes para que la nota signifique algo.
"""
import re

RESENAS_VOLUMEN = 20     # de ahi para arriba el negocio tiene agenda de verdad
RESENAS_MINIMAS = 10     # menos que esto y el rating no dice nada
RATING_MALO = 3.5


def es_celular(telefono):
    """13 digitos = area + 15 + numero. El fijo argentino tiene 11.

    En la base el corte es limpio: 87% cae en 11 digitos y 12% en 13, sin nada
    en el medio. Por eso alcanza el largo y no hay que parsear codigos de area,
    que van de 2 a 4 digitos segun la ciudad.
    """
    return len(re.sub(r'\D', '', telefono or '')) >= 13


def _entero(x):
    x = (x or '').strip()
    return int(x) if x.isdigit() else 0


def _rating(x):
    try:
        return float((x or '').replace(',', '.'))
    except ValueError:
        return 0.0


def calificar(prioridad_nicho, telefono='', web='', resenas='', rating=''):
    """-> ('A'|'B'|'C', motivo). A se llama primero, C es lo que sobra."""
    if prioridad_nicho >= 3:
        return 'C', 'el nicho no vende turnos'

    n = _entero(resenas)
    r = _rating(rating)
    cel, hay_web = es_celular(telefono), bool((web or '').strip())

    # Mal puntuado y con resenas suficientes para creerle: el vendedor pierde el
    # tiempo y ademas es un cliente que se va a quejar. Con 3 resenas no cuenta.
    if n >= RESENAS_MINIMAS and r and r < RATING_MALO:
        return 'C', f'{r:.1f} de puntaje con {n} reseñas'

    if cel and not hay_web:
        return 'C', 'celular y sin web: suele ser unipersonal'

    fuerte = []
    if n >= RESENAS_VOLUMEN:
        fuerte.append(f'{n} reseñas')
    if hay_web and not cel:
        fuerte.append('línea fija y con web')

    if prioridad_nicho == 1 and fuerte:
        return 'A', ', '.join(fuerte)

    razones = []
    if n:
        razones.append(f'{n} reseñas')
    if hay_web:
        razones.append('con web')
    if not cel:
        razones.append('línea fija')
    if prioridad_nicho == 2:
        razones.append('el nicho encaja con fricción')
    return 'B', ', '.join(razones) or 'sin señales para ordenarlo'


def demo():
    # El nicho apagado gana sobre todo lo demas: ya decidiste que no vende
    # turnos, y no hay cantidad de resenas que lo cambie.
    assert calificar(3, '03411234567', 'x.com', '500')[0] == 'C'
    assert 'no vende turnos' in calificar(3, '03411234567', 'x.com', '500')[1]

    # A: prioridad 1 con volumen probado, o con local y web.
    assert calificar(1, '03411234567', '', '162')[0] == 'A'
    assert calificar(1, '03411234567', 'x.com', '')[0] == 'A'
    assert '162 reseñas' in calificar(1, '03411234567', '', '162')[1]

    # Fijo sin web y sin resenas no alcanza para A: no sabemos nada de el.
    assert calificar(1, '03411234567', '', '')[0] == 'B'
    # Prioridad 2 nunca es A aunque tenga todo: el nicho compite con el turno.
    assert calificar(2, '03411234567', 'x.com', '300')[0] == 'B'

    # Celular sin web: el unipersonal desde la casa.
    assert calificar(1, '0111512345678', '', '')[0] == 'C'
    # ...pero con web deja de ser un cualquiera y vuelve a la lista.
    assert calificar(1, '0111512345678', 'x.com', '')[0] == 'B'

    # Mal puntuado con resenas suficientes para creerle.
    assert calificar(1, '03411234567', 'x.com', '80', '2,4')[0] == 'C'
    # Con 3 resenas un 2,4 no significa nada: no lo bajamos por eso.
    assert calificar(1, '03411234567', 'x.com', '3', '2,4')[0] == 'A'

    assert es_celular('0111526205229') and es_celular('0230154356392')
    assert not es_celular('03413539510') and not es_celular('02204863049')
    assert not es_celular('') and not es_celular(None)

    # El motivo nunca puede salir vacio: es lo unico que el vendedor puede leer
    # para estar en desacuerdo con la letra.
    for p in (1, 2, 3):
        for tel in ('03411234567', '0111512345678', ''):
            assert calificar(p, tel, '', '')[1].strip(), f'motivo vacio en {p}/{tel}'

    print('OK  calificar: A/B/C con motivo, y el nicho apagado manda')


if __name__ == '__main__':
    demo()
