"""Un telefono argentino, escrito de cualquier forma, en una sola clave.

El mismo negocio aparece como "0341 353-9510" en Maps, "+54 341 353 9510" en
una web y "3413539510" tipeado a mano. Para cruzar la planilla con los clientes
que pagan hace falta que las tres den lo mismo.

La clave es CARACTERISTICA + NUMERO, 10 digitos, sin pais, sin el 9 de celular
y sin el 15. Se tiran a proposito:

  +54 / 54    el pais. No distingue nada dentro de Argentina.
  9           marca celular en formato internacional (+54 9 11 ...).
  15          marca celular en formato local (011 15-...). Es el MISMO dato
              que el 9, escrito de la otra forma, y nunca aparecen los dos.

Por que tirar el 9 y el 15 en vez de normalizarlos a uno: porque cada lado
elige uno distinto y ninguno esta mal. Maps devuelve el 15, una web que pide
"+54" suele guardar el 9, y un formulario que pide "sin el 15" no guarda
ninguno de los dos. Tirandolos, los tres cruzan. Conservandolos hay que
adivinar cual uso cada lado, que es exactamente el bug que no queremos.

Se pierde poder distinguir celular de fijo a partir de la clave, pero eso ya se
sabe antes de normalizar (ver es_celular en calificar.py) y no hace falta para
cruzar.
"""
import re

# Caracteristicas de 2 digitos. Todas las demas son de 3 o 4, y ninguna de
# esas empieza con 11, asi que alcanza con mirar el arranque.
AREA_2 = ('11',)


def canonico(telefono):
    """-> 10 digitos (caracteristica + numero), o '' si no es un telefono.

    Es la clave para cruzar: dos escrituras del mismo numero dan el mismo
    string, sin importar de que lado vengan.
    """
    d = re.sub(r'\D', '', telefono or '')
    if not d:
        return ''

    if d.startswith('54'):
        d = d[2:]
        # El 9 de celular va pegado al pais: +54 9 11 2620-5229. Solo se saca
        # si lo que queda alcanza para un numero, porque hay caracteristicas
        # que empiezan con 9 (2901 Ushuaia no, pero 9 suelto si aparece).
        if d.startswith('9') and len(d) >= 11:
            d = d[1:]
    d = d.lstrip('0')

    # El 15 va DESPUES de la caracteristica, no al principio: 011 15-xxxx-xxxx,
    # 0341 15-xxx-xxxx. Un numero completo con 15 tiene 12 digitos (area 2/3/4
    # + 15 + numero 8/7/6), asi que se busca el 15 donde termina cada largo de
    # caracteristica posible. AREA_2 primero: 11 15 xxxx xxxx empieza con "1115"
    # y probar del mas corto al mas largo es lo que lo desambigua.
    if len(d) == 12:
        for n in (2, 3, 4):
            if d[n:n + 2] == '15':
                d = d[:n] + d[n + 2:]
                break

    return d if len(d) == 10 else ''


def demo():
    # El caso que importa: el MISMO negocio escrito como lo escribe cada lado.
    fijo = [
        '0341 353-9510',        # como lo muestra Maps
        '03413539510',          # como lo guarda la base
        '+54 341 353 9510',     # como lo guardaria una web con +54
        '+543413539510',
        '3413539510',           # tipeado a mano, sin el 0
        '(0341) 353-9510',
    ]
    assert len({canonico(t) for t in fijo}) == 1, {t: canonico(t) for t in fijo}
    assert canonico(fijo[0]) == '3413539510', canonico(fijo[0])

    celu = [
        '011 15-2620-5229',     # local con 15
        '0111526205229',        # como lo guarda la base
        '+54 9 11 2620-5229',   # internacional con 9
        '+5491126205229',
        '+54 11 2620 5229',     # lo que propusiste: sin 15 y sin 9
        '1126205229',
    ]
    assert len({canonico(t) for t in celu}) == 1, {t: canonico(t) for t in celu}
    assert canonico(celu[0]) == '1126205229', canonico(celu[0])

    # Celular del interior: caracteristica de 3 y el 15 en el medio.
    inter = ['0341 15-577-3501', '03411557735 01', '+5493415773501', '3415773501']
    assert len({canonico(t) for t in inter}) == 1, {t: canonico(t) for t in inter}

    # Caracteristica de 4 digitos (0220 Merlo), con y sin 15.
    assert canonico('02204863049') == '2204863049'
    assert canonico('0220 15-486-3049') == '2204863049'

    # Dos negocios distintos NO pueden colisionar.
    assert canonico('03413539510') != canonico('03413539511')

    # Basura: mejor vacio que un match equivocado, que paga una comision mal.
    for malo in ('', None, 'sin telefono', '123', '+34 91 123 4567', '0000000000'):
        c = canonico(malo)
        assert c == '' or len(c) == 10, f'{malo!r} -> {c!r}'
    assert canonico('+34 91 123 4567') != canonico('0341 123-4567')

    print('OK  telefono: 6 escrituras del fijo y 6 del celular dan una sola clave')


if __name__ == '__main__':
    demo()
