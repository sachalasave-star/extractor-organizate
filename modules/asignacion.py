"""Reparto de leads entre vendedores y regla de reposicion.

La idea es que un vendedor no vea la base: recibe un lote y trabaja eso. Cuando
lo termina de verdad, le llega otro. "De verdad" tiene una definicion estricta,
ver `puede_reponer`.

Todo lo de este modulo es logica pura sobre listas de diccionarios, sin tocar
Google. Lo que habla con Sheets vive en repartir_leads.py.
"""
from modules.telefono import canonico

LOTE = 30              # cuantos leads tiene un vendedor a la vez
MINIMO_HABLADOS = 20   # de esos, con cuantos tiene que haber hablado para reponer

SIN_CONTACTAR = 'Sin contactar'

# Estados que por si solos prueban que hubo una conversacion: no se llega a
# "Le interesó" sin que alguien te atienda.
ESTADOS_CON_CHARLA = {'Le interesó', 'Demo iniciada', 'Cliente activo'}

# Motivos que dicen explicitamente que NO se hablo con el negocio. El resto de
# los motivos ("Precio alto", "Ya tiene sistema", "Pidió que llame después",
# "No es el rubro") solo se pueden saber hablando.
MOTIVOS_SIN_CHARLA = {'No atiende', 'Número equivocado', 'Cerró'}

PRIORIDADES = ('A', 'B', 'C')


def contactado(estado):
    """Se intento el contacto: cualquier cosa que no sea 'Sin contactar'."""
    e = (estado or '').strip()
    return bool(e) and e != SIN_CONTACTAR


def hablo(estado, motivo):
    """Hubo una conversacion con una persona del negocio.

    Pide evidencia POSITIVA a proposito. Un lead contactado sin motivo cargado
    y sin estado avanzado NO cuenta como hablado: si contara, alcanzaria con
    marcar los 30 como "Volver a llamar" y dejar el motivo vacio para
    auto-reponerse sin haber hablado con nadie, que es justo lo que la regla
    quiere evitar.
    """
    e, m = (estado or '').strip(), (motivo or '').strip()
    if e in ESTADOS_CON_CHARLA:
        return True
    if not contactado(e):
        return False
    return bool(m) and m not in MOTIVOS_SIN_CHARLA


def resumen_lote(filas):
    """{asignados, contactados, hablados, sin_contactar} de un lote.

    filas: [{'estado':..., 'motivo':...}], como viene del archivo del vendedor.
    """
    contactados = sum(1 for f in filas if contactado(f.get('estado')))
    hablados = sum(1 for f in filas if hablo(f.get('estado'), f.get('motivo')))
    return {'asignados': len(filas), 'contactados': contactados,
            'hablados': hablados, 'sin_contactar': len(filas) - contactados}


def puede_reponer(filas, lote=LOTE, minimo_hablados=MINIMO_HABLADOS):
    """(bool, motivo) -> si corresponde darle un lote nuevo, y por que no.

    Las dos condiciones se piden juntas:
      1. contacto TODOS los que tiene, sin importar si le respondieron
      2. hablo con al menos `minimo_hablados` negocios

    La primera sola no alcanza: se cumple marcando todo como "No atiende" sin
    levantar el telefono. La segunda es la que obliga a que haya trabajo real.
    """
    r = resumen_lote(filas)
    if not filas:
        return True, 'no tiene leads asignados'
    if r['sin_contactar']:
        return False, f"le quedan {r['sin_contactar']} sin contactar"
    if r['hablados'] < minimo_hablados:
        return False, (f"contacto los {r['contactados']} pero hablo con "
                       f"{r['hablados']}, necesita {minimo_hablados}")
    return True, f"contacto {r['contactados']} y hablo con {r['hablados']}"


def _clave(lead):
    """La misma clave de dedup que usa el resto del sistema: el telefono.

    Cuando no hay telefono cae al nombre del negocio mas la ciudad, que es lo
    unico que queda. Sin clave el lead no se reparte: es preferible saltearlo
    antes que darle el mismo negocio a dos vendedores.
    """
    c = canonico(lead.get('telefono', ''))
    if c:
        return c
    nombre = (lead.get('negocio') or '').strip().lower()
    ciudad = (lead.get('ciudad') or '').strip().lower()
    return f'{nombre}|{ciudad}' if nombre else ''


def elegir_lote(disponibles, cantidad, ya_repartidos=None):
    """Elige `cantidad` leads lo mas parecidos entre si que se pueda.

    Agrupa por (nicho, ciudad) y arranca por el grupo con mas leads de
    prioridad A. Motivo: el equipo vende por telefono, y treinta peluquerias
    de Rosario se llaman con el mismo discurso y las mismas objeciones.
    Treinta negocios de rubros y ciudades distintas obligan a improvisar en
    cada llamada.

    Dentro del grupo se sirve A, despues B, despues C: la calificacion ya
    estimo cual vale mas la pena (ver modules/calificar.py).

    `ya_repartidos` es el conjunto de claves que ya tiene alguien. Un mismo
    negocio no puede caer en dos vendedores ni aunque figure duplicado en dos
    hojas de la base.
    """
    vistos = set(ya_repartidos or ())
    grupos = {}
    for lead in disponibles:
        clave = _clave(lead)
        if not clave or clave in vistos:
            continue
        vistos.add(clave)          # dedup dentro de la propia tanda tambien
        grupo = ((lead.get('nicho') or '').strip(), (lead.get('ciudad') or '').strip())
        grupos.setdefault(grupo, []).append(lead)

    def orden_grupo(item):
        _, leads = item
        aes = sum(1 for l in leads if (l.get('prioridad') or '').strip().upper() == 'A')
        return (-aes, -len(leads))

    def orden_lead(lead):
        p = (lead.get('prioridad') or '').strip().upper()
        return PRIORIDADES.index(p) if p in PRIORIDADES else len(PRIORIDADES)

    elegidos = []
    for _, leads in sorted(grupos.items(), key=orden_grupo):
        for lead in sorted(leads, key=orden_lead):
            if len(elegidos) >= cantidad:
                return elegidos
            elegidos.append(lead)
    return elegidos


def demo():
    # --- contactado / hablo -------------------------------------------------
    assert not contactado('Sin contactar') and not contactado('') and not contactado(None)
    assert contactado('No interesado') and contactado('Cliente activo')

    assert hablo('Le interesó', '')          # el estado ya lo prueba
    assert hablo('Demo iniciada', None)
    assert hablo('No interesado', 'Precio alto')
    assert hablo('Volver a llamar', 'Pidió que llame después')
    assert not hablo('Sin contactar', '')
    assert not hablo('No interesado', 'No atiende')
    assert not hablo('No interesado', 'Número equivocado')
    # El caso que la regla tiene que frenar: contactado, sin motivo, sin avance.
    assert not hablo('Volver a llamar', ''), 'sin motivo no se puede probar que hablo'
    print('OK  hablo: pide evidencia, no alcanza con marcar el estado')

    # --- reposicion ---------------------------------------------------------
    lote_a_medias = ([{'estado': 'No interesado', 'motivo': 'Precio alto'}] * 25 +
                     [{'estado': 'Sin contactar', 'motivo': ''}] * 5)
    ok, por_que = puede_reponer(lote_a_medias)
    assert not ok and '5 sin contactar' in por_que, por_que

    # Contacto los 30 pero a todos les puso "No atiende": no hablo con nadie.
    solo_intentos = [{'estado': 'No interesado', 'motivo': 'No atiende'}] * 30
    ok, por_que = puede_reponer(solo_intentos)
    assert not ok and 'hablo con 0' in por_que, por_que

    # 19 hablados: falta uno. La regla es 20, no "casi 20".
    casi = ([{'estado': 'No interesado', 'motivo': 'Precio alto'}] * 19 +
            [{'estado': 'No interesado', 'motivo': 'No atiende'}] * 11)
    ok, por_que = puede_reponer(casi)
    assert not ok and 'hablo con 19' in por_que, por_que

    # Las dos condiciones juntas.
    completo = ([{'estado': 'No interesado', 'motivo': 'Precio alto'}] * 20 +
                [{'estado': 'No interesado', 'motivo': 'No atiende'}] * 10)
    ok, por_que = puede_reponer(completo)
    assert ok, por_que
    assert puede_reponer([])[0], 'el vendedor nuevo tiene que recibir su primer lote'
    print('OK  puede_reponer: exige contactar todos Y hablar con 20, juntas')

    # --- eleccion del lote --------------------------------------------------
    disponibles = (
        [{'negocio': f'Peluqueria {i}', 'telefono': f'0341 100-{i:04d}',
          'nicho': 'Peluquerías', 'ciudad': 'Rosario', 'prioridad': 'A'} for i in range(40)] +
        [{'negocio': f'Gimnasio {i}', 'telefono': f'0341 200-{i:04d}',
          'nicho': 'Gimnasios', 'ciudad': 'Rosario', 'prioridad': 'C'} for i in range(50)] +
        [{'negocio': f'Spa {i}', 'telefono': f'011 300-{i:04d}',
          'nicho': 'Spas', 'ciudad': 'CABA', 'prioridad': 'B'} for i in range(5)]
    )
    lote = elegir_lote(disponibles, 30)
    assert len(lote) == 30
    # Gimnasios tiene mas leads, pero Peluquerias tiene mas prioridad A: gana.
    assert {l['nicho'] for l in lote} == {'Peluquerías'}, {l['nicho'] for l in lote}
    assert {l['ciudad'] for l in lote} == {'Rosario'}
    print('OK  elegir_lote: un solo nicho y una sola ciudad, priorizando los A')

    # Nadie recibe dos veces el mismo negocio, ni cuando esta duplicado.
    duplicados = [
        {'negocio': 'Barberia Uno', 'telefono': '0341 353-9510', 'nicho': 'B', 'ciudad': 'R'},
        {'negocio': 'Barberia Uno', 'telefono': '+54 341 353 9510', 'nicho': 'B', 'ciudad': 'R'},
        {'negocio': 'Barberia Dos', 'telefono': '0341 353-9511', 'nicho': 'B', 'ciudad': 'R'},
    ]
    assert len(elegir_lote(duplicados, 10)) == 2, 'no dedupeo el mismo telefono escrito distinto'

    # Lo que ya tiene otro vendedor no se vuelve a repartir.
    otro = elegir_lote(duplicados, 10, ya_repartidos={canonico('03413539510')})
    assert [l['negocio'] for l in otro] == ['Barberia Dos'], otro

    # Sin telefono cae al nombre + ciudad, y sin nombre se saltea.
    raros = [{'negocio': 'Sin tel', 'telefono': '', 'ciudad': 'Rosario', 'nicho': 'X'},
             {'negocio': '', 'telefono': '', 'ciudad': 'Rosario', 'nicho': 'X'}]
    assert [l['negocio'] for l in elegir_lote(raros, 10)] == ['Sin tel']

    # Si hay menos de los pedidos, devuelve los que hay y no rellena con nada.
    assert len(elegir_lote(disponibles, 500)) == 95
    print('OK  elegir_lote: dedup por telefono y respeta lo ya repartido')


if __name__ == '__main__':
    demo()
