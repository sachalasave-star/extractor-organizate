# Extractor Organizate

Scraper de Google Maps que junta **nombre, teléfono y tipo de negocio** de locales,
apuntando a nichos que trabajan con turnos (barberías, psicólogos, veterinarias...).

## Uso

```bash
pip install -r requirements.txt
python -m playwright install chromium

python verificar.py    # ~1 min: falla ruidoso si Google cambió los selectores
python scraper.py      # corre todas las búsquedas activas
```

Cortá con Ctrl+C cuando quieras: guarda después de cada búsqueda y al retomar
saltea las que ya hizo.

## Qué configurar

`config/busquedas.xlsx` — una fila por búsqueda:

| Nicho | Busqueda | Ciudad | Activo |
|---|---|---|---|
| Barberías | Barbería | Rosario | Si |
| Psicólogos | Psicólogo | Córdoba | Si |

El término que se busca es `"{Busqueda} en {Ciudad}"`. Poné `No` en Activo para saltear una fila.

`config/configuracion.json` — `headless` y `max_resultados_por_busqueda`.

## Qué genera

- **`output/negocios.db`** — la base maestra (SQLite). Es la fuente de verdad.
- **`output/Organizate.xlsx`** — export, una hoja por nicho. Se **regenera entero**
  en cada corrida: si lo editás a mano, se pierde.

Al Excel solo van los negocios **con teléfono**. Los que no tienen quedan en la base:
si en una pasada posterior aparece el número, entran solos.

## Cómo está armado

```
scraper.py          orquesta: lee las búsquedas, recorre, guarda
verificar.py        check de humo, correlo antes de una tanda larga
modules/
  buscador.py       navega a la url de búsqueda
  scroll.py         baja el listado hasta el final
  extractor.py      abre la ficha de cada negocio y lee los datos
  db.py             SQLite, dedup por place_id
  excel.py          regenera el xlsx desde la base
```

### Detalles que costaron encontrar

Si algo se rompe, empezá por acá — son los tres puntos frágiles:

1. **No se busca tipeando en el input.** Al apretar Enter, Maps toma la sugerencia
   resaltada del autocompletado, no lo tipeado: "Barbería en Rosario" terminaba en el
   mapa de la ciudad, con cero resultados. Se navega directo a `/maps/search/`.

2. **La ficha de un negocio NO es `div[role="dialog"]`.** Es un `div[role="main"]` con
   el nombre del negocio como `aria-label` (el listado es otro `role="main"`, con
   aria-label vacío). Esperar `role="dialog"` daba timeout y teléfono vacío en el 100%.

3. **El panel se busca con `get_by_role`, no con un selector CSS.** Un nombre con tilde
   no se puede escapar a mano para CSS (`ó` no es escape válido; `\f3 ` sí). Con
   CSS fallaba solo en los nombres acentuados: 5/5 en barberías, 9/82 en psicólogos.
   Por eso `verificar.py` prueba con psicólogos, cuyos nombres llevan tilde.

El teléfono, la dirección y la categoría salen de la ficha por `data-item-id`
(`phone:tel:`, `address`, `authority`), que no depende del idioma ni de las clases
ofuscadas de Google.

## Límites conocidos

- Google Maps corta cerca de **120 resultados por búsqueda**. Para más volumen hay que
  multiplicar búsquedas por ciudad, no esperar más de una sola.
- Alrededor del **20% de los negocios no publica teléfono** en Maps. No hay nada que
  extraer ahí.
