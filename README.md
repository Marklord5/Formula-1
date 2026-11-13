# Reporte de Acreditaciones · Gran Premio de México 2026

Reporte web de una sola página (`index.html`) que lee **directamente los Excel** de la carpeta `data/` cada vez que se abre. No hay que compilar nada.

```
index.html               ← estructura de la página (HTML + plantillas)
css/styles.css           ← todo el estilo
js/config.js             ← archivos, tipos de acreditación y glosario (lo único que se ajusta)
js/main.js               ← carga de los Excel, estado y eventos
js/excel.js              ← lectura de hojas y tablas; reglas de BODEGA y extras
js/overview.js           ← totales, resumen por tipo y pestañas
js/charts.js             ← gráficas de cada Excel
js/sheets.js             ← tarjetas de hoja, búsqueda y desglose
js/utils.js              ← formato, fechas y helpers del DOM
data/corte-final.xlsx    ← CORTE FINAL F1-2026
data/corte-diario.xlsx   ← CORTE DIARIO DE ACREDITACIONES 2026
data/devolucion.xlsx     ← DEVOLUCIÓN F1 2026
vendor/xlsx.full.min.js  ← lector de Excel (SheetJS)
```

## Publicar en GitHub Pages

1. Sube esta carpeta a un repositorio.
2. En GitHub: **Settings → Pages → Build and deployment → Source: Deploy from a branch**, rama `main`, carpeta `/ (root)`.
3. En 1–2 minutos queda en `https://<usuario>.github.io/<repositorio>/`.

> ⚠️ En un repositorio público, cualquiera con el link puede ver el reporte **y descargar los Excel** de `data/` (incluyen nombres de solicitantes). Si no debe ser público, usa un repositorio privado con Pages privado (GitHub Enterprise) u otro hosting con acceso restringido.

## Actualizar los datos

1. Reemplaza el archivo en `data/` **con el mismo nombre** (por ejemplo, el nuevo corte diario → `data/corte-diario.xlsx`).
2. `git add data && git commit -m "Actualiza corte" && git push`
3. Recarga la página. Siempre lee la versión más reciente (no usa caché).

## Ver en local

Da doble clic en **`Abrir reporte.command`**: levanta un servidor local y abre el reporte en `http://localhost:8765`, leyendo los Excel de `data/` automáticamente. Para detenerlo, cierra la ventana de Terminal.

(Si abres `index.html` directo con doble clic, el navegador bloquea la lectura de `data/` y la página te indica usar el lanzador.)

## Cómo está organizada la página

0. **Descargar los Excel**: un botón por archivo en el encabezado (sale de `FILES`; si agregas un Excel, aparece solo).
1. **Totales** arriba: PRODUCIDO, TOTAL ACREDITACIONES y TOTAL SOBRANTE (suma de esas columnas del Corte Final) y Devueltos (suma de CANTIDAD del Excel de devolución).
2. **Por tipo de acreditación**: el Corte Final separado en Montaje y Desmontaje, Hospitalidades y Operación.
3. **Paso 1 · Elige el Excel**: Corte Final, Corte Diario o Devolución.
4. **Paso 2 · Gráficas** del Excel elegido:
   - Corte Final → entregado vs sobrante, total y por categoría (agrupadas por tipo).
   - Corte Diario → entregas por día (sin filas de BODEGA) y quién recibió más.
   - Devolución → devueltos por hoja y lo que más se devolvió.
5. **Paso 3 · Hojas**: una tarjeta por hoja, agrupadas por tipo. Al darle clic se abre el desglose: las mismas tablas y columnas del Excel, con el título de cada bloque, lo producido, los folios y su fila de TOTALES. El buscador encuentra nombres, folios o áreas en todas las hojas.

### Reglas de lectura

- **BODEGA**: un movimiento con SOLICITANTE = BODEGA es stock que se quedó en bodega; no cuenta como acreditado. En el desglose esas filas se ven atenuadas.
- **Extras / 2ª producción**: un bloque cuyo título dice EXTRA, SEGUNDA o TERCERA PRODUCCIÓN, o "<bloque> 2" cuando existe "<bloque>" en la misma hoja (p. ej. LIMITADO ZONA 2), se marca como producción adicional.
- **Tipo de acreditación**: por el nombre de la categoría u hoja (o de sus bloques, p. ej. la hoja PADDOCK); los patrones están en `GROUPS` de `js/config.js`.

Las tablas se encuentran **por el texto de sus encabezados** (SOLICITANTE, FECHA, CANTIDAD, CATEGORIA, PRODUCIDO…), no por posición: puedes agregar filas, bloques u hojas nuevas y aparecen solas. No cambies el texto de esos encabezados.

## Ajustes

Todo en `js/config.js`:

- `FILES`: nombres de los archivos en `data/`.
- `GROUPS`: tipos de acreditación y qué categorías caen en cada uno (el último, Operación, recibe todo lo demás).
- `GLOSSARY`: lo que significa cada columna (aparece al pasar el mouse por el encabezado).
