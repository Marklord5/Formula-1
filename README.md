# Reporte de Acreditaciones · Gran Premio de México 2026

Reporte web de una sola página (`index.html`) que lee **directamente los Excel** de la carpeta `data/` cada vez que se abre. No hay que compilar nada.

```
index.html               ← el reporte
data/corte-final.xlsx    ← CORTE FINAL F1-2026
data/corte-diario.xlsx   ← CORTE DIARIO DE ACREDITACIONES 2026
data/devolucion.xlsx     ← DEVOLUCIÓN F1 2026
vendor/xlsx.full.min.js  ← lector de Excel de respaldo (si el CDN no responde)
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

Para **probar Excel nuevos sin subirlos**, usa el botón **Cargar Excel** en la página y elige uno o varios archivos: detecta solo cuál es cuál (corte final, diario o devolución) y redibuja todo. No modifica nada en el repositorio.

## Ver en local

Da doble clic en **`Abrir reporte.command`**: levanta un servidor local y abre el reporte en `http://localhost:8765`, leyendo los Excel de `data/` automáticamente. Para detenerlo, cierra la ventana de Terminal.

(Si abres `index.html` directo con doble clic, el navegador bloquea la lectura de `data/` y la página te indica usar el lanzador.)

## Qué toma de cada Excel

| Sección | Fuente |
|---|---|
| KPIs, utilización, tabla consolidada, dona | `corte-final.xlsx` → tabla CATEGORIA / PRODUCIDO / … / TOTAL SOBRANTE |
| Columna de folios | `corte-final.xlsx` → hojas FOLIOS OPERACION y FOLIOS STAFF HOSPITALIDADES |
| Genéricos activados | `corte-final.xlsx` → bloque GÉNERICOS ACTIVADOS |
| Curva acumulada, módulos, últimos movimientos | `corte-diario.xlsx` → todos los bloques SOLICITANTE / FECHA / CANTIDAD de todas las hojas (sin las filas de BODEGA) |
| Devoluciones | `devolucion.xlsx` → todas las hojas con CANTIDAD y CATEGORIA o TIPO DE BRAZALETE |

Las tablas se encuentran **por el texto de sus encabezados**, no por posición; puedes agregar filas, categorías u hojas nuevas mientras los encabezados se llamen igual.

## Ajustes

Al inicio del `<script>` de `index.html`:

- `FILES`: nombres de los archivos en `data/`.
- `GROUPS`: a qué grupo pertenece cada categoría (dona y filtro). Se evalúan en orden; lo que no coincida cae en "Otros".
- `PAGE`: filas por página de la tabla.
