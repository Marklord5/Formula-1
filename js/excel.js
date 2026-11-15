// Convierte cada hoja de Excel en tablas: { name, tables: [{ cols, rows, total, meta }], raw }
// Las tablas se encuentran por el texto de sus encabezados, no por posición.
import { GROUPS } from './config.js';
import { norm, blank, nothing, clean, num, fmt, shortDate, parseDate } from './utils.js';

// Una fila con 3 o más de estos textos es un encabezado de tabla
const HEAD = /^(NO\.?( SOLICITUD)?|ID|SOLICITANTE|RESPONSABLE|AREA|FECHA|DIA|CATEGORIA|TIPO( DE BRAZALETE)?|FOLIOS?( INICIAL| FINAL)?|F\. ?(INICIAL|FINAL)|CANTIDAD( DEVUELTA)?|EXISTENCIA EN BODEGA|OBSERVACIONES|(TOTAL )?PRODUCIDO|ENTREGADO PAQUETE|ACREDITADO MODULO|TOTAL ACREDITACIONES|SOBRANTES?( BODEGA| MODULO| ENTREGADO)?|TOTAL SOBRANTE|TOTAL)$/;
const FILLER = /^(FORMULA 1.*|ACREDITACIONES|F1)$/;
const CODE_COL = /FOLIO|^F\. |^NO\b|^ID$/;
const DATE_COL = /FECHA/;
const NUMERIC_COL = /CANTIDAD|EXISTENCIA|PRODUCIDO|SOBRANTE|TOTAL/;
const WAREHOUSE = /^(EN )?BODEGA$/;
// Fila que cierra un bloque: TOTAL o, en el Corte ECOS, ENTREGADOS (seguida de BODEGA con lo que quedó)
const TOTAL_ROW = /^(TOTAL|ENTREGADOS$)/;
const ORDINALS = { SEGUNDA: 2, TERCERA: 3, CUARTA: 4 };
const MAX_ROWS = 3000, MAX_COLS = 60, MAX_GAP = 10, TITLE_ROWS = 4;

// Movimiento con SOLICITANTE = BODEGA: stock que se quedó en bodega, no se usó ni se acreditó
export const isWarehouse = v => WAREHOUSE.test(norm(v));

// Bloque que es producción adicional: el título lo dice (EXTRA, SEGUNDA/TERCERA, "2 PRODUCCION")
// o es "<bloque> N" cuando en la misma hoja existe "<bloque>" sin número (p. ej. LIMITADO ZONA 2).
function extraLabel(title, titles) {
  const t = norm(title);
  if (/EXTRA/.test(t)) return 'Extra';
  const word = t.match(/\b(SEGUNDA|TERCERA|CUARTA)\b/);
  if (word) return `${ORDINALS[word[1]]}ª producción`;
  const numbered = t.match(/^(.*\S) (\d+)( PRODUCCION)?$/);
  if (numbered && +numbered[2] > 1 && (numbered[3] || titles.has(numbered[1]))) return `${numbered[2]}ª producción`;
  return '';
}

export const parseBook = wb => wb.SheetNames.map(name => {
  const sheet = wb.Sheets[name];
  // origin: celda donde empieza la hoja, para traducir índices a letras/filas reales de Excel
  const origin = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']).s : { r: 0, c: 0 };
  return { ...parseSheet(name, sheetGrid(sheet)), origin };
});

function sheetGrid(sheet) {
  if (!sheet['!ref']) return [];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  range.e.r = Math.min(range.e.r, MAX_ROWS);
  range.e.c = Math.min(range.e.c, MAX_COLS);
  const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, range, blankrows: true });
  // Filas ocultas (o filtradas) en Excel: se vacían y se marcan para no leerlas ni contarlas como hueco
  const props = sheet['!rows'] ?? [];
  return grid.map((row, i) => props[i + range.s.r]?.hidden ? Object.assign([], { hidden: true }) : row);
}

function parseSheet(name, g) {
  const heads = g.flatMap((row, r) => segments(row).map(seg => ({ r, ...seg })));
  const tables = heads.map(h => readTable(g, heads, h)).filter(Boolean);
  const titles = new Set(tables.map(t => norm(t.meta.title)));
  for (const t of tables) t.meta.extra = extraLabel(t.meta.title, titles);

  // Orden de lectura: primero la columna principal de arriba a abajo, luego las tablas laterales (vacías se ocultan)
  const minStart = Math.min(...tables.map(t => t.s));
  const isSide = t => t.s > minStart + 3;
  tables.sort((a, b) => isSide(a) - isSide(b) || a.r - b.r || a.s - b.s);
  const shown = tables.filter(t => t.rows.length || !isSide(t));
  // Las tablas laterales son cálculos auxiliares del Excel (p. ej. SOBRANTES): se marcan y se nombran si no tienen título
  for (const t of shown) {
    t.side = isSide(t);
    if (t.side && !t.meta.title) t.meta.title = `Resumen · ${t.cols.at(-1)?.name || 'auxiliar'}`;
  }

  // Hojas sin tablas: se muestran tal cual (solo las celdas con algo) y, si se puede, como pares nombre → número
  const raw = shown.length ? [] : g.map(row => (row || []).filter(v => !blank(v)).map(clean)).filter(row => row.length);
  return { name: name.trim(), tables: shown, raw, free: shown.length ? null : freeForm(g) };
}

// Hoja sin tablas (p. ej. ACREDITACIONES GLOBAL): cada texto con el número que tiene justo debajo en su columna
function freeForm(g) {
  const cells = g.flatMap((row, r) => (row || []).flatMap((v, c) => blank(v) ? [] : [{ r, c, v: clean(v) }]));
  const used = new Set(), items = [], texts = [];
  for (const cell of cells) {
    if (typeof cell.v !== 'string' || used.has(cell)) continue;
    const below = cells.find(o => o.c === cell.c && o.r > cell.r && !used.has(o));
    if (below && typeof below.v === 'number') { items.push({ label: cell.v, value: below.v }); used.add(below); }
    else texts.push(cell.v);
  }
  return { title: texts[0] || '', notes: texts.slice(1), items };
}

// En una fila de encabezados puede haber varias tablas lado a lado: se separan por 4+ columnas vacías o por un encabezado repetido
function segments(row) {
  const cells = row.flatMap((v, c) => typeof v === 'string' && v.trim() ? [{ c, t: norm(v) }] : []);
  if (cells.filter(x => HEAD.test(x.t)).length < 3) return [];
  const groups = [];
  let current = null;
  for (const cell of cells) {
    const prev = current?.at(-1);
    const repeated = current?.some(y => y.t === cell.t) && !(prev.t === cell.t && cell.c === prev.c + 1);
    if (!current || cell.c - prev.c > 3 || repeated) groups.push(current = []);
    current.push(cell);
  }
  return groups.filter(group => group.filter(x => HEAD.test(x.t)).length >= 2).map(group => ({ s: group[0].c, e: group.at(-1).c }));
}

const touches = (a, b) => a.s <= b.e + 1 && b.s <= a.e + 1;
const filledCells = (row = [], s, e) => row.slice(Math.max(0, s), e + 1).filter(v => !nothing(v));
const isText = v => typeof v === 'string' && v.trim() && !/^-?\d+$/.test(v.trim());

function readTable(g, heads, h) {
  const next = heads.filter(o => o.r > h.r && touches(o, h)).sort((a, b) => a.r - b.r)[0];
  const end = next ? (titleRows(g, heads, next)[0] ?? next.r) : g.length;
  const { rows, total } = readRows(g, h, end);
  const meta = readMeta(g, titleRows(g, heads, h), h);
  if (!rows.length && !meta.title) return null; // tabla auxiliar vacía
  const cols = rows.length ? dataColumns(g, h, rows) : headerColumns(g, h);
  return { r: h.r, s: h.s, cols, rows, total, meta };
}

// Filas de título justo arriba del encabezado (nombre del bloque, cantidad, folios).
// Se detiene en filas vacías, en otro encabezado o en el TOTAL del bloque anterior.
function titleRows(g, heads, h) {
  const rows = [];
  for (let r = h.r - 1; r >= 0 && h.r - r <= TITLE_ROWS; r--) {
    if (g[r].hidden) continue;
    const values = filledCells(g[r], h.s - 1, h.e + 1);
    if (heads.some(o => o.r === r && touches(o, h)) || values.some(v => TOTAL_ROW.test(norm(v)) || (values.length <= 2 && isWarehouse(v)))) break;
    if (!values.length) { if (rows.length) break; continue; }
    // Tras una fila vacía, solo cuenta como título un texto suelto
    if (!rows.length && r < h.r - 1 && !(values.length === 1 && typeof values[0] === 'string')) break;
    rows.unshift(r);
  }
  return rows;
}

function readRows(g, h, end) {
  const rows = [];
  const qtyCol = g[h.r].findIndex((v, c) => c >= h.s && c <= h.e && /^CANTIDAD$/.test(norm(v)));
  let gap = 0;
  for (let r = h.r + 1; r < end; r++) {
    if (g[r].hidden) continue;
    const cells = g[r].slice(h.s, h.e + 1);
    if (cells.every(nothing)) { if (++gap >= MAX_GAP) break; continue; }
    gap = 0;
    if (cells.some(v => typeof v === 'string' && TOTAL_ROW.test(norm(v)))) return { rows, total: g[r] };
    // Filas sin texto en una columna con encabezado son restos de fórmulas, no movimientos…
    // …salvo que traigan CANTIDAD > 0: es una entrega a la que no se le capturó el solicitante
    // ("Columna1", "Columna2"… son nombres automáticos de Excel, no cuentan como encabezado)
    const hasText = cells.some((v, i) => isText(v) && !/^(COLUMNA\d*)?$/.test(norm(g[h.r][h.s + i])));
    if (hasText || (qtyCol >= 0 && num(g[r][qtyCol]) > 0)) rows.push(g[r]);
  }
  return { rows, total: null };
}

// Columnas con algún dato; se omite una columna que repite exactamente a la anterior
function dataColumns(g, h, rows) {
  const cols = [];
  for (let c = h.s; c <= h.e; c++) {
    if (!rows.some(row => !blank(row[c]))) continue;
    const prev = cols.at(-1);
    if (prev && norm(g[h.r][prev.c]) === norm(g[h.r][c]) && rows.every(row => norm(row[prev.c]) === norm(row[c]))) continue;
    const name = clean(g[h.r][c]) ?? '', key = norm(name);
    const code = CODE_COL.test(key), date = DATE_COL.test(key);
    const mostlyNumbers = rows.filter(row => !isNaN(num(row[c]))).length > rows.filter(row => !blank(row[c])).length / 2;
    cols.push({ c, name: !key || /^COLUMNA\d*$/.test(key) ? 'Notas' : name, date, code, n: !code && !date && mostlyNumbers });
  }
  return cols;
}

// Bloque sin movimientos: columnas según su encabezado
function headerColumns(g, h) {
  const cols = [];
  for (let c = h.s; c <= h.e; c++) {
    const key = norm(g[h.r][c]);
    if (key && !/^COLUMNA/.test(key))
      cols.push({ c, name: clean(g[h.r][c]), date: DATE_COL.test(key), code: CODE_COL.test(key), n: NUMERIC_COL.test(key) });
  }
  return cols;
}

// Datos del título del bloque: título, folios ("FOLIO: …"), cantidades y notas
function readMeta(g, rows, h) {
  const meta = { title: '', folios: [], nums: [], notes: [] };
  let expectFolio = false;
  for (const r of rows) for (let c = Math.max(0, h.s - 1); c <= h.e + 1; c++) {
    const v = clean(g[r][c]);
    if (blank(v)) continue;
    if (typeof v === 'string' && /^FOLIOS?:?$/.test(norm(v))) { expectFolio = true; continue; }
    if (typeof v === 'number') {
      if (expectFolio) meta.folios.push(String(v)); else if (v) meta.nums.push(v);
      expectFolio = false;
      continue;
    }
    if (expectFolio) { if (!meta.folios.includes(v)) meta.folios.push(v); expectFolio = false; continue; }
    if (FILLER.test(norm(v)) || /^(SOLICITANTE|RESPONSABLE)\b/.test(norm(v))) continue; // nombres: dato sensible
    if (!meta.title && v !== '✅') meta.title = v;
    else if (!meta.notes.includes(v)) meta.notes.push(v);
  }
  return meta;
}

/* ---------- consultas sobre los libros ya leídos ---------- */

export const colOf = (table, pattern) => table.cols.find(col => pattern.test(norm(col.name)));

// Recorre las filas de las tablas que tienen todas las columnas pedidas; fn recibe { clave: valor }
export function eachRow(book, need, fn) {
  for (const sheet of book.sheets) for (const table of sheet.tables) {
    const cols = Object.entries(need).map(([key, pattern]) => [key, colOf(table, pattern)]);
    if (cols.some(([, col]) => !col)) continue;
    table.rows.forEach(row => fn(Object.fromEntries(cols.map(([key, col]) => [key, row[col.c]]))));
  }
}

// Dónde está cada columna pedida, tabla por tabla: [{ sheet, header: fila del encabezado, rows, cols: { clave: 'F' } }]
export function columnSources(book, need, skip = () => false) {
  const out = [];
  for (const sheet of book?.sheets ?? []) for (const table of sheet.tables) {
    if (skip(table)) continue;
    const cols = Object.entries(need).map(([key, pattern]) => [key, colOf(table, pattern)]);
    if (cols.some(([, col]) => !col)) continue;
    out.push({
      sheet: sheet.name, header: table.r + sheet.origin.r + 1, rows: table.rows.length,
      cols: Object.fromEntries(cols.map(([key, col]) => [key, XLSX.utils.encode_col(col.c + sheet.origin.c)]))
    });
  }
  return out;
}

// Hojas de folios: una fila sin CATEGORIA debajo de una categoría es un rango extra (otra producción) de esa categoría.
// Devuelve Map(fila → categoría a la que pertenece).
export function extraRows(table) {
  const catCol = colOf(table, /^CATEGORIA$/), extras = new Map();
  if (!catCol || !table.cols.some(c => /^FOLIO/.test(norm(c.name)))) return extras;
  let category = '';
  for (const row of table.rows) {
    if (!blank(row[catCol.c])) category = clean(row[catCol.c]);
    else if (category) extras.set(row, category);
  }
  return extras;
}

/* ---------- devoluciones ---------- */

// Columna devuelta: CANTIDAD o CANTIDAD DEVUELTA. Las tablas laterales son resúmenes del detalle
// (p. ej. "Resumen · CANTIDAD DEVUELTA" en MONTAJE Y DESMONTAJE): no se suman para no duplicar.
export const returnedCol = table => table.side ? null : colOf(table, /^CANTIDAD( DEVUELTA)?$/);
export const returnedProducedCol = table => table.side ? null : colOf(table, /^CANTIDAD PRODUCIDA$/);

export function returnedTotal(book) {
  let total = 0;
  for (const sheet of book.sheets) for (const table of sheet.tables) {
    const col = returnedCol(table);
    if (col) table.rows.forEach(row => total += num(row[col.c]) || 0);
  }
  return total;
}

/* ---------- tipos de acreditación ---------- */

const OTHER_GROUP = GROUPS.at(-1);
export const groupOf = text => GROUPS.find(g => g.pattern?.test(norm(text))) ?? OTHER_GROUP;

// Tipo de una hoja: por su nombre o, si no lo dice (p. ej. "PADDOCK"), por el título de sus bloques
export const sheetGroup = sheet => [sheet.name, ...sheet.tables.map(t => t.meta.title)].map(groupOf).find(g => g !== OTHER_GROUP) ?? OTHER_GROUP;

// Categorías de la tabla principal del Corte Final (hoja CORTE FINAL, tal cual):
// { cat, p: producido, e: entregado, s: total sobrante, w: sobrante en bodega, m: sobrante en módulo, group }
export function categoryRows(book) {
  const rows = [];
  const need = { cat: /^CATEGORIA$/, prod: /^PRODUCIDO$/, tot: /^TOTAL ACREDITACIONES$/, sob: /^TOTAL SOBRANTE$/, bod: /^SOBRANTE BODEGA$/, mod: /^SOBRANTE MODULO$/ };
  eachRow(book, need, x => {
    const [e, s, w, m] = [x.tot, x.sob, x.bod, x.mod].map(v => num(v) || 0);
    // Producido escrito como texto (p. ej. "180* X DIA 90* DISEÑO"): se toma entregado + sobrante de la fila
    const fixed = !blank(x.prod) && isNaN(num(x.prod)), p = fixed ? e + s : num(x.prod) || 0;
    if (p || e) rows.push({ cat: clean(x.cat), p, e, s, w, m, fixed, group: groupOf(x.cat) });
  });
  return rows;
}

// Categorías de la hoja CORTE FINAL del Corte Hospitalidades: { cat, p: producido, e: entregado, r: restante }
export function hospitalityRows(book) {
  const rows = [];
  eachRow(book, { cat: /^CATEGORIA$/, p: /^TOTAL PRODUCIDO$/, e: /^TOTAL ENTREGADO$/, r: /^RESTANTE$/ }, x => {
    const [e, r] = [x.e, x.r].map(v => num(v) || 0);
    // Producido escrito como texto (p. ej. "180* X DIA 90* DISEÑO"): se toma entregado + restante de la fila
    const fixed = !blank(x.p) && isNaN(num(x.p)), p = fixed ? e + r : num(x.p) || 0;
    if (p || e) rows.push({ cat: clean(x.cat), p, e, r, fixed });
  });
  return rows;
}

// Texto que se muestra en una celda según el tipo de columna
export function cellText(v, col) {
  if (blank(v)) return '';
  const date = col.date && parseDate(v);
  if (date) return `${shortDate(date)} ${date.getUTCFullYear()}`;
  if (typeof v === 'number' && !col.code) return fmt(v);
  return String(clean(v));
}
