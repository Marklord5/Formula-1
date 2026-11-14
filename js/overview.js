// Totales de arriba, resumen por tipo de acreditación y pestañas de archivo
import { GROUPS } from './config.js';
import { $, fill, clone, fmt, norm, num, pct, plural, pretty, sum, fileName, shortDate, weekday } from './utils.js';
import { categoryRows, hospitalityRows, returnedTotal, returnedCol, colOf, columnSources } from './excel.js';
import { dailyPoints, curveHtml, donutHtml } from './charts.js';

// Categorías de la hoja CORTE FINAL
const finalRows = books => categoryRows(books.find(b => b.kind === 'final'));

// Suma de PRODUCIDO de la hoja FOLIOS OPERACION del Corte Final
function foliosProduced(book) {
  const sheet = book.sheets.find(s => norm(s.name) === 'FOLIOS OPERACION');
  let total = 0;
  for (const table of sheet?.tables ?? []) {
    const col = colOf(table, /^PRODUCIDO$/);
    if (col) table.rows.forEach(row => total += num(row[col.c]) || 0);
  }
  return total;
}

// Columnas de la tabla principal de la hoja CORTE FINAL (Corte Final) y de la del Corte Hospitalidades
const FINAL_COLS = { p: /^PRODUCIDO$/, e: /^TOTAL ACREDITACIONES$/, s: /^TOTAL SOBRANTE$/, w: /^SOBRANTE BODEGA$/, m: /^SOBRANTE MODULO$/, cat: /^CATEGORIA$/ };
const HOSP_COLS = { p: /^TOTAL PRODUCIDO$/, e: /^TOTAL ENTREGADO$/, w: /^RESTANTE$/, cat: /^CATEGORIA$/ };
const COL_NAMES = { p: 'PRODUCIDO', e: 'TOTAL ACREDITACIONES', s: 'TOTAL SOBRANTE', w: 'SOBRANTE BODEGA', m: 'SOBRANTE MODULO' };
const HOSP_NAMES = { p: 'TOTAL PRODUCIDO', e: 'TOTAL ENTREGADO', w: 'RESTANTE' };

// Ubicación de las columnas: hojas y letras únicas ("CORTE FINAL", "H")
const sheetsOf = srcs => [...new Set(srcs.map(s => s.sheet))].join(', ') || 'no encontrada';
const letter = (srcs, key) => [...new Set(srcs.map(s => s.cols[key]))].join('/') || '?';
const headerOf = srcs => srcs.length ? `encabezado fila ${srcs.map(s => s.header).join(', ')}` : '';

// Nodo con clase y contenido (texto o nodos)
function el(tag, cls, ...kids) {
  const node = document.createElement(tag);
  if (cls) node.className = cls;
  node.append(...kids);
  return node;
}
const colTag = (l, name) => el('span', 'col', el('b', '', l), name);
const FILE_ICON = '<svg class="icon" viewBox="0 0 24 24"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/></svg>';

// Ruta "archivo / hoja / [F] COLUMNA" y una línea de detalle discreta
function trail(book, sheet, col) {
  const t = el('span', 'trail');
  t.innerHTML = FILE_ICON;
  t.append(el('span', 'file', fileName(book.url)), el('i', '', '/'), el('span', '', sheet), el('i', '', '/'), col);
  t.title = `Excel ${book.name} (${fileName(book.url)}) · hoja ${sheet}`;
  return t;
}

// Suma de cada hoja del Excel de Devolución: [{ sheet, letter, name, rows, total }]
function returnsBySheet(book) {
  return book.sheets.flatMap(sheet => sheet.tables.flatMap(table => {
    const col = returnedCol(table);
    if (!col) return [];
    const total = table.rows.reduce((t, row) => t + (num(row[col.c]) || 0), 0);
    return [{ sheet: sheet.name, letter: XLSX.utils.encode_col(col.c + sheet.origin.c), name: col.name, rows: table.rows.length, total }];
  }));
}

export function renderKpis(books) {
  const final = books.find(b => b.kind === 'final'), returns = books.find(b => b.kind === 'returns'), rows = finalRows(books);
  // Los % de entregado y sobrante siguen sobre lo producido de la hoja CORTE FINAL (entregado + sobrante = ese producido)
  const produced = sum(rows, 'p'), delivered = sum(rows, 'e'), leftover = sum(rows, 's');
  const folios = columnSources(final, { p: /^PRODUCIDO$/ }).filter(s => norm(s.sheet) === 'FOLIOS OPERACION');
  const cut = columnSources(final, FINAL_COLS);
  const cats = `${plural(rows.length, 'categoría')} · ${headerOf(cut)}`;
  fill($('#kpis'), {
    produced: fmt(foliosProduced(final)),
    delivered: fmt(delivered),
    deliveredShare: `${pct(delivered, produced)} · ${letter(cut, 'e')} ÷ ${letter(cut, 'p')}`,
    leftover: fmt(leftover),
    leftoverShare: `${pct(leftover, produced)} · ${letter(cut, 's')} ÷ ${letter(cut, 'p')}`,
    returned: fmt(returnedTotal(returns))
  });
  // El % se explica completo al pasar el mouse por el chip
  const why = key => `${COL_NAMES[key]} (col ${letter(cut, key)}) ÷ PRODUCIDO (col ${letter(cut, 'p')}) de la hoja ${sheetsOf(cut)}: ${fmt(sum(rows, key))} de ${fmt(produced)}`;
  $('#kpis [data-f="deliveredShare"]').title = why('e');
  $('#kpis [data-f="leftoverShare"]').title = why('s');

  const byReturn = returnsBySheet(returns), returnCols = [...new Set(byReturn.map(r => r.name))].join(' / ');
  const sources = {
    produced: [trail(final, sheetsOf(folios), colTag(letter(folios, 'p'), 'PRODUCIDO')),
      el('span', 'meta', `${plural(sum(folios, 'rows'), 'fila')} sumadas · ${headerOf(folios)}`)],
    delivered: [trail(final, sheetsOf(cut), colTag(letter(cut, 'e'), COL_NAMES.e)), el('span', 'meta', cats)],
    leftover: [trail(final, sheetsOf(cut), colTag(letter(cut, 's'), COL_NAMES.s)),
      el('span', 'meta', '= ', colTag(letter(cut, 'w'), 'bodega'), ' + ', colTag(letter(cut, 'm'), 'módulo'), ` · ${plural(rows.length, 'categoría')}`)],
    returned: [trail(returns, plural(byReturn.length, 'hoja'), colTag('', returnCols)),
      el('span', 'sheets', ...byReturn.map(r => el('span', '', el('span', '', r.sheet), colTag(r.letter, ''), el('span', 'num', fmt(r.total)))))]
  };
  for (const [key, nodes] of Object.entries(sources)) $(`#src-${key}`).replaceChildren(...nodes);
}


// Categorías de un tipo: del Corte Final o, si el tipo tiene `book`, de ese Excel (restante va como bodega)
function groupRows(books, group) {
  if (!group.book) return finalRows(books).filter(r => r.group === group);
  return hospitalityRows(books.find(b => b.kind === group.book)).map(r => ({ ...r, s: r.r, w: r.r, m: 0 }));
}

// Montaje y Desmontaje / Operación del Corte Final y Hospitalidades de su propio Excel, con desglose por concepto
export function renderGroups(books) {
  $('#groups').replaceChildren(...GROUPS.map(group => {
    const own = groupRows(books, group);
    const produced = sum(own, 'p'), delivered = sum(own, 'e'), warehouse = sum(own, 'w'), module = sum(own, 'm');
    const card = clone('tpl-group', {
      name: group.name,
      warehouseLabel: group.book ? 'Restante' : 'En bodega',
      categories: plural(own.length, 'categoría'),
      produced: fmt(produced),
      delivered: fmt(delivered),
      deliveredShare: pct(delivered, produced),
      warehouse: fmt(warehouse),
      warehouseShare: pct(warehouse, produced),
      module: fmt(module),
      moduleShare: pct(module, produced)
    });
    card.classList.toggle('no-module', !module);
    // Fuente plegada: ruta en el resumen y qué columna alimenta cada número
    const book = books.find(b => b.kind === (group.book ?? 'final'));
    const srcs = columnSources(book, group.book ? HOSP_COLS : FINAL_COLS), names = group.book ? HOSP_NAMES : COL_NAMES;
    const map = [['Entregado', 'e'], [group.book ? 'Restante' : 'En bodega', 'w'], ...(group.book ? [] : [['En módulo', 'm']]), ['Total', 'p']];
    const origin = card.querySelector('.origin');
    origin.querySelector('summary').append(trail(book, sheetsOf(srcs), el('span', '', `${plural(map.length, 'columna')}`)));
    origin.querySelector('.origin-body').append(
      ...map.map(([label, key]) => el('span', 'row', el('span', 'k', label), colTag(letter(srcs, key), names[key]))),
      el('span', 'row', el('span', 'k', 'Categorías'), el('span', '', group.rule)));
    const fixed = own.filter(r => r.fixed);
    if (fixed.length) origin.querySelector('.origin-body').append(el('span', 'row', el('span', 'k', 'Ajuste'),
      el('span', '', `${plural(fixed.length, 'fila')} con producido en texto: se usó entregado + ${group.book ? 'restante' : 'sobrante'}`)));
    card.querySelector('.meter i').style.width = `${produced ? Math.min(100, delivered / produced * 100) : 0}%`;
    card.querySelector('tbody').replaceChildren(...own.map(r => clone('tpl-concept', {
      name: pretty(r.cat), delivered: fmt(r.e), warehouse: fmt(r.w), module: fmt(r.m), produced: fmt(r.p)
    })));
    return card;
  }));
}

// Curva acumulada del Corte Diario con sus tres hitos (primer día, día pico y último día)
export function renderTrend(books) {
  const points = dailyPoints(books.find(b => b.kind === 'daily'));
  $('#curve').innerHTML = curveHtml(points, $('#curve').clientWidth);
  if (!points.length) return;
  const last = points.at(-1), peak = points.reduce((a, p) => p.v > a.v ? p : a);
  fill($('#trend'), { cutDate: `Último: ${shortDate(new Date(last.t))}` });
  for (const [key, point] of Object.entries({ first: points[0], peak, last })) {
    const date = new Date(point.t);
    fill($(`#milestone-${key}`), { date: shortDate(date), qty: fmt(point.v), weekday: weekday(date) });
  }
}

// Dona de lo entregado por tipo de acreditación (mismas fuentes que las tarjetas)
export function renderDistribution(books) {
  const slices = GROUPS.map((group, tone) => ({ name: group.name, tone, value: sum(groupRows(books, group), 'e') })).filter(s => s.value);
  const total = sum(slices, 'value');
  $('#donut').innerHTML = donutHtml(slices, total);
  fill($('#distribution'), { donutTotal: `Total ${fmt(total)}` });
  $('#legend').replaceChildren(...slices.map(s => {
    const item = clone('tpl-legend', { name: s.name, value: fmt(s.value), share: pct(s.value, total) });
    item.classList.add(`tone-${s.tone}`);
    return item;
  }));
}

// Un botón de descarga por archivo (se pinta antes de leer los datos)
export function renderDownloads(files) {
  $('#downloads').replaceChildren(...files.map(({ name, url }) => {
    const link = clone('tpl-download', { name });
    Object.assign(link, { href: url, download: fileName(url), title: `Descargar ${fileName(url)}` });
    return link;
  }));
}

export function renderTabs(books, state) {
  $('#tabs').replaceChildren(...books.map((book, i) => {
    const tab = clone('tpl-tab', { name: book.name, detail: `${plural(book.sheets.length, 'hoja')} · ${fileName(book.url)}` });
    tab.dataset.file = i;
    tab.classList.toggle('on', i === state.file);
    tab.setAttribute('aria-pressed', i === state.file);
    return tab;
  }));
  const { url } = books[state.file];
  Object.assign($('#download'), { href: url, download: fileName(url) });
}
