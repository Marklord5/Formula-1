// Totales de arriba, resumen por tipo de acreditación y pestañas de archivo
import { GROUPS } from './config.js';
import { $, fill, clone, fmt, norm, num, pct, plural, pretty, sum, fileName, shortDate, weekday } from './utils.js';
import { categoryRows, returnedTotal, colOf } from './excel.js';
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

export function renderKpis(books) {
  const returns = books.find(b => b.kind === 'returns'), rows = finalRows(books);
  // Los % de entregado y sobrante siguen sobre lo producido de la hoja CORTE FINAL (entregado + sobrante = ese producido)
  const produced = sum(rows, 'p'), delivered = sum(rows, 'e'), leftover = sum(rows, 's');
  fill($('#kpis'), {
    produced: fmt(foliosProduced(books.find(b => b.kind === 'final'))),
    delivered: fmt(delivered),
    deliveredShare: `${pct(delivered, produced)} de lo producido`,
    leftover: fmt(leftover),
    leftoverShare: `${pct(leftover, produced)} de lo producido`,
    returned: fmt(returnedTotal(returns)),
    returnSheets: `${returns.sheets.length} hojas`
  });
}


// Corte Final separado en Montaje y Desmontaje / Hospitalidades / Operación, con desglose por concepto
export function renderGroups(books) {
  const rows = finalRows(books);
  $('#groups').replaceChildren(...GROUPS.map(group => {
    const own = rows.filter(r => r.group === group);
    const produced = sum(own, 'p'), delivered = sum(own, 'e'), warehouse = sum(own, 'w'), module = sum(own, 'm');
    const card = clone('tpl-group', {
      name: group.name,
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

// Dona de lo entregado por tipo de acreditación (Corte Final)
export function renderDistribution(books) {
  const rows = finalRows(books);
  const slices = GROUPS.map((group, tone) => ({ name: group.name, tone, value: sum(rows.filter(r => r.group === group), 'e') })).filter(s => s.value);
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
