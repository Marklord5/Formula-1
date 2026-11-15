// Gráficas de cada Excel. Devuelven HTML/SVG; los colores viven en css/styles.css (.c-ent, .c-sob, .c-dev)
import { GROUPS, TOP_PER_GROUP, TOP_HOSPITALITY } from './config.js';
import { $, esc, fmt, pct, plural, norm, num, clean, pretty, sum, parseDate, shortDate, weekday, DAY_MS } from './utils.js';
import { colOf, eachRow, isWarehouse, categoryRows, returnedCol, returnedProducedCol, hospitalityRows } from './excel.js';

const DATE_WINDOW = 120 * DAY_MS;

const CHARTS = { final: finalCharts, daily: dailyCharts, returns: returnCharts, hospitality: hospitalityCharts, ecos: ecosCharts,
  global: () => card('Las gráficas del Reporte Global están arriba', 'Ve a <a href="#global">Panorama global del sistema</a>. Abajo están las tablas de resumen de cada hoja; los nombres de personas y los folios no se muestran.', '', '', 'wide') };
export const chartsHtml = (book, options) => CHARTS[book.kind](book, options);

/* ---------- piezas comunes ---------- */

const tipAttr = html => `data-tip="${esc(html)}"`;
const tipRow = (key, value) => `<div><span class="k">${key}:</span> ${value}</div>`;
const legend = items => `<div class="legend">${items.map(([tone, label]) => `<span><i class="c-${tone}"></i>${label}</span>`).join('')}</div>`;
const card = (title, how, body, extra = '', cls = '') => `<section class="card lg ${cls}">
  <div class="head"><div><h2>${title}</h2><p class="how">${how}</p></div>${extra}</div>${body}</section>`;

// Barras horizontales; cada fila: { label, value, tip, segs: [{ v, tone }] }. `max` permite compartir escala entre grupos.
export function hbars(rows, max = Math.max(...rows.map(r => sum(r.segs, 'v')), 1)) {
  return `<div class="hb">${rows.map(r => `<div class="hb-row" ${tipAttr(r.tip)}>
      <span class="hb-l">${esc(r.label)}</span>
      <span class="hb-t">${r.segs.filter(s => s.v > 0).map(s => `<i class="c-${s.tone}" style="width:${s.v / max * 100}%"></i>`).join('')}</span>
      <span class="hb-v num">${r.value}</span></div>`).join('')}</div>`;
}

// Paso "redondo" del eje Y (1, 2, 2.5 o 5 × 10ⁿ) para ~4 divisiones
function niceStep(max) {
  const raw = max / 4, magnitude = 10 ** Math.floor(Math.log10(raw || 1));
  return [1, 2, 2.5, 5, 10].find(m => m * magnitude >= raw) * magnitude;
}

// Columnas por día sobre un eje de tiempo real; points = [{ t, v, n }] ordenados.
// Se dibuja al ancho real del contenedor para que el texto no se agrande ni se encoja.
function columns(points, containerWidth) {
  if (!points.length) return '<div class="empty">No hay fechas en este Excel.</div>';
  const W = Math.max(300, (containerWidth || 800) - 50), H = 240, L = 44, R = 12, T = 24, B = 30;
  const t0 = points[0].t, t1 = points.at(-1).t, days = Math.max(1, (t1 - t0) / DAY_MS + 1);
  const slot = (W - L - R) / days, barWidth = Math.max(3, Math.min(24, slot * .7));
  const maxV = Math.max(...points.map(p => p.v)), step = niceStep(maxV), yMax = Math.ceil(maxV / step) * step || 1;
  const x = t => L + ((t - t0) / DAY_MS + .5) * slot, y = v => T + (1 - v / yMax) * (H - T - B);
  const peak = points.reduce((a, p) => p.v > a.v ? p : a);

  const bar = p => {
    const left = x(p.t) - barWidth / 2, top = y(p.v), r = Math.min(4, barWidth / 2, y(0) - top);
    return `<path class="col" d="M${left},${y(0)}V${top + r}Q${left},${top} ${left + r},${top}H${left + barWidth - r}Q${left + barWidth},${top} ${left + barWidth},${top + r}V${y(0)}Z"/>`;
  };
  const hit = p => {
    const d = new Date(p.t);
    return `<rect class="col-hit" x="${x(p.t) - slot / 2}" y="${T}" width="${slot}" height="${H - T - B}"
      ${tipAttr(`<b>${weekday(d)} ${shortDate(d)}</b>${tipRow('Salieron ese día', fmt(p.v))}${tipRow('Movimientos', p.n)}`)}/>`;
  };

  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Folios que salieron de bodega por día">
    <g class="axis">${yAxis(yMax, step, y, L, W - R)}${xAxis(t0, days, x, W, H)}</g>
    ${points.map(bar).join('')}
    <text class="peak-label" x="${x(peak.t)}" y="${y(peak.v) - 8}" text-anchor="middle">${fmt(peak.v)}</text>
    ${points.map(hit).join('')}
  </svg>`;
}

/* ---------- ejes compartidos ---------- */

const compact = v => v >= 10000 ? `${fmt(v / 1000)}K` : fmt(v);

// Líneas guía y etiquetas del eje Y
const yAxis = (yMax, step, y, left, right) => Array.from({ length: Math.round(yMax / step) + 1 }, (_, i) => i * step).map(v =>
  `<line class="grid-l" x1="${left}" x2="${right}" y1="${y(v)}" y2="${y(v)}"/>
   <text x="${left - 8}" y="${y(v) + 4}" text-anchor="end">${compact(v)}</text>`).join('');

const LABEL_SPACING = 64; // px mínimos entre fechas del eje X

// Fechas del eje X: tantas como quepan (cada N días), siempre con el primer y el último día
function xAxis(t0, days, x, width, height) {
  const every = Math.max(1, Math.ceil((days - 1) / Math.max(1, Math.floor(width / LABEL_SPACING) - 1)));
  const offsets = [];
  for (let i = 0; i < days - 1; i += every) offsets.push(i);
  if (offsets.length > 1 && days - 1 - offsets.at(-1) < every * .7) offsets.pop(); // no encimar con el último
  offsets.push(days - 1);
  const ticks = offsets.map(i => t0 + i * DAY_MS);
  const anchor = i => i === 0 ? 'start' : i === ticks.length - 1 ? 'end' : 'middle';
  return ticks.map((t, i) => `<line class="tick" x1="${x(t)}" x2="${x(t)}" y1="${height - 30}" y2="${height - 25}"/>
    <text x="${x(t)}" y="${height - 8}" text-anchor="${anchor(i)}">${shortDate(new Date(t))}</text>`).join('');
}

/* ---------- resumen general: curva acumulada y dona por tipo ---------- */

// Curva acumulada de entregas; points = [{ t, v, n }] por día. Guía y punto de cada día aparecen con :hover (CSS).
export function curveHtml(points, containerWidth) {
  if (!points.length) return '<div class="empty">No hay fechas en el Corte Diario.</div>';
  const W = Math.max(300, (containerWidth || 640) - 2), H = 330, L = 44, R = 16, T = 30, B = 30;
  let running = 0;
  const pts = points.map(p => ({ ...p, total: running += p.v }));
  const first = pts[0], last = pts.at(-1), peak = pts.reduce((a, p) => p.v > a.v ? p : a);
  const days = (last.t - first.t) / DAY_MS + 1, step = niceStep(last.total), yMax = Math.ceil(last.total / step) * step;
  const x = t => L + (t - first.t) / ((last.t - first.t) || 1) * (W - L - R), y = v => T + (1 - v / yMax) * (H - T - B);
  const line = pts.map((p, i) => `${i ? 'L' : 'M'}${x(p.t).toFixed(1)},${y(p.total).toFixed(1)}`).join('');
  const xs = pts.map(p => x(p.t));

  const hover = (p, i) => {
    const from = i ? (xs[i - 1] + xs[i]) / 2 : L, to = i < xs.length - 1 ? (xs[i] + xs[i + 1]) / 2 : W - R, d = new Date(p.t);
    return `<g class="pt" ${tipAttr(`<b>${weekday(d)} ${shortDate(d)}</b>${tipRow('Acumulado', fmt(p.total))}${tipRow('Ese día', `+${fmt(p.v)}`)}${tipRow('Movimientos', p.n)}`)}>
      <rect class="col-hit" x="${from}" y="${T}" width="${to - from}" height="${H - T - B}"/>
      <line class="guide" x1="${xs[i]}" x2="${xs[i]}" y1="${T}" y2="${y(0)}"/><circle class="hover-dot" cx="${xs[i]}" cy="${y(p.total)}" r="5"/></g>`;
  };

  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Folios que salieron de bodega, acumulados por fecha">
    <defs><linearGradient id="curveFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-opacity=".28"/><stop offset="1" stop-opacity="0"/></linearGradient></defs>
    <g class="axis">${yAxis(yMax, step, y, L, W - R)}${xAxis(first.t, days, x, W, H)}</g>
    <path class="curve-area" d="${line}L${x(last.t)},${y(0)}L${x(first.t)},${y(0)}Z"/>
    <path class="curve-line" d="${line}"/>
    <line class="cut-line" x1="${x(last.t)}" x2="${x(last.t)}" y1="${T - 10}" y2="${y(0)}"/>
    <text class="cut-label" x="${x(last.t) - 6}" y="${T - 14}" text-anchor="end">${fmt(last.total)} folios</text>
    ${[peak, last].map(p => `<circle class="dot" cx="${x(p.t)}" cy="${y(p.total)}" r="4.5"/>`).join('')}
    ${pts.map(hover).join('')}
  </svg>`;
}

// Dona; slices = [{ name, value, tone }] — tone es el índice del tipo (.tone-N en CSS)
export function donutHtml(slices, total) {
  const size = 200, radius = 82, gap = 2, center = size / 2, circumference = 2 * Math.PI * radius;
  let offset = 0;
  const arcs = slices.map(s => {
    const length = s.value / total * circumference;
    const arc = `<circle class="seg tone-${s.tone}" cx="${center}" cy="${center}" r="${radius}"
      stroke-dasharray="${Math.max(0, length - gap)} ${circumference}" stroke-dashoffset="${-offset}" transform="rotate(-90 ${center} ${center})"
      ${tipAttr(`<b>${esc(s.name)}</b>${tipRow('Entregadas', fmt(s.value))}${tipRow('Del total', pct(s.value, total))}`)}/>`;
    offset += length;
    return arc;
  }).join('');
  return `<svg viewBox="0 0 ${size} ${size}" role="img" aria-label="Acreditaciones entregadas por tipo">${arcs}
    <text class="donut-total" x="${center}" y="${center - 2}">${fmt(total)}</text>
    <text class="donut-label" x="${center}" y="${center + 18}">ENTREGADAS</text>
  </svg>`;
}

/* ---------- Corte Final: entregado + sobrante por categoría ---------- */

function finalCharts(book, { all }) {
  const rows = categoryRows(book).sort((a, b) => b.p - a.p);
  const key = legend([['ent', 'Entregado'], ['sob', 'Sobrante']]);

  // Una sección por tipo de acreditación, todas con la misma escala
  const max = Math.max(...rows.map(r => r.e + r.s), 1);
  const bar = r => ({
    label: r.cat, value: fmt(r.p), segs: [{ v: r.e, tone: 'ent' }, { v: r.s, tone: 'sob' }],
    tip: `<b>${esc(r.cat)}</b>${tipRow('Producido', fmt(r.p))}${tipRow('Entregado', `${fmt(r.e)} (${pct(r.e, r.p)})`)}${tipRow('Sobrante', `${fmt(r.s)} (${pct(r.s, r.p)})`)}`
  });
  const byGroup = GROUPS.map(group => ({ group, own: rows.filter(r => r.group === group) })).filter(g => g.own.length);
  const sections = byGroup.map(({ group, own }) => {
    const shown = all ? own : own.slice(0, TOP_PER_GROUP), hidden = own.length - shown.length;
    return `<div class="hb-group"><span class="label">${esc(group.name)}</span><span class="sub">${fmt(sum(own, 'p'))} producidas · ${pct(sum(own, 'e'), sum(own, 'p'))} entregado${hidden ? ` · ${plural(hidden, 'categoría')} más` : ''}</span></div>
      ${hbars(shown.map(bar), max)}`;
  }).join('');
  const toggle = byGroup.some(g => g.own.length > TOP_PER_GROUP)
    ? `<button class="btn more" data-all>${all ? `Ver solo las ${TOP_PER_GROUP} principales de cada tipo` : `Ver las ${rows.length} categorías`}</button>` : '';
  const bars = card('Entregado y sobrante por categoría',
    'Cada barra es una categoría de la hoja CORTE FINAL, separadas por tipo de acreditación. La parte verde es lo entregado y la dorada lo que sobró. El número es lo producido.',
    sections + toggle, key, 'wide');

  return bars;
}

/* ---------- Corte Diario: entregas por día + quién recibió más ---------- */

function dailyCharts(book, { width }) {
  const moves = [];
  eachRow(book, { sol: /^SOLICITANTE$/, q: /^CANTIDAD$/ }, x => {
    const q = num(x.q);
    if (q > 0 && !isWarehouse(x.sol)) moves.push({ sol: clean(x.sol) || 'Sin nombre', q });
  });
  const total = sum(moves, 'q');

  const bySolicitor = {};
  for (const m of moves) {
    const entry = bySolicitor[norm(m.sol)] ||= { label: pretty(m.sol), v: 0, n: 0 };
    entry.v += m.q;
    entry.n++;
  }
  const top = Object.values(bySolicitor).sort((a, b) => b.v - a.v).slice(0, 10);

  return card('Salidas de bodega por día',
      `Suma de la columna CANTIDAD de todas las hojas, agrupada por FECHA: folios que salieron de bodega hacia módulos o solicitantes. No es lo mismo que acreditado: lo que sobró en los módulos el Corte Final lo cuenta como sobrante módulo. No cuenta las filas de BODEGA. En total salieron <b>${fmt(total)}</b>.`,
      `<div class="chart">${columns(dailyPoints(book), width)}</div>`, '', 'wide') +
    card('¿Quién recibió más?', 'Los 10 solicitantes o módulos que más folios recibieron de bodega.',
      hbars(top.map(s => ({
        label: s.label, value: fmt(s.v), segs: [{ v: s.v, tone: 'ent' }],
        tip: `<b>${esc(s.label)}</b>${tipRow('Recibidos', fmt(s.v))}${tipRow('Movimientos', s.n)}${tipRow('Del total', pct(s.v, total))}`
      }))), '', 'wide');
}

// Entregas por fecha. Se descartan fechas mal capturadas: más de 120 días lejos de la mediana.
export function dailyPoints(book) {
  const dated = [];
  for (const sheet of book.sheets) for (const table of sheet.tables) {
    const solCol = colOf(table, /^SOLICITANTE$/), qtyCol = colOf(table, /^CANTIDAD$/), dateCol = colOf(table, /^FECHA$/);
    if (!qtyCol || !dateCol) continue;
    for (const row of table.rows) {
      const q = num(row[qtyCol.c]), d = parseDate(row[dateCol.c]);
      if (q > 0 && d && !(solCol && isWarehouse(row[solCol.c]))) dated.push({ t: d.getTime(), q });
    }
  }
  const times = dated.map(d => d.t).sort((a, b) => a - b), median = times[times.length >> 1];
  const byDay = {};
  for (const d of dated.filter(d => Math.abs(d.t - median) <= DATE_WINDOW)) {
    const day = byDay[d.t] ||= { t: d.t, v: 0, n: 0 };
    day.v += d.q;
    day.n++;
  }
  return Object.values(byDay).sort((a, b) => a.t - b.t);
}

/* ---------- Devolución: por hoja y por categoría ---------- */

function returnCharts(book) {
  const bySheet = [], byCategory = {};
  for (const sheet of book.sheets) {
    let sheetTotal = 0;
    for (const table of sheet.tables) {
      const qtyCol = returnedCol(table), prodCol = returnedProducedCol(table), catCol = colOf(table, /^CATEGORIA$|^TIPO DE BRAZALETE$/);
      if (!qtyCol) continue;
      for (const row of table.rows) {
        const q = num(row[qtyCol.c]) || 0;
        sheetTotal += q;
        if (!catCol || !q) continue;
        const entry = byCategory[norm(row[catCol.c]) || 'SIN CATEGORÍA'] ||= { label: pretty(row[catCol.c] || 'Sin categoría'), v: 0, p: 0, vp: 0 };
        entry.v += q;
        const produced = prodCol ? num(row[prodCol.c]) || 0 : 0;
        if (produced) { entry.p += produced; entry.vp += q; } // % solo sobre filas con CANTIDAD PRODUCIDA
      }
    }
    if (sheetTotal) bySheet.push({ label: sheet.name, v: sheetTotal });
  }
  const total = sum(bySheet, 'v');
  const top = Object.values(byCategory).sort((a, b) => b.v - a.v).slice(0, 10);
  // CANTIDAD PRODUCIDA (cuando la hoja la trae) permite decir qué parte de lo producido regresó
  const bars = rows => hbars(rows.map(s => ({
    label: s.label, value: fmt(s.v), segs: [{ v: s.v, tone: 'dev' }],
    tip: `<b>${esc(s.label)}</b>${tipRow('Devueltos', fmt(s.v))}${s.p ? tipRow('De lo producido', `${fmt(s.vp)} de ${fmt(s.p)} (${pct(s.vp, s.p)})`) : ''}${tipRow('Del total', pct(s.v, total))}`
  })));

  return card('Devueltos por hoja', `Suma de CANTIDAD (o CANTIDAD DEVUELTA) en cada hoja del Excel de devolución, sin las tablas de resumen. En total regresaron <b>${fmt(total)}</b>.`, bars(bySheet)) +
    card('Lo que más se devolvió', 'Las 10 categorías o tipos de brazalete con más piezas devueltas.', bars(top));
}

/* ---------- Corte Hospitalidades: entregado vs restante y entregas por día ---------- */

function hospitalityCharts(book, { all, width }) {
  const rows = hospitalityRows(book);
  rows.sort((a, b) => b.p - a.p);
  const P = sum(rows, 'p'), E = sum(rows, 'e'), R = sum(rows, 'r');
  const key = legend([['ent', 'Entregado'], ['sob', 'Restante']]);
  // Categorías con más entregado que producido: el TOTAL PRODUCIDO del Excel está incompleto
  const overDelivered = rows.filter(r => r.e > r.p);
  const warning = overDelivered.length
    ? ` <b>Ojo:</b> ${plural(overDelivered.length, 'categoría')} tienen más entregado que producido (p. ej. ${esc(overDelivered[0].cat)}); revisa su TOTAL PRODUCIDO en el Excel.` : '';

  const total = card('¿Cuánto se entregó de lo que se produjo?',
    `Hoja CORTE FINAL del Corte Hospitalidades: se produjeron <b>${fmt(P)}</b>, se entregaron <b>${fmt(E)}</b> (${pct(E, P)}) y quedaron <b>${fmt(R)}</b> sin entregar.${warning}`,
    `<div class="hb-t total" ${tipAttr(`<b>Total Hospitalidades</b>${tipRow('Producido', fmt(P))}${tipRow('Entregado', `${fmt(E)} (${pct(E, P)})`)}${tipRow('Restante', fmt(R))}`)}>
      <i class="c-ent" style="width:${E / (E + R) * 100}%"></i><i class="c-sob" style="width:${R / (E + R) * 100}%"></i></div>`,
    key, 'wide');

  const shown = all ? rows : rows.slice(0, TOP_HOSPITALITY);
  const toggle = rows.length > TOP_HOSPITALITY
    ? `<button class="btn more" data-all>${all ? `Ver solo las ${TOP_HOSPITALITY} principales` : `Ver las ${rows.length} categorías`}</button>` : '';
  const bars = card('Entregado y restante por categoría',
    'Cada barra es una categoría; la parte verde es lo entregado y la dorada lo que quedó. El número es lo producido.',
    hbars(shown.map(r => ({
      label: r.cat, value: fmt(r.p), segs: [{ v: r.e, tone: 'ent' }, { v: r.r, tone: 'sob' }],
      tip: `<b>${esc(r.cat)}</b>${tipRow('Producido', fmt(r.p))}${tipRow('Entregado', `${fmt(r.e)} (${pct(r.e, r.p)})`)}${tipRow('Restante', fmt(r.r))}`
    }))) + toggle,
    key, 'wide');

  return total + bars + card('Entregas por día', ...hospitalityDays(book, width), '', 'wide');
}

// Hoja ENTREGAS: TOTAL ENTREGADO agrupado por FECHA → [descripción, gráfica]
function hospitalityDays(book, width) {
  const byDay = {};
  let undated = 0;
  eachRow(book, { d: /^FECHA$/, q: /^TOTAL ENTREGADO$/ }, x => {
    const q = num(x.q), date = parseDate(x.d);
    if (!(q > 0)) return;
    if (!date) { undated += q; return; }
    const day = byDay[+date] ||= { t: +date, v: 0, n: 0 };
    day.v += q;
    day.n++;
  });
  const points = Object.values(byDay).sort((a, b) => a.t - b.t);
  const how = `Suma de TOTAL ENTREGADO de la hoja ENTREGAS por FECHA: <b>${fmt(sum(points, 'v'))}</b> en ${plural(points.length, 'día')}.`
    + (undated ? ` Otras <b>${fmt(undated)}</b> no tienen fecha y no aparecen aquí.` : '');
  return [how, `<div class="chart">${columns(points, width)}</div>`];
}

/* ---------- Corte ECOS: pases de estacionamiento y candados ---------- */

// Hoja CONTEO FINAL (producido, entregado y sobrante en bodega por categoría) + entregas por día de los bloques
function ecosCharts(book, { all, width }) {
  const rows = [];
  eachRow(book, { cat: /^CATEGORIA$/, p: /^TOTAL PRODUCIDO$/, e: /^CANTIDAD ENTREGADA$/, s: /^SOBRANTE BODEGA$/ }, x => {
    const [p, e, s] = [x.p, x.e, x.s].map(v => num(v) || 0);
    if (p || e) rows.push({ cat: clean(x.cat), p, e, s });
  });
  rows.sort((a, b) => b.p - a.p);
  const P = sum(rows, 'p'), E = sum(rows, 'e'), S = sum(rows, 's');
  const key = legend([['ent', 'Entregado'], ['sob', 'Sobrante bodega']]);

  const total = card('¿Cuánto se entregó de lo que se produjo?',
    `Hoja CONTEO FINAL del Corte ECOS: se produjeron <b>${fmt(P)}</b> pases y candados, se entregaron <b>${fmt(E)}</b> (${pct(E, P)}) y quedaron <b>${fmt(S)}</b> en bodega.`,
    `<div class="hb-t total" ${tipAttr(`<b>Total ECOS</b>${tipRow('Producido', fmt(P))}${tipRow('Entregado', `${fmt(E)} (${pct(E, P)})`)}${tipRow('Sobrante bodega', fmt(S))}`)}>
      <i class="c-ent" style="width:${E / (E + S) * 100}%"></i><i class="c-sob" style="width:${S / (E + S) * 100}%"></i></div>`,
    key, 'wide');

  const shown = all ? rows : rows.slice(0, TOP_HOSPITALITY);
  const toggle = rows.length > TOP_HOSPITALITY
    ? `<button class="btn more" data-all>${all ? `Ver solo las ${TOP_HOSPITALITY} principales` : `Ver las ${rows.length} categorías`}</button>` : '';
  const bars = card('Entregado y sobrante por categoría',
    'Cada barra es una categoría de la hoja CONTEO FINAL; la parte verde es lo entregado y la dorada lo que quedó en bodega. El número es lo producido.',
    hbars(shown.map(r => ({
      label: r.cat, value: fmt(r.p), segs: [{ v: r.e, tone: 'ent' }, { v: r.s, tone: 'sob' }],
      tip: `<b>${esc(r.cat)}</b>${tipRow('Producido', fmt(r.p))}${tipRow('Entregado', `${fmt(r.e)} (${pct(r.e, r.p)})`)}${tipRow('Sobrante bodega', fmt(r.s))}`
    }))) + toggle,
    key, 'wide');

  const points = dailyPoints(book);
  const days = card('Entregas por día',
    `Suma de CANTIDAD de las hojas CORTE DIARIO PASES y CANDADOS, agrupada por FECHA: <b>${fmt(sum(points, 'v'))}</b> en ${plural(points.length, 'día')}.`,
    `<div class="chart">${columns(points, width)}</div>`, '', 'wide');

  return total + bars + days;
}

/* ---------- Tooltip compartido: cualquier elemento con data-tip ---------- */

export function initTooltip() {
  const tip = $('#tip');
  document.addEventListener('pointermove', e => {
    const target = e.target.closest?.('[data-tip]');
    tip.classList.toggle('on', !!target);
    if (!target) return;
    tip.innerHTML = target.dataset.tip;
    const x = Math.min(e.clientX + 14, innerWidth - tip.offsetWidth - 8), y = e.clientY - tip.offsetHeight - 12;
    tip.style.left = `${x}px`;
    tip.style.top = `${y < 8 ? e.clientY + 18 : y}px`;
  });
}
