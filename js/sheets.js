// Paso 3: tarjetas de resumen por hoja y desglose con las tablas del Excel
import { GLOSSARY, GROUPS, OPERATIVE } from './config.js';
import { $, clone, esc, fmt, pct, norm, num, plural } from './utils.js';
import { colOf, cellText, isWarehouse, sheetGroup, groupOf, extraRows, returnedCol, returnedProducedCol } from './excel.js';

/* ---------- búsqueda ---------- */

// Datos sensibles que no se muestran ni se buscan: nombres de personas (OBSERVACIONES también los trae) y folios entregados
const SENSITIVE = /^(SOLICITANTE|RESPONSABLE|OBSERVACIONES|FOLIOS?( INICIAL| FINAL)?|F\. ?(INICIAL|FINAL))$/;
const visibleCols = table => table.cols.filter(c => !SENSITIVE.test(norm(c.name)));

// Filas que coinciden con la búsqueda (si coincide el título del bloque, se muestran todas)
const matchRows = (table, q) => !q || norm(table.meta.title).includes(q)
  ? table.rows
  : table.rows.filter(row => norm(visibleCols(table).map(col => cellText(row[col.c], col)).join(' ')).includes(q));

const matchRaw = (sheet, q) => sheet.raw.filter(row => !q || norm(row.join(' ')).includes(q));

const matchCount = (sheet, q) => sheet.tables.length ? sheet.tables.reduce((n, t) => n + matchRows(t, q).length, 0) : matchRaw(sheet, q).length;

// Resalta la búsqueda (solo si normalizar no cambió la longitud, para no desfasar el corte)
function highlight(text, q) {
  const i = norm(text).indexOf(q);
  if (!q || !text || i < 0 || norm(text).length !== text.length) return esc(text);
  return `${esc(text.slice(0, i))}<mark>${esc(text.slice(i, i + q.length))}</mark>${esc(text.slice(i + q.length))}`;
}

/* ---------- tarjetas de hoja ---------- */

function sheetSummary(sheet, isReturns) {
  let prod = 0, extra = 0, ent = 0, bod = 0, dev = 0, devMatched = 0, devProd = 0, rows = 0, ranges = 0, outflow = false;
  for (const table of sheet.tables) {
    rows += table.rows.length;
    const qtyCol = table.cols.find(c => norm(c.name) === 'CANTIDAD'), solCol = colOf(table, /^SOLICITANTE$/);
    if (isReturns) {
      const devCol = returnedCol(table), devProdCol = returnedProducedCol(table);
      for (const r of table.rows) {
        const returned = devCol ? num(r[devCol.c]) || 0 : 0, produced = devProdCol ? num(r[devProdCol.c]) || 0 : 0;
        dev += returned;
        // El % devuelto solo compara filas que traen CANTIDAD PRODUCIDA
        if (produced) { devMatched += returned; devProd += produced; }
      }
      continue;
    }
    const prodCol = colOf(table, /^(TOTAL )?PRODUCIDO$/), deliveredCol = colOf(table, /^TOTAL (ACREDITACIONES|ENTREGADO)$/) || qtyCol;
    if (prodCol) table.rows.forEach(r => prod += num(r[prodCol.c]) || 0);
    else if (table.meta.nums.length) {
      prod += table.meta.nums[0];
      if (table.meta.extra) extra += table.meta.nums[0];
    }
    // Tabla de folios (FOLIO INICIAL / FINAL sin entregas): cuenta rangos, no entregas
    if (!deliveredCol && colOf(table, /^FOLIO INICIAL$/)) ranges += table.rows.length;
    if (!deliveredCol || !(prodCol || table.meta.nums.length)) continue; // solo bloques con cantidad producida
    // Sin TOTAL ACREDITACIONES (Corte Diario) la CANTIDAD es lo que salió de bodega, no lo acreditado
    if (deliveredCol === qtyCol) outflow = true;
    for (const r of table.rows) {
      const v = num(r[deliveredCol.c]) || 0;
      if (solCol && isWarehouse(r[solCol.c])) bod += v; else ent += v;
    }
  }
  return { prod, extra, ent, bod, dev, devMatched, devProd, rows, ranges, outflow, blocks: sheet.tables.length };
}

// Cifra principal de la tarjeta según lo que tenga la hoja
function headline(sheet, s) {
  if (sheet.headline) return { ...sheet.headline }; // hojas armadas a mano (Reporte Global)
  if (!sheet.tables.length) {
    const items = sheet.free?.items ?? [];
    return items.length
      ? { value: items.length, unit: 'cifras', detail: items.slice(0, 2).map(i => `${i.label} ${fmt(i.value)}`).join(' · ') }
      : { value: sheet.raw.length, unit: 'datos', detail: 'Hoja sin tabla' };
  }
  if (s.dev) return { value: s.dev, unit: 'devueltos', detail: s.devProd
    ? `${fmt(s.devMatched)} de ${fmt(s.devProd)} producidas (${pct(s.devMatched, s.devProd)})${s.dev > s.devMatched ? ` · ${fmt(s.dev - s.devMatched)} sin dato de producido` : ''}`
    : plural(s.rows, 'registro') };
  // En bodega = producido − lo que salió (filas BODEGA + folios que el Excel no registra)
  const stock = s.outflow ? Math.max(0, s.prod - s.ent) : s.bod;
  if (s.ent && s.prod) return {
    value: s.ent, unit: s.outflow ? 'salieron de bodega' : 'entregadas', share: Math.min(100, s.ent / s.prod * 100),
    detail: `${pct(s.ent, s.prod)} de ${fmt(s.prod)} producidas${stock ? ` · ${fmt(stock)} en bodega` : ''}`
  };
  if (s.ranges) return { value: s.ranges, unit: 'rangos de folios', detail: 'Folio inicial y final de cada categoría' };
  if (s.prod) return { value: s.prod, unit: 'producidas', detail: plural(s.rows, 'fila') };
  return { value: s.rows, unit: s.rows ? 'filas' : '', detail: s.rows ? '' : 'Sin movimientos' };
}

function sheetCard(sheet, index, book, state) {
  const q = norm(state.query), open = !q && index === state.sheet;
  const summary = sheetSummary(sheet, book.kind === 'returns'), info = headline(sheet, summary);
  if (q) {
    const matches = matchCount(sheet, q);
    info.detail = matches ? `${plural(matches, 'coincidencia')} con “${state.query}”` : 'Sin coincidencias';
    info.dim = !matches;
  }
  const card = clone('tpl-sheet', {
    name: sheet.name, blocks: `${summary.blocks} bloques`, value: fmt(info.value), unit: info.unit,
    extra: `Incluye ${fmt(summary.extra)} de extras / 2ª producción`,
    detail: info.detail, action: open ? 'Abierta ↓' : 'Ver desglose →'
  });
  card.dataset.sheet = index;
  card.classList.toggle('on', open);
  card.classList.toggle('dim', !!info.dim);
  card.querySelector('[data-f="blocks"]').hidden = summary.blocks <= 1;
  card.querySelector('[data-f="extra"]').hidden = !summary.extra;
  const meter = card.querySelector('.meter');
  meter.hidden = info.share == null;
  meter.firstElementChild.style.width = `${info.share}%`;
  return card;
}

// Tarjetas agrupadas por tipo de acreditación. Cada tipo es un recuadro; los recuadros se acomodan lado a lado
// (crecen según cuántas hojas tienen) para no dejar huecos. Si todas son del mismo tipo, van sin recuadro.
export function renderSheets(book, state) {
  // Solo el Corte Diario y Devolución tienen una hoja por tipo; el Corte Final mezcla tipos en cada hoja
  // y el Corte Hospitalidades es todo de un tipo, así que esos no se agrupan
  const grouped = ['daily', 'returns'].includes(book.kind);
  const indexed = book.sheets.map((sheet, i) => ({ sheet, i, group: grouped ? sheetGroup(sheet) : GROUPS[0] }));
  const sections = GROUPS.map(group => ({ group, items: indexed.filter(x => x.group === group) })).filter(s => s.items.length);
  const cards = items => items.map(({ sheet, i }) => sheetCard(sheet, i, book, state));
  const boxes = sections.length > 1 ? sections.map(({ group, items }) => {
    const box = clone('tpl-sheet-group', { name: group.name, count: plural(items.length, 'hoja') });
    box.style.setProperty('--n', items.length);
    box.querySelector('.sg-cards').replaceChildren(...cards(items));
    return box;
  }) : [Object.assign(document.createElement('div'), { className: 'sg-cards' })];
  if (sections.length === 1) boxes[0].replaceChildren(...cards(sections[0].items));
  $('#sheets').replaceChildren(...boxes);
  $('#sheetInfo').textContent = state.query ? `Resultados de “${state.query}” en ${book.name}` : `${book.name} · ${plural(book.sheets.length, 'hoja')}`;
}

/* ---------- desglose: las columnas cambian con cada Excel, por eso se arman aquí ---------- */

// Total entregado de un bloque del Corte Diario:
// entregado = CANTIDAD sin las filas BODEGA; en bodega = producido − entregado (filas BODEGA + folios sin registrar)
function ownTotal(table) {
  const qtyCol = colOf(table, /^CANTIDAD$/), solCol = colOf(table, /^SOLICITANTE$/), produced = table.meta.nums[0];
  if (!qtyCol || !solCol || !produced || !table.rows.length || table.side) return null;
  const delivered = table.rows.reduce((t, r) => t + (isWarehouse(r[solCol.c]) ? 0 : num(r[qtyCol.c]) || 0), 0);
  const hasWarehouse = table.rows.some(r => isWarehouse(r[solCol.c]));
  return { qtyCol, stockCol: colOf(table, /^EXISTENCIA/), delivered, warehouse: Math.max(0, produced - delivered), hasWarehouse };
}

// Distintivo por categoría: hospitalidad (plata) u operación (oro); Montaje no lleva.
// En el Corte Hospitalidades y en hojas de hospitalidades todo es hospitalidad salvo staff, talento e invitados zona verde.
const KINDS = { hosp: 'Hospitalidad', oper: 'Operación' };
function kindOf(name, allHosp) {
  if (!norm(name)) return null;
  if (allHosp) return OPERATIVE.test(norm(name)) ? 'oper' : 'hosp';
  const group = groupOf(name);
  return group.book ? 'hosp' : group.pattern ? null : 'oper';
}
const kindDot = kind => kind ? `<i class="kind ${kind}" title="${KINDS[kind]}"></i>` : '';

function tableCard(table, sheet, q, book) {
  const rows = matchRows(table, q);
  if (!rows.length && (q || table.rows.length)) return '';
  const { meta } = table, cols = visibleCols(table);
  const solCol = colOf(table, /^SOLICITANTE$/), catCol = colOf(table, /^CATEGORIA$/);
  const inWarehouse = row => solCol && isWarehouse(row[solCol.c]);
  const extras = extraRows(table);
  const allHosp = book.kind === 'hospitality' || !!sheetGroup(sheet).book;
  const marked = catCol && book.kind !== 'ecos' && !/ECOS/.test(norm(sheet.name)); // ECOS no es hospitalidad ni operación
  const kinds = new Map(marked ? rows.map(r => [r, kindOf(extras.get(r) ?? r[catCol.c], allHosp)]) : []);
  const used = Object.keys(KINDS).filter(k => [...kinds.values()].includes(k));
  // Se muestra si el Excel no trae fila de TOTALES o si la suya incluye filas BODEGA
  const own = ownTotal(table), showOwn = own && (!table.total || own.hasWarehouse);
  const chips = [
    ...(meta.extra ? [`<span class="chip amber">${meta.extra}</span>`] : []),
    ...meta.nums.map(n => `<span class="chip cyan">Producido ${fmt(n)}</span>`),
    ...used.map(k => `<span class="chip kind-chip">${kindDot(k)}${KINDS[k]}</span>`),
    ...(rows.some(inWarehouse) ? ['<span class="chip warehouse">Filas BODEGA = quedaron en bodega, no acreditadas</span>'] : []),
    ...(rows.some(r => extras.has(r)) ? ['<span class="chip extra">Extra = rango de folios adicional (otra producción) de la categoría de arriba</span>'] : []),
    ...meta.notes.map(n => n === '✅' ? '<span class="chip green">✅ Completo</span>' : `<span class="chip amber">${esc(n)}</span>`)
  ].join('');
  const firstText = cols.find(c => !c.n && !c.date && !c.code);
  const isIndex = c => /^NO\b|^ID$/.test(norm(c.name)); // número de fila o de solicitud, no un folio
  const cellClass = c => [c.n ? 'n num' : c.date ? 'd' : isIndex(c) ? 'mono idx' : c.code ? 'mono' : '', c === firstText ? 'name' : ''].join(' ').trim();
  const header = c => {
    const meaning = GLOSSARY[norm(c.name)];
    return `<th class="${c.n ? 'n' : ''}"${meaning ? ` title="${esc(meaning)}"` : ''}>${esc(c.name)}</th>`;
  };
  const rowClass = r => inWarehouse(r) ? ' class="warehouse"' : extras.has(r) ? ' class="extra"' : '';
  const cellHtml = (r, c) => (c === catCol ? kindDot(kinds.get(r)) : '') + (c === catCol && extras.has(r)
    ? `<span class="chip extra">Extra</span> <span class="extra-of">de ${esc(extras.get(r))}</span>`
    : highlight(cellText(r[c.c], c), q));
  const body = rows.map(r => `<tr${rowClass(r)}>${cols.map(c => `<td class="${cellClass(c)}">${cellHtml(r, c)}</td>`).join('')}</tr>`).join('')
    || `<tr><td colspan="${cols.length}" class="empty">Sin movimientos registrados</td></tr>`;
  // Fila de TOTALES del Excel y, debajo, el total entregado sin BODEGA
  const ownCell = c => c === cols[0] ? 'TOTAL ENTREGADO (sin BODEGA)'
    : c === own.qtyCol ? fmt(own.delivered) : c === own.stockCol ? fmt(own.warehouse) : '';
  const footRow = cell => `<tr>${cols.map(c => `<td class="${c.n ? 'n num' : ''}">${esc(cell(c))}</td>`).join('')}</tr>`;
  const footRows = [...(table.total ? [footRow(c => cellText(table.total[c.c], c))] : []), ...(showOwn ? [footRow(ownCell)] : [])];
  const foot = footRows.length && !q ? `<tfoot>${footRows.join('')}</tfoot>` : '';

  return `<section class="card lg">
    <div class="head">
      <div><h2>${esc(meta.title || sheet.name)}</h2>${chips ? `<div class="chips">${chips}</div>` : ''}</div>
      <span class="chip slate">${rows.length ? plural(rows.length, 'fila') : 'Sin movimientos'}</span>
    </div>
    <div class="tbl"><table><thead><tr>${cols.map(header).join('')}</tr></thead><tbody>${body}</tbody>${foot}</table></div>
  </section>`;
}

function rawCard(sheet, q) {
  const rows = matchRaw(sheet, q);
  if (!rows.length) return q ? '' : '<section class="card lg"><div class="empty">Esta hoja está vacía.</div></section>';
  // Pares nombre → número: se muestran como tarjetas (como se leen en el Excel, el número debajo del nombre)
  const free = sheet.free;
  if (free?.items.length) {
    const items = free.items.filter(i => !q || norm(`${i.label} ${fmt(i.value)}`).includes(q) || norm(free.title).includes(q));
    if (!items.length) return '';
    return `<section class="card lg"><div class="head"><div><h2>${esc(free.title || sheet.name)}</h2>
      <p class="how">Esta hoja no tiene tabla: cada cifra es el número que aparece debajo de su nombre en el Excel.</p>
      ${free.notes.length ? `<div class="chips">${free.notes.map(n => `<span class="chip amber">${esc(n)}</span>`).join('')}</div>` : ''}</div></div>
      <div class="stats">${items.map(i => `<div class="stat"><span class="label">${highlight(i.label, q)}</span><b class="num">${highlight(fmt(i.value), q)}</b></div>`).join('')}</div></section>`;
  }
  const cell = v => typeof v === 'number' ? `<td class="n num">${highlight(fmt(v), q)}</td>` : `<td>${highlight(String(v), q)}</td>`;
  return `<section class="card lg"><div class="head"><h2>${esc(sheet.name)}</h2></div>
    <div class="tbl"><table><tbody>${rows.map(r => `<tr>${r.map(cell).join('')}</tr>`).join('')}</tbody></table></div></section>`;
}

export function renderBody(book, state) {
  const q = norm(state.query), sheet = book.sheets[state.sheet];
  $('#hint').hidden = !!(q || sheet);
  $('#detail').hidden = !!q || !sheet;
  if (sheet) $('#detailName').textContent = sheet.name;
  if (!q && !sheet) return $('#body').replaceChildren();

  const html = (q ? book.sheets : [sheet]).map(s => {
    // Tablas laterales (cálculos auxiliares del Excel) van plegadas al final para no confundirse con los movimientos
    const main = s.tables.filter(t => !t.side).map(t => tableCard(t, s, q, book)).join('');
    const side = s.tables.filter(t => t.side).map(t => tableCard(t, s, q, book)).filter(Boolean);
    const aux = side.length ? `<details class="aux"${q ? ' open' : ''}><summary>Cálculos auxiliares del Excel (${side.length}) <span class="sub">tablas de apoyo al lado de los movimientos, como los sobrantes</span></summary>${side.join('')}</details>` : '';
    const part = s.tables.length ? main + aux : rawCard(s, q);
    return part && q ? `<div class="label sheet-label">Hoja ${esc(s.name)}</div>${part}` : part;
  }).join('');
  $('#body').innerHTML = html || `<section class="card lg"><div class="empty">No encontré “${esc(state.query)}” en ${esc(book.name)}.</div></section>`;
}
