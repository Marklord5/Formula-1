// Panorama global del sistema (Excel Reporte Global): 4 totales y 3 gráficas simples al abrir la página
import { GLOBAL } from './config.js';
import { $, fill, fmt, pct, plural, norm, num, pretty, parseDate, shortDate, weekday } from './utils.js';
import { hbars } from './charts.js';

const TOP = 6;
const SCAN = [['PASA', 'ent', 'Pasa'], ['NO ESTA EN BASE', 'sob', 'No está en base'], ['NO PASA', 'dev', 'No pasa']];

// Filas de la primera hoja cuyo nombre coincide, con las llaves normalizadas (FECHA_CORTA, ESTATUS…)
function rowsOf(wb, pattern) {
  const name = wb.SheetNames.find(n => pattern.test(norm(n)));
  if (!name) return [];
  return XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '', raw: false })
    .map(row => Object.fromEntries(Object.entries(row).map(([k, v]) => [norm(k), norm(v)])));
}

// Conteo por valor de una columna, de mayor a menor: [[valor, n]]
function countBy(rows, key) {
  const counts = new Map();
  rows.forEach(r => r[key] && counts.set(r[key], (counts.get(r[key]) || 0) + 1));
  return [...counts].sort((a, b) => b[1] - a[1]);
}

const top = (pairs, tone, unit) => hbars(pairs.slice(0, TOP).map(([label, v]) => ({
  label: pretty(label), value: fmt(v), segs: [{ v, tone }], tip: `<b>${pretty(label)}</b><div>${fmt(v)} ${unit}</div>`
})));

const card = (title, sub, body) => `<section class="card lg"><div class="head"><div><h2>${title}</h2><p class="how">${sub}</p></div></div>${body}</section>`;

// Lee el Excel una sola vez: lo usan el panorama de arriba y la pestaña del Paso 1. Si falla, devuelve el Error.
export const loadGlobal = stamp => fetch(GLOBAL.url + stamp)
  .then(res => { if (!res.ok) throw new Error(`No encontré ${GLOBAL.url}`); return res.arrayBuffer(); })
  .then(buf => XLSX.read(buf, { dense: true }))
  .catch(err => err);

export async function renderGlobal(loading) {
  const box = $('#global');
  try {
    const wb = await loading;
    if (wb instanceof Error) throw wb;
    const requests = rowsOf(wb, /SOLICITUDES/).filter(r => r.ID_SOLICITUD), scans = rowsOf(wb, /ESCANEO/), assigned = rowsOf(wb, /ASIGNADO/);

    const authorized = requests.filter(r => r.ESTATUS === 'AUTORIZADO');
    const requested = authorized.reduce((t, r) => t + (num(r['TOTAL POR SOLICITANTE']) || 0), 0);
    const given = assigned.filter(r => r.ESTATUS === 'ASIGNADO');
    const result = Object.fromEntries(countBy(scans, 'AUTORIZADO'));
    const passed = result.PASA || 0;

    fill(box, {
      requests: fmt(requests.length),
      authorized: `${fmt(authorized.length)} autorizadas`,
      rejected: `${fmt(requests.filter(r => r.ESTATUS === 'RECHAZADO').length)} rechazadas`,
      requested: fmt(requested),
      assigned: fmt(given.length),
      modules: plural(countBy(given, 'MODULO').length, 'módulo'),
      scans: fmt(scans.length),
      passShare: `${pct(passed, scans.length)} pasan`,
      folios: `${fmt(new Set(scans.map(r => r.FOLIO)).size)} folios distintos`
    });

    // Escaneos por día, partidos por resultado
    const byDay = new Map();
    for (const r of scans) {
      const day = byDay.get(r.FECHA_CORTA) ?? byDay.set(r.FECHA_CORTA, {}).get(r.FECHA_CORTA);
      day[r.AUTORIZADO] = (day[r.AUTORIZADO] || 0) + 1;
    }
    const days = [...byDay].map(([d, c]) => ({ date: parseDate(d), c })).filter(d => d.date).sort((a, b) => a.date - b.date);
    const scanRow = (label, c) => {
      const total = SCAN.reduce((t, [k]) => t + (c[k] || 0), 0);
      return { label, value: fmt(total), segs: SCAN.map(([k, tone]) => ({ v: c[k] || 0, tone })),
        tip: `<b>${label}</b>${SCAN.map(([k, , name]) => `<div>${name}: ${fmt(c[k] || 0)} (${pct(c[k] || 0, total)})</div>`).join('')}` };
    };
    const legend = `<div class="legend">${SCAN.map(([k, tone, name]) => `<span><i class="c-${tone}"></i>${name} · ${fmt(result[k] || 0)}</span>`).join('')}</div>`;

    $('#globalCharts').innerHTML =
      card('Escaneos en accesos', 'Cada lectura de gafete en una puerta, por día y resultado',
        hbars(days.map(d => scanRow(`${weekday(d.date)} ${shortDate(d.date)}`, d.c))) + legend) +
      card('Accesos con más escaneos', `Top ${TOP} de ${plural(countBy(scans, 'NOMBRE_ACCESO').length, 'acceso')}`,
        top(countBy(scans, 'NOMBRE_ACCESO'), 'ent', 'escaneos')) +
      card('Categorías más asignadas', `Top ${TOP} de ${fmt(given.length)} acreditaciones asignadas`,
        top(countBy(given, 'CATEGORIA'), 'alt', 'asignadas'));
    box.dataset.state = 'ready';
  } catch (err) {
    $('#globalCharts').innerHTML = `<section class="card lg"><div class="empty">No pude leer el Reporte Global (${err.message}). El resto del reporte sigue funcionando.</div></section>`;
    box.dataset.state = 'error';
  }
}

/* ---------- pestaña del Paso 1: tablas de resumen (sin nombres ni folios, que son datos sensibles) ---------- */

const RESULTS = SCAN.map(([k]) => k);
const isoToDay = v => v.replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$3/$2/$1'); // 2026-10-30 → 30/10/2026

// Tabla en el formato que pinta sheets.js; la primera columna es la etiqueta y las demás son números
function table(title, head, rows, note) {
  const cols = head.map((name, c) => ({ c, name, n: c > 0, date: c === 0 && /FECHA/.test(name) }));
  const total = ['TOTAL', ...head.slice(1).map((_, i) => rows.reduce((t, r) => t + r[i + 1], 0))];
  return { cols, rows, total, side: false, meta: { title, nums: [], folios: [], notes: note ? [note] : [], extra: '' } };
}

// Conteo cruzado: una fila por valor de `key`, una columna por cada valor de `by` y el total
function crossTab(rows, key, by, values, byDate = false) {
  const map = new Map();
  for (const r of rows) {
    if (!r[key]) continue;
    const counts = map.get(r[key]) ?? map.set(r[key], values.map(() => 0)).get(r[key]);
    const i = values.indexOf(r[by]);
    if (i >= 0) counts[i]++;
  }
  const out = [...map].map(([k, counts]) => [k, ...counts, counts.reduce((a, b) => a + b, 0)]);
  return byDate ? out.sort((a, b) => parseDate(a[0]) - parseDate(b[0])) : out.sort((a, b) => b.at(-1) - a.at(-1));
}

const count = (rows, key, byDate) => crossTab(rows.map(r => ({ ...r, _: 1 })), key, '_', [1], byDate).map(([k, n]) => [k, n]);
const sheetOf = (name, headline, tables) => ({ name, headline, tables, raw: [], origin: { r: 0, c: 0 } });

// Cada hoja del Reporte Global resumida en tablas de conteo
export function globalBook(wb) {
  const requests = rowsOf(wb, /SOLICITUDES/).filter(r => r.ID_SOLICITUD);
  const scans = rowsOf(wb, /ESCANEO/), assigned = rowsOf(wb, /ASIGNADO/), segments = rowsOf(wb, /SEGMENTACION/);
  assigned.forEach(r => r.DIA = isoToDay(r['FECHA ASIGNADO']));

  // Solicitudes: las columnas entre AREAS y TOTAL POR SOLICITANTE son las categorías pedidas
  const keys = Object.keys(requests[0] ?? {});
  const categories = keys.slice(keys.indexOf('AREAS') + 1, keys.indexOf('TOTAL POR SOLICITANTE'));
  const authorized = requests.filter(r => r.ESTATUS === 'AUTORIZADO');
  const requestedBy = categories.map(c => [c, authorized.reduce((t, r) => t + (num(r[c]) || 0), 0)]).filter(([, n]) => n).sort((a, b) => b[1] - a[1]);
  const statuses = [...new Set(requests.map(r => r.ESTATUS).filter(Boolean))];
  const byStatus = statuses.map(st => {
    const own = requests.filter(r => r.ESTATUS === st);
    return [st, own.length, own.reduce((t, r) => t + (num(r['TOTAL POR SOLICITANTE']) || 0), 0)];
  }).sort((a, b) => b[1] - a[1]);

  // Asignado: MODULO a veces trae el nombre o correo de quien asignó; esos se juntan en un solo renglón
  const given = assigned.filter(r => r.ESTATUS === 'ASIGNADO');
  given.forEach(r => r.PUNTO = /^(MODULO|REGISTRO EXPRES)/.test(r.MODULO) ? r.MODULO : 'OTRO RESPONSABLE (DATO PERSONAL, NO SE MUESTRA)');
  const assignStatus = [...new Set(assigned.map(r => r.ESTATUS).filter(Boolean))];

  const scanHead = name => [name, 'PASA', 'NO ESTA EN BASE', 'NO PASA', 'TOTAL'];
  const passed = scans.filter(r => r.AUTORIZADO === 'PASA').length;
  return {
    kind: 'global', name: GLOBAL.name, url: GLOBAL.url,
    sheets: [
      sheetOf('Solicitudes', { value: requests.length, unit: 'solicitudes', detail: `${fmt(authorized.length)} autorizadas · ${fmt(requestedBy.reduce((t, r) => t + r[1], 0))} acreditaciones pedidas` }, [
        table('Solicitudes por estatus', ['ESTATUS', 'SOLICITUDES', 'ACREDITACIONES SOLICITADAS'], byStatus),
        table('Acreditaciones solicitadas por categoría', ['CATEGORIA', 'CANTIDAD SOLICITADA'], requestedBy, 'Solo solicitudes autorizadas'),
        table('Solicitudes por tipo', ['TIPO', 'SOLICITUDES'], count(requests, 'NOMBRE_TIPO')),
        table('Solicitudes por día', ['FECHA', 'SOLICITUDES'], count(requests, 'FECHA_CORTA', true))
      ]),
      sheetOf('Asignaciones', { value: given.length, unit: 'asignadas', detail: `${plural(new Set(given.map(r => r.CATEGORIA)).size, 'categoría')} · a una persona cada una` }, [
        table('Acreditaciones por categoría y estatus', ['CATEGORIA', ...assignStatus, 'TOTAL'], crossTab(assigned, 'CATEGORIA', 'ESTATUS', assignStatus)),
        table('Asignadas por día', ['FECHA', 'ASIGNADAS'], count(given, 'DIA', true)),
        table('Asignadas por módulo', ['MODULO', 'ASIGNADAS'], count(given, 'PUNTO'))
      ]),
      sheetOf('Escaneos en accesos', { value: scans.length, unit: 'escaneos', share: passed / (scans.length || 1) * 100, detail: `${pct(passed, scans.length)} pasan · ${plural(new Set(scans.map(r => r.NOMBRE_ACCESO)).size, 'acceso')}` }, [
        table('Escaneos por día', scanHead('FECHA'), crossTab(scans, 'FECHA_CORTA', 'AUTORIZADO', RESULTS, true), 'Cada escaneo es una lectura de gafete en una puerta'),
        table('Escaneos por acceso', scanHead('ACCESO'), crossTab(scans, 'NOMBRE_ACCESO', 'AUTORIZADO', RESULTS)),
        table('Escaneos por categoría', scanHead('CATEGORIA'), crossTab(scans, 'NOMBRE_CATEGORIA', 'AUTORIZADO', RESULTS))
      ]),
      sheetOf('Segmentación por módulo', { value: segments.length, unit: 'personas', detail: `${plural(new Set(segments.map(r => r.NOMBRE_MODULO)).size, 'módulo')}` }, [
        table('Personas por módulo', ['MODULO', 'PERSONAS'], count(segments, 'NOMBRE_MODULO'))
      ])
    ]
  };
}
