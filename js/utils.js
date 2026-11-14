export const DAY_MS = 864e5;
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const MONTHS = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
const WEEKDAYS = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
const HTML_ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };

/* ---------- valores de celda ---------- */
export const norm = v => String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
export const blank = v => v == null || typeof v === 'boolean' || (typeof v === 'string' && !v.trim());
export const nothing = v => blank(v) || v === 0;
export const clean = v => typeof v === 'string' ? v.replace(/\s+/g, ' ').trim() : v;
// Número de una celda. Acepta números guardados como texto ("90", "005", "1,390") y con asterisco de nota ("390*")
const NUM_TEXT = /^\s*(-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)\s*\**\s*$/;
export const num = v => typeof v === 'number' ? v
  : typeof v === 'string' && NUM_TEXT.test(v) ? +v.match(NUM_TEXT)[1].replace(/,/g, '') : NaN;

/* ---------- formato ---------- */
export const esc = v => String(v ?? '').replace(/[&<>"]/g, c => HTML_ESCAPES[c]);
export const fmt = n => (Math.round((n || 0) * 100) / 100).toLocaleString('es-MX');
export const pct = (part, total) => total ? `${(part / total * 100).toFixed(1)}%` : '—';
export const plural = (n, word) => `${n.toLocaleString('es-MX')} ${word}${n === 1 ? '' : 's'}`;
export const pretty = v => String(v).trim().toLowerCase().replace(/(^|[\s(/-])\S/g, m => m.toUpperCase());
export const fileName = url => url.split('/').pop();
export const sum = (items, field) => items.reduce((total, item) => total + item[field], 0);

/* ---------- fechas (siempre UTC) ---------- */
export const shortDate = d => `${d.getUTCDate()} ${MONTHS[d.getUTCMonth()]}`;
export const weekday = d => WEEKDAYS[d.getUTCDay()];

// Número de serie de Excel → Date
export const serialDate = v => typeof v === 'number' && v > 30000 && v < 80000 ? new Date(EXCEL_EPOCH + Math.round(v) * DAY_MS) : null;

// Serie de Excel o texto dd-mm-aa(aa) → Date
export function parseDate(v) {
  const match = typeof v === 'string' && v.trim().match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/);
  if (!match) return serialDate(v);
  const [, day, month, year] = match;
  return new Date(Date.UTC(year.length === 2 ? 2000 + +year : +year, month - 1, +day));
}

/* ---------- DOM ---------- */
export const $ = selector => document.querySelector(selector);

// Escribe cada valor como texto en los elementos [data-f="clave"] dentro de root
export function fill(root, data) {
  for (const [key, value] of Object.entries(data))
    root.querySelectorAll(`[data-f="${key}"]`).forEach(el => el.textContent = value);
  return root;
}

// Clona un <template id="..."> y lo rellena
export const clone = (id, data) => fill(document.getElementById(id).content.firstElementChild.cloneNode(true), data);
