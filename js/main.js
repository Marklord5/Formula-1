// Punto de entrada: estado, navegación y carga de los Excel
import { FILES } from './config.js';
import { $, fill, esc } from './utils.js';
import { parseBook } from './excel.js';
import { chartsHtml, initTooltip } from './charts.js';
import { renderSheets, renderBody } from './sheets.js';
import { renderKpis, renderGroups, renderTrend, renderDistribution, renderTabs, renderDownloads } from './overview.js';

let books = [];
const state = { file: 0, sheet: -1, query: '', all: false }; // sheet = hoja abierta (-1: ninguna)
const currentBook = () => books[state.file];
const scrollTo = (selector, block = 'start') => $(selector).scrollIntoView({ behavior: 'smooth', block });

/* ---------- pintado ---------- */

function renderCharts() {
  const book = currentBook();
  $('#chartsFor').textContent = `de ${book.name}`;
  try {
    $('#charts').innerHTML = chartsHtml(book, { all: state.all, width: $('#charts').clientWidth });
  } catch (err) {
    $('#charts').innerHTML = `<section class="card lg wide"><div class="empty">No pude armar las gráficas de este Excel (${esc(err.message)}). Las tablas de abajo siguen funcionando.</div></section>`;
  }
}

function renderSearch() {
  renderSheets(currentBook(), state);
  renderBody(currentBook(), state);
}

function render() {
  renderTabs(books, state);
  renderCharts();
  renderSearch();
  history.replaceState(null, '', `#${state.file}/${state.sheet}`);
}

function go(file, sheet = -1) {
  if (file !== state.file) state.all = false;
  Object.assign(state, { file, sheet, query: '' });
  $('#query').value = '';
  render();
}

/* ---------- eventos ---------- */

function bindEvents() {
  initTooltip();
  $('#btnPrint').addEventListener('click', () => print());
  $('#tabs').addEventListener('click', e => {
    const tab = e.target.closest('[data-file]');
    if (tab) go(+tab.dataset.file);
  });
  $('#kpis').addEventListener('click', e => {
    const kpi = e.target.closest('[data-file]');
    if (!kpi) return;
    go(books.findIndex(b => b.kind === kpi.dataset.file));
    scrollTo('#tabs');
  });
  $('#charts').addEventListener('click', e => {
    if (!e.target.closest('[data-all]')) return;
    state.all = !state.all;
    renderCharts();
  });
  $('#sheets').addEventListener('click', e => {
    const card = e.target.closest('[data-sheet]');
    if (!card) return;
    go(state.file, +card.dataset.sheet);
    scrollTo('#detail');
  });
  $('#back').addEventListener('click', () => scrollTo('#sheets', 'center'));
  $('#query').addEventListener('input', e => { state.query = e.target.value; renderSearch(); });

  let resizeTimer;
  addEventListener('resize', () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => { renderCharts(); renderTrend(books); }, 150);
  });
}

/* ---------- carga ---------- */

function setStatus(status, error = '') {
  fill($('#status'), { error });
  $('#status').dataset.state = status;
}

async function loadBook(file, stamp) {
  const res = await fetch(file.url + stamp);
  if (!res.ok) throw new Error(`No encontré ${file.url} (${res.status}).`);
  return { ...file, sheets: parseBook(XLSX.read(await res.arrayBuffer(), { cellStyles: true })) };
}

// Abre la pestaña y hoja guardadas en la URL (#archivo/hoja)
function restoreFromHash() {
  const [file, sheet] = location.hash.slice(1).split('/').map(Number);
  if (!books[file]) return;
  state.file = file;
  if (books[file].sheets[sheet]) state.sheet = sheet;
}

async function start() {
  const urls = FILES.map(f => f.url);
  fill(document, { requiredFiles: urls.join(', ') });
  renderDownloads(FILES);
  try {
    const stamp = `?v=${Date.now()}`; // evita la caché al actualizar los Excel
    books = await Promise.all(FILES.map(file => loadBook(file, stamp)));
    restoreFromHash();
    renderKpis(books);
    renderGroups(books);
    renderTrend(books);
    renderDistribution(books);
    render();
    bindEvents();
    fill(document, {
      readAt: `Leído ${new Date().toLocaleString('es-MX', { dateStyle: 'medium', timeStyle: 'short' })}`,
      source: `Fuente: ${urls.join(' · ')}`
    });
    setStatus('ready');
  } catch (err) {
    setStatus('error', err.message);
  }
}

start();
