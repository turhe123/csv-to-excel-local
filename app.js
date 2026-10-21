// app.js — CSV to Excel Converter (ES module). Parsing and workbook building run in worker.js;
// this file only handles the page: file cards, options, progress and the download.
import { toast, downloadBlob, formatBytes, wireDropzone } from './site.js';

const MAX_ROWS = 1048576;
const MAX_COLS = 16384;
const PREBUILD_LIMIT = 10 * 1024 * 1024; // build ahead of the click so the download keeps the user gesture
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const ACCEPTED = /\.(csv|tsv|tab|txt)$/i;
const DELIM_NAMES = { ',': 'comma', ';': 'semicolon', '\t': 'tab', '|': 'pipe' };

const zone = document.getElementById('dropzone');
const input = document.getElementById('file-input');
const list = document.getElementById('file-list');
const options = document.getElementById('options');
const results = document.getElementById('results');
const downloadBtn = document.getElementById('download');
const clearBtn = document.getElementById('clear');
const optHeader = document.getElementById('opt-header');
const optText = document.getElementById('opt-text');
const optWidth = document.getElementById('opt-width');

const state = { files: [], version: 0, nextId: 1, nextJob: 1, cached: null, pending: null, savedUrl: null };
const jobs = new Map();
const worker = new Worker('./worker.js');

// ---------- Worker plumbing ----------
worker.onmessage = (e) => {
  const msg = e.data;
  if (msg.jobId) {
    const job = jobs.get(msg.jobId);
    if (!job) return;
    if (msg.type === 'progress') job.onProgress?.(msg.done / msg.total);
    else if (msg.type === 'built') { jobs.delete(msg.jobId); job.resolve(msg); }
    else if (msg.type === 'error') { jobs.delete(msg.jobId); job.reject(new Error(msg.message)); }
    return;
  }
  const f = state.files.find((x) => x.id === msg.id);
  if (!f) return;
  if (msg.type === 'progress') setCardProgress(f, msg.done / msg.total);
  else if (msg.type === 'parsed') {
    f.status = 'ready';
    f.meta = msg.meta;
    renderCard(f);
    changed();
  } else if (msg.type === 'error') {
    showError(msg.message);
    removeFile(f, false);
  }
};
worker.onerror = (e) => { e.preventDefault(); showError('The converter stopped unexpectedly. Reload the page and try again.'); };

function runBuild(onProgress) {
  const jobId = state.nextJob++;
  const version = state.version;
  const ready = state.files.filter((f) => f.status === 'ready');
  const opts = getOptions();
  const promise = new Promise((resolve, reject) => jobs.set(jobId, { resolve, reject, onProgress }))
    .then((msg) => ({ version, name: outputName(), blob: new Blob([msg.buffer], { type: XLSX_MIME }), sheetNames: msg.sheetNames, stats: msg.stats }));
  worker.postMessage({ type: 'build', jobId, options: opts, files: ready.map((f) => ({ id: f.id, name: f.name, sheetName: f.sheetName, split: f.split })) });
  return promise;
}

// ---------- Helpers ----------
const getOptions = () => ({ header: optHeader.checked, keepText: optText.checked, autoWidth: optWidth.checked });
const baseName = (name) => name.replace(/\.[^.]+$/, '') || name;
const fmt = (n) => n.toLocaleString('en-US');

function outputName() {
  const ready = state.files.filter((f) => f.status !== 'error');
  if (ready.length !== 1) return 'workbook.xlsx';
  const safe = baseName(ready[0].name).replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '-').trim() || 'workbook';
  return `${safe}.xlsx`;
}

function sheetBase(name, index) {
  let s = baseName(name).replace(/[\\/:*?[\]\u0000-\u001f]+/g, ' ').replace(/\s+/g, ' ').trim().replace(/^'+|'+$/g, '').slice(0, 31).trim();
  if (!s) s = `Sheet${index + 1}`;
  if (s.toLowerCase() === 'history') s = 'History 1';
  return s;
}

function assignSheetNames() {
  const used = new Set();
  state.files.forEach((f, i) => {
    const base = sheetBase(f.name, i);
    let name = base;
    let n = 2;
    while (used.has(name.toLowerCase())) { const suffix = ` (${n++})`; name = base.slice(0, 31 - suffix.length) + suffix; }
    used.add(name.toLowerCase());
    f.sheetName = name;
  });
}

function parts(meta, header) {
  const rowParts = meta.rows > MAX_ROWS ? Math.ceil((meta.rows - (header ? 1 : 0)) / (MAX_ROWS - (header ? 1 : 0))) : 1;
  return { rowParts, colParts: Math.ceil(meta.cols / MAX_COLS) || 1 };
}

const tooBig = (f) => { const p = parts(f.meta, optHeader.checked); return p.rowParts * p.colParts > 1; };

function el(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') node.className = v;
    else if (k === 'text') node.textContent = v;
    else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
    else if (v !== false && v != null) node.setAttribute(k, v === true ? '' : v);
  }
  for (const c of children) if (c != null) node.append(c);
  return node;
}

function showError(message) {
  const alert = el('div', { class: 'alert alert--error', role: 'alert', text: message });
  results.prepend(alert);
  alert.scrollIntoView({ block: 'nearest' });
}

function clearMessages() {
  results.querySelectorAll('.alert').forEach((a) => a.remove());
}

// ---------- Adding files ----------
function handleFiles(incoming) {
  clearMessages();
  const added = [];
  for (const file of incoming) {
    if (!ACCEPTED.test(file.name) && !/^text\/(csv|tab-separated-values|plain)$/.test(file.type)) {
      showError(`${file.name} was skipped: only CSV, TSV or TXT text files can be converted.`);
      continue;
    }
    if (file.size === 0) {
      showError(`${file.name} is empty, so there is nothing to convert.`);
      continue;
    }
    const f = { id: state.nextId++, file, name: file.name, size: file.size, status: 'loading', delimiter: 'auto', split: false, meta: null };
    state.files.push(f);
    added.push(f);
  }
  if (!added.length) return;
  assignSheetNames();
  renderList();
  added.forEach(readFile);
  added[0].card.scrollIntoView({ block: 'nearest' });
  changed();
}

async function readFile(f) {
  try {
    const buffer = await f.file.arrayBuffer();
    worker.postMessage({ type: 'add', id: f.id, name: f.name, buffer }, [buffer]);
  } catch (err) {
    showError(`${f.name} could not be read: ${err.message}`);
    removeFile(f, false);
  }
}

function removeFile(f, announce = true) {
  state.files = state.files.filter((x) => x !== f);
  worker.postMessage({ type: 'remove', id: f.id });
  f.card?.remove();
  assignSheetNames();
  renderList();
  changed();
  if (announce) toast(`Removed ${f.name}`);
}

function move(f, dir, focusClass) {
  const i = state.files.indexOf(f);
  const j = i + dir;
  if (j < 0 || j >= state.files.length) return;
  [state.files[i], state.files[j]] = [state.files[j], state.files[i]];
  assignSheetNames();
  renderList();
  changed();
  const target = f.card.querySelector(focusClass);
  (target.disabled ? f.card.querySelector(dir < 0 ? '.js-down' : '.js-up') : target)?.focus();
}

// ---------- Rendering ----------
function renderList() {
  state.files.forEach((f, i) => {
    if (!f.card) f.card = buildCard(f);
    if (list.children[i] !== f.card) list.insertBefore(f.card, list.children[i] || null);
    f.card.querySelector('.js-up').disabled = i === 0;
    f.card.querySelector('.js-down').disabled = i === state.files.length - 1;
    f.card.querySelector('.file-card__sheet').textContent = `Sheet: ${f.sheetName}`;
  });
  options.hidden = state.files.length === 0;
}

function buildCard(f) {
  const selId = `delim-${f.id}`;
  const select = el('select', { id: selId, class: 'file-card__delim', onchange: () => {
    f.delimiter = select.value;
    f.status = 'loading';
    f.split = false;
    renderCard(f);
    worker.postMessage({ type: 'reparse', id: f.id, name: f.name, delimiter: f.delimiter });
    changed();
  } },
  el('option', { value: 'auto', text: 'Detect automatically' }),
  el('option', { value: ',', text: 'Comma ( , )' }),
  el('option', { value: ';', text: 'Semicolon ( ; )' }),
  el('option', { value: '\t', text: 'Tab' }),
  el('option', { value: '|', text: 'Pipe ( | )' }));
  const card = el('li', { class: 'file-card', 'data-name': f.name },
    el('div', { class: 'file-card__head' },
      el('div', { class: 'file-card__title' },
        el('div', { class: 'file-card__name', text: f.name, title: f.name }),
        el('div', { class: 'file-card__meta' }),
        el('div', { class: 'file-card__sheet' })),
      el('div', { class: 'file-card__actions' },
        el('button', { type: 'button', class: 'btn btn--small btn--ghost js-up', 'aria-label': `Move ${f.name} up`, title: 'Move up', onclick: () => move(f, -1, '.js-up'), text: '↑' }),
        el('button', { type: 'button', class: 'btn btn--small btn--ghost js-down', 'aria-label': `Move ${f.name} down`, title: 'Move down', onclick: () => move(f, 1, '.js-down'), text: '↓' }),
        el('button', { type: 'button', class: 'btn btn--small btn--ghost btn--danger js-remove', 'aria-label': `Remove ${f.name}`, title: 'Remove', onclick: () => removeFile(f), text: '✕' }))),
    el('div', { class: 'field file-card__field' }, el('label', { for: selId, text: 'Delimiter' }), select),
    el('div', { class: 'file-card__body' }));
  f.card = card;
  renderCard(f);
  return card;
}

function setCardProgress(f, ratio) {
  const bar = f.card?.querySelector('.progress > span');
  if (!bar) return;
  const pct = Math.floor(ratio * 1000) / 10;
  bar.style.width = `${pct}%`;
  f.card.querySelector('.progress').setAttribute('aria-valuenow', String(pct));
  f.card.querySelector('.file-card__status').textContent = `Reading ${pct}%`;
}

function renderCard(f) {
  const card = f.card;
  if (!card) return;
  const metaEl = card.querySelector('.file-card__meta');
  const body = card.querySelector('.file-card__body');
  body.replaceChildren();
  card.classList.toggle('is-loading', f.status === 'loading');
  if (f.status === 'loading') {
    metaEl.textContent = `${formatBytes(f.size)} · reading…`;
    body.append(el('div', { class: 'file-card__status', text: 'Reading…' }),
      el('div', { class: 'progress', role: 'progressbar', 'aria-label': `Reading ${f.name}`, 'aria-valuemin': '0', 'aria-valuemax': '100', 'aria-valuenow': '0' }, el('span')));
    return;
  }
  const m = f.meta;
  const delimName = DELIM_NAMES[m.delimiter] || 'comma';
  const detected = f.delimiter === 'auto' ? (m.guessFailed ? 'single column (no delimiter found)' : `${delimName} detected`) : `${delimName} (manual)`;
  metaEl.textContent = `${detected} · ${fmt(m.rows)} rows · ${fmt(m.cols)} columns · ${m.encoding}`;
  card.querySelector('.file-card__delim').value = f.delimiter;
  card.dataset.delimiter = delimName;
  if (tooBig(f)) {
    const p = parts(m, optHeader.checked);
    const why = [];
    if (p.rowParts > 1) why.push(`${fmt(m.rows)} rows (Excel allows ${fmt(MAX_ROWS)} per sheet)`);
    if (p.colParts > 1) why.push(`${fmt(m.cols)} columns (Excel allows ${fmt(MAX_COLS)})`);
    const n = p.rowParts * p.colParts;
    const box = el('div', { class: `alert ${f.split ? 'alert--ok' : 'alert--warn'} file-card__limit` },
      el('span', { text: f.split ? `Will be split into ${n} sheets: the file has ${why.join(' and ')}.` : `Too large for one sheet: the file has ${why.join(' and ')}.` }));
    if (!f.split) box.append(' ', el('button', { type: 'button', class: 'btn btn--small', text: `Split into ${n} sheets`, onclick: () => { f.split = true; renderCard(f); changed(); } }));
    body.append(box);
  }
  if (m.longCells) body.append(el('div', { class: 'alert alert--warn', text: `${fmt(m.longCells)} cell(s) are longer than the 32,767 characters Excel allows and will be truncated.` }));
  body.append(previewTable(m));
}

function previewTable(m) {
  const header = optHeader.checked;
  const rows = m.preview;
  const table = el('table');
  table.append(el('caption', { class: 'visually-hidden', text: 'First rows of the file' }));
  const width = Math.max(...rows.map((r) => r.length));
  let start = 0;
  if (header && rows.length) {
    table.append(el('thead', {}, el('tr', {}, ...rows[0].map((v) => el('th', { scope: 'col', text: v })))));
    start = 1;
  }
  const tbody = el('tbody');
  rows.slice(start, start + 5).forEach((r) => {
    const tr = el('tr');
    for (let c = 0; c < width; c++) tr.append(el('td', { text: r[c] ?? '' }));
    tbody.append(tr);
  });
  table.append(tbody);
  const more = m.rows - start - Math.min(5, rows.length - start);
  const wrap = el('div', { class: 'table-wrap file-card__preview', tabindex: '0', role: 'region', 'aria-label': 'Preview of the first rows' }, table);
  return more > 0 ? el('div', {}, wrap, el('p', { class: 'file-card__more', text: `+ ${fmt(more)} more row${more === 1 ? '' : 's'}` })) : wrap;
}

// ---------- State changes, prebuild and download ----------
let prebuildTimer = 0;
function changed() {
  state.version++;
  state.cached = null;
  state.pending = null;
  results.querySelector('.convert-result')?.remove();
  const ready = state.files.filter((f) => f.status === 'ready');
  const loading = state.files.some((f) => f.status === 'loading');
  const blocked = ready.filter((f) => tooBig(f) && !f.split);
  downloadBtn.disabled = !ready.length || loading || blocked.length > 0;
  downloadBtn.textContent = `Download ${outputName()}`;
  clearTimeout(prebuildTimer);
  const total = ready.reduce((s, f) => s + f.size, 0);
  if (!downloadBtn.disabled && total <= PREBUILD_LIMIT) {
    prebuildTimer = setTimeout(() => {
      const p = runBuild();
      state.pending = p;
      p.then((r) => { if (r.version === state.version) state.cached = r; }).catch(() => { /* reported on click */ });
    }, 200);
  }
}

function showResult(r, downloaded) {
  results.querySelector('.convert-result')?.remove();
  if (state.savedUrl) URL.revokeObjectURL(state.savedUrl);
  state.savedUrl = URL.createObjectURL(r.blob);
  const notes = [];
  if (r.stats.truncated) notes.push(`${fmt(r.stats.truncated)} long cell(s) were cut to 32,767 characters.`);
  if (r.stats.formulas) notes.push(`${fmt(r.stats.formulas)} value(s) starting with = were kept as text, not formulas.`);
  const card = el('div', { class: 'result convert-result' },
    el('div', { class: 'result__meta' },
      el('div', { class: 'result__name', text: r.name }),
      el('div', { class: 'result__info', text: `${formatBytes(r.blob.size)} · ${r.sheetNames.length} sheet${r.sheetNames.length === 1 ? '' : 's'}: ${r.sheetNames.join(', ')}` }),
      notes.length ? el('div', { class: 'result__info convert-result__note', text: notes.join(' ') }) : null),
    el('a', { class: 'btn', href: state.savedUrl, download: r.name, text: downloaded ? 'Save again' : `Save ${r.name}` }));
  card.prepend(sheetIcon());
  results.append(card);
  card.scrollIntoView({ block: 'nearest' });
}

function sheetIcon() {
  const ns = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(ns, 'svg');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('class', 'convert-result__icon');
  svg.setAttribute('aria-hidden', 'true');
  svg.innerHTML = '<rect x="3" y="3" width="18" height="18" rx="3" fill="none" stroke="currentColor" stroke-width="2"/><path d="M3 9h18M3 15h18M9 3v18" stroke="currentColor" stroke-width="2"/>';
  return svg;
}

async function onDownload() {
  clearMessages();
  // Ready already: download synchronously inside the click so mobile Safari accepts it.
  if (state.cached && state.cached.version === state.version) {
    downloadBlob(state.cached.blob, state.cached.name);
    showResult(state.cached, true);
    toast(`Saved ${state.cached.name}`);
    return;
  }
  const version = state.version;
  downloadBtn.disabled = true;
  const bar = el('div', { class: 'progress', role: 'progressbar', 'aria-label': 'Building the workbook', 'aria-valuemin': '0', 'aria-valuemax': '100', 'aria-valuenow': '0' }, el('span'));
  const label = el('p', { class: 'convert-progress__label', text: 'Building the workbook…' });
  const wrap = el('div', { class: 'convert-progress' }, label, bar);
  results.append(wrap);
  try {
    const p = state.pending && state.version === version ? state.pending : runBuild((ratio) => {
      const pct = Math.floor(ratio * 1000) / 10;
      bar.firstChild.style.width = `${pct}%`;
      bar.setAttribute('aria-valuenow', String(pct));
      label.textContent = `Building the workbook… ${pct}%`;
    });
    const r = await p;
    if (version !== state.version) return; // files changed meanwhile; the button reflects the new state
    state.cached = r;
    downloadBlob(r.blob, r.name);
    showResult(r, true);
    toast(`Saved ${r.name}`);
  } catch (err) {
    showError(`The workbook could not be built: ${err.message}`);
  } finally {
    wrap.remove();
    if (version === state.version) downloadBtn.disabled = false;
  }
}

function clearAll() {
  for (const f of state.files) { worker.postMessage({ type: 'remove', id: f.id }); f.card?.remove(); }
  state.files = [];
  results.replaceChildren();
  if (state.savedUrl) { URL.revokeObjectURL(state.savedUrl); state.savedUrl = null; }
  renderList();
  changed();
  input.focus();
}

function onOptionChange() {
  state.files.forEach((f) => { if (f.status === 'ready') renderCard(f); });
  changed();
}

wireDropzone(zone, input, handleFiles);
downloadBtn.addEventListener('click', onDownload);
clearBtn.addEventListener('click', clearAll);
[optHeader, optText, optWidth].forEach((o) => o.addEventListener('change', onOptionChange));
renderList();

const whenReady = () => new Promise((resolve) => {
  const check = () => (state.files.length && state.files.every((f) => f.status === 'ready') ? resolve() : setTimeout(check, 50));
  check();
});

// Called by the factory's screenshot script: loads two fictional Maple Crumb exports and shows a built workbook.
window.__demo = async () => {
  const load = async (name) => new File([await (await fetch(`./tests/fixtures/${name}`)).blob()], name, { type: 'text/csv' });
  handleFiles([await load('sample-sales.csv'), await load('delivery-routes.csv')]);
  await whenReady();
  const r = await runBuild();
  state.cached = r;
  showResult(r, false);
};

export { handleFiles };
