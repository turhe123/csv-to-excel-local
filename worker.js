// worker.js — parses CSV/TSV files and builds the XLSX workbook off the main thread.
// Classic worker: the vendored libraries are loaded with importScripts. Nothing leaves the device.
/* global Papa, XLSX */
importScripts('vendor/papaparse/papaparse.min.js', 'vendor/xlsx/xlsx.full.min.js');

const MAX_ROWS = 1048576;
const MAX_COLS = 16384;
const MAX_CELL = 32767;
const MAX_WIDTH = 60;
const PREVIEW_ROWS = 6; // header + five rows
const DELIMITERS = [',', ';', '\t', '|'];
const CHUNK = 1 << 20;

const files = new Map(); // id -> { text, rows, cols, delimiter }

const post = (msg, transfer) => self.postMessage(msg, transfer || []);

// ---------- Decoding ----------
function decode(buffer) {
  const bytes = new Uint8Array(buffer);
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return { text: new TextDecoder('utf-16le').decode(bytes), encoding: 'UTF-16' };
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return { text: new TextDecoder('utf-16be').decode(bytes), encoding: 'UTF-16' };
  let text;
  let encoding = 'UTF-8';
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    text = new TextDecoder('windows-1252').decode(bytes);
    encoding = 'Windows-1252';
  }
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  return { text, encoding };
}

// Finds the line of a quoted field that is never closed (papaparse only reports the row index).
function unclosedQuoteLine(text, delimiter) {
  let inQuote = false;
  let start = -1;
  let fieldStart = true;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuote) {
      if (ch === '"') {
        if (text[i + 1] === '"') i++;
        else inQuote = false;
      }
      continue;
    }
    if (fieldStart && ch === '"') { inQuote = true; start = i; fieldStart = false; continue; }
    fieldStart = ch === delimiter || ch === '\n' || ch === '\r';
  }
  if (!inQuote) return 0;
  let line = 1;
  for (let i = 0; i < start; i++) if (text.charCodeAt(i) === 10) line++;
  return line;
}

// ---------- Parsing ----------
function parse(id, name, delimiterChoice) {
  const entry = files.get(id);
  const { text } = entry;
  if (!text.trim()) throw new Error(`${name} is empty.`);
  const rows = [];
  let delimiter = '';
  let guessFailed = false;
  let chunkIndex = 0;
  let missingQuotes = false;
  Papa.parse(text, {
    delimiter: delimiterChoice === 'auto' ? '' : delimiterChoice,
    delimitersToGuess: DELIMITERS,
    skipEmptyLines: true,
    chunkSize: CHUNK,
    chunk(res) {
      for (const r of res.data) rows.push(r);
      if (!delimiter) delimiter = res.meta.delimiter;
      for (const e of res.errors) {
        if (e.code === 'MissingQuotes') missingQuotes = true;
        if (e.code === 'UndetectableDelimiter') guessFailed = true;
      }
      chunkIndex++;
      post({ type: 'progress', id, done: Math.min(chunkIndex * CHUNK, text.length), total: text.length });
    },
  });
  if (missingQuotes) {
    const line = unclosedQuoteLine(text, delimiter);
    throw new Error(`${name} could not be parsed: a quoted field is never closed${line ? ` (the quote opens on line ${line})` : ''}. Check for a missing " character.`);
  }
  if (!rows.length) throw new Error(`${name} has no rows.`);
  let cols = 0;
  let longCells = 0;
  for (const r of rows) {
    if (r.length > cols) cols = r.length;
    for (const v of r) if (v.length > MAX_CELL) longCells++;
  }
  entry.rows = rows;
  entry.cols = cols;
  entry.delimiter = delimiter;
  const preview = rows.slice(0, PREVIEW_ROWS).map((r) => {
    const out = [];
    for (let c = 0; c < Math.min(cols, 50); c++) out.push((r[c] ?? '').slice(0, 120));
    return out;
  });
  return { delimiter, guessFailed: delimiterChoice === 'auto' && guessFailed, rows: rows.length, cols, longCells, preview, encoding: entry.encoding };
}

// ---------- Building ----------
const NUMERIC = /^-?(?:\d+|\d*\.\d+)(?:[eE][+-]?\d+)?$/;
const LEADING_ZERO = /^[+-]?0\d/;

function textColumns(rows, cols, firstData) {
  const flagged = new Array(cols).fill(false);
  for (let r = firstData; r < rows.length; r++) {
    const row = rows[r];
    for (let c = 0; c < row.length; c++) {
      if (flagged[c]) continue;
      const v = row[c].trim();
      if (!v) continue;
      if (LEADING_ZERO.test(v) || (v.replace(/\D/g, '').length > 15 && /^[\d\s+\-.()]+$/.test(v))) flagged[c] = true;
    }
  }
  return flagged;
}

const displayWidth = (s) => {
  let w = 0;
  for (const ch of s) w += ch.codePointAt(0) > 0x2e80 ? 2 : 1;
  return w;
};

function makeCell(raw, isHeader, isText, keepText, stats) {
  let v = raw;
  if (v.length > MAX_CELL) { v = v.slice(0, MAX_CELL); stats.truncated++; }
  if (isHeader) return { t: 's', v };
  if (keepText && isText) return { t: 's', v, z: '@' };
  const trimmed = v.trim();
  if (trimmed && NUMERIC.test(trimmed)) {
    const n = Number(trimmed);
    if (Number.isFinite(n)) {
      const cell = { t: 'n', v: n };
      const dot = trimmed.indexOf('.');
      if (dot >= 0 && !/[eE]/.test(trimmed)) {
        const decimals = trimmed.length - dot - 1;
        if (decimals > 0 && decimals <= 10) cell.z = '0.' + '0'.repeat(decimals);
      }
      return cell;
    }
  }
  if (v.startsWith('=')) stats.formulas++;
  return { t: 's', v }; // formula-looking values stay plain text
}

function buildSheet(rows, cols, r0, r1, c0, c1, opts, stats, flagged) {
  const ws = {};
  const header = opts.header;
  const widths = new Array(c1 - c0).fill(0);
  const sourceRows = [];
  if (header && r0 > 0) sourceRows.push(0); // repeat header on split parts
  for (let r = r0; r < r1; r++) sourceRows.push(r);
  for (let i = 0; i < sourceRows.length; i++) {
    const r = sourceRows[i];
    const row = rows[r];
    const isHeader = header && r === 0;
    for (let c = c0; c < c1; c++) {
      const raw = row[c];
      if (raw === undefined || raw === '') continue; // ragged rows are padded with empty cells
      const cell = makeCell(raw, isHeader, flagged[c], opts.keepText, stats);
      ws[XLSX.utils.encode_cell({ r: i, c: c - c0 })] = cell;
      if (opts.autoWidth) {
        const w = Math.min(MAX_WIDTH, displayWidth(String(cell.v).slice(0, MAX_WIDTH * 2)) + (isHeader ? 3 : 2));
        if (w > widths[c - c0]) widths[c - c0] = w;
      }
    }
  }
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, sourceRows.length - 1), c: Math.max(0, c1 - c0 - 1) } });
  if (opts.autoWidth) ws['!cols'] = widths.map((w) => ({ wch: Math.max(6, w) }));
  return ws;
}

function uniqueName(base, used) {
  let name = base;
  let n = 2;
  while (used.has(name.toLowerCase())) {
    const suffix = ` (${n++})`;
    name = base.slice(0, 31 - suffix.length) + suffix;
  }
  used.add(name.toLowerCase());
  return name;
}

// The community build of SheetJS does not write fonts or panes, so the header styling and the
// frozen top row are patched into the generated XML.
const enc = new TextEncoder();
const dec = new TextDecoder();

function indexOfBytes(hay, needle, from = 0) {
  outer: for (let i = from; i <= hay.length - needle.length; i++) {
    for (let j = 0; j < needle.length; j++) if (hay[i + j] !== needle[j]) continue outer;
    return i;
  }
  return -1;
}

function patchHeader(buffer, sheetCount) {
  const cfb = XLSX.CFB.read(new Uint8Array(buffer), { type: 'array' });
  const styles = XLSX.CFB.find(cfb, '/xl/styles.xml');
  let css = dec.decode(styles.content);
  const fontCount = Number(/<fonts count="(\d+)"/.exec(css)[1]);
  const xfCount = Number(/<cellXfs count="(\d+)"/.exec(css)[1]);
  css = css
    .replace(/<fonts count="\d+">([\s\S]*?)<\/fonts>/, (_, inner) => `<fonts count="${fontCount + 1}">${inner}<font><b/><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>`)
    .replace(/<cellXfs count="\d+">([\s\S]*?)<\/cellXfs>/, (_, inner) => `<cellXfs count="${xfCount + 1}">${inner}<xf numFmtId="0" fontId="${fontCount}" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>`);
  styles.content = enc.encode(css);
  styles.size = styles.content.length;
  const rowEnd = enc.encode('</row>');
  for (let i = 1; i <= sheetCount; i++) {
    const entry = XLSX.CFB.find(cfb, `/xl/worksheets/sheet${i}.xml`);
    if (!entry) continue;
    const bytes = entry.content instanceof Uint8Array ? entry.content : new Uint8Array(entry.content);
    let end = indexOfBytes(bytes, rowEnd);
    end = end < 0 ? Math.min(bytes.length, 4096) : end + rowEnd.length;
    let head = dec.decode(bytes.subarray(0, end));
    head = head
      .replace('<sheetView workbookViewId="0"/>', '<sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView>')
      .replace(/<row r="1"([^>]*)>([\s\S]*?)<\/row>/, (_, attrs, inner) => `<row r="1"${attrs}>${inner.replace(/<c r="([A-Z]+1)"(?![^>]*\ss=)/g, `<c r="$1" s="${xfCount}"`)}</row>`);
    const headBytes = enc.encode(head);
    const out = new Uint8Array(headBytes.length + bytes.length - end);
    out.set(headBytes, 0);
    out.set(bytes.subarray(end), headBytes.length);
    entry.content = out;
    entry.size = out.length;
  }
  return XLSX.CFB.write(cfb, { fileType: 'zip', type: 'array', compression: true });
}

function build(jobId, list, opts) {
  const wb = XLSX.utils.book_new();
  const used = new Set();
  const stats = { truncated: 0, formulas: 0 };
  let sheets = 0;
  list.forEach((item, idx) => {
    const entry = files.get(item.id);
    if (!entry || !entry.rows) throw new Error(`${item.name} is not ready yet.`);
    const { rows, cols } = entry;
    const headerRows = opts.header ? 1 : 0;
    const flagged = textColumns(rows, cols, headerRows);
    const rowsPerPart = MAX_ROWS - (opts.header ? 1 : 0);
    const dataRows = rows.length - headerRows;
    const rowParts = rows.length > MAX_ROWS ? Math.ceil(dataRows / rowsPerPart) : 1;
    const colParts = Math.ceil(cols / MAX_COLS) || 1;
    if ((rowParts > 1 || colParts > 1) && !item.split) throw new Error(`${item.name} is larger than an Excel sheet allows.`);
    const total = rowParts * colParts;
    let part = 0;
    for (let rp = 0; rp < rowParts; rp++) {
      const r0 = rowParts === 1 ? 0 : headerRows + rp * rowsPerPart;
      const r1 = rowParts === 1 ? rows.length : Math.min(rows.length, r0 + rowsPerPart);
      for (let cp = 0; cp < colParts; cp++) {
        const c0 = cp * MAX_COLS;
        const c1 = Math.min(cols, c0 + MAX_COLS);
        const ws = buildSheet(rows, cols, rowParts === 1 ? 0 : r0, r1, c0, c1, opts, stats, flagged);
        part++;
        const base = total > 1 ? `${item.sheetName.slice(0, 25)} ${part}` : item.sheetName;
        XLSX.utils.book_append_sheet(wb, ws, uniqueName(base, used));
        sheets++;
      }
    }
    post({ type: 'progress', jobId, done: idx + 1, total: list.length + 1 });
  });
  let buffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression: true });
  if (opts.header) buffer = patchHeader(buffer, sheets);
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const out = bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength ? bytes.buffer : bytes.slice().buffer;
  post({ type: 'built', jobId, buffer: out, sheetNames: wb.SheetNames, stats }, [out]);
}

self.onmessage = (e) => {
  const msg = e.data;
  try {
    if (msg.type === 'add') {
      const { text, encoding } = decode(msg.buffer);
      files.set(msg.id, { text, encoding });
      post({ type: 'parsed', id: msg.id, meta: parse(msg.id, msg.name, msg.delimiter || 'auto') });
    } else if (msg.type === 'reparse') {
      post({ type: 'parsed', id: msg.id, meta: parse(msg.id, msg.name, msg.delimiter) });
    } else if (msg.type === 'remove') {
      files.delete(msg.id);
    } else if (msg.type === 'build') {
      build(msg.jobId, msg.files, msg.options);
    }
  } catch (err) {
    if (msg.type === 'add') files.delete(msg.id);
    post({ type: 'error', id: msg.id, jobId: msg.jobId, message: err.message || String(err) });
  }
};
