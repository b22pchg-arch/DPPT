'use strict';

const state = {
  headers: [], rawRows: [], rows: [], colMap: {}, model: null, trainingResult: null, forecastRows: [], sampleLoaded: false,
  sourceFileName: '', delimiter: ',', workbook: null, currentSheet: '', editor: {page: 1, pageSize: 100, query: '', filter: 'all', selected: new Set(), dirty: false, dateMode: 'all', dateSingle: '', dateMulti: '', dateFrom: '', dateTo: '', focus: {mode:'none', label:'', indices:[]}}
};

const $ = id => document.getElementById(id);
const logBox = $('log');

function log(msg, type='') {
  const t = new Date().toLocaleTimeString('vi-VN');
  logBox.textContent += `[${t}] ${msg}\n`;
  logBox.scrollTop = logBox.scrollHeight;
}

window.addEventListener('pwa-status', e => log(e.detail));

function escapeHtml(v) {
  return String(v ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function safeFileBase(name) {
  return String(name || 'du_lieu_phu_tai')
    .replace(/\.[^.]+$/, '')
    .replace(/[^a-zA-Z0-9_\-]+/g, '_')
    .slice(0, 80) || 'du_lieu_phu_tai';
}

function norm(s) {
  return String(s ?? '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function countDelimiterOutsideQuotes(line, delimiter) {
  let count = 0, inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i], next = line[i + 1];
    if (ch === '"') {
      if (inQuotes && next === '"') i++;
      else inQuotes = !inQuotes;
    } else if (ch === delimiter && !inQuotes) count++;
  }
  return count;
}

function detectDelimiter(text) {
  const candidates = [',', ';', '\t', '|'];
  const lines = text.replace(/^\uFEFF/, '').split(/\r?\n/).filter(l => l.trim()).slice(0, 20);
  let best = ',', bestScore = -1;
  for (const d of candidates) {
    const counts = lines.map(l => countDelimiterOutsideQuotes(l, d));
    const score = counts.reduce((a,b)=>a+b,0) - Math.abs(Math.max(...counts,0) - Math.min(...counts,0)) * 0.2;
    if (score > bestScore) { bestScore = score; best = d; }
  }
  return best;
}

function parseDelimited(text, delimiter=null) {
  text = text.replace(/^\uFEFF/, '');
  delimiter = delimiter || detectDelimiter(text);
  const rows = [];
  let row = [], cell = '', inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i], next = text[i + 1];
    if (ch === '"') {
      if (inQuotes && next === '"') { cell += '"'; i++; }
      else inQuotes = !inQuotes;
    } else if (ch === delimiter && !inQuotes) {
      row.push(cell); cell = '';
    } else if ((ch === '\n' || ch === '\r') && !inQuotes) {
      if (ch === '\r' && next === '\n') i++;
      row.push(cell); cell = '';
      if (row.some(v => String(v).trim() !== '')) rows.push(row);
      row = [];
    } else {
      cell += ch;
    }
  }
  row.push(cell);
  if (row.some(v => String(v).trim() !== '')) rows.push(row);
  if (!rows.length) return {headers: [], data: [], delimiter};
  const headers = rows[0].map((h, i) => String(h || `cot_${i+1}`).trim() || `cot_${i+1}`);
  const seen = new Map();
  const uniqueHeaders = headers.map(h => {
    const n = seen.get(h) || 0; seen.set(h, n + 1);
    return n ? `${h}_${n+1}` : h;
  });
  const data = rows.slice(1).map(r => {
    const o = {};
    uniqueHeaders.forEach((h, i) => o[h] = r[i] ?? '');
    return o;
  });
  return {headers: uniqueHeaders, data, delimiter};
}

function parseCSV(text) {
  return parseDelimited(text, ',');
}

function parseDataFileText(text, filename='') {
  const ext = String(filename).toLowerCase().split('.').pop();
  if (ext === 'json') {
    const obj = JSON.parse(text);
    let headers = [], data = [];
    if (Array.isArray(obj)) {
      data = obj.map(r => ({...r}));
      headers = [...new Set(data.flatMap(r => Object.keys(r)))];
    } else if (Array.isArray(obj.rawRows)) {
      data = obj.rawRows.map(r => ({...r}));
      headers = Array.isArray(obj.headers) && obj.headers.length ? obj.headers.slice() : [...new Set(data.flatMap(r => Object.keys(r)))];
    } else if (Array.isArray(obj.rows)) {
      data = obj.rows.map(r => ({...r}));
      headers = Array.isArray(obj.headers) && obj.headers.length ? obj.headers.slice() : [...new Set(data.flatMap(r => Object.keys(r)))];
    } else {
      throw new Error('JSON phải là mảng object hoặc có trường rawRows/rows.');
    }
    return {headers, data, delimiter: ',', meta: obj.meta || {}};
  }
  return parseDelimited(text, ext === 'tsv' ? '\t' : null);
}


function isSpreadsheetFile(filename='') {
  return /\.(xlsx|xlsm|xlsb|xls)$/i.test(String(filename));
}

function showSheetSelector(show) {
  const label = $('sheetLabel'), sel = $('sheetSelect'), info = $('sheetInfo');
  if (!label || !sel || !info) return;
  label.style.display = show ? '' : 'none';
  sel.style.display = show ? '' : 'none';
  if (!show) { sel.innerHTML = ''; info.textContent = ''; }
}

function workbookToParsed(workbook, sheetName) {
  if (!window.XLSX || !XLSX.utils || !XLSX.utils.sheet_to_json) throw new Error('Thiếu thư viện đọc Excel XLSX.');
  if (!workbook || !workbook.Sheets) throw new Error('Workbook không hợp lệ.');
  sheetName = sheetName || workbook.SheetNames?.[0];
  if (!sheetName || !workbook.Sheets[sheetName]) throw new Error('Không tìm thấy sheet Excel.');
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {defval: '', raw: false});
  const headers = [...new Set(data.flatMap(r => Object.keys(r)))];
  return {headers, data, delimiter: ',', meta: {sheetName}};
}

function populateSheetSelector(workbook) {
  const sel = $('sheetSelect');
  showSheetSelector(!!(workbook && workbook.SheetNames && workbook.SheetNames.length));
  if (!sel || !workbook) return;
  sel.innerHTML = '';
  workbook.SheetNames.forEach(name => {
    const opt = document.createElement('option'); opt.value = name; opt.textContent = name; sel.appendChild(opt);
  });
  sel.value = workbook.SheetNames[0] || '';
  state.currentSheet = sel.value;
  const info = $('sheetInfo');
  if (info) {
    const engine = window.XLSX?.version || 'unknown';
    info.innerHTML = `<span class="pill">${workbook.SheetNames.length} sheet</span> <span class="pill">Excel reader: ${escapeHtml(engine)}</span>`;
  }
}

function applyParsedDataset(parsed, filename='', sourceLabel='file') {
  state.headers = parsed.headers || [];
  state.rawRows = parsed.data || [];
  state.rows = [];
  state.sampleLoaded = false;
  state.sourceFileName = filename || sourceLabel;
  state.delimiter = parsed.delimiter || ',';
  state.editor.selected.clear();
  state.editor.page = 1;
  state.editor.dirty = false;
  fillColumnSelects(state.headers);
  applySavedColumnMapIfPossible();
  try { normalizeRows(); applyDataInfo(); previewData(); }
  catch(err) { log('Cần kiểm tra ánh xạ cột: ' + err.message); }
  renderEditorTable();
}

async function loadSpreadsheetFile(file) {
  if (!window.XLSX || !XLSX.read || !XLSX.utils) throw new Error('Thiếu thư viện SheetJS/XLSX để đọc file Excel.');
  const ext = String(file.name).toLowerCase().split('.').pop();
  const usingLite = String(XLSX.version || '').includes('lite-xlsx-reader');
  if (ext === 'xls' && usingLite) {
    throw new Error('File .xls nhị phân cũ cần thư viện SheetJS chính thức xlsx.full.min.js. Bản LV3 offline hiện đọc trực tiếp .xlsx/.xlsm; hãy lưu lại file thành .xlsx hoặc CSV UTF-8.');
  }
  const buf = await file.arrayBuffer();
  const workbook = XLSX.read(buf, {type: 'array', cellDates: false});
  if (!workbook.SheetNames || !workbook.SheetNames.length) throw new Error('File Excel không có sheet dữ liệu.');
  state.workbook = workbook;
  populateSheetSelector(workbook);
  const sheetName = state.currentSheet || workbook.SheetNames[0];
  const parsed = workbookToParsed(workbook, sheetName);
  applyParsedDataset(parsed, file.name, 'Excel');
  state.currentSheet = sheetName;
  $('sheetSelect').value = sheetName;
  log(`Đã nạp Excel: ${file.name}, sheet '${sheetName}', ${state.rawRows.length} dòng thô, ${state.headers.length} cột.`);
}

function loadWorkbookSheet(sheetName) {
  if (!state.workbook) return;
  const parsed = workbookToParsed(state.workbook, sheetName);
  state.currentSheet = sheetName;
  applyParsedDataset(parsed, state.sourceFileName, 'Excel');
  log(`Đã chuyển sang sheet '${sheetName}', ${state.rawRows.length} dòng thô, ${state.headers.length} cột.`);
}

function csvEscape(v) {
  if (v == null) return '';
  const s = String(v);
  if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function toCSV(rows, headers) {
  const out = [headers.map(csvEscape).join(',')];
  rows.forEach(r => out.push(headers.map(h => csvEscape(r[h])).join(',')));
  return out.join('\n');
}

function exportCSVContent(rows, headers) {
  return '\uFEFF' + toCSV(rows, headers);
}

function exportEditedCSV() {
  if (!state.headers.length) throw new Error('Chưa có dữ liệu để xuất.');
  saveTextFile(safeFileBase(state.sourceFileName) + '_edited.csv', exportCSVContent(state.rawRows, state.headers), 'text/csv;charset=utf-8');
  log('Đã xuất CSV đã hiệu chỉnh.');
}

function exportEditedJSON() {
  if (!state.headers.length) throw new Error('Chưa có dữ liệu để xuất.');
  const payload = {
    type: 'SCADA_LOAD_DATA_EDITED_LV4',
    exportedAt: new Date().toISOString(),
    sourceFileName: state.sourceFileName,
    headers: state.headers,
    colMap: state.colMap,
    rawRows: state.rawRows
  };
  saveTextFile(safeFileBase(state.sourceFileName) + '_edited.json', JSON.stringify(payload, null, 2), 'application/json');
  log('Đã xuất JSON đã hiệu chỉnh.');
}

function saveTextFile(filename, content, type='application/octet-stream') {
  const blob = new Blob([content], {type});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

function parseNumber(v) {
  if (v == null) return NaN;
  let s = String(v).trim();
  if (!s) return NaN;
  s = s.replace(/\s/g, '');
  if (s.includes(',') && !s.includes('.')) s = s.replace(',', '.');
  s = s.replace(/[^0-9+\-eE.]/g, '');
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}

function parseFlag(v) {
  const s = norm(v);
  if (!s) return 0;
  if (['1','true','yes','y','co','có','x','bat thuong','su co','cat dien','chuyen tai','nghi','le'].includes(s)) return 1;
  const n = parseNumber(v);
  return Number.isFinite(n) && n !== 0 ? 1 : 0;
}

function parseTime(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (v == null) return null;
  const s0 = String(v).trim();
  if (!s0) return null;
  const n = Number(s0);
  if (Number.isFinite(n) && n > 20000 && n < 90000) {
    const ms = Math.round((n - 25569) * 86400 * 1000);
    const d = new Date(ms);
    return isNaN(d) ? null : d;
  }
  let s = s0.replace('T', ' ').replace(/\//g, '-');
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3], +(m[4]||0), +(m[5]||0), +(m[6]||0));
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1], +(m[4]||0), +(m[5]||0), +(m[6]||0));
  const d = new Date(s0);
  return isNaN(d) ? null : d;
}

function fmtTime(d) {
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}


function fmtDateKey(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
}

function parseDateKey(v) {
  if (v instanceof Date) return fmtDateKey(v);
  const s0 = String(v ?? '').trim();
  if (!s0) return '';
  let m = s0.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return `${m[1]}-${String(+m[2]).padStart(2,'0')}-${String(+m[3]).padStart(2,'0')}`;
  m = s0.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return `${m[3]}-${String(+m[2]).padStart(2,'0')}-${String(+m[1]).padStart(2,'0')}`;
  const d = parseTime(s0);
  return d ? fmtDateKey(d) : '';
}

function getRawDateKey(raw) {
  const m = readMap();
  if (!m.time) return '';
  const d = parseTime(raw[m.time]);
  return d ? fmtDateKey(d) : '';
}

function parseDateList(text) {
  return new Set(String(text || '').split(/[;,\n]+/).map(parseDateKey).filter(Boolean));
}

function readEditorDateFilterFromUI() {
  if (!$('dateFilterMode')) return;
  state.editor.dateMode = $('dateFilterMode').value || 'all';
  state.editor.dateSingle = $('dateFilterSingle').value || '';
  state.editor.dateMulti = $('dateFilterMulti').value || '';
  state.editor.dateFrom = $('dateFilterFrom').value || '';
  state.editor.dateTo = $('dateFilterTo').value || '';
}

function editorDateFilterPass(raw) {
  const mode = state.editor.dateMode || 'all';
  if (mode === 'all') return true;
  const key = getRawDateKey(raw);
  if (!key) return false;
  if (mode === 'single') return key === parseDateKey(state.editor.dateSingle);
  if (mode === 'multi') return parseDateList(state.editor.dateMulti).has(key);
  if (mode === 'range') {
    const from = parseDateKey(state.editor.dateFrom);
    const to = parseDateKey(state.editor.dateTo);
    if (from && key < from) return false;
    if (to && key > to) return false;
    return true;
  }
  return true;
}

function dateFromKey(key) {
  const m = String(key || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(+m[1], +m[2]-1, +m[3]) : null;
}

function parseHolidayRules(text) {
  const rules = [];
  const tokens = String(text || '').split(/[;,\n]+/).map(s => s.trim()).filter(Boolean);
  for (const token of tokens) {
    const range = token.split(/\.\.|\s+den\s+|\s+đến\s+|\s+to\s+/i).map(s => s.trim());
    if (range.length === 2) {
      const a = parseDateKey(range[0]), b = parseDateKey(range[1]);
      if (a && b) rules.push({type:'range', from: a <= b ? a : b, to: a <= b ? b : a});
      continue;
    }
    const md = token.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
    if (md) { rules.push({type:'monthday', month:+md[2], day:+md[1]}); continue; }
    const exact = parseDateKey(token);
    if (exact) { rules.push({type:'exact', date: exact}); continue; }
  }
  return rules;
}

function holidayByRules(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj)) return false;
  const dow = dateObj.getDay();
  if ($('holidaySat')?.checked && dow === 6) return true;
  if ($('holidaySun')?.checked && dow === 0) return true;
  let text = '';
  if ($('holidayFixed')?.checked) text += '01-01,30-04,01-05,02-09,';
  text += $('holidayExtraDates')?.value || '';
  const key = fmtDateKey(dateObj);
  const month = dateObj.getMonth() + 1;
  const day = dateObj.getDate();
  for (const rule of parseHolidayRules(text)) {
    if (rule.type === 'exact' && rule.date === key) return true;
    if (rule.type === 'range' && key >= rule.from && key <= rule.to) return true;
    if (rule.type === 'monthday' && rule.month === month && rule.day === day) return true;
  }
  return false;
}

function formatNum(n, digits=3) {
  return Number.isFinite(n) ? Number(n).toFixed(digits) : '-';
}



// ===== LV4.1 FIX: helper functions for Vietnamese headers and quick-fill column list =====
const COLUMN_KEY_TO_SELECT = {
  time: 'colTime',
  p: 'colP',
  station: 'colStation',
  temp: 'colTemp',
  rain: 'colRain',
  holiday: 'colHoliday',
  abnormal: 'colAbnormal',
  outage: 'colOutage',
  transfer: 'colTransfer'
};

const VI_HEADER_BY_KEY = {
  time: 'Thời gian',
  p: 'Công suất P',
  station: 'Trạm/Lộ/Khu vực',
  temp: 'Nhiệt độ',
  rain: 'Mưa',
  holiday: 'Ngày nghỉ/lễ',
  abnormal: 'Bất thường',
  outage: 'Cắt điện/sự cố',
  transfer: 'Chuyển tải'
};

const VI_HEADER_ALIASES = [
  ['time', ['time','timestamp','datetime','date','ngay gio','thoi gian','ngày giờ','thời gian']],
  ['p', ['p','p mw','p_mw','mw','load','phu tai','phụ tải','cong suat','công suất','active power']],
  ['station', ['station','feeder','area','tram','trạm','lo','lộ','xuat tuyen','xuất tuyến','khu vuc','khu vực']],
  ['temp', ['temperature','temp','tmax','nhiet do','nhiệt độ']],
  ['rain', ['rain','rainfall','mua','mưa']],
  ['holiday', ['holiday','nghi','nghỉ','le','lễ','ngay le','ngày lễ','ngay nghi','ngày nghỉ']],
  ['abnormal', ['abnormal','bat thuong','bất thường','loi du lieu','lỗi dữ liệu']],
  ['outage', ['outage','fault','cat dien','cắt điện','su co','sự cố']],
  ['transfer', ['transfer','chuyen tai','chuyển tải','ket luoi','kết lưới']]
];

function makeUniqueName(base, usedNames) {
  base = String(base || 'Cột mới').trim() || 'Cột mới';
  const used = new Set((usedNames || []).map(v => String(v)));
  if (!used.has(base)) return base;
  let i = 2;
  while (used.has(`${base}_${i}`)) i++;
  return `${base}_${i}`;
}

function getColumnSelectValues() {
  const out = {};
  for (const id of Object.values(COLUMN_KEY_TO_SELECT)) out[id] = $(id)?.value || '';
  out.quickCustomCol = $('quickCustomCol')?.value || '';
  return out;
}

function restoreColumnSelectValues(values) {
  values = values || {};
  for (const [id, value] of Object.entries(values)) {
    const el = $(id);
    if (!el || value == null) continue;
    if (el.options && el.options.length) {
      const ok = [...el.options].some(o => o.value === value);
      if (ok) el.value = value;
    } else {
      el.value = value;
    }
  }
  refreshQuickCustomColumns(values.quickCustomCol);
}

function headerDisplayName(header) {
  const raw = String(header || '');
  const n = norm(raw);
  for (const [key, aliases] of VI_HEADER_ALIASES) {
    if (aliases.some(a => n === norm(a))) return VI_HEADER_BY_KEY[key];
  }
  return raw;
}

function refreshQuickCustomColumns(preferredValue=null) {
  const sel = $('quickCustomCol');
  if (!sel) return;
  const oldValue = preferredValue != null ? preferredValue : sel.value;
  sel.innerHTML = '';
  const none = document.createElement('option');
  none.value = '';
  none.textContent = '-- Không dùng --';
  sel.appendChild(none);
  for (const h of state.headers || []) {
    const opt = document.createElement('option');
    opt.value = h;
    opt.textContent = headerDisplayName(h);
    if (headerDisplayName(h) !== h) opt.textContent += ` (${h})`;
    sel.appendChild(opt);
  }
  if ([...sel.options].some(o => o.value === oldValue)) sel.value = oldValue;
}
// ===== END LV4.1 FIX =====

function fillColumnSelects(headers) {
  const ids = ['colTime','colP','colStation','colTemp','colRain','colHoliday','colAbnormal','colOutage','colTransfer'];
  ids.forEach(id => {
    const sel = $(id);
    sel.innerHTML = '<option value="">-- Không dùng --</option>';
    headers.forEach(h => {
      const opt = document.createElement('option'); opt.value = h; opt.textContent = h; sel.appendChild(opt);
    });
  });
  const find = aliases => {
    const ns = headers.map(h => [h, norm(h)]);
    for (const a of aliases) {
      const na = norm(a);
      const hit = ns.find(([h, nh]) => nh === na || nh.includes(na) || na.includes(nh));
      if (hit) return hit[0];
    }
    return '';
  };
  $('colTime').value = find(['thoi gian','ngay gio','timestamp','datetime','time','date']);
  $('colP').value = find(['p','p mw','mw','load','phu tai','cong suat','congsuat','active power']);
  $('colStation').value = find(['tram','lo','xuat tuyen','station','feeder','area','khu vuc']);
  $('colTemp').value = find(['nhiet do','temperature','temp','tmax']);
  $('colRain').value = find(['mua','rain','rainfall']);
  $('colHoliday').value = find(['ngay le','le','holiday','nghi']);
  $('colAbnormal').value = find(['bat thuong','abnormal','loi du lieu']);
  $('colOutage').value = find(['cat dien','su co','outage','fault']);
  $('colTransfer').value = find(['chuyen tai','transfer','ket luoi']);
  refreshQuickCustomColumns();
}


function readMap() {
  state.colMap = {
    time: $('colTime').value, p: $('colP').value, station: $('colStation').value,
    temp: $('colTemp').value, rain: $('colRain').value, holiday: $('colHoliday').value,
    abnormal: $('colAbnormal').value, outage: $('colOutage').value, transfer: $('colTransfer').value
  };
  return state.colMap;
}


function ensureMappedColumn(key, preferredHeader=null) {
  const id = COLUMN_KEY_TO_SELECT[key];
  if (!id || !$(id)) throw new Error('Không xác định được cột cần tạo: ' + key);
  let col = $(id).value;
  if (col && state.headers.includes(col)) return col;
  const current = getColumnSelectValues();
  const base = preferredHeader || VI_HEADER_BY_KEY[key] || 'Cột mới';
  col = makeUniqueName(base, state.headers);
  state.headers.push(col);
  state.rawRows.forEach(r => { if (!(col in r)) r[col] = ''; });
  fillColumnSelects(state.headers);
  restoreColumnSelectValues(current);
  $(id).value = col;
  readMap();
  refreshQuickCustomColumns();
  return col;
}

function vietnamizeHeaders() {
  if (!state.headers.length) { log('Chưa có dữ liệu để đổi tiêu đề.'); return; }
  const map = readMap();
  const desired = {...VI_HEADER_BY_KEY};
  const renameMap = new Map();
  for (const [key, oldName] of Object.entries(map)) {
    if (!oldName || !state.headers.includes(oldName) || !desired[key]) continue;
    renameMap.set(oldName, desired[key]);
  }
  for (const h of state.headers) {
    const vi = headerDisplayName(h);
    if (vi && vi !== h && !renameMap.has(h)) renameMap.set(h, vi);
  }
  const used = [];
  const finalNames = state.headers.map(h => {
    const base = renameMap.get(h) || h;
    const unique = makeUniqueName(base, used);
    used.push(unique);
    return unique;
  });
  const oldHeaders = state.headers.slice();
  state.rawRows = state.rawRows.map(row => {
    const out = {};
    oldHeaders.forEach((h, i) => out[finalNames[i]] = row[h] ?? '');
    return out;
  });
  state.headers = finalNames;
  fillColumnSelects(state.headers);
  for (const [key, oldName] of Object.entries(map)) {
    const idx = oldHeaders.indexOf(oldName);
    const id = COLUMN_KEY_TO_SELECT[key];
    if (idx >= 0 && id && $(id)) $(id).value = finalNames[idx];
  }
  readMap();
  markEditorDirty(true);
  renderEditorTable();
  log('Đã đổi các tiêu đề cột sang tiếng Việt và giữ lại ánh xạ cột hiện có.');
}

function normalizeRows() {
  const m = readMap();
  if (!m.time || !m.p) throw new Error('Cần chọn cột Thời gian và Công suất P.');
  state.rows = state.rawRows.map((r, idx) => {
    const time = parseTime(r[m.time]);
    const p = parseNumber(r[m.p]);
    const station = m.station ? String(r[m.station] ?? '').trim() : 'ALL';
    const temp = m.temp ? parseNumber(r[m.temp]) : NaN;
    const rain = m.rain ? parseNumber(r[m.rain]) : 0;
    const holiday = m.holiday ? parseFlag(r[m.holiday]) : 0;
    const abnormal = m.abnormal ? parseFlag(r[m.abnormal]) : 0;
    const outage = m.outage ? parseFlag(r[m.outage]) : 0;
    const transfer = m.transfer ? parseFlag(r[m.transfer]) : 0;
    const excludeTrain = parseFlag(r['bo_khoi_huan_luyen']);
    return {idx, raw: r, time, p, station: station || 'ALL', temp, rain, holiday, abnormal, outage, transfer, excludeTrain};
  }).filter(r => r.time && Number.isFinite(r.p));
  state.rows.sort((a,b) => a.time - b.time);
  buildStationSelect();
  log(`Đã chuẩn hóa ${state.rows.length} dòng hợp lệ.`);
  if (state.rows.length < 20) log('Cảnh báo: dữ liệu quá ít, model sẽ kém ổn định.');
}

function buildStationSelect() {
  const selected = $('stationSelect').value;
  const stations = [...new Set(state.rows.map(r => r.station || 'ALL'))].sort();
  const sel = $('stationSelect');
  sel.innerHTML = '<option value="__ALL__">Tất cả / không lọc</option>';
  stations.forEach(s => {
    const opt = document.createElement('option'); opt.value = s; opt.textContent = s; sel.appendChild(opt);
  });
  if ([...sel.options].some(o => o.value === selected)) sel.value = selected;
}

function getSelectedRows() {
  const station = $('stationSelect').value;
  return state.rows.filter(r => station === '__ALL__' || r.station === station).sort((a,b) => a.time - b.time);
}

function detectIntervalMinutes(rows) {
  const diffs = [];
  for (let i=1;i<rows.length;i++) {
    const d = (rows[i].time - rows[i-1].time) / 60000;
    if (d > 0 && d < 1440 * 7) diffs.push(d);
  }
  if (!diffs.length) return 60;
  diffs.sort((a,b)=>a-b);
  return Math.max(1, Math.round(diffs[Math.floor(diffs.length/2)]));
}

const FEATURE_NAMES = [
  'hour','hour_sin','hour_cos','dow','dow_sin','dow_cos','month','is_weekend','is_holiday',
  'temp','rain','lag1','lag2','lag4','lag_day','lag_week','avg4','avg_day','avg_same_hour_7d','trend_1_4'
];

function mean(arr) {
  const v = arr.filter(Number.isFinite);
  return v.length ? v.reduce((a,b)=>a+b,0)/v.length : NaN;
}

function computeFeatureVector(series, idx, nPerDay, futureDefaults={}) {
  const r = series[idx];
  const d = r.time;
  const hour = d.getHours() + d.getMinutes()/60;
  const dow = d.getDay();
  const month = d.getMonth() + 1;
  const valAt = k => (idx-k >= 0 && Number.isFinite(series[idx-k].p)) ? series[idx-k].p : NaN;
  const avgLast = k => {
    const arr=[];
    for (let j=1; j<=k && idx-j>=0; j++) if (Number.isFinite(series[idx-j].p)) arr.push(series[idx-j].p);
    return mean(arr);
  };
  const avgSame = days => {
    const arr=[];
    for (let d=1; d<=days; d++) {
      const k = nPerDay * d;
      if (idx-k >= 0 && Number.isFinite(series[idx-k].p)) arr.push(series[idx-k].p);
    }
    return mean(arr);
  };
  const lag1 = valAt(1), lag4 = valAt(4);
  let temp = Number.isFinite(r.temp) ? r.temp : parseNumber(futureDefaults.temp);
  let rain = Number.isFinite(r.rain) ? r.rain : parseNumber(futureDefaults.rain);
  if (!Number.isFinite(rain)) rain = 0;
  const isHoliday = r.holiday || 0;
  return [
    hour,
    Math.sin(2*Math.PI*hour/24), Math.cos(2*Math.PI*hour/24),
    dow, Math.sin(2*Math.PI*dow/7), Math.cos(2*Math.PI*dow/7),
    month, (dow===0 || dow===6) ? 1 : 0, isHoliday,
    temp, rain,
    lag1, valAt(2), valAt(4), valAt(nPerDay), valAt(nPerDay*7),
    avgLast(4), avgLast(nPerDay), avgSame(7),
    Number.isFinite(lag1) && Number.isFinite(lag4) ? lag1-lag4 : NaN
  ];
}

function buildDataset(rows, featureMeansFromModel=null) {
  const intervalMinutes = detectIntervalMinutes(rows);
  const nPerDay = Math.max(1, Math.round(1440 / intervalMinutes));
  const Xraw=[], y=[], times=[], cleanRows=[];
  const tempMean = mean(rows.map(r => r.temp));
  const filled = rows.map(r => ({...r, temp: Number.isFinite(r.temp) ? r.temp : tempMean}));
  for (let i=1; i<filled.length; i++) {
    const r = filled[i];
    if (!Number.isFinite(r.p)) continue;
    if (r.abnormal || r.outage || r.transfer || r.excludeTrain) continue;
    const x = computeFeatureVector(filled, i, nPerDay, {});
    if (!Number.isFinite(x[11])) continue; // cần lag1 tối thiểu
    Xraw.push(x); y.push(r.p); times.push(r.time); cleanRows.push(r);
  }
  let featureMeans = featureMeansFromModel;
  if (!featureMeans) {
    featureMeans = FEATURE_NAMES.map((_, j) => mean(Xraw.map(x => x[j])));
    featureMeans = featureMeans.map(v => Number.isFinite(v) ? v : 0);
  }
  const X = Xraw.map(x => x.map((v,j) => Number.isFinite(v) ? v : featureMeans[j]));
  return {X, y, times, cleanRows, featureNames: FEATURE_NAMES.slice(), featureMeans, intervalMinutes, nPerDay};
}

function varianceSSE(vals) {
  let n=0, sum=0, sumSq=0;
  for (const v of vals) { n++; sum+=v; sumSq+=v*v; }
  return n ? sumSq - sum*sum/n : 0;
}

function fitTree(X, y, params, indices=null, depth=0) {
  const idxs = indices || X.map((_,i)=>i);
  let sum=0, sumSq=0;
  for (const i of idxs) { sum += y[i]; sumSq += y[i]*y[i]; }
  const pred = sum / idxs.length;
  if (depth >= params.maxDepth || idxs.length < params.minLeaf * 2) return {leaf:true, value:pred, n:idxs.length};

  let best = null;
  const baseSSE = sumSq - sum*sum/idxs.length;
  const nFeatures = X[0].length;
  for (let f=0; f<nFeatures; f++) {
    const vals = idxs.map(i => X[i][f]).filter(Number.isFinite).sort((a,b)=>a-b);
    if (vals.length < params.minLeaf*2) continue;
    const min = vals[0], max = vals[vals.length-1];
    if (min === max) continue;
    const thresholds = [];
    const bins = Math.min(params.maxBins, vals.length-1);
    for (let b=1; b<=bins; b++) {
      const pos = Math.floor(b * vals.length / (bins + 1));
      const th = vals[pos];
      if (Number.isFinite(th) && th > min && th < max && thresholds[thresholds.length-1] !== th) thresholds.push(th);
    }
    for (const th of thresholds) {
      let nl=0,nr=0,sl=0,sr=0,sql=0,sqr=0;
      for (const i of idxs) {
        const yy = y[i];
        if (X[i][f] <= th) {nl++; sl+=yy; sql+=yy*yy;} else {nr++; sr+=yy; sqr+=yy*yy;}
      }
      if (nl < params.minLeaf || nr < params.minLeaf) continue;
      const sse = (sql - sl*sl/nl) + (sqr - sr*sr/nr);
      if (!best || sse < best.sse) best = {f, th, sse};
    }
  }
  if (!best || best.sse >= baseSSE * 0.999) return {leaf:true, value:pred, n:idxs.length};
  const left=[], right=[];
  for (const i of idxs) (X[i][best.f] <= best.th ? left : right).push(i);
  return {
    leaf:false, feature:best.f, threshold:best.th, value:pred, n:idxs.length,
    left: fitTree(X, y, params, left, depth+1),
    right: fitTree(X, y, params, right, depth+1)
  };
}

function predictTree(tree, x) {
  let node = tree;
  while (!node.leaf) node = x[node.feature] <= node.threshold ? node.left : node.right;
  return node.value;
}

function predictModel(model, xRaw) {
  const x = xRaw.map((v,j) => Number.isFinite(v) ? v : (model.featureMeans[j] ?? 0));
  let p = model.initPred;
  for (const tr of model.trees) p += model.learningRate * predictTree(tr, x);
  return p;
}

function metrics(actual, pred) {
  let n=0, ae=0, se=0, ape=0, apeN=0;
  for (let i=0;i<actual.length;i++) {
    const a=actual[i], p=pred[i];
    if (!Number.isFinite(a) || !Number.isFinite(p)) continue;
    const e = a-p; n++; ae += Math.abs(e); se += e*e;
    if (Math.abs(a) > 1e-9) { ape += Math.abs(e/a); apeN++; }
  }
  return {n, mae: n?ae/n:NaN, rmse:n?Math.sqrt(se/n):NaN, mape:apeN?ape/apeN*100:NaN};
}

async function trainGBDT() {
  if (!state.rows.length) throw new Error('Chưa có dữ liệu.');
  const rows = getSelectedRows();
  if (rows.length < 50) throw new Error('Dữ liệu sau khi lọc quá ít. Cần tối thiểu khoảng 50 dòng, tốt nhất vài tuần trở lên.');
  const ds = buildDataset(rows);
  if (ds.X.length < 40) throw new Error('Không đủ mẫu sau khi tạo đặc trưng lag. Kiểm tra cột thời gian/P hoặc dữ liệu bị thiếu quá nhiều.');
  const valPercent = parseNumber($('valPercent').value) || 20;
  const nVal = Math.max(1, Math.floor(ds.X.length * valPercent / 100));
  const nTrain = ds.X.length - nVal;
  const Xtr = ds.X.slice(0,nTrain), ytr = ds.y.slice(0,nTrain);
  const Xval = ds.X.slice(nTrain), yval = ds.y.slice(nTrain);
  const initPred = mean(ytr);
  let predTr = ytr.map(() => initPred);
  const params = {
    nTrees: Math.max(1, Math.floor(parseNumber($('nTrees').value) || 70)),
    maxDepth: Math.max(1, Math.floor(parseNumber($('maxDepth').value) || 3)),
    learningRate: Math.max(0.001, parseNumber($('learningRate').value) || 0.08),
    minLeaf: Math.max(2, Math.floor(parseNumber($('minLeaf').value) || 12)),
    maxBins: Math.max(4, Math.floor(parseNumber($('maxBins').value) || 28))
  };
  const trees = [];
  log(`Bắt đầu huấn luyện: ${Xtr.length} mẫu train, ${Xval.length} mẫu validation, interval ≈ ${ds.intervalMinutes} phút.`);
  $('trainBtn').disabled = true;
  try {
    for (let t=0; t<params.nTrees; t++) {
      const residual = ytr.map((yy,i) => yy - predTr[i]);
      const tree = fitTree(Xtr, residual, params);
      trees.push(tree);
      for (let i=0;i<Xtr.length;i++) predTr[i] += params.learningRate * predictTree(tree, Xtr[i]);
      if ((t+1)%5===0 || t===params.nTrees-1) {
        const mt = metrics(ytr, predTr);
        log(`Cây ${t+1}/${params.nTrees} - MAE train ${formatNum(mt.mae,3)}`);
        await new Promise(r => setTimeout(r, 0));
      }
    }
  } finally {
    $('trainBtn').disabled = false;
  }
  const model = {
    type: 'GBDT_REGRESSION_JS_OFFLINE_LV5',
    createdAt: new Date().toISOString(),
    station: $('stationSelect').value,
    featureNames: ds.featureNames,
    featureMeans: ds.featureMeans,
    intervalMinutes: ds.intervalMinutes,
    nPerDay: ds.nPerDay,
    initPred, learningRate: params.learningRate,
    trees,
    params,
    colMap: state.colMap,
    note: 'Model huấn luyện ngoài mạng SCADA. Trong SCADA chỉ cần nạp model_gbdt.json và dữ liệu gần nhất để dự báo.'
  };
  const predVal = Xval.map(x => predictModel(model, x));
  const predTrain = Xtr.map(x => predictModel(model, x));
  const mVal = metrics(yval, predVal);
  const mTrain = metrics(ytr, predTrain);
  model.metrics = {validation: mVal, train: mTrain};
  state.model = model;
  state.trainingResult = {ds, nTrain, yval, predVal, timesVal: ds.times.slice(nTrain), yTrain: ytr, predTrain};
  updateMetrics(mVal, ds.intervalMinutes);
  drawSeries(ds.times.slice(nTrain), yval, predVal, 'Validation thực tế', 'Validation dự báo');
  updateValidationErrorViews();
  $('exportModelBtn').disabled = false;
  $('forecastBtn').disabled = false;
  if ($('forecastAllBtn')) $('forecastAllBtn').disabled = false;
  renderModelInfo();
  log(`Hoàn tất huấn luyện. Validation MAE=${formatNum(mVal.mae,3)}, MAPE=${formatNum(mVal.mape,2)}%, RMSE=${formatNum(mVal.rmse,3)}.`);
}

function updateMetrics(m, interval) {
  $('maeVal').textContent = formatNum(m?.mae,3);
  $('mapeVal').textContent = Number.isFinite(m?.mape) ? formatNum(m.mape,2) + '%' : '-';
  $('rmseVal').textContent = formatNum(m?.rmse,3);
  $('intervalVal').textContent = interval ? interval + ' phút' : '-';
}

function renderModelInfo() {
  if (!state.model) { $('modelInfo').innerHTML = ''; return; }
  const m = state.model;
  $('modelInfo').innerHTML = `
    <span class="pill">Model: ${m.type}</span>
    <span class="pill">Số cây: ${m.trees.length}</span>
    <span class="pill">Interval: ${m.intervalMinutes} phút</span>
    <span class="pill">Station: ${m.station || 'ALL'}</span>
    <span class="pill">MAPE: ${formatNum(m.metrics?.validation?.mape,2)}%</span>`;
}

function drawSeries(times, actual, pred, label1='Thực tế', label2='Dự báo') {
  const canvas = $('chart');
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0,0,canvas.width,canvas.height);
  ctx.fillStyle = '#071120'; ctx.fillRect(0,0,canvas.width,canvas.height);
  const W = canvas.width, H = canvas.height, padL=54, padR=20, padT=30, padB=42;
  const vals = [...actual, ...pred].filter(Number.isFinite);
  if (!vals.length) return;
  let minY = Math.min(...vals), maxY = Math.max(...vals);
  const span = maxY-minY || 1; minY -= span*0.08; maxY += span*0.08;
  const n = Math.max(actual.length, pred.length);
  const xAt = i => padL + (W-padL-padR) * (n <= 1 ? 0 : i/(n-1));
  const yAt = v => padT + (H-padT-padB) * (1 - (v-minY)/(maxY-minY));
  ctx.strokeStyle = '#29415f'; ctx.lineWidth = 1;
  ctx.beginPath();
  for (let g=0; g<=4; g++) { const y = padT + (H-padT-padB)*g/4; ctx.moveTo(padL,y); ctx.lineTo(W-padR,y); }
  ctx.stroke();
  ctx.fillStyle = '#9fb0c6'; ctx.font = '13px system-ui';
  for (let g=0; g<=4; g++) { const v = maxY - (maxY-minY)*g/4; ctx.fillText(formatNum(v,1), 8, padT + (H-padT-padB)*g/4 + 4); }
  function line(arr, color) {
    ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.beginPath(); let started=false;
    arr.forEach((v,i)=>{ if(!Number.isFinite(v)) return; const x=xAt(i), y=yAt(v); if(!started){ctx.moveTo(x,y); started=true;} else ctx.lineTo(x,y); });
    ctx.stroke();
  }
  line(actual, '#38bdf8'); line(pred, '#f59e0b');
  ctx.fillStyle = '#38bdf8'; ctx.fillText(label1, padL, 18);
  ctx.fillStyle = '#f59e0b'; ctx.fillText(label2, padL+160, 18);
  if (times && times.length) {
    ctx.fillStyle = '#9fb0c6';
    ctx.fillText(fmtTime(times[0]), padL, H-14);
    ctx.textAlign = 'right'; ctx.fillText(fmtTime(times[times.length-1]), W-padR, H-14); ctx.textAlign='left';
  }
}

function renderTable(rows, headers=null, max=200) {
  const box = $('tableBox');
  if (!rows || !rows.length) { box.innerHTML = '<table><tbody><tr><td>Không có dữ liệu</td></tr></tbody></table>'; return; }
  headers = headers || Object.keys(rows[0]);
  const shown = rows.slice(0, max);
  let html = '<table><thead><tr>' + headers.map(h=>`<th title="${escapeHtml(h)}">${escapeHtml(headerDisplayName(h))}</th>`).join('') + '</tr></thead><tbody>';
  shown.forEach(r => {
    html += '<tr>' + headers.map(h => `<td>${escapeHtml(r[h] ?? '')}</td>`).join('') + '</tr>';
  });
  html += '</tbody></table>';
  if (rows.length > max) html += `<div class="note" style="padding:8px">Đang hiển thị ${max}/${rows.length} dòng.</div>`;
  box.innerHTML = html;
}

function validateRawRow(raw) {
  const m = readMap();
  const issues = [];
  if (!m.time || !parseTime(raw[m.time])) issues.push('time');
  if (!m.p || !Number.isFinite(parseNumber(raw[m.p]))) issues.push('p');
  return issues;
}

function rawRowIsAbnormal(raw) {
  const m = readMap();
  return (m.abnormal && parseFlag(raw[m.abnormal])) || (m.outage && parseFlag(raw[m.outage])) || (m.transfer && parseFlag(raw[m.transfer]));
}

function getEditorFilteredIndices() {
  const q = norm(state.editor.query);
  const filter = state.editor.filter;
  const indices = [];
  for (let i=0; i<state.rawRows.length; i++) {
    const raw = state.rawRows[i];
    if (q) {
      const joined = norm(state.headers.map(h => raw[h]).join(' '));
      if (!joined.includes(q)) continue;
    }
    if (!editorDateFilterPass(raw)) continue;
    if (filter === 'invalid' && validateRawRow(raw).length === 0) continue;
    if (filter === 'abnormal' && !rawRowIsAbnormal(raw)) continue;
    indices.push(i);
  }
  return indices;
}

function markEditorDirty(dirty=true) {
  state.editor.dirty = dirty;
  renderEditorStatus();
}

function renderEditorStatus(extra='') {
  const total = state.rawRows.length;
  const invalid = state.rawRows.reduce((n, r) => n + (validateRawRow(r).length ? 1 : 0), 0);
  const abnormal = state.rawRows.reduce((n, r) => n + (rawRowIsAbnormal(r) ? 1 : 0), 0);
  const selected = state.editor.selected.size;
  const dirty = state.editor.dirty;
  const status = [];
  status.push(`<span class="pill">${total} dòng thô</span>`);
  status.push(`<span class="pill ${invalid?'bad':'ok'}">${invalid} dòng lỗi</span>`);
  status.push(`<span class="pill ${abnormal?'warn':''}">${abnormal} dòng bất thường</span>`);
  if (selected) status.push(`<span class="pill warn">đã chọn ${selected}</span>`);
  status.push(`<span class="pill ${dirty?'warn':'ok'}">${dirty?'có thay đổi chưa áp dụng':'đã áp dụng'}</span>`);
  if (extra) status.push(`<span class="pill">${escapeHtml(extra)}</span>`);
  $('editorStatus').innerHTML = status.join('');
}

function renderEditorTable() {
  const box = $('editorBox');
  renderEditorStatus();
  if (!state.headers.length || !state.rawRows.length) {
    box.innerHTML = '<table><tbody><tr><td>Chưa có dữ liệu để hiệu chỉnh</td></tr></tbody></table>';
    return;
  }
  const indices = getEditorFilteredIndices();
  const pageSize = Math.max(10, Math.floor(parseNumber($('editorPageSize').value) || state.editor.pageSize || 100));
  state.editor.pageSize = pageSize;
  const totalPages = Math.max(1, Math.ceil(indices.length / pageSize));
  state.editor.page = Math.min(Math.max(1, Math.floor(parseNumber($('editorPage').value) || state.editor.page || 1)), totalPages);
  $('editorPage').value = state.editor.page;
  const start = (state.editor.page - 1) * pageSize;
  const shown = indices.slice(start, start + pageSize);
  const headers = state.headers;
  let html = '<table class="editorTable"><thead><tr><th>Chọn</th><th>#</th>' + headers.map(h => `<th title="${escapeHtml(h)}">${escapeHtml(headerDisplayName(h))}</th>`).join('') + '</tr></thead><tbody>';
  for (const idx of shown) {
    const raw = state.rawRows[idx];
    const invalid = validateRawRow(raw).length ? 'invalid' : '';
    const selectedClass = state.editor.selected.has(idx) ? 'selected' : '';
    const cls = [invalid, selectedClass].filter(Boolean).join(' ');
    html += `<tr${cls ? ' class="' + cls + '"' : ''}><td><input class="rowSelect" type="checkbox" data-row="${idx}" ${state.editor.selected.has(idx)?'checked':''}></td><td>${idx+1}</td>`;
    for (const h of headers) {
      html += `<td contenteditable="true" data-row="${idx}" data-col="${escapeHtml(h)}">${escapeHtml(raw[h])}</td>`;
    }
    html += '</tr>';
  }
  html += '</tbody></table>';
  html += `<div class="note" style="padding:8px">Đang hiển thị ${shown.length}/${indices.length} dòng lọc, trang ${state.editor.page}/${totalPages}. Sửa trực tiếp trong ô rồi bấm “Lưu thay đổi vào dữ liệu”.</div>`;
  box.innerHTML = html;
  renderEditorStatus(`trang ${state.editor.page}/${totalPages}`);
}

function addEditorRow() {
  if (!state.headers.length) {
    state.headers = ['time','station','p_mw','temperature','rain','is_holiday','is_abnormal','is_outage','is_transfer'];
    fillColumnSelects(state.headers);
  }
  const r = {};
  for (const h of state.headers) r[h] = '';
  if (state.rawRows.length) {
    const last = state.rawRows[state.rawRows.length - 1];
    for (const h of state.headers) r[h] = '';
    const m = readMap();
    if (m.station) r[m.station] = last[m.station] || '';
    if (m.rain) r[m.rain] = last[m.rain] || '0';
    if (m.holiday) r[m.holiday] = '0';
    if (m.abnormal) r[m.abnormal] = '0';
    if (m.outage) r[m.outage] = '0';
    if (m.transfer) r[m.transfer] = '0';
  }
  state.rawRows.push(r);
  state.editor.page = Math.ceil(state.rawRows.length / state.editor.pageSize);
  $('editorPage').value = state.editor.page;
  markEditorDirty(true);
  renderEditorTable();
  log('Đã thêm 1 dòng trống để nhập/sửa.');
}

function deleteSelectedEditorRows() {
  const selected = state.editor.selected;
  if (!selected.size) { log('Chưa chọn dòng để xóa.'); return; }
  const before = state.rawRows.length;
  state.rawRows = state.rawRows.filter((_, i) => !selected.has(i));
  selected.clear();
  state.editor.page = 1;
  $('editorPage').value = 1;
  markEditorDirty(true);
  renderEditorTable();
  log(`Đã xóa ${before - state.rawRows.length} dòng đã chọn.`);
}

function applyEditorEdits() {
  if (!state.headers.length) throw new Error('Chưa có dữ liệu.');
  normalizeRows();
  applyDataInfo();
  renderEditorTable();
  previewData();
  saveSettingsToLocal();
  markEditorDirty(false);
  log('Đã áp dụng thay đổi bảng vào dữ liệu chuẩn hóa.');
}


function getCurrentPageEditorIndices() {
  const indices = getEditorFilteredIndices();
  const pageSize = Math.max(10, Math.floor(parseNumber($('editorPageSize')?.value) || state.editor.pageSize || 100));
  const totalPages = Math.max(1, Math.ceil(indices.length / pageSize));
  const page = Math.min(Math.max(1, Math.floor(parseNumber($('editorPage')?.value) || state.editor.page || 1)), totalPages);
  const start = (page - 1) * pageSize;
  return indices.slice(start, start + pageSize);
}

function selectVisibleEditorRows() {
  for (const idx of getCurrentPageEditorIndices()) state.editor.selected.add(idx);
  renderEditorTable();
  log(`Đã chọn ${getCurrentPageEditorIndices().length} dòng đang hiển thị.`);
}

function selectFilteredEditorRows() {
  const indices = getEditorFilteredIndices();
  for (const idx of indices) state.editor.selected.add(idx);
  renderEditorTable();
  log(`Đã chọn ${indices.length} dòng sau bộ lọc hiện tại.`);
}

function clearEditorSelection() {
  state.editor.selected.clear();
  renderEditorTable();
  log('Đã bỏ chọn tất cả dòng.');
}

function resetDateFilter() {
  state.editor.dateMode = 'all'; state.editor.dateSingle = ''; state.editor.dateMulti = ''; state.editor.dateFrom = ''; state.editor.dateTo = '';
  if ($('dateFilterMode')) $('dateFilterMode').value = 'all';
  if ($('dateFilterSingle')) $('dateFilterSingle').value = '';
  if ($('dateFilterMulti')) $('dateFilterMulti').value = '';
  if ($('dateFilterFrom')) $('dateFilterFrom').value = '';
  if ($('dateFilterTo')) $('dateFilterTo').value = '';
  state.editor.page = 1; $('editorPage').value = 1;
  renderEditorTable();
  log('Đã bỏ bộ lọc ngày.');
}

function selectedEditorIndicesOrLog() {
  const arr = [...state.editor.selected].filter(i => i >= 0 && i < state.rawRows.length).sort((a,b)=>a-b);
  if (!arr.length) log('Chưa chọn dòng. Có thể dùng nút “Chọn tất cả dòng đang hiển thị” hoặc “Chọn tất cả dòng sau lọc ngày”.');
  return arr;
}

function applyQuickFillToSelected() {
  const indices = selectedEditorIndicesOrLog();
  if (!indices.length) return;
  const tempVal = $('quickTemp')?.value.trim() ?? '';
  const rainVal = $('quickRain')?.value ?? '';
  const holidayVal = $('quickHoliday')?.value ?? '';
  const abnormalVal = $('quickAbnormal')?.value ?? '';
  const outageVal = $('quickOutage')?.value ?? '';
  const transferVal = $('quickTransfer')?.value ?? '';
  const stationVal = $('quickStation')?.value.trim() ?? '';
  const customCol = $('quickCustomCol')?.value ?? '';
  const customVal = $('quickCustomValue')?.value ?? '';
  let changed = 0;
  const setKey = (idx, key, value) => {
    const col = ensureMappedColumn(key);
    state.rawRows[idx][col] = value;
    changed++;
  };
  for (const idx of indices) {
    const raw = state.rawRows[idx];
    if (!raw) continue;
    if (tempVal !== '') setKey(idx, 'temp', tempVal);
    if (rainVal !== '') setKey(idx, 'rain', rainVal);
    if (holidayVal !== '') {
      if (holidayVal === 'auto') {
        const key = getRawDateKey(raw);
        const d = dateFromKey(key);
        setKey(idx, 'holiday', holidayByRules(d) ? '1' : '0');
      } else setKey(idx, 'holiday', holidayVal);
    }
    if (abnormalVal !== '') setKey(idx, 'abnormal', abnormalVal);
    if (outageVal !== '') setKey(idx, 'outage', outageVal);
    if (transferVal !== '') setKey(idx, 'transfer', transferVal);
    if (stationVal !== '') setKey(idx, 'station', stationVal);
    if (customCol && customVal !== '') { raw[customCol] = customVal; changed++; }
  }
  if (!changed) { log('Chưa nhập giá trị nào để điền nhanh.'); return; }
  markEditorDirty(true);
  renderEditorTable();
  log(`Đã điền nhanh ${changed} ô cho ${indices.length} dòng đã chọn.`);
}

function autoHolidayForIndices(indices) {
  if (!indices.length) { log('Không có dòng nào để tự nhận dạng ngày nghỉ/lễ.'); return; }
  const col = ensureMappedColumn('holiday');
  const clearNonMatch = $('holidayClearNonMatch')?.checked ?? true;
  let one = 0, zero = 0, skipped = 0;
  for (const idx of indices) {
    const raw = state.rawRows[idx];
    const key = getRawDateKey(raw);
    const d = dateFromKey(key);
    if (!d) { skipped++; continue; }
    if (holidayByRules(d)) { raw[col] = '1'; one++; }
    else if (clearNonMatch) { raw[col] = '0'; zero++; }
  }
  markEditorDirty(true);
  renderEditorTable();
  log(`Đã tự nhận dạng ngày nghỉ/lễ: ${one} dòng = 1, ${zero} dòng = 0, bỏ qua ${skipped} dòng lỗi ngày.`);
}

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('SCADA_LOAD_FORECAST_LV4_DB', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('kv');
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbSet(key, value) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction('kv', 'readwrite');
    tx.objectStore('kv').put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function idbGet(key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction('kv', 'readonly');
    const req = tx.objectStore('kv').get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function saveSettingsToLocal() {
  try { localStorage.setItem('SCADA_LOAD_FORECAST_LV4_SETTINGS', JSON.stringify({colMap: state.colMap, sourceFileName: state.sourceFileName})); } catch(_) {}
}

function loadSettingsFromLocal() {
  try { return JSON.parse(localStorage.getItem('SCADA_LOAD_FORECAST_LV4_SETTINGS') || '{}'); } catch(_) { return {}; }
}

async function saveDatasetOffline() {
  if (!state.headers.length) throw new Error('Chưa có dữ liệu để lưu.');
  const payload = {
    type: 'SCADA_LOAD_DATA_OFFLINE_LV4', savedAt: new Date().toISOString(), sourceFileName: state.sourceFileName,
    headers: state.headers, rawRows: state.rawRows, colMap: readMap(), delimiter: state.delimiter
  };
  await idbSet('currentDataset', payload);
  saveSettingsToLocal();
  markEditorDirty(false);
  log(`Đã lưu dữ liệu vào trình duyệt offline: ${state.rawRows.length} dòng.`);
}

async function loadDatasetOffline() {
  const payload = await idbGet('currentDataset');
  if (!payload || !payload.rawRows) { log('Chưa có bản dữ liệu offline đã lưu.'); return; }
  state.headers = payload.headers || [...new Set(payload.rawRows.flatMap(r => Object.keys(r)))];
  state.rawRows = payload.rawRows.map(r => ({...r}));
  state.sourceFileName = payload.sourceFileName || 'offline_saved_dataset.json';
  state.delimiter = payload.delimiter || ',';
  fillColumnSelects(state.headers);
  if (payload.colMap) {
    for (const [k, v] of Object.entries(payload.colMap)) {
      const id = {time:'colTime',p:'colP',station:'colStation',temp:'colTemp',rain:'colRain',holiday:'colHoliday',abnormal:'colAbnormal',outage:'colOutage',transfer:'colTransfer'}[k];
      if (id && $(id) && [...$(id).options].some(o => o.value === v)) $(id).value = v;
    }
  }
  normalizeRows();
  applyDataInfo();
  renderEditorTable();
  previewData();
  markEditorDirty(false);
  log(`Đã nạp bản dữ liệu offline đã lưu: ${state.rawRows.length} dòng.`);
}

function applySavedColumnMapIfPossible() {
  const saved = loadSettingsFromLocal();
  if (!saved.colMap) return;
  const mapId = {time:'colTime',p:'colP',station:'colStation',temp:'colTemp',rain:'colRain',holiday:'colHoliday',abnormal:'colAbnormal',outage:'colOutage',transfer:'colTransfer'};
  for (const [k, v] of Object.entries(saved.colMap)) {
    const id = mapId[k];
    if (id && $(id) && [...$(id).options].some(o => o.value === v)) $(id).value = v;
  }
}

function forecastNext() {
  if (!state.model) throw new Error('Chưa có model.');
  const baseRows = getSelectedRows();
  if (baseRows.length < 5) throw new Error('Chưa có đủ dữ liệu gần nhất để dự báo.');
  const model = state.model;
  const station = $('stationSelect').value;
  if (model.station && model.station !== '__ALL__' && station !== model.station) {
    log(`Cảnh báo: model được huấn luyện cho '${model.station}', hiện đang chọn '${station}'.`);
  }
  const tempDefault = parseNumber($('futureTemp').value);
  const rainDefault = parseNumber($('futureRain').value) || 0;
  const interval = model.intervalMinutes || detectIntervalMinutes(baseRows);
  const nPerDay = model.nPerDay || Math.round(1440/interval);
  const lastTemp = [...baseRows].reverse().find(r => Number.isFinite(r.temp))?.temp;
  const tempUse = Number.isFinite(tempDefault) ? tempDefault : lastTemp;
  const series = baseRows.map(r => ({...r})).sort((a,b)=>a.time-b.time);
  const steps = Math.max(1, Math.floor(parseNumber($('forecastSteps').value) || 24));
  const forecast = [];
  let lastTime = series[series.length-1].time;
  for (let s=1; s<=steps; s++) {
    const t = new Date(lastTime.getTime() + interval*60000);
    const row = {time:t, p:NaN, station: station === '__ALL__' ? 'ALL' : station, temp:tempUse, rain:rainDefault, holiday:0, abnormal:0, outage:0, transfer:0};
    series.push(row);
    const idx = series.length - 1;
    const x = computeFeatureVector(series, idx, nPerDay, {temp: tempUse, rain: rainDefault});
    const pred = Math.max(0, predictModel(model, x));
    row.p = pred;
    forecast.push({
      step:s,
      time:fmtTime(t),
      station:row.station,
      forecast_p_mw:formatNum(pred,3),
      temp:Number.isFinite(tempUse)?formatNum(tempUse,1):'',
      rain:rainDefault
    });
    lastTime = t;
  }
  state.forecastRows = forecast;
  renderTable(forecast, ['step','time','station','forecast_p_mw','temp','rain'], 500);
  drawSeries(series.slice(-Math.min(steps*2, 200)).map(r=>r.time), series.slice(-Math.min(steps*2, 200)).map((r,i,arr)=> i < arr.length-steps ? r.p : NaN), series.slice(-Math.min(steps*2, 200)).map((r,i,arr)=> i >= arr.length-steps ? r.p : NaN), 'Lịch sử gần nhất', 'Dự báo');
  $('exportForecastBtn').disabled = false;
  log(`Đã dự báo ${steps} bước tiếp theo, mỗi bước ${interval} phút.`);
}

function exportModel() {
  if (!state.model) return;
  saveTextFile('model_gbdt.json', JSON.stringify(state.model, null, 2), 'application/json');
  log('Đã xuất model_gbdt.json. Copy file này vào mạng SCADA để dùng offline.');
}

function exportForecast() {
  if (!state.forecastRows.length) return;
  saveTextFile('forecast.csv', toCSV(state.forecastRows, ['step','time','station','forecast_p_mw','temp','rain']), 'text/csv;charset=utf-8');
  log('Đã xuất forecast.csv.');
}

function previewData() {
  if (!state.rows.length) normalizeRows();
  const rows = getSelectedRows().slice(0, 200).map(r => ({
    time: fmtTime(r.time), station:r.station, p:formatNum(r.p,3), temp:formatNum(r.temp,1), rain: r.rain ?? '', holiday:r.holiday, abnormal:r.abnormal, outage:r.outage, transfer:r.transfer
  }));
  renderTable(rows, ['time','station','p','temp','rain','holiday','abnormal','outage','transfer']);
  log(`Xem trước ${rows.length} dòng.`);
}

function applyDataInfo() {
  const rows = state.rows.length ? state.rows : [];
  if (!rows.length) { $('dataInfo').innerHTML = ''; return; }
  const interval = detectIntervalMinutes(rows);
  const minT = fmtTime(rows[0].time), maxT = fmtTime(rows[rows.length-1].time);
  $('dataInfo').innerHTML = `
    <span class="pill">${rows.length} dòng hợp lệ</span>
    <span class="pill">${minT} → ${maxT}</span>
    <span class="pill">Interval ≈ ${interval} phút</span>`;
}

function buildSampleCSV() {
  const headers = ['time','station','p_mw','temperature','rain','is_holiday','is_abnormal','is_outage','is_transfer'];
  const rows=[];
  const start = new Date(2026, 3, 1, 0, 0, 0);
  for (let i=0;i<24*75;i++) {
    const t = new Date(start.getTime() + i*60*60000);
    const hour = t.getHours();
    const dow = t.getDay();
    const day = Math.floor(i/24);
    const temp = 26 + 7*Math.sin(2*Math.PI*(hour-11)/24) + 3*Math.sin(2*Math.PI*day/30);
    const weekend = (dow===0 || dow===6) ? 1 : 0;
    const eveningPeak = Math.exp(-Math.pow((hour-20)/3,2))*12;
    const morningPeak = Math.exp(-Math.pow((hour-9)/3,2))*5;
    const industrial = weekend ? -5 : 4;
    const heat = Math.max(0,temp-30)*1.2;
    const trend = day*0.035;
    const p = 35 + eveningPeak + morningPeak + industrial + heat + trend + (Math.sin(i*1.7)*0.7);
    rows.push([fmtTime(t),'E22.1',p.toFixed(3),temp.toFixed(1),0,0,0,0,0]);
  }
  return [headers.join(','), ...rows.map(r=>r.join(','))].join('\n');
}

function loadTextFile(file, cb) {
  const reader = new FileReader();
  reader.onload = () => cb(String(reader.result || ''));
  reader.onerror = () => log('Không đọc được file: ' + reader.error?.message);
  reader.readAsText(file, 'utf-8');
}


// ======================== LV5 EXTENSIONS ========================
state.qualityIssues = state.qualityIssues || [];
state.thresholds = state.thresholds || [];
state.appMode = state.appMode || 'external';
state.lv5 = state.lv5 || {version:'LV5.4'};

function getHybridWeights() {
  const read = id => parseNumber($(id)?.value);
  let w = {
    gbdt: Number.isFinite(read('wGbdt')) ? read('wGbdt') : 0.50,
    similar: Number.isFinite(read('wSimilar')) ? read('wSimilar') : 0.25,
    week: Number.isFinite(read('wWeek')) ? read('wWeek') : 0.15,
    trend: Number.isFinite(read('wTrend')) ? read('wTrend') : 0.10
  };
  const mode = $('forecastBlend')?.value || 'hybrid';
  if (mode === 'gbdt') w = {gbdt:1, similar:0, week:0, trend:0};
  if (mode === 'similar') w = {gbdt:0, similar:1, week:0, trend:0};
  const sum = Object.values(w).filter(Number.isFinite).reduce((a,b)=>a+Math.max(0,b),0) || 1;
  Object.keys(w).forEach(k => w[k] = Math.max(0, w[k]) / sum);
  return w;
}

function addProcessColumns() {
  let added = [];
  for (const name of ['p_goc','du_lieu_noi_suy','ghi_chu_xu_ly','bo_khoi_huan_luyen']) {
    if (!state.headers.includes(name)) {
      state.headers.push(name);
      state.rawRows.forEach(r => { r[name] = ''; });
      added.push(name);
    }
  }
  fillColumnSelects(state.headers);
  refreshQuickCustomColumns();
  if (added.length) log('Đã thêm cột xử lý: ' + added.join(', '));
  return added;
}

function renderQualityReport() {
  const box = $('qualityBox');
  const sumBox = $('qualitySummary');
  if (!box || !sumBox) return;
  const issues = state.qualityIssues || [];
  const byType = issues.reduce((m, x) => { m[x.type] = (m[x.type] || 0) + 1; return m; }, {});
  const parts = [`<span class="pill ${issues.length?'warn':'ok'}">${issues.length} cảnh báo</span>`];
  Object.entries(byType).forEach(([k,v]) => parts.push(`<span class="pill">${escapeHtml(k)}: ${v}</span>`));
  sumBox.innerHTML = parts.join('');
  if (!issues.length) { box.innerHTML = '<table><tbody><tr><td>Không phát hiện lỗi lớn</td></tr></tbody></table>'; return; }
  const max = Math.max(1, Math.floor(parseNumber($('qualityMaxRows')?.value) || 300));
  const shown = issues.slice(0, max);
  let html = '<table><thead><tr><th>Dòng</th><th>Thời gian</th><th>Trạm/Lộ</th><th>Loại lỗi</th><th>Giá trị</th><th>Gợi ý xử lý</th></tr></thead><tbody>';
  for (const it of shown) {
    html += `<tr><td>${it.rowIndex != null ? it.rowIndex + 2 : ''}</td><td>${escapeHtml(it.time || '')}</td><td>${escapeHtml(it.station || '')}</td><td>${escapeHtml(it.type)}</td><td>${escapeHtml(it.value ?? '')}</td><td>${escapeHtml(it.suggestion || '')}</td></tr>`;
  }
  html += '</tbody></table>';
  if (issues.length > max) html += `<div class="note" style="padding:8px">Đang hiển thị ${max}/${issues.length} cảnh báo.</div>`;
  box.innerHTML = html;
}

function runQualityCheckBaseLV73() {
  const m = readMap();
  const issues = [];
  const spikePct = Math.max(1, parseNumber($('spikePercent')?.value) || 35);
  const minValidP = parseNumber($('minValidP')?.value);
  for (let i=0; i<state.rawRows.length; i++) {
    const raw = state.rawRows[i];
    const time = m.time ? parseTime(raw[m.time]) : null;
    const p = m.p ? parseNumber(raw[m.p]) : NaN;
    const station = m.station ? String(raw[m.station] ?? '').trim() : 'ALL';
    if (!time) issues.push({rowIndex:i, time:'', station, type:'Sai thời gian', value: raw[m.time], suggestion:'Sửa định dạng ngày giờ hoặc xóa khỏi huấn luyện'});
    if (!Number.isFinite(p)) issues.push({rowIndex:i, time: time?fmtTime(time):'', station, type:'Thiếu/sai P', value: raw[m.p], suggestion:'Nội suy P hoặc kiểm tra nguồn dữ liệu'});
    if (Number.isFinite(p) && p < 0) issues.push({rowIndex:i, time:fmtTime(time), station, type:'P âm', value:p, suggestion:'Đánh dấu bất thường hoặc sửa theo số đo đúng'});
    if (Number.isFinite(p) && Number.isFinite(minValidP) && p <= minValidP) {
      const outage = m.outage ? parseFlag(raw[m.outage]) : 0;
      const transfer = m.transfer ? parseFlag(raw[m.transfer]) : 0;
      const abnormal = m.abnormal ? parseFlag(raw[m.abnormal]) : 0;
      let type = 'P thấp bất thường';
      let suggestion = 'Nếu không có cờ cắt điện/chuyển tải thì nên nội suy hoặc đánh dấu bất thường/bỏ khỏi huấn luyện';
      if (outage) { type = 'P thấp do cắt điện/sự cố'; suggestion = 'Giữ làm dữ liệu vận hành, nhưng bỏ khỏi huấn luyện; chỉ nội suy nếu cần chuỗi P liên tục để dự báo'; }
      else if (transfer) { type = 'P thấp do chuyển tải'; suggestion = 'Kiểm tra lộ/trạm nhận chuyển tải có tăng tương ứng; bỏ khỏi huấn luyện hoặc nội suy khi cần chuỗi liên tục'; }
      else if (abnormal) { type = 'P thấp đã đánh dấu bất thường'; suggestion = 'Bỏ khỏi huấn luyện; chỉ nội suy nếu cần chuỗi liên tục'; }
      issues.push({rowIndex:i, time:time?fmtTime(time):'', station, type, value:p, suggestion});
    }
  }
  let rows = [];
  try { normalizeRows(); rows = getSelectedRows(); } catch(_) { rows = state.rows || []; }
  const groups = new Map();
  rows.forEach(r => { const k = r.station || 'ALL'; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(r); });
  const expected = $('expectedInterval')?.value && $('expectedInterval').value !== 'auto' ? parseNumber($('expectedInterval').value) : null;
  for (const [station, arr0] of groups) {
    const arr = arr0.slice().sort((a,b)=>a.time-b.time);
    const interval = expected || detectIntervalMinutes(arr);
    const seen = new Map();
    for (let i=0;i<arr.length;i++) {
      const r = arr[i], key = fmtTime(r.time);
      if (seen.has(key)) issues.push({rowIndex:r.idx, time:key, station, type:'Trùng mốc thời gian', value:r.p, suggestion:'Giữ một dòng đúng, xóa hoặc đánh dấu dòng trùng'});
      else seen.set(key, r.idx);
      if (i > 0) {
        const prev = arr[i-1];
        const gap = (r.time - prev.time)/60000;
        if (gap > interval * 1.5) issues.push({rowIndex:r.idx, time:fmtTime(r.time), station, type:'Mất mốc thời gian', value:`thiếu khoảng ${Math.round(gap/interval)-1} mốc`, suggestion:'Bổ sung mốc thiếu bằng nội suy'});
        const base = Math.max(Math.abs(prev.p), 0.001);
        const pct = Math.abs(r.p - prev.p)/base*100;
        if (Number.isFinite(pct) && pct >= spikePct) issues.push({rowIndex:r.idx, time:fmtTime(r.time), station, type:'P tăng/giảm đột biến', value:`${formatNum(pct,1)}%`, suggestion:'Kiểm tra sự cố/chuyển tải hoặc đánh dấu bất thường'});
      }
    }
  }
  // Kiểm tra nhanh logic chuyển tải: một ngăn lộ giảm thấp do chuyển tải thường phải có ngăn lộ/khu vực khác tăng lên cùng thời điểm.
  try {
    const byTime = new Map();
    const byStation = new Map();
    for (const r of rows) {
      const tk = fmtTime(r.time);
      if (!byTime.has(tk)) byTime.set(tk, []);
      byTime.get(tk).push(r);
      const sk = r.station || 'ALL';
      if (!byStation.has(sk)) byStation.set(sk, []);
      byStation.get(sk).push(r);
    }
    for (const arr of byStation.values()) arr.sort((a,b)=>a.time-b.time);
    const prevByIdx = new Map();
    for (const arr of byStation.values()) {
      for (let i=1;i<arr.length;i++) prevByIdx.set(arr[i].idx, arr[i-1]);
    }
    for (const r of rows) {
      if (!Number.isFinite(minValidP) || !(r.transfer && Number.isFinite(r.p) && r.p <= minValidP)) continue;
      const sameTime = byTime.get(fmtTime(r.time)) || [];
      let hasReceiverIncrease = false;
      for (const other of sameTime) {
        if (other.idx === r.idx || other.station === r.station) continue;
        const prev = prevByIdx.get(other.idx);
        if (!prev || !Number.isFinite(prev.p) || !Number.isFinite(other.p)) continue;
        const delta = other.p - prev.p;
        const pct = Math.abs(prev.p) > 0.001 ? delta / Math.abs(prev.p) * 100 : 0;
        if (delta > 0 && pct >= 5) { hasReceiverIncrease = true; break; }
      }
      if (!hasReceiverIncrease) {
        issues.push({rowIndex:r.idx, time:fmtTime(r.time), station:r.station, type:'Chuyển tải chưa thấy lộ nhận tăng', value:r.p, suggestion:'Kiểm tra lại dữ liệu cùng thời điểm: nếu là chuyển tải thật, nên có trạm/lộ khác tăng tải; nếu không có thì có thể là lỗi đo hoặc cắt điện'});
      }
    }
  } catch(_) {}
  state.qualityIssues = issues;
  renderQualityReport();
  log(`Kiểm tra chất lượng dữ liệu hoàn tất: ${issues.length} cảnh báo.`);
  return issues;
}

function selectQualityRows() {
  const rows = [...new Set((state.qualityIssues || []).map(x => x.rowIndex).filter(Number.isInteger))];
  rows.forEach(i => state.editor.selected.add(i));
  renderEditorTable();
  log(`Đã chọn ${rows.length} dòng lỗi trên bảng hiệu chỉnh.`);
}

function markQualityAbnormal() {
  if (!state.headers.length) return;
  const abnormalCol = ensureMappedColumn('abnormal', 'Bất thường');
  addProcessColumns();
  let n=0;
  for (const it of state.qualityIssues || []) {
    if (Number.isInteger(it.rowIndex) && state.rawRows[it.rowIndex]) {
      state.rawRows[it.rowIndex][abnormalCol] = '1';
      state.rawRows[it.rowIndex]['ghi_chu_xu_ly'] = `${state.rawRows[it.rowIndex]['ghi_chu_xu_ly'] || ''} ${it.type}`.trim();
      n++;
    }
  }
  markEditorDirty(true);
  renderEditorTable();
  log(`Đã đánh dấu bất thường ${n} dòng lỗi.`);
}

function exportQualityReport() {
  const rows = (state.qualityIssues || []).map(x => ({dong: x.rowIndex != null ? x.rowIndex + 2 : '', thoi_gian:x.time||'', tram_lo:x.station||'', loai_loi:x.type, gia_tri:x.value??'', goi_y:x.suggestion||''}));
  if (!rows.length) { log('Chưa có báo cáo chất lượng để xuất.'); return; }
  saveTextFile('quality_report.csv', toCSV(rows, ['dong','thoi_gian','tram_lo','loai_loi','gia_tri','goi_y']), 'text/csv;charset=utf-8');
}

function groupRowsByStationRaw() {
  const m = readMap();
  const groups = new Map();
  for (let i=0; i<state.rawRows.length; i++) {
    const raw = state.rawRows[i];
    const t = m.time ? parseTime(raw[m.time]) : null;
    if (!t) continue;
    const st = m.station ? String(raw[m.station] ?? '').trim() || 'ALL' : 'ALL';
    const key = st;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({raw, rawIndex:i, time:t, p:m.p?parseNumber(raw[m.p]):NaN, station:st});
  }
  for (const arr of groups.values()) arr.sort((a,b)=>a.time-b.time);
  return groups;
}

function inferPForMissing(method, arr, leftIdx, rightIdx, targetTime, interval) {
  const left = arr[leftIdx], right = arr[rightIdx];
  if (method === 'prev') return left?.p;
  if (method === 'day' || method === 'week') {
    const backMs = (method === 'day' ? 24 : 24*7) * 3600*1000;
    const want = targetTime.getTime() - backMs;
    const found = arr.find(x => Math.abs(x.time.getTime() - want) <= interval*60000*0.51 && Number.isFinite(x.p));
    if (found) return found.p;
  }
  if (left && right && Number.isFinite(left.p) && Number.isFinite(right.p)) {
    const f = (targetTime - left.time) / (right.time - left.time);
    return left.p + (right.p - left.p) * f;
  }
  return Number.isFinite(left?.p) ? left.p : NaN;
}


function editorDateFilterPassTime(time) {
  const mode = state.editor.dateMode || 'all';
  if (mode === 'all') return true;
  const key = fmtDateKey(time);
  if (mode === 'single') return key === parseDateKey(state.editor.dateSingle);
  if (mode === 'multi') return parseDateList(state.editor.dateMulti).has(key);
  if (mode === 'range') {
    const from = parseDateKey(state.editor.dateFrom);
    const to = parseDateKey(state.editor.dateTo);
    if (from && key < from) return false;
    if (to && key > to) return false;
    return true;
  }
  return true;
}

function getInvalidPThreshold() {
  const v = parseNumber($('minValidP')?.value);
  return Number.isFinite(v) ? v : NaN;
}

function rawOperationalFlags(raw, m) {
  return {
    abnormal: m?.abnormal ? parseFlag(raw?.[m.abnormal]) : 0,
    outage: m?.outage ? parseFlag(raw?.[m.outage]) : 0,
    transfer: m?.transfer ? parseFlag(raw?.[m.transfer]) : 0,
    excludeTrain: parseFlag(raw?.['bo_khoi_huan_luyen'])
  };
}

function isLowPValue(p, threshold=getInvalidPThreshold()) {
  return Number.isFinite(p) && Number.isFinite(threshold) && p <= threshold;
}

function shouldInterpolateLowP(raw, m, p, threshold=getInvalidPThreshold()) {
  if (!isLowPValue(p, threshold)) return false;
  const mode = $('lowPHandlingMode')?.value || 'auto';
  const flags = rawOperationalFlags(raw, m);
  const hasOperationFlag = !!(flags.abnormal || flags.outage || flags.transfer || flags.excludeTrain);
  if (mode === 'fill_all') return true;
  if (mode === 'exclude_events' && hasOperationFlag) return false;
  // auto: P thấp/P=0 chỉ coi là lỗi cần nội suy khi chưa có cờ vận hành.
  // Nếu đã có cờ cắt điện/chuyển tải/bất thường, giữ số đo gốc và loại khỏi huấn luyện.
  return !hasOperationFlag;
}

function isBadPForInterpolation(p, raw, m, threshold=getInvalidPThreshold()) {
  const original = m?.p ? String(raw?.[m.p] ?? '').trim() : '';
  if (!original) return true;
  if (!Number.isFinite(p)) return true;
  if (isLowPValue(p, threshold)) return shouldInterpolateLowP(raw, m, p, threshold);
  return false;
}

function markLowPOperationalEventsForTraining() {
  const m = readMap();
  if (!m.p) throw new Error('Cần ánh xạ cột Công suất P.');
  addProcessColumns();
  const targetSet = getInterpolationTargetIndexSet({});
  const threshold = getInvalidPThreshold();
  let marked=0, lowNoFlag=0, checked=0;
  for (const idx of targetSet) {
    const raw = state.rawRows[idx];
    if (!raw) continue;
    const p = parseNumber(raw[m.p]);
    if (!isLowPValue(p, threshold)) continue;
    checked++;
    const flags = rawOperationalFlags(raw, m);
    if (flags.outage || flags.transfer || flags.abnormal) {
      raw['bo_khoi_huan_luyen'] = '1';
      const reason = flags.outage ? 'cat dien/su co' : (flags.transfer ? 'chuyen tai' : 'bat thuong');
      raw['ghi_chu_xu_ly'] = `${raw['ghi_chu_xu_ly'] || ''} LV5.3 giu P thap do ${reason}, bo khoi huan luyen`.trim();
      marked++;
    } else {
      lowNoFlag++;
    }
  }
  normalizeRows(); renderEditorTable(); previewData(); markEditorDirty(true);
  $('interpolationInfo').innerHTML = `<span class="pill ok">Đã bỏ khỏi huấn luyện ${marked} dòng P thấp có cờ vận hành</span><span class="pill ${lowNoFlag?'warn':''}">${lowNoFlag} dòng P thấp chưa có cờ: nên nội suy hoặc đánh dấu sự kiện</span><span class="pill">Đã xét ${checked} dòng P thấp</span>`;
  log(`LV5.3 xử lý P thấp theo cờ: bỏ huấn luyện ${marked}, còn ${lowNoFlag} dòng P thấp chưa có cờ.`);
}

function getInterpolationTargetIndexSet(options={}) {
  const selected = [...(state.editor?.selected || new Set())].filter(i => i >= 0 && i < state.rawRows.length);
  if (selected.length) return new Set(selected);
  if (options.selectedOnly) return new Set();
  return new Set(getEditorFilteredIndices());
}

function interpolateMissingTimestamps() {
  if (!state.headers.length) throw new Error('Chưa có dữ liệu.');
  const m = readMap();
  if (!m.time || !m.p) throw new Error('Cần ánh xạ cột thời gian và P.');
  const stationFilter = $('stationSelect')?.value || '__ALL__';
  const scope = $('interpScope')?.value || 'current';
  const method = $('interpMethod')?.value || 'linear';
  const maxGap = Math.max(1, Math.floor(parseNumber($('interpMaxGap')?.value) || 12));
  const syntheticCol = 'du_lieu_noi_suy', noteCol='ghi_chu_xu_ly';
  addProcessColumns();

  // Khi người dùng đang chọn dòng có P = 0/trống/lỗi rồi bấm "Bổ sung mốc thiếu",
  // xử lý luôn các dòng đã chọn để tránh hiểu nhầm giữa "thiếu mốc thời gian" và "thiếu giá trị P".
  const fixedSelectedBadP = state.editor.selected.size ? fillInvalidP({silent:true, selectedOnly:true}) : 0;

  const groups = groupRowsByStationRaw();
  let added=0, skipped=0, skippedByDate=0;
  for (const [station, arr] of groups) {
    if (scope === 'current' && stationFilter !== '__ALL__' && station !== stationFilter) continue;
    const expected = $('expectedInterval')?.value !== 'auto' ? parseNumber($('expectedInterval')?.value) : detectIntervalMinutes(arr.map(x => ({time:x.time,p:x.p})));
    const interval = expected || 60;
    for (let i=1; i<arr.length; i++) {
      const prev = arr[i-1], next = arr[i];
      const gap = (next.time - prev.time)/60000;
      const miss = Math.round(gap/interval) - 1;
      if (miss <= 0) continue;
      if (miss > maxGap) { skipped += miss; continue; }
      for (let k=1; k<=miss; k++) {
        const t = new Date(prev.time.getTime() + k*interval*60000);
        if (!editorDateFilterPassTime(t)) { skippedByDate++; continue; }
        const row = {};
        state.headers.forEach(h => row[h] = '');
        row[m.time] = fmtTime(t);
        row[m.p] = formatNum(inferPForMissing(method, arr, i-1, i, t, interval), 3);
        if (m.station) row[m.station] = station;
        if (m.temp) {
          const a = parseNumber(prev.raw[m.temp]), b = parseNumber(next.raw[m.temp]);
          if (Number.isFinite(a) && Number.isFinite(b)) row[m.temp] = formatNum(a + (b-a)*(k/(miss+1)), 1);
          else if (Number.isFinite(a)) row[m.temp] = formatNum(a,1);
        }
        if (m.rain) row[m.rain] = '0';
        if (m.holiday) row[m.holiday] = holidayByRules(t) ? '1' : '0';
        if (m.abnormal) row[m.abnormal] = '0';
        if (m.outage) row[m.outage] = '0';
        if (m.transfer) row[m.transfer] = '0';
        row[syntheticCol] = '1';
        row[noteCol] = `LV5.3 noi suy moc thieu ${method}`;
        state.rawRows.push(row); added++;
      }
    }
  }
  state.rawRows.sort((a,b) => {
    const ta = parseTime(a[m.time]) || 0, tb = parseTime(b[m.time]) || 0;
    if (ta - tb) return ta - tb;
    const sa = m.station ? String(a[m.station] || '') : '';
    const sb = m.station ? String(b[m.station] || '') : '';
    return sa.localeCompare(sb, 'vi');
  });
  normalizeRows(); applyDataInfo(); renderEditorTable(); previewData(); markEditorDirty(true);
  const parts = [
    `<span class="pill ok">Đã bổ sung ${added} mốc thời gian thiếu</span>`,
    `<span class="pill ok">Đã nội suy ${fixedSelectedBadP} giá trị P trống/lỗi/P thấp cần nội suy trong dòng đã chọn</span>`,
    `<span class="pill ${skipped?'warn':''}">Bỏ qua ${skipped} mốc do gap quá lớn</span>`
  ];
  if (skippedByDate) parts.push(`<span class="pill warn">Bỏ qua ${skippedByDate} mốc ngoài bộ lọc ngày</span>`);
  $('interpolationInfo').innerHTML = parts.join('');
  log(`Nội suy hoàn tất: thêm ${added} mốc, sửa ${fixedSelectedBadP} P trống/lỗi/P thấp cần nội suy, bỏ qua ${skipped}.`);
}

function fillInvalidP(options={}) {
  const m = readMap();
  if (!m.time || !m.p) throw new Error('Cần ánh xạ thời gian/P.');
  addProcessColumns();
  const targetSet = getInterpolationTargetIndexSet(options);
  if (!targetSet.size) {
    if (!options.silent) {
      $('interpolationInfo').innerHTML = '<span class="pill warn">Chưa có dòng nào trong phạm vi xử lý. Hãy chọn dòng hoặc lọc ngày trước.</span>';
      log('Chưa có dòng nào để nội suy P.');
    }
    return 0;
  }
  const stationFilter = $('stationSelect')?.value || '__ALL__';
  const scope = $('interpScope')?.value || 'current';
  const threshold = getInvalidPThreshold();
  const groups = groupRowsByStationRaw();
  let fixed=0, considered=0;
  for (const [station, arr] of groups) {
    if (scope === 'current' && stationFilter !== '__ALL__' && station !== stationFilter) continue;
    const interval = detectIntervalMinutes(arr.map(x=>({time:x.time,p:x.p}))) || 60;
    for (let i=0;i<arr.length;i++) {
      const item = arr[i];
      if (!targetSet.has(item.rawIndex)) continue;
      if (!editorDateFilterPassTime(item.time)) continue;
      considered++;
      if (!isBadPForInterpolation(item.p, item.raw, m, threshold)) continue;
      let l=i-1; while(l>=0 && isBadPForInterpolation(arr[l].p, arr[l].raw, m, threshold)) l--;
      let r=i+1; while(r<arr.length && isBadPForInterpolation(arr[r].p, arr[r].raw, m, threshold)) r++;
      let val=NaN;
      const method = $('interpMethod')?.value || 'linear';
      if (method === 'day' || method === 'week') val = inferPForMissing(method, arr, l>=0?l:i, r<arr.length?r:i, item.time, interval);
      if (!Number.isFinite(val)) {
        if (l>=0 && r<arr.length) val = inferPForMissing('linear', arr, l, r, item.time, interval);
        else if (l>=0) val = arr[l].p;
        else if (r<arr.length) val = arr[r].p;
      }
      if (Number.isFinite(val)) {
        item.raw[m.p] = formatNum(val,3);
        item.raw['du_lieu_noi_suy'] = '1';
        item.raw['ghi_chu_xu_ly'] = `LV5.3 noi suy P loi/trong/<=nguong tu ${Number.isFinite(threshold)?threshold:'NaN'}`;
        fixed++;
      }
    }
  }
  normalizeRows(); renderEditorTable(); previewData(); markEditorDirty(true);
  if (!options.silent) {
    $('interpolationInfo').innerHTML = `<span class="pill ok">Đã nội suy ${fixed} giá trị P trống/lỗi/P thấp cần nội suy</span><span class="pill">Đã xét ${considered} dòng trong phạm vi chọn/lọc</span>`;
    log(`Đã nội suy ${fixed} giá trị P trống/lỗi/P thấp cần nội suy trên ${considered} dòng được xét.`);
  }
  return fixed;
}

function similarDayPrediction(series, idx, nPerDay) {
  const target = series[idx];
  const cand=[];
  for (let d=1; d<=56; d++) {
    const j = idx - nPerDay*d;
    if (j < 0) break;
    const r = series[j];
    if (!Number.isFinite(r.p)) continue;
    let score = 1;
    if (r.time.getDay() === target.time.getDay()) score += 3;
    if ((r.holiday||0) === (target.holiday||0)) score += 2;
    if (Number.isFinite(r.temp) && Number.isFinite(target.temp)) score += Math.max(0, 3 - Math.abs(r.temp-target.temp)/2);
    score += Math.max(0, 2 - d/14);
    cand.push({p:r.p, score, time:r.time});
  }
  cand.sort((a,b)=>b.score-a.score);
  const top = cand.slice(0, Math.min(10, cand.length));
  const sw = top.reduce((a,b)=>a+b.score,0);
  return sw ? top.reduce((a,b)=>a+b.p*b.score,0)/sw : NaN;
}

function lastWeekPrediction(series, idx, nPerDay) {
  const j = idx - nPerDay*7;
  return j >= 0 && Number.isFinite(series[j].p) ? series[j].p : NaN;
}

function trendPrediction(series, idx) {
  const vals=[];
  for (let k=1;k<=5 && idx-k>=0;k++) if (Number.isFinite(series[idx-k].p)) vals.push(series[idx-k].p);
  if (!vals.length) return NaN;
  if (vals.length < 2) return vals[0];
  const newest = vals[0], oldest = vals[vals.length-1];
  return Math.max(0, newest + (newest-oldest)/(vals.length-1));
}

function resolveModelForStation(station) {
  if (!state.model) return null;
  if (state.model.modelsByStation) return state.model.modelsByStation[station] || state.model.modelsByStation['__ALL__'] || null;
  return state.model;
}

function estimateRecentBias(model, rows, windowN=24) {
  if (!model || !rows || rows.length < 10 || !windowN) return 0;
  const ds = buildDataset(rows, model.featureMeans);
  const n = Math.min(Math.max(0, windowN), ds.X.length);
  if (!n) return 0;
  let errors=[];
  for (let i=ds.X.length-n; i<ds.X.length; i++) {
    const p = predictModel(model, ds.X[i]);
    if (Number.isFinite(p) && Number.isFinite(ds.y[i])) errors.push(ds.y[i]-p);
  }
  return mean(errors) || 0;
}

function updateForecastMetrics(rows) {
  const pVals = rows.map(r => parseNumber(r.forecast_p_mw)).filter(Number.isFinite);
  if (pVals.length) {
    const max = Math.max(...pVals); const idx = rows.findIndex(r => parseNumber(r.forecast_p_mw) === max);
    $('pmaxForecastVal').textContent = formatNum(max,3);
    $('pmaxTimeVal').textContent = rows[idx]?.time || '-';
  }
}

function drawErrorChart(times, actual, pred) {
  const canvas = $('errorChart'); if (!canvas) return;
  const ctx = canvas.getContext('2d'); ctx.clearRect(0,0,canvas.width,canvas.height);
  ctx.fillStyle = '#071120'; ctx.fillRect(0,0,canvas.width,canvas.height);
  const err = actual.map((a,i)=> Number.isFinite(a)&&Number.isFinite(pred[i]) ? pred[i]-a : NaN);
  const vals = err.filter(Number.isFinite); if (!vals.length) return;
  const W=canvas.width,H=canvas.height,padL=54,padR=20,padT=26,padB=36;
  const maxAbs = Math.max(...vals.map(v=>Math.abs(v)), 1);
  const n = err.length; const xAt=i=>padL+(W-padL-padR)*(n<=1?0:i/(n-1)); const yAt=v=>padT+(H-padT-padB)*(0.5-v/(2*maxAbs));
  ctx.strokeStyle='#29415f'; ctx.lineWidth=1; ctx.beginPath(); ctx.moveTo(padL,yAt(0)); ctx.lineTo(W-padR,yAt(0)); ctx.stroke();
  ctx.strokeStyle='#ef4444'; ctx.lineWidth=2; ctx.beginPath(); let started=false;
  err.forEach((v,i)=>{ if(!Number.isFinite(v)) return; const x=xAt(i), y=yAt(v); if(!started){ctx.moveTo(x,y);started=true}else ctx.lineTo(x,y); }); ctx.stroke();
  ctx.fillStyle='#ef4444'; ctx.font='13px system-ui'; ctx.fillText('Sai số = Dự báo - Thực tế', padL, 18);
  ctx.fillStyle='#9fb0c6'; ctx.fillText('+'+formatNum(maxAbs,1), 8, padT+5); ctx.fillText('-'+formatNum(maxAbs,1), 8, H-padB+5);
  const absVals = vals.map(Math.abs); const maxE = Math.max(...absVals); const idx = absVals.indexOf(maxE);
  $('maxErrorVal').textContent = formatNum(maxE,3);
  $('maxErrorTimeVal').textContent = times?.[idx] ? fmtTime(times[idx]) : '-';
}

function applyThresholdsToForecast(rows) {
  const thresholds = parseThresholdText();
  const alerts=[];
  for (const r of rows) {
    const st = r.station || '__ALL__';
    const th = thresholds.find(x => x.station === st) || thresholds.find(x => x.station === '__ALL__' || x.station.toUpperCase() === 'ALL');
    const p = parseNumber(r.forecast_p_mw);
    let status = 'Bình thường';
    if (th && Number.isFinite(p)) {
      if (Number.isFinite(th.danger) && p >= th.danger) { status = 'NGUY HIỂM'; alerts.push({time:r.time, station:st, p, level:'NGUY HIỂM', threshold:th.danger}); }
      else if (Number.isFinite(th.warn) && p >= th.warn) { status = 'CẢNH BÁO'; alerts.push({time:r.time, station:st, p, level:'CẢNH BÁO', threshold:th.warn}); }
    }
    r.nguong_canh_bao = th ? `${th.warn ?? ''}/${th.danger ?? ''}` : '';
    r.trang_thai_nguong = status;
  }
  renderThresholdAlerts(alerts);
}

function forecastForStation(station, stepsOverride=null) {
  const model = resolveModelForStation(station);
  if (!model) throw new Error('Chưa có model phù hợp cho ' + station);
  const sourceRows = state.rows.filter(r => station === '__ALL__' || r.station === station).sort((a,b)=>a.time-b.time);
  if (sourceRows.length < 5) throw new Error('Chưa có đủ dữ liệu gần nhất cho ' + station);
  const tempDefault = parseNumber($('futureTemp')?.value);
  const rainDefault = parseNumber($('futureRain')?.value) || 0;
  const interval = model.intervalMinutes || detectIntervalMinutes(sourceRows);
  const nPerDay = model.nPerDay || Math.round(1440/interval);
  const lastTemp = [...sourceRows].reverse().find(r => Number.isFinite(r.temp))?.temp;
  const tempUse = Number.isFinite(tempDefault) ? tempDefault : lastTemp;
  const series = sourceRows.map(r => ({...r}));
  const steps = Math.max(1, Math.floor(stepsOverride || parseNumber($('forecastSteps')?.value) || 24));
  const w = getHybridWeights();
  const bias = estimateRecentBias(model, sourceRows, Math.max(0, Math.floor(parseNumber($('biasWindow')?.value) || 0)));
  const forecast=[];
  let lastTime = series[series.length-1].time;
  for (let s=1; s<=steps; s++) {
    const t = new Date(lastTime.getTime() + interval*60000);
    const row = {time:t, p:NaN, station: station === '__ALL__' ? 'ALL' : station, temp:tempUse, rain:rainDefault, holiday:holidayByRules(t)?1:0, abnormal:0, outage:0, transfer:0};
    series.push(row);
    const idx = series.length - 1;
    const x = computeFeatureVector(series, idx, nPerDay, {temp: tempUse, rain: rainDefault});
    const gbdt = Math.max(0, predictModel(model, x));
    const similar = similarDayPrediction(series, idx, nPerDay);
    const week = lastWeekPrediction(series, idx, nPerDay);
    const trend = trendPrediction(series, idx);
    const comps = {gbdt, similar, week, trend};
    let pred=0, sw=0;
    for (const [k, val] of Object.entries(comps)) if (Number.isFinite(val) && Number.isFinite(w[k]) && w[k] > 0) { pred += val*w[k]; sw += w[k]; }
    pred = sw ? pred/sw : gbdt;
    let rawPredBeforeCalibration = Math.max(0, pred + bias);
    const cal = applyCalibrationLV85(rawPredBeforeCalibration, station, t);
    pred = cal.value;
    row.p = pred;
    forecast.push({step:s, time:fmtTime(t), station:row.station, forecast_p_mw:formatNum(pred,3), temp:Number.isFinite(tempUse)?formatNum(tempUse,1):'', rain:rainDefault, holiday:row.holiday, gbdt:formatNum(gbdt,3), similar_day:formatNum(similar,3), same_hour_last_week:formatNum(week,3), trend:formatNum(trend,3), bias:formatNum(bias,3), calibration_lv85:cal.applied?1:0, calibration_mw:formatNum(cal.delta,3), calibration_source:cal.source || '', forecast_before_calibration_mw:formatNum(rawPredBeforeCalibration,3)});
    lastTime = t;
  }
  return {forecast, series, interval, model, station, weights:w};
}

function forecastNext() {
  if (!state.model) throw new Error('Chưa có model.');
  const selected = $('stationSelect')?.value || '__ALL__';
  const station = selected === '__ALL__' && state.model.modelsByStation ? Object.keys(state.model.modelsByStation)[0] : selected;
  const out = forecastForStation(station);
  state.forecastRows = out.forecast;
  applyThresholdsToForecast(state.forecastRows);
  renderTable(state.forecastRows, ['step','time','station','forecast_p_mw','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong'], 1000);
  const steps = out.forecast.length;
  const actualHist = out.series.slice(-Math.min(steps*2, 240)).map((r,i,arr)=> i < arr.length-steps ? r.p : NaN);
  const predHist = out.series.slice(-Math.min(steps*2, 240)).map((r,i,arr)=> i >= arr.length-steps ? r.p : NaN);
  const times = out.series.slice(-Math.min(steps*2, 240)).map(r=>r.time);
  drawSeries(times, actualHist, predHist, 'Lịch sử gần nhất', 'Dự báo LV5');
  updateForecastMetrics(state.forecastRows);
  renderForecastExplain(out);
  $('exportForecastBtn').disabled = false;
  log(`Đã dự báo LV5 cho ${out.station}, ${steps} bước, mỗi bước ${out.interval} phút.`);
}

function forecastAllStations() {
  if (!state.model) throw new Error('Chưa có model.');
  const stations = state.model.modelsByStation ? Object.keys(state.model.modelsByStation) : [...new Set(state.rows.map(r=>r.station||'ALL'))];
  let all=[];
  const perStation = [];
  for (const st of stations) {
    try {
      const out = forecastForStation(st);
      all = all.concat(out.forecast);
      perStation.push(out);
    }
    catch(e) { log('Bỏ qua ' + st + ': ' + e.message); }
  }
  state.forecastRows = all;
  applyThresholdsToForecast(state.forecastRows);
  const cols = ['step','time','station','forecast_p_mw','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong'];
  renderTable(all, cols, 5000);
  updateForecastMetrics(all);
  drawForecastAllSummaryChart(all);
  renderForecastAllExplain(all, perStation);
  $('exportForecastBtn').disabled = false;
  log(`Đã dự báo tất cả trạm/lộ: ${all.length} dòng, ${perStation.length} trạm/lộ có kết quả.`);
}

function drawForecastAllSummaryChart(rows) {
  if (!rows || !rows.length) { drawSeries([], [], [], 'Tổng dự báo', ''); return; }
  const byTime = new Map();
  for (const r of rows) {
    const key = String(r.time || '');
    const p = parseNumber(r.forecast_p_mw);
    if (!key || !Number.isFinite(p)) continue;
    byTime.set(key, (byTime.get(key) || 0) + p);
  }
  const sorted = Array.from(byTime.entries()).sort((a,b) => String(a[0]).localeCompare(String(b[0])));
  const times = sorted.map(x => parseTime(x[0]) || x[0]);
  const total = sorted.map(x => x[1]);
  drawSeries(times, [], total, 'Tất cả trạm/lộ', 'Tổng P dự báo');
}

function renderForecastAllExplain(rows, perStation=[]) {
  const box = $('forecastExplainBox');
  if (!box) return;
  if (!rows || !rows.length) {
    box.innerHTML = '<span class="pill warn">Không có dòng dự báo nào được tạo</span>';
    return;
  }
  const stationCount = new Set(rows.map(r => r.station || '')).size;
  const pVals = rows.map(r => parseNumber(r.forecast_p_mw)).filter(Number.isFinite);
  const max = pVals.length ? Math.max(...pVals) : NaN;
  const maxRow = Number.isFinite(max) ? rows.find(r => parseNumber(r.forecast_p_mw) === max) : null;
  const totalByStep = new Map();
  for (const r of rows) {
    const step = String(r.step || '');
    const p = parseNumber(r.forecast_p_mw);
    if (!step || !Number.isFinite(p)) continue;
    totalByStep.set(step, (totalByStep.get(step) || 0) + p);
  }
  const maxTotalEntry = Array.from(totalByStep.entries()).sort((a,b)=>b[1]-a[1])[0];
  box.innerHTML = `<b>Kết quả dự báo tất cả trạm/lộ:</b>` +
    `<span class="pill ok">${rows.length} dòng forecast</span>` +
    `<span class="pill">${stationCount} trạm/lộ</span>` +
    (maxRow ? `<span class="pill">Pmax từng dòng ${formatNum(max,3)} MW tại ${escapeHtml(maxRow.time)} - ${escapeHtml(maxRow.station)}</span>` : '') +
    (maxTotalEntry ? `<span class="pill modeBadge">Tổng P lớn nhất bước ${escapeHtml(maxTotalEntry[0])}: ${formatNum(maxTotalEntry[1],3)} MW</span>` : '') +
    `<span class="pill">Bảng kết quả đang hiển thị ở khung Bảng dữ liệu bên dưới</span>`;
}

function renderForecastExplain(out) {
  const box = $('forecastExplainBox'); if (!box || !out?.forecast?.length) return;
  const first = out.forecast[0];
  box.innerHTML = `<b>Giải thích bước đầu:</b> <span class="pill">GBDT ${first.gbdt}</span><span class="pill">Similar Day ${first.similar_day}</span><span class="pill">Tuần trước ${first.same_hour_last_week}</span><span class="pill">Xu hướng ${first.trend}</span><span class="pill">Bù sai số ${first.bias}</span><span class="pill ok">Dự báo cuối ${first.forecast_p_mw} MW</span>`;
}

function parseThresholdText() {
  const text = $('thresholdText')?.value || '';
  const rows=[];
  for (const line of text.split(/\r?\n/)) {
    const s = line.trim(); if (!s || s.startsWith('#')) continue;
    const parts = s.split(/[;,\t|]/).map(x=>x.trim());
    if (parts.length < 2) continue;
    rows.push({station: parts[0], warn: parseNumber(parts[1]), danger: parseNumber(parts[2])});
  }
  state.thresholds = rows;
  return rows;
}

function renderThresholdAlerts(alerts=[]) {
  const box = $('thresholdAlerts'); if (!box) return;
  if (!alerts.length) { box.innerHTML = '<span class="pill ok">Không có dự báo vượt ngưỡng</span>'; return; }
  const max = alerts.slice(0,20).map(a => `<div class="pill ${a.level==='NGUY HIỂM'?'bad':'warn'}">${escapeHtml(a.station)} ${escapeHtml(a.time)}: ${formatNum(a.p,2)} MW ≥ ${formatNum(a.threshold,2)} (${a.level})</div>`).join('');
  box.innerHTML = max + (alerts.length>20 ? `<div class="note">... còn ${alerts.length-20} cảnh báo khác</div>` : '');
}

function exportThresholds() {
  const rows = parseThresholdText();
  if (!rows.length) { log('Chưa có ngưỡng để xuất.'); return; }
  saveTextFile('thresholds.csv', toCSV(rows, ['station','warn','danger']), 'text/csv;charset=utf-8');
}

function updateValidationErrorViews() {
  if (!state.trainingResult) return;
  const {timesVal, yval, predVal} = state.trainingResult;
  drawErrorChart(timesVal, yval, predVal);
  const errRows = yval.map((a,i)=>({time:fmtTime(timesVal[i]), actual:formatNum(a,3), forecast:formatNum(predVal[i],3), error:formatNum(predVal[i]-a,3), abs_error:formatNum(Math.abs(predVal[i]-a),3)}));
  if (errRows.length) {
    const abs = errRows.map(r=>parseNumber(r.abs_error)); const max=Math.max(...abs); const idx=abs.indexOf(max);
    $('maxErrorVal').textContent = formatNum(max,3); $('maxErrorTimeVal').textContent = errRows[idx]?.time || '-';
  }
}

function renderModelInfo() {
  if (!state.model) { $('modelInfo').innerHTML = ''; return; }
  const m = state.model;
  if (m.modelsByStation) {
    const n = Object.keys(m.modelsByStation).length;
    $('modelInfo').innerHTML = `<span class="pill">Model bundle LV5</span><span class="pill">${n} trạm/lộ</span><span class="pill">Created: ${escapeHtml(m.createdAt||'')}</span>`;
    $('forecastBtn').disabled = false; $('forecastAllBtn').disabled = false;
    return;
  }
  $('modelInfo').innerHTML = `
    <span class="pill">Model: ${escapeHtml(m.type || '')}</span>
    <span class="pill">Số cây: ${m.trees?.length || 0}</span>
    <span class="pill">Interval: ${m.intervalMinutes} phút</span>
    <span class="pill">Station: ${escapeHtml(m.station || 'ALL')}</span>
    <span class="pill">MAPE: ${formatNum(m.metrics?.validation?.mape,2)}%</span>`;
}

async function trainAllStations() {
  if (!state.rows.length) throw new Error('Chưa có dữ liệu.');
  const original = $('stationSelect').value;
  const stations = [...new Set(state.rows.map(r => r.station || 'ALL'))].filter(Boolean);
  if (!stations.length) throw new Error('Không có trạm/lộ để huấn luyện.');
  const modelsByStation = {};
  for (const st of stations) {
    $('stationSelect').value = st;
    try {
      await trainGBDT();
      if (state.model && state.model.trees) { modelsByStation[st] = state.model; log(`Đã huấn luyện xong model riêng cho ${st}.`); }
    } catch(e) { log(`Không huấn luyện được ${st}: ${e.message}`); }
  }
  $('stationSelect').value = original;
  if (!Object.keys(modelsByStation).length) throw new Error('Không tạo được model riêng nào.');
  state.model = {type:'GBDT_STATION_BUNDLE_JS_OFFLINE_LV5', createdAt:new Date().toISOString(), modelsByStation, colMap:state.colMap, note:'Bundle model theo từng trạm/lộ, dùng offline trong SCADA.'};
  $('exportModelBtn').disabled = false; $('forecastBtn').disabled = false; $('forecastAllBtn').disabled = false; renderModelInfo();
  log(`Hoàn tất huấn luyện bundle LV5: ${Object.keys(modelsByStation).length} model riêng.`);
}

function exportForecast() {
  if (!state.forecastRows.length) return;
  const headers = ['step','time','station','forecast_p_mw','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong'];
  saveTextFile('forecast.csv', toCSV(state.forecastRows, headers), 'text/csv;charset=utf-8');
  log('Đã xuất forecast.csv LV5.');
}

function exportModel() {
  if (!state.model) return;
  const payload = {...state.model, lv5Config: collectFullConfig()};
  saveTextFile('model_gbdt.json', JSON.stringify(payload, null, 2), 'application/json');
  log('Đã xuất model_gbdt.json LV5 kèm cấu hình. Copy file này vào mạng SCADA để dùng offline.');
}

function collectFullConfig() {
  return {
    version:'LV5', savedAt:new Date().toISOString(), name:$('configName')?.value || 'SCADA_LOAD_FORECAST_LV5',
    appMode:$('appMode')?.value || 'external', colMap:readMap(), sourceFileName:state.sourceFileName,
    holiday:{sat:$('holidaySat')?.checked, sun:$('holidaySun')?.checked, fixed:$('holidayFixed')?.checked, extra:$('holidayExtraDates')?.value || '', clearNonMatch:$('holidayClearNonMatch')?.checked},
    gbdt:{nTrees:$('nTrees')?.value, maxDepth:$('maxDepth')?.value, learningRate:$('learningRate')?.value, minLeaf:$('minLeaf')?.value, maxBins:$('maxBins')?.value, valPercent:$('valPercent')?.value},
    hybrid:{wGbdt:$('wGbdt')?.value, wSimilar:$('wSimilar')?.value, wWeek:$('wWeek')?.value, wTrend:$('wTrend')?.value, biasWindow:$('biasWindow')?.value, forecastBlend:$('forecastBlend')?.value},
    quality:{spikePercent:$('spikePercent')?.value, minValidP:$('minValidP')?.value, expectedInterval:$('expectedInterval')?.value},
    thresholdsText:$('thresholdText')?.value || ''
  };
}

function applyFullConfig(cfg={}) {
  if (cfg.name && $('configName')) $('configName').value = cfg.name;
  if (cfg.appMode && $('appMode')) $('appMode').value = cfg.appMode;
  const setVal = (id,v) => { if ($(id) && v !== undefined && v !== null) $(id).value = v; };
  const setChk = (id,v) => { if ($(id) && v !== undefined && v !== null) $(id).checked = !!v; };
  if (cfg.holiday) { setChk('holidaySat',cfg.holiday.sat); setChk('holidaySun',cfg.holiday.sun); setChk('holidayFixed',cfg.holiday.fixed); setVal('holidayExtraDates',cfg.holiday.extra); setChk('holidayClearNonMatch',cfg.holiday.clearNonMatch); }
  if (cfg.gbdt) Object.entries(cfg.gbdt).forEach(([k,v]) => setVal(k,v));
  if (cfg.hybrid) Object.entries(cfg.hybrid).forEach(([k,v]) => setVal(k,v));
  if (cfg.quality) { setVal('spikePercent',cfg.quality.spikePercent); setVal('minValidP',cfg.quality.minValidP); setVal('expectedInterval',cfg.quality.expectedInterval); }
  setVal('thresholdText', cfg.thresholdsText);
  if (cfg.colMap) {
    const mapId = {time:'colTime',p:'colP',station:'colStation',temp:'colTemp',rain:'colRain',holiday:'colHoliday',abnormal:'colAbnormal',outage:'colOutage',transfer:'colTransfer'};
    Object.entries(cfg.colMap).forEach(([k,v]) => { const id=mapId[k]; if ($(id) && [...$(id).options].some(o=>o.value===v)) $(id).value=v; });
  }
  applyAppMode(); parseThresholdText();
}

function saveSettingsToLocal() {
  try { localStorage.setItem('SCADA_LOAD_FORECAST_LV5_SETTINGS', JSON.stringify(collectFullConfig())); } catch(_) {}
}

function loadSettingsFromLocal() {
  try { return JSON.parse(localStorage.getItem('SCADA_LOAD_FORECAST_LV5_SETTINGS') || localStorage.getItem('SCADA_LOAD_FORECAST_LV4_SETTINGS') || '{}'); } catch(_) { return {}; }
}

async function saveFullConfigOffline() { await idbSet('fullConfigLV5', collectFullConfig()); saveSettingsToLocal(); log('Đã lưu cấu hình LV5 vào trình duyệt offline.'); }
async function loadFullConfigOffline() { const cfg = await idbGet('fullConfigLV5') || loadSettingsFromLocal(); applyFullConfig(cfg); log('Đã nạp cấu hình LV5 offline.'); }
function exportFullConfig() { saveTextFile('config_lv5.json', JSON.stringify(collectFullConfig(), null, 2), 'application/json'); }
function importFullConfigFile(file) { loadTextFile(file, text => { try { applyFullConfig(JSON.parse(text)); saveSettingsToLocal(); log('Đã nạp config_lv5.json.'); } catch(e) { log('Lỗi nạp cấu hình: '+e.message); } }); }

function applyAppMode() {
  const mode = $('appMode')?.value || 'external'; state.appMode = mode;
  document.querySelectorAll('.modeExternal').forEach(el => el.classList.toggle('hideByMode', mode === 'scada'));
  const txt = mode === 'scada' ? 'SCADA: chỉ dự báo/cảnh báo' : 'Mạng ngoài: hiệu chỉnh/huấn luyện';
  $('versionInfo').innerHTML = `<span class="pill modeBadge">LV5.3</span><span class="pill">${txt}</span>`;
}

async function forceUpdateApp() {
  log('Đang ép cập nhật: bỏ cache PWA/Service Worker cũ...');
  try {
    if ('serviceWorker' in navigator) {
      const regs = await navigator.serviceWorker.getRegistrations();
      for (const r of regs) await r.unregister();
    }
    if ('caches' in window) {
      const keys = await caches.keys();
      for (const k of keys) if (/scada|load|forecast|lv/i.test(k)) await caches.delete(k);
    }
    log('Đã xóa cache ứng dụng. Trang sẽ tải lại.');
    setTimeout(()=>location.reload(true), 600);
  } catch(e) { log('Lỗi ép cập nhật: ' + e.message); }
}

function lv5BindEvents() {
  const on = (id, ev, fn) => { if ($(id)) $(id).addEventListener(ev, fn); };
  on('appMode','change', () => { applyAppMode(); saveSettingsToLocal(); });
  on('forceUpdateBtn','click', forceUpdateApp);
  on('saveFullConfigBtn','click', () => saveFullConfigOffline().catch(e=>log('Lỗi lưu cấu hình: '+e.message)));
  on('loadFullConfigBtn','click', () => loadFullConfigOffline().catch(e=>log('Lỗi nạp cấu hình: '+e.message)));
  on('exportConfigBtn','click', exportFullConfig);
  on('importConfigBtn','click', () => $('configFile')?.click());
  on('configFile','change', e => { const f=e.target.files[0]; if(f) importFullConfigFile(f); });
  on('runQualityBtn','click', () => { try { runQualityCheck(); } catch(e) { log('Lỗi kiểm tra dữ liệu: '+e.message); } });
  on('selectQualityRowsBtn','click', selectQualityRows);
  on('markQualityAbnormalBtn','click', () => { try { markQualityAbnormal(); } catch(e) { log('Lỗi đánh dấu bất thường: '+e.message); } });
  on('exportQualityBtn','click', exportQualityReport);
  on('addProcessColumnsBtn','click', () => { addProcessColumns(); renderEditorTable(); });
  on('interpolateBtn','click', () => { try { interpolateMissingTimestamps(); } catch(e) { log('Lỗi nội suy mốc thiếu: '+e.message); } });
  on('fillInvalidPBtn','click', () => { try { fillInvalidP(); } catch(e) { log('Lỗi nội suy P: '+e.message); } });
  on('handleLowPEventsBtn','click', () => { try { markLowPOperationalEventsForTraining(); } catch(e) { log('Lỗi xử lý P thấp theo cờ vận hành: '+e.message); } });
  on('trainAllStationsBtn','click', () => trainAllStations().catch(e => { $('trainBtn').disabled=false; log('Lỗi huấn luyện bundle: ' + e.message); }));
  on('forecastAllBtn','click', () => { try { forecastAllStations(); } catch(e) { log('Lỗi dự báo tất cả: '+e.message); } });
  on('applyThresholdBtn','click', () => { const rows=parseThresholdText(); renderThresholdAlerts([]); saveSettingsToLocal(); log(`Đã áp dụng ${rows.length} dòng ngưỡng cảnh báo.`); });
  on('exportThresholdBtn','click', exportThresholds);
  ['wGbdt','wSimilar','wWeek','wTrend','biasWindow','forecastBlend','thresholdText','expectedInterval','spikePercent','minValidP'].forEach(id => on(id, id==='thresholdText'?'input':'change', saveSettingsToLocal));
  applyFullConfig(loadSettingsFromLocal());
  applyAppMode();
}

setTimeout(lv5BindEvents, 0);
// ====================== END LV5 EXTENSIONS ======================

$('csvFile').addEventListener('change', e => {
  const file = e.target.files[0]; if (!file) return;
  state.workbook = null; state.currentSheet = '';
  if (isSpreadsheetFile(file.name)) {
    loadSpreadsheetFile(file).catch(err => log('Lỗi nạp Excel: ' + err.message));
    return;
  }
  showSheetSelector(false);
  loadTextFile(file, text => {
    try {
      const parsed = parseDataFileText(text, file.name);
      applyParsedDataset(parsed, file.name, 'file');
      log(`Đã nạp file: ${file.name}, ${state.rawRows.length} dòng thô, ${state.headers.length} cột, định dạng ${file.name.toLowerCase().endsWith('.json')?'JSON':'delimiter '+JSON.stringify(state.delimiter)}.`);
    } catch(err) {
      log('Lỗi nạp file dữ liệu: ' + err.message);
    }
  });
});

if ($('sheetSelect')) $('sheetSelect').addEventListener('change', e => { try { loadWorkbookSheet(e.target.value); } catch(err) { log('Lỗi đổi sheet: ' + err.message); } });


function validateOperationalModelPackage(m) {
  const errors = [];
  if (!m || typeof m !== 'object') errors.push('File JSON không phải object model.');
  if (m?.scadaModelPackage !== 'SCADA_LOAD_FORECAST_OPERATIONAL_MODEL') errors.push('Thiếu chữ ký scadaModelPackage của model vận hành.');
  if (m?.exportSource !== 'MUC_8_EXPORT_MODEL') errors.push('File không được xuất từ Mục 8.');
  if (m?.allowedImportSection !== '10') errors.push('File không khai báo được phép nạp ở Mục 10.');
  if (m?.modelKind !== 'OPERATIONAL_FORECAST_MODEL') errors.push('Không phải gói model dự báo vận hành.');
  const hasSingle = !!(m?.trees && m?.featureNames);
  const hasBundle = !!m?.modelsByStation;
  if (!hasSingle && !hasBundle) errors.push('Không tìm thấy GBDT hoặc bundle model theo trạm/lộ.');
  return {ok: errors.length === 0, errors};
}

function renderModelImportGuard(status, messages=[]) {
  const box = $('modelImportGuard');
  if (!box) return;
  if (status === 'ok') box.innerHTML = `<span class="pill ok">Model hợp lệ từ Mục 8</span>${messages.map(x=>`<span class="pill">${escapeHtml(x)}</span>`).join('')}`;
  else if (status === 'bad') box.innerHTML = `<span class="pill bad">Từ chối model</span>${messages.map(x=>`<span class="pill bad">${escapeHtml(x)}</span>`).join('')}`;
  else box.innerHTML = '<span class="pill">Chưa nạp model vận hành</span>';
}

$('modelFile').addEventListener('change', e => {
  const file = e.target.files[0]; if (!file) return;
  loadTextFile(file, text => {
    try {
      const m = JSON.parse(text);
      const guard = validateOperationalModelPackage(m);
      if (!guard.ok) {
        renderModelImportGuard('bad', guard.errors.slice(0, 4));
        $('forecastBtn').disabled = true;
        if ($('forecastAllBtn')) $('forecastAllBtn').disabled = true;
        e.target.value = '';
        throw new Error('Model bị từ chối. Chỉ được nạp file model vận hành xuất từ Mục 8. ' + guard.errors.join(' '));
      }
      state.model = m;
      updateMetrics(m.metrics?.validation || {}, m.intervalMinutes);
      renderModelInfo();
      renderModelImportGuard('ok', [file.name, m.appVersion || '', m.exportedAt || ''].filter(Boolean));
      $('forecastBtn').disabled = false;
      if ($('forecastAllBtn')) $('forecastAllBtn').disabled = false;
      $('exportModelBtn').disabled = false;
      log(`Đã nạp model vận hành hợp lệ từ Mục 8: ${file.name}, ${m.modelsByStation ? Object.keys(m.modelsByStation).length + ' model trạm/lộ' : (m.trees?.length||0) + ' cây'}.`);
    } catch(err) { log('Lỗi nạp model: ' + err.message); }
  });
});

$('applyMapBtn').addEventListener('click', () => { try { normalizeRows(); applyDataInfo(); previewData(); renderEditorTable(); saveSettingsToLocal(); log('Đã áp dụng và lưu tạm ánh xạ cột.'); } catch(e) { log('Lỗi: ' + e.message); }});
$('previewBtn').addEventListener('click', () => { try { if(!state.rows.length) normalizeRows(); previewData(); renderEditorTable(); } catch(e) { log('Lỗi: ' + e.message); }});
$('saveMapBtn').addEventListener('click', () => { readMap(); saveSettingsToLocal(); log('Đã lưu cấu hình ánh xạ cột vào trình duyệt.'); });
$('trainBtn').addEventListener('click', () => trainGBDT().catch(e => { $('trainBtn').disabled=false; log('Lỗi huấn luyện: ' + e.message); }));
$('exportModelBtn').addEventListener('click', exportModel);
$('forecastBtn').addEventListener('click', () => { try { forecastNext(); } catch(e) { log('Lỗi dự báo: ' + e.message); }});
$('exportForecastBtn').addEventListener('click', exportForecast);
$('clearBtn').addEventListener('click', () => { state.headers=[]; state.rawRows=[]; state.rows=[]; state.model=null; state.forecastRows=[]; state.sourceFileName=''; state.workbook=null; state.currentSheet=''; showSheetSelector(false); state.editor.selected.clear(); state.editor.page=1; state.editor.dirty=false; fillColumnSelects([]); renderTable([]); renderEditorTable(); updateMetrics({}, null); renderModelInfo(); applyDataInfo(); log('Đã xóa dữ liệu trong phiên hiện tại.'); });
$('loadSampleBtn').addEventListener('click', () => {
  const parsed = parseCSV(buildSampleCSV());
  showSheetSelector(false); state.workbook=null; state.currentSheet=''; state.headers = parsed.headers; state.rawRows = parsed.data; state.rows=[]; state.sampleLoaded=true; state.sourceFileName='sample_load_data.csv'; state.delimiter=','; state.editor.selected.clear(); state.editor.page=1; state.editor.dirty=false;
  fillColumnSelects(state.headers);
  normalizeRows(); applyDataInfo(); previewData(); renderEditorTable();
  log('Đã nạp dữ liệu mẫu 75 ngày theo giờ để kiểm tra nhanh.');
});

$('editorSearch').addEventListener('input', e => { state.editor.query = e.target.value; state.editor.page = 1; renderEditorTable(); });
$('editorFilter').addEventListener('change', e => { state.editor.filter = e.target.value; state.editor.page = 1; renderEditorTable(); });
['dateFilterMode','dateFilterSingle','dateFilterMulti','dateFilterFrom','dateFilterTo'].forEach(id => {
  if ($(id)) $(id).addEventListener(id === 'dateFilterMulti' ? 'input' : 'change', () => { readEditorDateFilterFromUI(); state.editor.page = 1; $('editorPage').value = 1; renderEditorTable(); });
});
$('resetDateFilterBtn')?.addEventListener('click', resetDateFilter);
$('selectVisibleRowsBtn')?.addEventListener('click', selectVisibleEditorRows);
$('selectFilteredRowsBtn')?.addEventListener('click', selectFilteredEditorRows);
$('clearSelectedRowsBtn')?.addEventListener('click', clearEditorSelection);
$('vietnamizeHeadersBtn')?.addEventListener('click', vietnamizeHeaders);
$('quickFillSelectedBtn')?.addEventListener('click', () => { try { applyQuickFillToSelected(); } catch(e) { log('Lỗi điền nhanh: ' + e.message); } });
$('autoHolidaySelectedBtn')?.addEventListener('click', () => { try { autoHolidayForIndices(selectedEditorIndicesOrLog()); } catch(e) { log('Lỗi tự nhận dạng ngày nghỉ/lễ: ' + e.message); } });
$('autoHolidayFilteredBtn')?.addEventListener('click', () => { try { autoHolidayForIndices(getEditorFilteredIndices()); } catch(e) { log('Lỗi tự nhận dạng ngày nghỉ/lễ: ' + e.message); } });
$('editorPageSize').addEventListener('change', e => { state.editor.pageSize = parseNumber(e.target.value) || 100; state.editor.page = 1; renderEditorTable(); });
$('editorPage').addEventListener('change', e => { state.editor.page = parseNumber(e.target.value) || 1; renderEditorTable(); });
$('editorPrevBtn').addEventListener('click', () => { state.editor.page = Math.max(1, state.editor.page - 1); $('editorPage').value = state.editor.page; renderEditorTable(); });
$('editorNextBtn').addEventListener('click', () => { state.editor.page += 1; $('editorPage').value = state.editor.page; renderEditorTable(); });
$('addRowBtn').addEventListener('click', addEditorRow);
$('deleteRowsBtn').addEventListener('click', deleteSelectedEditorRows);
$('applyEditsBtn').addEventListener('click', () => { try { applyEditorEdits(); } catch(e) { log('Lỗi lưu thay đổi: ' + e.message); } });
$('saveOfflineBtn').addEventListener('click', () => saveDatasetOffline().catch(e => log('Lỗi lưu offline: ' + e.message)));
$('loadSavedDataBtn').addEventListener('click', () => loadDatasetOffline().catch(e => log('Lỗi nạp offline: ' + e.message)));
$('exportEditedCsvBtn').addEventListener('click', () => { try { exportEditedCSV(); } catch(e) { log('Lỗi xuất CSV: ' + e.message); } });
$('exportEditedJsonBtn').addEventListener('click', () => { try { exportEditedJSON(); } catch(e) { log('Lỗi xuất JSON: ' + e.message); } });
$('editorBox').addEventListener('input', e => {
  const td = e.target.closest('td[contenteditable="true"]');
  if (!td) return;
  const row = Number(td.dataset.row);
  const col = td.dataset.col;
  if (state.rawRows[row] && state.headers.includes(col)) { state.rawRows[row][col] = td.textContent.trim(); markEditorDirty(true); }
});
$('editorBox').addEventListener('change', e => {
  const cb = e.target.closest('.rowSelect');
  if (!cb) return;
  const row = Number(cb.dataset.row);
  if (cb.checked) state.editor.selected.add(row); else state.editor.selected.delete(row);
  renderEditorStatus();
  const tr = cb.closest('tr'); if (tr) tr.classList.toggle('selected', cb.checked);
});

fillColumnSelects([]);
renderEditorTable();
log('Sẵn sàng LV6. Bước 1: nạp Excel/CSV, ánh xạ cột, kiểm tra/tách chỉ danh Đơn vị/Trạm/Lộ/Nối vòng rồi hiệu chỉnh, huấn luyện hoặc nạp model.');


// ======================== LV5.4 INTERPOLATION OVERRIDE ========================
// Mục tiêu LV5.4:
// 1) P = 0/P thấp không mặc định là lỗi; xét theo cờ vận hành.
// 2) Nếu người dùng chọn "Nội suy cả P thấp dù có cờ vận hành" thì ÉP nội suy dòng đã chọn/đang lọc.
// 3) Nếu giữ sự kiện vận hành thì đánh dấu bo_khoi_huan_luyen=1 để model không học sự kiện như quy luật bình thường.
// 4) Khi nội suy P, lưu P cũ vào p_goc để truy vết.

function addProcessColumns() {
  let added = [];
  for (const name of ['p_goc','du_lieu_noi_suy','ghi_chu_xu_ly','bo_khoi_huan_luyen']) {
    if (!state.headers.includes(name)) {
      state.headers.push(name);
      state.rawRows.forEach(r => { r[name] = ''; });
      added.push(name);
    }
  }
  fillColumnSelects(state.headers);
  if (typeof refreshQuickCustomColumns === 'function') refreshQuickCustomColumns();
  if (added.length) log('Đã thêm cột xử lý LV5.4: ' + added.join(', '));
  return added;
}

function getOperationReason(flags) {
  const reasons = [];
  if (flags.outage) reasons.push('cat dien/su co');
  if (flags.transfer) reasons.push('chuyen tai');
  if (flags.abnormal) reasons.push('bat thuong');
  if (flags.excludeTrain) reasons.push('da bo khoi huan luyen');
  return reasons.join(' + ') || '';
}

function shouldUseAsReferenceP(item, m, threshold) {
  if (!item || !Number.isFinite(item.p)) return false;
  if (Number.isFinite(threshold) && item.p <= threshold) return false;
  const flags = rawOperationalFlags(item.raw, m);
  // Không dùng các dòng sự kiện vận hành làm điểm neo nội suy, tránh kéo sai xu hướng.
  if (flags.abnormal || flags.outage || flags.transfer || flags.excludeTrain) return false;
  return true;
}

function findLeftReference(arr, startIdx, m, threshold) {
  for (let i=startIdx; i>=0; i--) if (shouldUseAsReferenceP(arr[i], m, threshold)) return i;
  return -1;
}

function findRightReference(arr, startIdx, m, threshold) {
  for (let i=startIdx; i<arr.length; i++) if (shouldUseAsReferenceP(arr[i], m, threshold)) return i;
  return -1;
}

function inferCleanPValue(method, arr, idx, m, threshold, interval) {
  const item = arr[idx];
  const leftIdx = findLeftReference(arr, idx - 1, m, threshold);
  const rightIdx = findRightReference(arr, idx + 1, m, threshold);

  if (method === 'day' || method === 'week') {
    const backMs = (method === 'day' ? 24 : 24*7) * 3600*1000;
    const want = item.time.getTime() - backMs;
    const found = arr.find(x => Math.abs(x.time.getTime() - want) <= interval*60000*0.51 && shouldUseAsReferenceP(x, m, threshold));
    if (found) return found.p;
  }
  if (method === 'prev' && leftIdx >= 0) return arr[leftIdx].p;

  if (leftIdx >= 0 && rightIdx >= 0) {
    const left = arr[leftIdx], right = arr[rightIdx];
    const denom = right.time - left.time;
    if (denom > 0) {
      const f = (item.time - left.time) / denom;
      return left.p + (right.p - left.p) * f;
    }
  }
  if (leftIdx >= 0) return arr[leftIdx].p;
  if (rightIdx >= 0) return arr[rightIdx].p;
  return NaN;
}

function lowPActionForRow(raw, m, p, threshold) {
  const original = m?.p ? String(raw?.[m.p] ?? '').trim() : '';
  const flags = rawOperationalFlags(raw, m);
  const hasOp = !!(flags.abnormal || flags.outage || flags.transfer || flags.excludeTrain);
  const mode = $('lowPHandlingMode')?.value || 'auto';
  if (!original || !Number.isFinite(p)) return {action:'interpolate', reason:'P trong/sai dinh dang', hasOp, flags};
  if (Number.isFinite(threshold) && p <= threshold) {
    if (mode === 'fill_all') return {action:'interpolate', reason: hasOp ? ('P thap co co van hanh: ' + getOperationReason(flags)) : 'P thap khong co co van hanh', hasOp, flags};
    if (hasOp) return {action:'exclude', reason:'P thap co co van hanh: ' + getOperationReason(flags), hasOp, flags};
    return {action:'interpolate', reason:'P thap khong co co van hanh', hasOp, flags};
  }
  return {action:'keep', reason:'P hop le', hasOp, flags};
}

function fillInvalidP(options={}) {
  const m = readMap();
  if (!m.time || !m.p) throw new Error('Cần ánh xạ thời gian/P.');
  addProcessColumns();
  const targetSet = getInterpolationTargetIndexSet(options);
  if (!targetSet.size) {
    if (!options.silent) {
      $('interpolationInfo').innerHTML = '<span class="pill warn">Chưa có dòng nào trong phạm vi xử lý. Hãy chọn dòng hoặc lọc ngày trước.</span>';
      log('Chưa có dòng nào để nội suy P.');
    }
    return 0;
  }
  const stationFilter = $('stationSelect')?.value || '__ALL__';
  const scope = $('interpScope')?.value || 'current';
  const method = $('interpMethod')?.value || 'linear';
  const threshold = getInvalidPThreshold();
  const groups = groupRowsByStationRaw();
  let fixed=0, considered=0, excluded=0, kept=0, noRef=0;
  for (const [station, arr] of groups) {
    if (scope === 'current' && stationFilter !== '__ALL__' && station !== stationFilter) continue;
    const interval = detectIntervalMinutes(arr.map(x=>({time:x.time,p:x.p}))) || 60;
    for (let i=0; i<arr.length; i++) {
      const item = arr[i];
      if (!targetSet.has(item.rawIndex)) continue;
      if (!editorDateFilterPassTime(item.time)) continue;
      considered++;
      const action = lowPActionForRow(item.raw, m, item.p, threshold);
      if (action.action === 'keep') { kept++; continue; }

      if (action.action === 'exclude') {
        item.raw['bo_khoi_huan_luyen'] = '1';
        item.raw['ghi_chu_xu_ly'] = `${item.raw['ghi_chu_xu_ly'] || ''} LV5.4 giu P goc, bo khoi huan luyen (${action.reason})`.trim();
        excluded++;
        continue;
      }

      const val = inferCleanPValue(method, arr, i, m, threshold, interval);
      if (!Number.isFinite(val)) { noRef++; continue; }
      if (!item.raw['p_goc']) item.raw['p_goc'] = String(item.raw[m.p] ?? '');
      item.raw[m.p] = formatNum(val, 3);
      item.raw['du_lieu_noi_suy'] = '1';
      // Nếu dòng có cờ vận hành, vẫn bỏ khỏi huấn luyện, nhưng P sạch giúp các lag/trend xung quanh không bị kéo về 0.
      if (action.hasOp) item.raw['bo_khoi_huan_luyen'] = '1';
      item.raw['ghi_chu_xu_ly'] = `${item.raw['ghi_chu_xu_ly'] || ''} LV5.4 noi suy P tu ${method}; ${action.reason}`.trim();
      fixed++;
    }
  }
  normalizeRows(); renderEditorTable(); previewData(); markEditorDirty(true);
  if (!options.silent) {
    $('interpolationInfo').innerHTML = [
      `<span class="pill ok">Đã nội suy ${fixed} giá trị P</span>`,
      `<span class="pill ${excluded?'warn':''}">Giữ P gốc & bỏ huấn luyện ${excluded} dòng có cờ vận hành</span>`,
      `<span class="pill">Đã xét ${considered} dòng</span>`,
      `<span class="pill ${noRef?'warn':''}">Không đủ điểm neo ${noRef} dòng</span>`,
      `<span class="pill">Giữ nguyên ${kept} dòng P hợp lệ</span>`
    ].join('');
    log(`LV5.4 nội suy P: sửa ${fixed}, bỏ huấn luyện ${excluded}, không đủ điểm neo ${noRef}, xét ${considered}.`);
  }
  return fixed;
}

function markLowPOperationalEventsForTraining() {
  const m = readMap();
  if (!m.p) throw new Error('Cần ánh xạ cột Công suất P.');
  addProcessColumns();
  const targetSet = getInterpolationTargetIndexSet({});
  const threshold = getInvalidPThreshold();
  const mode = $('lowPHandlingMode')?.value || 'auto';
  // Nếu chọn fill_all, nút này cũng sẽ nội suy luôn, đúng như tên chế độ.
  if (mode === 'fill_all') {
    const fixed = fillInvalidP({silent:true});
    $('interpolationInfo').innerHTML = `<span class="pill ok">Chế độ ép nội suy: đã nội suy ${fixed} dòng P=0/P thấp trong phạm vi chọn/lọc</span><span class="pill">Giá trị cũ đã lưu vào p_goc</span>`;
    log(`LV5.4 áp dụng quy tắc P thấp ở chế độ ép nội suy: ${fixed} dòng.`);
    return;
  }
  let marked=0, lowNoFlag=0, checked=0, nonLow=0;
  for (const idx of targetSet) {
    const raw = state.rawRows[idx];
    if (!raw) continue;
    const p = parseNumber(raw[m.p]);
    if (!isLowPValue(p, threshold)) { nonLow++; continue; }
    checked++;
    const flags = rawOperationalFlags(raw, m);
    if (flags.outage || flags.transfer || flags.abnormal) {
      raw['bo_khoi_huan_luyen'] = '1';
      const reason = getOperationReason(flags);
      raw['ghi_chu_xu_ly'] = `${raw['ghi_chu_xu_ly'] || ''} LV5.4 giu P thap do ${reason}, bo khoi huan luyen`.trim();
      marked++;
    } else {
      // Không có cờ vận hành: nên nội suy ngay để tránh model học P=0 giả.
      lowNoFlag++;
    }
  }
  let fixedNoFlag = 0;
  if (lowNoFlag) fixedNoFlag = fillInvalidP({silent:true});
  normalizeRows(); renderEditorTable(); previewData(); markEditorDirty(true);
  $('interpolationInfo').innerHTML = `<span class="pill ok">Đã bỏ khỏi huấn luyện ${marked} dòng P thấp có cờ vận hành</span><span class="pill ${fixedNoFlag?'ok':'warn'}">Đã nội suy ${fixedNoFlag} dòng P thấp không có cờ/hoặc theo quy tắc</span><span class="pill">Đã xét ${checked} dòng P thấp</span><span class="pill">Bỏ qua ${nonLow} dòng P hợp lệ</span>`;
  log(`LV5.4 xử lý P thấp: bỏ huấn luyện ${marked}, nội suy ${fixedNoFlag}, xét ${checked}.`);
}

function interpolateMissingTimestamps() {
  if (!state.headers.length) throw new Error('Chưa có dữ liệu.');
  const m = readMap();
  if (!m.time || !m.p) throw new Error('Cần ánh xạ cột thời gian và P.');
  const stationFilter = $('stationSelect')?.value || '__ALL__';
  const scope = $('interpScope')?.value || 'current';
  const method = $('interpMethod')?.value || 'linear';
  const maxGap = Math.max(1, Math.floor(parseNumber($('interpMaxGap')?.value) || 12));
  const syntheticCol = 'du_lieu_noi_suy', noteCol='ghi_chu_xu_ly';
  addProcessColumns();

  // Luôn xử lý P trống/lỗi/P=0/P thấp trong phạm vi chọn/lọc trước khi bổ sung mốc thời gian.
  // Nhờ vậy người dùng chọn dòng P=0 rồi bấm nút màu cam cũng sẽ có kết quả rõ ràng.
  const fixedBadP = fillInvalidP({silent:true});

  const groups = groupRowsByStationRaw();
  let added=0, skipped=0, skippedByDate=0;
  for (const [station, arr] of groups) {
    if (scope === 'current' && stationFilter !== '__ALL__' && station !== stationFilter) continue;
    const expected = $('expectedInterval')?.value !== 'auto' ? parseNumber($('expectedInterval')?.value) : detectIntervalMinutes(arr.map(x => ({time:x.time,p:x.p})));
    const interval = expected || 60;
    for (let i=1; i<arr.length; i++) {
      const prev = arr[i-1], next = arr[i];
      const gap = (next.time - prev.time)/60000;
      const miss = Math.round(gap/interval) - 1;
      if (miss <= 0) continue;
      if (miss > maxGap) { skipped += miss; continue; }
      for (let k=1; k<=miss; k++) {
        const t = new Date(prev.time.getTime() + k*interval*60000);
        if (!editorDateFilterPassTime(t)) { skippedByDate++; continue; }
        const row = {};
        state.headers.forEach(h => row[h] = '');
        row[m.time] = fmtTime(t);
        row[m.p] = formatNum(inferPForMissing(method, arr, i-1, i, t, interval), 3);
        if (m.station) row[m.station] = station;
        if (m.temp) {
          const a = parseNumber(prev.raw[m.temp]), b = parseNumber(next.raw[m.temp]);
          if (Number.isFinite(a) && Number.isFinite(b)) row[m.temp] = formatNum(a + (b-a)*(k/(miss+1)), 1);
          else if (Number.isFinite(a)) row[m.temp] = formatNum(a,1);
        }
        if (m.rain) row[m.rain] = '0';
        if (m.holiday) row[m.holiday] = holidayByRules(t) ? '1' : '0';
        if (m.abnormal) row[m.abnormal] = '0';
        if (m.outage) row[m.outage] = '0';
        if (m.transfer) row[m.transfer] = '0';
        row[syntheticCol] = '1';
        row[noteCol] = `LV5.4 noi suy moc thieu ${method}`;
        state.rawRows.push(row); added++;
      }
    }
  }
  state.rawRows.sort((a,b) => {
    const ta = parseTime(a[m.time]) || 0, tb = parseTime(b[m.time]) || 0;
    if (ta - tb) return ta - tb;
    const sa = m.station ? String(a[m.station] || '') : '';
    const sb = m.station ? String(b[m.station] || '') : '';
    return sa.localeCompare(sb, 'vi');
  });
  normalizeRows(); applyDataInfo(); renderEditorTable(); previewData(); markEditorDirty(true);
  const parts = [
    `<span class="pill ok">Đã nội suy ${fixedBadP} giá trị P trống/lỗi/P=0/P thấp</span>`,
    `<span class="pill ok">Đã bổ sung ${added} mốc thời gian thiếu</span>`,
    `<span class="pill ${skipped?'warn':''}">Bỏ qua ${skipped} mốc do gap quá lớn</span>`
  ];
  if (skippedByDate) parts.push(`<span class="pill warn">Bỏ qua ${skippedByDate} mốc ngoài bộ lọc ngày</span>`);
  $('interpolationInfo').innerHTML = parts.join('');
  log(`LV5.4 nội suy/bổ sung hoàn tất: sửa ${fixedBadP} P, thêm ${added} mốc, bỏ qua ${skipped}.`);
}

try {
  state.lv5.version = 'LV5.4';
  log('Đã nạp lớp nội suy LV5.4 nền tảng: P=0/P thấp được xử lý theo lựa chọn cờ vận hành, có lưu p_goc.');
} catch(_) {}

// ======================== LV6 DESIGNATION + OPERATION ANALYSIS ========================
// Quy tắc LV6:
//   Đơn vị/Trạm/Lộ
//   Đơn vị/Trạm/Lộ_Đơn vị nối vòng/Trạm nối vòng/Lộ nối vòng
// Ví dụ: Tuyên Quang/E22.1/473_Hà Giang/E14.1/471
//        Hà Giang/E14.2/473

function cleanDesignationPart(v) {
  return String(v ?? '').trim().replace(/\s+/g, ' ');
}

function parseDesignationPath(text) {
  const raw = cleanDesignationPart(text);
  const parts = raw.split('/').map(cleanDesignationPart).filter(Boolean);
  if (parts.length >= 3) {
    return {
      valid: true,
      unit: parts[0],
      substation: parts[1],
      feeder: parts.slice(2).join('/'),
      key: `${parts[0]}/${parts[1]}/${parts.slice(2).join('/')}`
    };
  }
  return {
    valid: false,
    unit: '',
    substation: '',
    feeder: raw,
    key: raw || 'ALL'
  };
}

function parseDesignation(text) {
  const raw = cleanDesignationPart(text);
  if (!raw || raw.toUpperCase() === 'ALL' || raw === '__ALL__') {
    return {raw, valid:false, primary:parseDesignationPath('ALL'), ring:null, primaryKey:'ALL', forecastKey:'ALL', display:'ALL', hasRing:false};
  }
  const us = raw.split('_');
  const left = us.shift();
  const right = us.join('_').trim();
  const primary = parseDesignationPath(left);
  let ring = null;
  if (right) {
    if (right.includes('/')) ring = parseDesignationPath(right);
    else {
      // Hỗ trợ định dạng cũ: Đơn vị/Trạm/Lộ_LộNốiVòng
      ring = {
        valid: !!primary.valid,
        unit: primary.unit,
        substation: primary.substation,
        feeder: cleanDesignationPart(right),
        key: primary.valid ? `${primary.unit}/${primary.substation}/${cleanDesignationPart(right)}` : cleanDesignationPart(right)
      };
    }
  }
  const primaryKey = primary.valid ? primary.key : (raw || 'ALL');
  const forecastKey = ring && ring.key ? `${primaryKey}_${ring.key}` : primaryKey;
  return {
    raw,
    valid: !!primary.valid,
    primary,
    ring,
    primaryKey,
    forecastKey,
    display: forecastKey,
    hasRing: !!(ring && ring.key),
    unit: primary.unit,
    substation: primary.substation,
    feeder: primary.feeder,
    ringUnit: ring?.unit || '',
    ringSubstation: ring?.substation || '',
    ringFeeder: ring?.feeder || ''
  };
}

function stationLabelFromKey(key) {
  const d = parseDesignation(key);
  if (d.valid) {
    if (d.hasRing) return `${d.primary.unit}/${d.primary.substation}/${d.primary.feeder} ⇄ ${d.ring.unit}/${d.ring.substation}/${d.ring.feeder}`;
    return `${d.primary.unit}/${d.primary.substation}/${d.primary.feeder}`;
  }
  return key || 'ALL';
}

function enrichRawWithDesignation(raw, m) {
  const txt = m.station ? String(raw?.[m.station] ?? '').trim() : 'ALL';
  return parseDesignation(txt || 'ALL');
}

function normalizeRows() {
  const m = readMap();
  if (!m.time || !m.p) throw new Error('Cần chọn cột Thời gian và Công suất P.');
  state.rows = state.rawRows.map((r, idx) => {
    const time = parseTime(r[m.time]);
    const p = parseNumber(r[m.p]);
    const des = enrichRawWithDesignation(r, m);
    const station = des.forecastKey || des.primaryKey || 'ALL';
    const temp = m.temp ? parseNumber(r[m.temp]) : NaN;
    const rain = m.rain ? parseNumber(r[m.rain]) : 0;
    const holiday = m.holiday ? parseFlag(r[m.holiday]) : 0;
    const abnormal = m.abnormal ? parseFlag(r[m.abnormal]) : 0;
    const outage = m.outage ? parseFlag(r[m.outage]) : 0;
    const transfer = m.transfer ? parseFlag(r[m.transfer]) : 0;
    const excludeTrain = parseFlag(r['bo_khoi_huan_luyen']);
    return {
      idx, raw: r, time, p,
      station: station || 'ALL',
      stationRaw: m.station ? String(r[m.station] ?? '').trim() : 'ALL',
      designation: des,
      unit: des.unit || '', substation: des.substation || '', feeder: des.feeder || '',
      primaryKey: des.primaryKey || station,
      ringKey: des.ring?.key || '', ringUnit: des.ringUnit || '', ringSubstation: des.ringSubstation || '', ringFeeder: des.ringFeeder || '', hasRing: des.hasRing ? 1 : 0,
      temp, rain, holiday, abnormal, outage, transfer, excludeTrain
    };
  }).filter(r => r.time && Number.isFinite(r.p));
  state.rows.sort((a,b) => a.time - b.time || String(a.station).localeCompare(String(b.station), 'vi'));
  buildStationSelect();
  updateDesignationInfo();
  log(`Đã chuẩn hóa LV6 ${state.rows.length} dòng hợp lệ, đã tách chỉ danh Đơn vị/Trạm/Lộ/Nối vòng.`);
  if (state.rows.length < 20) log('Cảnh báo: dữ liệu quá ít, model sẽ kém ổn định.');
}

function buildStationSelect() {
  const sel = $('stationSelect');
  if (!sel) return;
  const selected = sel.value;
  const stations = [...new Set(state.rows.map(r => r.station || 'ALL'))].sort((a,b)=>stationLabelFromKey(a).localeCompare(stationLabelFromKey(b),'vi'));
  sel.innerHTML = '<option value="__ALL__">Tất cả / không lọc</option>';
  stations.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = stationLabelFromKey(s);
    sel.appendChild(opt);
  });
  if ([...sel.options].some(o => o.value === selected)) sel.value = selected;
}

function getSelectedRows() {
  const station = $('stationSelect')?.value || '__ALL__';
  return state.rows.filter(r => station === '__ALL__' || r.station === station).sort((a,b) => a.time - b.time);
}

function groupRowsByStationRaw() {
  const m = readMap();
  const groups = new Map();
  for (let i=0; i<state.rawRows.length; i++) {
    const raw = state.rawRows[i];
    const t = m.time ? parseTime(raw[m.time]) : null;
    if (!t) continue;
    const des = enrichRawWithDesignation(raw, m);
    const st = des.forecastKey || des.primaryKey || 'ALL';
    if (!groups.has(st)) groups.set(st, []);
    groups.get(st).push({raw, rawIndex:i, time:t, p:m.p?parseNumber(raw[m.p]):NaN, station:st, designation:des});
  }
  for (const arr of groups.values()) arr.sort((a,b)=>a.time-b.time);
  return groups;
}

function updateDesignationInfo() {
  const box = $('designationInfo');
  if (!box) return;
  if (!state.rawRows.length) { box.innerHTML = '<span class="pill">Chưa có dữ liệu</span>'; return; }
  const m = readMap();
  const rows = state.rawRows.map(r => enrichRawWithDesignation(r, m));
  const valid = rows.filter(d => d.valid).length;
  const ring = rows.filter(d => d.hasRing).length;
  const invalid = rows.length - valid;
  const units = [...new Set(rows.map(d => d.unit).filter(Boolean))].sort();
  const examples = [...new Set(rows.map(d => d.raw).filter(Boolean))].slice(0,4).map(x => `<span class="pill">${escapeHtml(x)}</span>`).join('');
  box.innerHTML = [
    `<span class="pill ok">${valid} dòng đúng dạng Đơn vị/Trạm/Lộ</span>`,
    `<span class="pill warn">${ring} dòng có nối vòng</span>`,
    `<span class="pill ${invalid?'bad':''}">${invalid} dòng chưa đúng dạng LV6</span>`,
    `<span class="pill">${units.length} đơn vị</span>`,
    examples ? `<div style="margin-top:6px">${examples}</div>` : ''
  ].join('');
}

function ensureRawColumn(name) {
  if (!state.headers.includes(name)) {
    state.headers.push(name);
    state.rawRows.forEach(r => { if (!(name in r)) r[name] = ''; });
  }
  return name;
}

function applyDesignationColumns() {
  if (!state.headers.length) throw new Error('Chưa có dữ liệu.');
  const m = readMap();
  if (!m.station) throw new Error('Cần ánh xạ cột Trạm/Lộ/Khu vực chứa chỉ danh LV6.');
  const current = getColumnSelectValues();
  const cols = ['lv6_don_vi','lv6_tram','lv6_lo','lv6_don_vi_noi_vong','lv6_tram_noi_vong','lv6_lo_noi_vong','lv6_chi_danh_chuan','lv6_chi_danh_du_bao','lv6_co_noi_vong'];
  cols.forEach(ensureRawColumn);
  let ok=0, ring=0, bad=0;
  for (const raw of state.rawRows) {
    const d = enrichRawWithDesignation(raw, m);
    raw.lv6_don_vi = d.unit || '';
    raw.lv6_tram = d.substation || '';
    raw.lv6_lo = d.feeder || '';
    raw.lv6_don_vi_noi_vong = d.ringUnit || '';
    raw.lv6_tram_noi_vong = d.ringSubstation || '';
    raw.lv6_lo_noi_vong = d.ringFeeder || '';
    raw.lv6_chi_danh_chuan = d.primaryKey || '';
    raw.lv6_chi_danh_du_bao = d.forecastKey || '';
    raw.lv6_co_noi_vong = d.hasRing ? '1' : '0';
    if (d.valid) ok++; else bad++;
    if (d.hasRing) ring++;
  }
  fillColumnSelects(state.headers); restoreColumnSelectValues(current);
  normalizeRows(); renderEditorTable(); previewData(); markEditorDirty(true);
  updateDesignationInfo();
  log(`LV6 đã tách chỉ danh: ${ok} dòng hợp lệ, ${ring} dòng có nối vòng, ${bad} dòng chưa đúng dạng.`);
}

function exportDesignationMap() {
  const m = readMap();
  const seen = new Set();
  const rows = [];
  for (const raw of state.rawRows) {
    const d = enrichRawWithDesignation(raw, m);
    const k = d.forecastKey || d.raw;
    if (seen.has(k)) continue;
    seen.add(k);
    rows.push({
      chi_danh_goc: d.raw,
      don_vi: d.unit,
      tram: d.substation,
      lo: d.feeder,
      don_vi_noi_vong: d.ringUnit,
      tram_noi_vong: d.ringSubstation,
      lo_noi_vong: d.ringFeeder,
      chi_danh_chuan: d.primaryKey,
      chi_danh_du_bao: d.forecastKey,
      co_noi_vong: d.hasRing ? 1 : 0,
      hop_le_lv6: d.valid ? 1 : 0
    });
  }
  saveTextFile('designation_map.csv', toCSV(rows, ['chi_danh_goc','don_vi','tram','lo','don_vi_noi_vong','tram_noi_vong','lo_noi_vong','chi_danh_chuan','chi_danh_du_bao','co_noi_vong','hop_le_lv6']), 'text/csv;charset=utf-8');
  log(`Đã xuất designation_map.csv gồm ${rows.length} chỉ danh duy nhất.`);
}

function previewData() {
  if (!state.rows.length) normalizeRows();
  const rows = getSelectedRows().slice(0, 200).map(r => ({
    time: fmtTime(r.time),
    don_vi: r.unit || '', tram: r.substation || '', lo: r.feeder || '',
    noi_vong: r.hasRing ? `${r.ringUnit}/${r.ringSubstation}/${r.ringFeeder}` : '',
    station: r.station,
    p:formatNum(r.p,3), temp:formatNum(r.temp,1), rain: r.rain ?? '', holiday:r.holiday, abnormal:r.abnormal, outage:r.outage, transfer:r.transfer
  }));
  renderTable(rows, ['time','don_vi','tram','lo','noi_vong','station','p','temp','rain','holiday','abnormal','outage','transfer']);
  log(`Xem trước LV6 ${rows.length} dòng.`);
}

function previousItemInGroup(groups, station, time) {
  const arr = groups.get(station) || [];
  let prev = null;
  for (const it of arr) {
    if (it.time < time && Number.isFinite(it.p)) prev = it;
    if (it.time >= time) break;
  }
  return prev;
}

function previousCleanItemInGroup(groups, station, time, minP) {
  const arr = groups.get(station) || [];
  let prev = null;
  for (const it of arr) {
    if (it.time >= time) break;
    const badForBaseline = !Number.isFinite(it.p) || isLowPValue(it.p, minP) || it.outage || it.transfer || it.abnormal || it.excludeTrain;
    if (!badForBaseline) prev = it;
  }
  return prev;
}

function inferIntervalMsForStation(groups, station) {
  const arr = groups.get(station) || [];
  const mins = detectIntervalMinutes(arr.map(x => ({time:x.time, p:x.p}))) || 60;
  return mins * 60 * 1000;
}

function findTimeWindowCandidates(rows, t, sourceStation, windowMs) {
  const center = t.getTime();
  return rows.filter(r => {
    if (r.station === sourceStation) return false;
    if (!Number.isFinite(r.p)) return false;
    return Math.abs(r.time.getTime() - center) <= windowMs;
  });
}

function isExpectedReceiver(c, expectedKey, sourcePrimaryKey) {
  if (!c) return false;
  const expected = !!(expectedKey && (c.primaryKey === expectedKey || c.station === expectedKey || String(c.station).startsWith(expectedKey + '_')));
  const reciprocal = !!(sourcePrimaryKey && c.ringKey && c.ringKey === sourcePrimaryKey);
  return expected || reciprocal;
}

function chooseTransferReceiver(candidates, groups, sourceRow, drop, minP, tolPct) {
  const expectedKey = sourceRow.ringKey || '';
  const sourcePrimaryKey = sourceRow.primaryKey || '';
  const scored = [];
  for (const c of candidates) {
    if (isLowPValue(c.p, minP) || c.outage || c.abnormal) continue;
    const cpClean = previousCleanItemInGroup(groups, c.station, sourceRow.time, minP) || previousItemInGroup(groups, c.station, sourceRow.time);
    const recvDelta = cpClean && Number.isFinite(cpClean.p) ? c.p - cpClean.p : NaN;
    if (!Number.isFinite(recvDelta) || recvDelta <= 0) continue;
    const expected = isExpectedReceiver(c, expectedKey, sourcePrimaryKey);
    const diff = Number.isFinite(drop) && drop > 0 ? Math.abs(recvDelta - drop) : 0;
    const withinTol = !(Number.isFinite(drop) && drop > 0) || diff <= drop * tolPct / 100;

    // Nếu không phải lộ nối vòng dự kiến, chỉ coi là ứng viên khi mức tăng gần bằng mức giảm.
    // Quy tắc này tránh chọn nhầm các lộ dao động tăng nhẹ cùng thời điểm.
    if (!expected && Number.isFinite(drop) && drop > 0 && !withinTol) continue;

    const ratioPenalty = Number.isFinite(drop) && drop > 0 ? diff / Math.max(drop, 1e-9) : 0;
    const score = (expected ? 10000 : 0) + (withinTol ? 1000 : 0) + Math.max(0, 500 - ratioPenalty * 500) + Math.min(recvDelta, 999);
    scored.push({...c, recvPrevP: cpClean?.p, recvDelta, recvDiff: diff, receiverExpected: expected ? 1 : 0, receiverWithinTol: withinTol ? 1 : 0, receiverScore: score});
  }
  scored.sort((a,b) => b.receiverScore - a.receiverScore);
  return scored[0] || null;
}

function analyzeOperationEvents() {
  if (!state.rows.length) normalizeRows();
  const minP = getInvalidPThreshold();
  const dropPctMin = Math.max(1, parseNumber($('operationDropPercent')?.value) || 35);
  const tolPct = Math.max(1, parseNumber($('transferTolerancePercent')?.value) || 30);
  const windowMode = $('operationTimeWindow')?.value || 'same';
  const groups = new Map();
  for (const r of state.rows) {
    if (!groups.has(r.station)) groups.set(r.station, []);
    groups.get(r.station).push(r);
  }
  for (const arr of groups.values()) arr.sort((a,b)=>a.time-b.time);

  const events = [];
  for (const [station, arr] of groups) {
    for (let i=0; i<arr.length; i++) {
      const r = arr[i];
      const prevImmediate = i > 0 ? arr[i-1] : null;
      const baseline = previousCleanItemInGroup(groups, station, r.time, minP) || prevImmediate;
      const baselineP = baseline && Number.isFinite(baseline.p) ? baseline.p : NaN;
      const delta = Number.isFinite(baselineP) && Number.isFinite(r.p) ? r.p - baselineP : NaN;
      const drop = Number.isFinite(delta) && delta < 0 ? -delta : 0;
      const dropPct = Number.isFinite(baselineP) && Math.abs(baselineP) > 1e-9 ? drop / Math.abs(baselineP) * 100 : NaN;
      const low = isLowPValue(r.p, minP);
      const significantDrop = Number.isFinite(dropPct) && dropPct >= dropPctMin && drop > 0;

      // LV6.1: Không coi mọi dòng có cờ chuyển tải là "lộ nghi vấn".
      // Dòng có cờ chuyển tải nhưng P đang tăng/ổn định thường là lộ nhận tải, dùng làm ứng viên nhận tải chứ không đưa vào báo cáo nguồn sự kiện.
      const sourceLike = low || significantDrop || ((r.outage || r.abnormal) && (low || significantDrop));
      if (!sourceLike) continue;

      const intervalMs = inferIntervalMsForStation(groups, station);
      const windowMs = windowMode === 'near' ? intervalMs * 1.1 : 0;
      const candidates = findTimeWindowCandidates(state.rows, r.time, station, windowMs);
      const receiver = r.transfer ? chooseTransferReceiver(candidates, groups, r, drop, minP, tolPct) : null;

      let conclusion = '', suggestion = '';
      const flags = [r.outage?'cắt điện/sự cố':'', r.transfer?'chuyển tải':'', r.abnormal?'bất thường':''].filter(Boolean).join(', ');
      const expectedKey = r.ringKey || '';
      const receiverExpected = receiver ? isExpectedReceiver(receiver, expectedKey, r.primaryKey) : false;
      const diff = receiver ? Math.abs((receiver.recvDelta || 0) - drop) : NaN;
      const okBalance = receiver && drop > 0 ? diff <= drop * tolPct/100 : !!receiver;

      if (r.transfer) {
        if (receiver) {
          if (receiverExpected && okBalance) {
            conclusion = 'Chuyển tải hợp lý: đúng lộ nối vòng dự kiến và cân bằng tải';
            suggestion = 'Giữ cờ chuyển tải, bỏ khỏi huấn luyện nền; không cần chọn lộ nghi vấn khác.';
          } else if (receiverExpected) {
            conclusion = 'Đúng lộ nối vòng dự kiến nhưng chưa cân bằng tải';
            suggestion = 'Kiểm tra tổng P khu vực, thời điểm thao tác hoặc dữ liệu đo lộ nhận.';
          } else if (okBalance) {
            conclusion = 'Có lộ nhận tăng tải nhưng khác lộ nối vòng dự kiến';
            suggestion = 'Kiểm tra lại chỉ danh nối vòng hoặc kết lưới thực tế.';
          } else {
            conclusion = 'Có lộ nhận tăng nhưng chưa cân bằng tải';
            suggestion = 'Kiểm tra thêm kết lưới, tổng P khu vực hoặc dữ liệu đo lộ nhận.';
          }
        } else {
          conclusion = 'Chuyển tải chưa thấy lộ nhận tăng phù hợp';
          suggestion = 'Kiểm tra lại chỉ danh lộ nối vòng, dữ liệu lộ nhận cùng thời điểm hoặc chọn cửa sổ ±1 mốc.';
        }
      } else if (r.outage) {
        conclusion = 'Cắt điện/sự cố làm P thấp/sụt tải';
        suggestion = 'Giữ sự kiện vận hành, bỏ khỏi huấn luyện nền; có thể nội suy P sạch nếu cần chuỗi liên tục.';
      } else if (low) {
        conclusion = 'P thấp/P=0 không có cờ chuyển tải/cắt điện';
        suggestion = 'Nên nội suy hoặc đánh dấu bất thường trước khi huấn luyện.';
      } else {
        conclusion = 'P sụt bất thường';
        suggestion = 'Kiểm tra sự kiện vận hành, chuyển tải chéo hoặc lỗi đo.';
      }

      events.push({
        rowIndex: r.idx,
        time: fmtTime(r.time),
        chi_danh: r.stationRaw || r.station,
        don_vi: r.unit, tram: r.substation, lo: r.feeder,
        vai_tro: r.transfer ? 'nguồn giảm/chuyển tải' : (r.outage ? 'nguồn cắt điện/sự cố' : 'nguồn bất thường'),
        p_moc_chuan: Number.isFinite(baselineP) ? formatNum(baselineP,3) : '',
        p_hien_tai: formatNum(r.p,3),
        giam_mw: Number.isFinite(drop) ? formatNum(drop,3) : '',
        giam_pct: Number.isFinite(dropPct) ? formatNum(dropPct,1) : '',
        co_van_hanh: flags,
        noi_vong_du_kien: expectedKey,
        lo_nhan_nghi_van: receiver ? (receiver.stationRaw || receiver.station) : '',
        tang_mw_lo_nhan: receiver ? formatNum(receiver.recvDelta,3) : '',
        do_lech_can_bang_mw: receiver && Number.isFinite(diff) ? formatNum(diff,3) : '',
        khop_lo_noi_vong: receiver ? (receiverExpected ? 1 : 0) : '',
        ket_luan: conclusion,
        goi_y: suggestion
      });
    }
  }
  state.operationEvents = events;
  renderOperationReport();
  log(`LV6.1 phân tích sự kiện vận hành: ${events.length} nguồn sự kiện cần xem xét. Dòng lộ nhận tải chỉ dùng làm ứng viên, không còn bị báo nhầm là lộ nghi vấn.`);
  return events;
}

function renderOperationReport() {
  const box = $('operationBox');
  const sum = $('operationSummary');
  if (!box || !sum) return;
  const rows = state.operationEvents || [];
  if (!rows.length) {
    sum.innerHTML = '<span class="pill ok">Không có sự kiện cần cảnh báo theo ngưỡng hiện tại</span>';
    box.innerHTML = '<table><tbody><tr><td>Không có báo cáo</td></tr></tbody></table>';
    return;
  }
  const nTransferMissing = rows.filter(r => String(r.ket_luan).includes('chưa thấy')).length;
  const nLow = rows.filter(r => String(r.ket_luan).includes('P thấp')).length;
  const nOk = rows.filter(r => String(r.ket_luan).includes('hợp lý')).length;
  sum.innerHTML = `<span class="pill warn">${rows.length} sự kiện</span><span class="pill ok">${nOk} chuyển tải hợp lý</span><span class="pill ${nTransferMissing?'bad':''}">${nTransferMissing} chuyển tải chưa thấy lộ nhận</span><span class="pill ${nLow?'warn':''}">${nLow} P thấp chưa có cờ</span>`;
  renderTableInBox(box, rows, ['rowIndex','time','chi_danh','don_vi','tram','lo','vai_tro','p_moc_chuan','p_hien_tai','giam_mw','giam_pct','co_van_hanh','noi_vong_du_kien','lo_nhan_nghi_van','tang_mw_lo_nhan','do_lech_can_bang_mw','khop_lo_noi_vong','ket_luan','goi_y'], 600);
}

function renderTableInBox(box, rows, headers, max=500) {
  const data = rows.slice(0, max);
  let html = '<table><thead><tr>' + headers.map(h => `<th>${escapeHtml(h)}</th>`).join('') + '</tr></thead><tbody>';
  for (const r of data) html += '<tr>' + headers.map(h => `<td>${escapeHtml(r[h] ?? '')}</td>`).join('') + '</tr>';
  html += '</tbody></table>';
  if (rows.length > max) html += `<div class="note">Đang hiển thị ${max}/${rows.length} dòng.</div>`;
  box.innerHTML = html;
}

function selectOperationRows() {
  const rows = [...new Set((state.operationEvents || []).map(x => x.rowIndex).filter(Number.isInteger))];
  rows.forEach(i => state.editor.selected.add(i));
  renderEditorTable();
  log(`Đã chọn ${rows.length} dòng sự kiện vận hành trên bảng hiệu chỉnh.`);
}

function exportOperationEvents() {
  const rows = state.operationEvents || [];
  saveTextFile('operation_event_report.csv', toCSV(rows, ['rowIndex','time','chi_danh','don_vi','tram','lo','vai_tro','p_moc_chuan','p_hien_tai','giam_mw','giam_pct','co_van_hanh','noi_vong_du_kien','lo_nhan_nghi_van','tang_mw_lo_nhan','do_lech_can_bang_mw','khop_lo_noi_vong','ket_luan','goi_y']), 'text/csv;charset=utf-8');
  log(`Đã xuất operation_event_report.csv gồm ${rows.length} dòng.`);
}

function restoreOriginalPForSelected() {
  const m = readMap();
  if (!m.p) throw new Error('Cần ánh xạ cột Công suất P.');
  const selected = [...(state.editor?.selected || new Set())].filter(i => i >= 0 && i < state.rawRows.length);
  if (!selected.length) { log('Chưa chọn dòng để khôi phục P gốc.'); return; }
  addProcessColumns();
  let n=0;
  for (const idx of selected) {
    const raw = state.rawRows[idx];
    if (raw && raw.p_goc !== undefined && String(raw.p_goc).trim() !== '') {
      raw[m.p] = raw.p_goc;
      raw.du_lieu_noi_suy = '0';
      raw.ghi_chu_xu_ly = `${raw.ghi_chu_xu_ly || ''} LV6 khoi phuc P goc`.trim();
      n++;
    }
  }
  normalizeRows(); renderEditorTable(); previewData(); markEditorDirty(true);
  log(`LV6 đã khôi phục P gốc cho ${n}/${selected.length} dòng đã chọn.`);
}

function compareEditedOriginal() {
  const m = readMap();
  const rows = [];
  for (let i=0; i<state.rawRows.length; i++) {
    const raw = state.rawRows[i];
    if (raw.p_goc !== undefined && String(raw.p_goc).trim() !== '') {
      const p0 = parseNumber(raw.p_goc);
      const p1 = m.p ? parseNumber(raw[m.p]) : NaN;
      const d = enrichRawWithDesignation(raw, m);
      rows.push({
        rowIndex:i,
        time:m.time ? raw[m.time] : '',
        chi_danh:m.station ? raw[m.station] : '',
        don_vi:d.unit, tram:d.substation, lo:d.feeder,
        p_goc:Number.isFinite(p0)?formatNum(p0,3):raw.p_goc,
        p_hien_tai:Number.isFinite(p1)?formatNum(p1,3):'',
        chenhlech:Number.isFinite(p0)&&Number.isFinite(p1)?formatNum(p1-p0,3):'',
        ghi_chu:raw.ghi_chu_xu_ly || ''
      });
    }
  }
  state.operationEvents = rows.map(r => ({...r, ket_luan:'So sánh trước/sau hiệu chỉnh', goi_y:'Kiểm tra lại trước khi huấn luyện'}));
  renderTableInBox($('operationBox'), rows, ['rowIndex','time','chi_danh','don_vi','tram','lo','p_goc','p_hien_tai','chenhlech','ghi_chu'], 1000);
  $('operationSummary').innerHTML = `<span class="pill warn">${rows.length} dòng có p_goc để so sánh</span>`;
  log(`LV6 so sánh trước/sau hiệu chỉnh: ${rows.length} dòng có p_goc.`);
}

function exportForecast() {
  if (!state.forecastRows.length) return;
  const rows = state.forecastRows.map(r => {
    const d = parseDesignation(r.station);
    return {...r, don_vi:d.unit||'', tram:d.substation||'', lo:d.feeder||'', noi_vong:d.hasRing?`${d.ringUnit}/${d.ringSubstation}/${d.ringFeeder}`:''};
  });
  saveTextFile('forecast.csv', toCSV(rows, ['step','time','don_vi','tram','lo','noi_vong','station','forecast_p_mw','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong']), 'text/csv;charset=utf-8');
  log('Đã xuất forecast.csv theo cấu trúc LV6 có Đơn vị/Trạm/Lộ/Nối vòng.');
}

function lv6BindEvents() {
  $('parseDesignationBtn')?.addEventListener('click', () => { try { applyDesignationColumns(); } catch(e) { log('Lỗi tách chỉ danh LV6: ' + e.message); } });
  $('exportDesignationBtn')?.addEventListener('click', () => { try { exportDesignationMap(); } catch(e) { log('Lỗi xuất designation_map.csv: ' + e.message); } });
  $('analyzeOperationBtn')?.addEventListener('click', () => { try { analyzeOperationEvents(); } catch(e) { log('Lỗi phân tích vận hành LV6: ' + e.message); } });
  $('selectOperationRowsBtn')?.addEventListener('click', selectOperationRows);
  $('exportOperationBtn')?.addEventListener('click', exportOperationEvents);
  $('restoreOriginalPBtn')?.addEventListener('click', () => { try { restoreOriginalPForSelected(); } catch(e) { log('Lỗi khôi phục P gốc: ' + e.message); } });
  $('compareEditsBtn')?.addEventListener('click', () => { try { compareEditedOriginal(); } catch(e) { log('Lỗi so sánh hiệu chỉnh: ' + e.message); } });
}

setTimeout(() => {
  try {
    lv6BindEvents();
    if (state.lv5) state.lv5.version = 'LV6';
    log('Sẵn sàng LV6: đã bật quy tắc Đơn vị/Trạm/Lộ/Lộ nối vòng và phân tích chuyển tải chéo.');
  } catch(e) { log('Lỗi khởi tạo LV6: ' + e.message); }
}, 0);
// ====================== END LV6 EXTENSIONS ======================

// ======================== LV7 MODEL SELECTION EXTENSIONS ========================
// LV7 tập trung nâng cấp mô hình: so sánh nhiều thuật toán dự báo, tự chọn mô hình tốt nhất
// theo từng chỉ danh Đơn vị/Trạm/Lộ/Lộ nối vòng và lưu chiến lược dự báo vào gói model vận hành xuất từ Mục 8.
state.lv7 = state.lv7 || {version:'LV7', modelReports:[], hourlyReports:[], strategyByStation:{}};

const LV7_METHOD_LABELS = {
  gbdt: 'GBDT',
  similar_day: 'Similar Day',
  same_hour_last_week: 'Cùng giờ tuần trước',
  trend: 'Xu hướng gần nhất',
  hybrid_current: 'Hybrid trọng số hiện tại',
  auto_blend: 'Auto Blend theo sai số'
};

function getStationRowsForV7(station) {
  const st = station || $('stationSelect')?.value || '__ALL__';
  const rows = (state.rows || []).filter(r => st === '__ALL__' || r.station === st).sort((a,b) => a.time - b.time);
  return rows;
}

function getValidationTargetIndicesV7(series, valPercent=null) {
  const interval = detectIntervalMinutes(series);
  const nPerDay = Math.max(1, Math.round(1440 / interval));
  const tempMean = mean(series.map(r => r.temp));
  const filled = series.map(r => ({...r, temp: Number.isFinite(r.temp) ? r.temp : tempMean}));
  const targets = [];
  for (let i=1; i<filled.length; i++) {
    const r = filled[i];
    if (!Number.isFinite(r.p)) continue;
    if (r.abnormal || r.outage || r.transfer || r.excludeTrain) continue;
    const x = computeFeatureVector(filled, i, nPerDay, {});
    if (!Number.isFinite(x[11])) continue;
    targets.push(i);
  }
  const pct = Number.isFinite(valPercent) ? valPercent : (parseNumber($('valPercent')?.value) || 20);
  const nVal = Math.max(1, Math.floor(targets.length * pct / 100));
  return {targets, evalIdx: targets.slice(-nVal), filled, interval, nPerDay};
}

function combinePredictionV7(values, weights) {
  let sw = 0, s = 0;
  for (const [k, w0] of Object.entries(weights || {})) {
    const v = values[k];
    const w = Number(w0);
    if (Number.isFinite(v) && Number.isFinite(w) && w > 0) { s += v*w; sw += w; }
  }
  return sw ? s / sw : NaN;
}

function makeInverseErrorWeightsV7(metricMap) {
  const candidates = ['gbdt','similar_day','same_hour_last_week','trend'];
  const raw = {};
  let sum = 0;
  for (const m of candidates) {
    const e = metricMap[m]?.mape;
    if (Number.isFinite(e) && e >= 0) {
      // MAPE càng nhỏ thì trọng số càng lớn. Chặn dưới để tránh chia quá lớn.
      raw[m] = 1 / Math.max(0.5, e);
      sum += raw[m];
    }
  }
  if (!sum) return {gbdt:0.5, similar_day:0.25, same_hour_last_week:0.15, trend:0.10};
  const out = {};
  for (const [k,v] of Object.entries(raw)) out[k] = v / sum;
  return out;
}

function scoreModelCandidatesV7(station=null, model=null) {
  if (!state.rows.length) normalizeRows();
  const st = station || $('stationSelect')?.value || '__ALL__';
  const mdl = model || resolveModelForStation(st) || state.model;
  const series = getStationRowsForV7(st);
  if (series.length < 40) throw new Error('Dữ liệu quá ít để so sánh mô hình cho ' + st);
  const {evalIdx, filled, nPerDay, interval} = getValidationTargetIndicesV7(series);
  if (evalIdx.length < 5) throw new Error('Không đủ điểm validation sạch để so sánh mô hình cho ' + st);

  const actual = [];
  const times = [];
  const hours = [];
  const preds = {gbdt:[], similar_day:[], same_hour_last_week:[], trend:[], hybrid_current:[], auto_blend:[]};
  const wCurrent0 = getHybridWeights();
  const wCurrent = {gbdt:wCurrent0.gbdt, similar_day:wCurrent0.similar, same_hour_last_week:wCurrent0.week, trend:wCurrent0.trend};

  // Vòng 1: tạo dự báo cho các mô hình đơn và hybrid trọng số hiện tại.
  for (const idx of evalIdx) {
    const r = filled[idx];
    const x = computeFeatureVector(filled, idx, nPerDay, {});
    const values = {
      gbdt: (mdl && mdl.trees) ? Math.max(0, predictModel(mdl, x)) : NaN,
      similar_day: similarDayPrediction(filled, idx, nPerDay),
      same_hour_last_week: lastWeekPrediction(filled, idx, nPerDay),
      trend: trendPrediction(filled, idx)
    };
    values.hybrid_current = combinePredictionV7(values, wCurrent);
    actual.push(r.p);
    times.push(r.time);
    hours.push(r.time.getHours());
    for (const k of ['gbdt','similar_day','same_hour_last_week','trend','hybrid_current']) preds[k].push(values[k]);
  }

  const metricMap = {};
  for (const k of ['gbdt','similar_day','same_hour_last_week','trend','hybrid_current']) metricMap[k] = metrics(actual, preds[k]);
  const autoWeights = makeInverseErrorWeightsV7(metricMap);

  // Vòng 2: Auto Blend dùng trọng số tự tính theo sai số validation.
  for (let i=0; i<actual.length; i++) {
    const values = {
      gbdt: preds.gbdt[i],
      similar_day: preds.similar_day[i],
      same_hour_last_week: preds.same_hour_last_week[i],
      trend: preds.trend[i]
    };
    preds.auto_blend.push(combinePredictionV7(values, autoWeights));
  }
  metricMap.auto_blend = metrics(actual, preds.auto_blend);

  const summary = Object.entries(metricMap).map(([method, mt]) => ({
    station: st,
    method,
    ten_mo_hinh: LV7_METHOD_LABELS[method] || method,
    n: mt.n,
    mae: Number.isFinite(mt.mae) ? mt.mae : NaN,
    mape: Number.isFinite(mt.mape) ? mt.mape : NaN,
    rmse: Number.isFinite(mt.rmse) ? mt.rmse : NaN,
    interval_minutes: interval,
    w_gbdt: method === 'auto_blend' ? autoWeights.gbdt : (method === 'hybrid_current' ? wCurrent.gbdt : ''),
    w_similar: method === 'auto_blend' ? autoWeights.similar_day : (method === 'hybrid_current' ? wCurrent.similar_day : ''),
    w_week: method === 'auto_blend' ? autoWeights.same_hour_last_week : (method === 'hybrid_current' ? wCurrent.same_hour_last_week : ''),
    w_trend: method === 'auto_blend' ? autoWeights.trend : (method === 'hybrid_current' ? wCurrent.trend : '')
  })).sort((a,b) => {
    const ma = Number.isFinite(a.mape) ? a.mape : Infinity;
    const mb = Number.isFinite(b.mape) ? b.mape : Infinity;
    return ma - mb;
  });
  summary.forEach((r,i) => { r.xep_hang = i + 1; r.khuyen_nghi = i === 0 ? 'Dùng cho dự báo' : ''; });

  const hourly = [];
  for (const method of Object.keys(preds)) {
    for (let h=0; h<24; h++) {
      const a=[], p=[];
      for (let i=0; i<actual.length; i++) if (hours[i] === h) { a.push(actual[i]); p.push(preds[method][i]); }
      if (a.length) {
        const mt = metrics(a,p);
        hourly.push({station:st, method, ten_mo_hinh:LV7_METHOD_LABELS[method]||method, hour:h, n:mt.n, mae:mt.mae, mape:mt.mape, rmse:mt.rmse});
      }
    }
  }

  const detail = evalIdx.map((idx,i) => ({
    station: st,
    time: fmtTime(times[i]),
    actual: actual[i],
    gbdt: preds.gbdt[i],
    similar_day: preds.similar_day[i],
    same_hour_last_week: preds.same_hour_last_week[i],
    trend: preds.trend[i],
    hybrid_current: preds.hybrid_current[i],
    auto_blend: preds.auto_blend[i]
  }));

  return {station: st, model: mdl, summary, hourly, detail, autoWeights, best: summary[0] || null, interval};
}

function formatMetricRowsV7(rows) {
  return rows.map(r => ({
    station:r.station,
    method:r.method,
    ten_mo_hinh:r.ten_mo_hinh,
    n:r.n,
    MAE:formatNum(r.mae,3),
    MAPE_pct:formatNum(r.mape,2),
    RMSE:formatNum(r.rmse,3),
    xep_hang:r.xep_hang || '',
    khuyen_nghi:r.khuyen_nghi || '',
    w_gbdt:Number.isFinite(r.w_gbdt)?formatNum(r.w_gbdt,3):r.w_gbdt,
    w_similar:Number.isFinite(r.w_similar)?formatNum(r.w_similar,3):r.w_similar,
    w_week:Number.isFinite(r.w_week)?formatNum(r.w_week,3):r.w_week,
    w_trend:Number.isFinite(r.w_trend)?formatNum(r.w_trend,3):r.w_trend
  }));
}

function renderV7Reports(result) {
  const sumBox = $('lv7ModelSummary');
  const modelBox = $('lv7ModelBox');
  const hourBox = $('lv7HourlyBox');
  if (!result) return;
  state.lv7.modelReports = result.summary || [];
  state.lv7.hourlyReports = result.hourly || [];
  const best = result.best;
  if (sumBox) {
    if (best) {
      sumBox.innerHTML = `<span class="pill modeBadge">LV7</span><span class="pill">${escapeHtml(result.station)}</span><span class="pill ok">Khuyến nghị: ${escapeHtml(best.ten_mo_hinh)}</span><span class="pill">MAPE ${formatNum(best.mape,2)}%</span><span class="pill">MAE ${formatNum(best.mae,3)} MW</span>`;
    } else sumBox.innerHTML = '<span class="pill warn">Chưa chọn được mô hình tốt nhất</span>';
  }
  if (modelBox) renderTableInBox(modelBox, formatMetricRowsV7(result.summary), ['station','method','ten_mo_hinh','n','MAE','MAPE_pct','RMSE','xep_hang','khuyen_nghi','w_gbdt','w_similar','w_week','w_trend'], 200);
  if (hourBox) {
    const hrows = result.hourly.map(r => ({station:r.station, method:r.method, hour:r.hour, n:r.n, MAE:formatNum(r.mae,3), MAPE_pct:formatNum(r.mape,2), RMSE:formatNum(r.rmse,3)}));
    renderTableInBox(hourBox, hrows, ['station','method','hour','n','MAE','MAPE_pct','RMSE'], 400);
  }
}

function evaluateModelsV7() {
  const st = $('stationSelect')?.value || '__ALL__';
  const result = scoreModelCandidatesV7(st, resolveModelForStation(st));
  state.lv7.lastResult = result;
  renderV7Reports(result);
  log(`LV7 đã so sánh ${result.summary.length} mô hình cho ${st}. Mô hình khuyến nghị: ${result.best?.ten_mo_hinh || '-'}; MAPE ${formatNum(result.best?.mape,2)}%.`);
  return result;
}

function ensureStrategyContainerV7() {
  // LV7.6: chiến lược so sánh phải được áp dụng vào một gói model đã huấn luyện ở Mục 8.
  // Không tạo model strategy-only để tránh người dùng xuất nhầm gói không có GBDT.
  if (!state.model || (!state.model.trees && !state.model.modelsByStation)) {
    throw new Error('Cần huấn luyện GBDT ở Mục 8 trước, sau đó mới áp dụng chiến lược so sánh vào model.');
  }
  if (!state.model.strategyByStation) state.model.strategyByStation = {};
  if (!state.lv7.strategyByStation) state.lv7.strategyByStation = {};
}

function selectBestModelV7(result=null) {
  const res = result || state.lv7.lastResult || evaluateModelsV7();
  if (!res?.best) throw new Error('Chưa có kết quả so sánh mô hình để chọn.');
  ensureStrategyContainerV7();
  const method = res.best.method;
  const weights = method === 'auto_blend' ? res.autoWeights : weightsFromReportV72(res.best);
  const strategy = makeStrategyV72(res.station, method, res.best.ten_mo_hinh, {mae:res.best.mae, mape:res.best.mape, rmse:res.best.rmse, n:res.best.n}, weights, res.interval);
  applyStrategyToModelV72(res.station, strategy);
  log(`LV7.6 đã áp dụng gợi ý tốt nhất cho ${res.station}: ${strategy.label}, MAPE ${formatNum(strategy.validation.mape,2)}%.`);
  return strategy;
}

async function trainSelectBestAllStationsV7() {
  if (!state.rows.length) normalizeRows();
  const original = $('stationSelect')?.value || '__ALL__';
  const stations = [...new Set((state.rows || []).map(r => r.station || 'ALL'))].filter(Boolean).sort();
  if (!stations.length) throw new Error('Không có chỉ danh để huấn luyện LV7.');
  const modelsByStation = {};
  const strategyByStation = {};
  const allReports = [];
  const allHourly = [];
  for (const st of stations) {
    if ($('stationSelect')) $('stationSelect').value = st;
    try {
      await trainGBDT();
      const mdl = state.model && state.model.trees ? state.model : null;
      if (!mdl) throw new Error('Không tạo được GBDT.');
      modelsByStation[st] = mdl;
      const res = scoreModelCandidatesV7(st, mdl);
      allReports.push(...res.summary);
      allHourly.push(...res.hourly);
      const best = res.best;
      strategyByStation[st] = {
        version:'LV7', selectedAt:new Date().toISOString(), station:st,
        method:best.method, label:best.ten_mo_hinh,
        validation:{mae:best.mae, mape:best.mape, rmse:best.rmse, n:best.n},
        weights: best.method === 'auto_blend' ? res.autoWeights : null,
        intervalMinutes: res.interval
      };
      log(`LV7 ${st}: chọn ${best.ten_mo_hinh}, MAPE ${formatNum(best.mape,2)}%.`);
    } catch(e) {
      log(`LV7 bỏ qua ${st}: ${e.message}`);
    }
  }
  if ($('stationSelect')) $('stationSelect').value = original;
  if (!Object.keys(modelsByStation).length) throw new Error('Không tạo được model nào trong LV7.');
  state.model = {
    type:'GBDT_STATION_BUNDLE_JS_OFFLINE_LV7',
    createdAt:new Date().toISOString(),
    modelsByStation,
    strategyByStation,
    colMap:state.colMap,
    lv7Reports: allReports,
    note:'Bundle LV7: mỗi chỉ danh có GBDT riêng và chiến lược dự báo tự chọn theo validation. Dùng offline trong SCADA.'
  };
  state.lv7.modelReports = allReports;
  state.lv7.hourlyReports = allHourly;
  state.lv7.strategyByStation = strategyByStation;
  $('exportModelBtn').disabled = false;
  $('forecastBtn').disabled = false;
  if ($('forecastAllBtn')) $('forecastAllBtn').disabled = false;
  renderModelInfo();
  const bestRows = Object.values(strategyByStation).map(s => ({station:s.station, method:s.method, ten_mo_hinh:s.label, n:s.validation.n, MAE:formatNum(s.validation.mae,3), MAPE_pct:formatNum(s.validation.mape,2), RMSE:formatNum(s.validation.rmse,3), khuyen_nghi:'Dùng cho dự báo'}));
  renderTableInBox($('lv7ModelBox'), bestRows, ['station','method','ten_mo_hinh','n','MAE','MAPE_pct','RMSE','khuyen_nghi'], 500);
  if ($('lv7ModelSummary')) $('lv7ModelSummary').innerHTML = `<span class="pill ok">LV7 đã huấn luyện ${Object.keys(modelsByStation).length} model riêng</span><span class="pill">Đã tự chọn chiến lược cho từng chỉ danh</span>`;
  log(`LV7 hoàn tất huấn luyện + tự chọn mô hình: ${Object.keys(modelsByStation).length} chỉ danh.`);
}

function getForecastStrategyV7(station) {
  const mode = $('appMode')?.value || 'external';
  const manual = (mode === 'scada' ? ($('forecastStrategyScada')?.value || $('forecastStrategy')?.value) : ($('forecastStrategy')?.value || $('forecastStrategyScada')?.value)) || 'auto';
  if (manual !== 'auto') return {method: manual, label: LV7_METHOD_LABELS[manual] || manual, weights: manual === 'auto_blend' ? null : null};
  const s = state.model?.strategyByStation?.[station] || state.model?.strategyByStation?.['__ALL__'] || state.model?.v7Strategy || state.lv7?.strategyByStation?.[station];
  if (s) return s;
  return {method:'hybrid_current', label:'Hybrid trọng số hiện tại', weights:null};
}

function forecastValueByStrategyV7(values, strategy) {
  const method = strategy?.method || 'hybrid_current';
  if (method === 'gbdt') return values.gbdt;
  if (method === 'similar_day') return values.similar_day;
  if (method === 'same_hour_last_week') return values.same_hour_last_week;
  if (method === 'trend') return values.trend;
  if (method === 'auto_blend') {
    const w = strategy.weights || {gbdt:0.5, similar_day:0.25, same_hour_last_week:0.15, trend:0.10};
    return combinePredictionV7(values, w);
  }
  const w0 = getHybridWeights();
  const w = {gbdt:w0.gbdt, similar_day:w0.similar, same_hour_last_week:w0.week, trend:w0.trend};
  return combinePredictionV7(values, w);
}

function forecastForStation(station, stepsOverride=null) {
  const model = resolveModelForStation(station);
  if (!model) throw new Error('Chưa có model phù hợp cho ' + station);
  const sourceRows = (state.rows || []).filter(r => station === '__ALL__' || r.station === station).sort((a,b)=>a.time-b.time);
  if (sourceRows.length < 5) throw new Error('Chưa có đủ dữ liệu gần nhất cho ' + station);
  const tempDefault = parseNumber($('futureTemp')?.value);
  const rainDefault = parseNumber($('futureRain')?.value) || 0;
  const interval = model.intervalMinutes || detectIntervalMinutes(sourceRows);
  const nPerDay = model.nPerDay || Math.round(1440/interval);
  const lastTemp = [...sourceRows].reverse().find(r => Number.isFinite(r.temp))?.temp;
  const tempUse = Number.isFinite(tempDefault) ? tempDefault : lastTemp;
  const series = sourceRows.map(r => ({...r}));
  const steps = Math.max(1, Math.floor(stepsOverride || parseNumber($('forecastSteps')?.value) || 24));
  const strategy = getForecastStrategyV7(station);
  const bias = estimateRecentBias(model, sourceRows, Math.max(0, Math.floor(parseNumber($('biasWindow')?.value) || 0)));
  const forecast=[];
  let lastTime = series[series.length-1].time;
  for (let s=1; s<=steps; s++) {
    const t = new Date(lastTime.getTime() + interval*60000);
    const row = {time:t, p:NaN, station: station === '__ALL__' ? 'ALL' : station, temp:tempUse, rain:rainDefault, holiday:holidayByRules(t)?1:0, abnormal:0, outage:0, transfer:0};
    series.push(row);
    const idx = series.length - 1;
    const x = computeFeatureVector(series, idx, nPerDay, {temp: tempUse, rain: rainDefault});
    const values = {
      gbdt: Math.max(0, predictModel(model, x)),
      similar_day: similarDayPrediction(series, idx, nPerDay),
      same_hour_last_week: lastWeekPrediction(series, idx, nPerDay),
      trend: trendPrediction(series, idx)
    };
    let pred = forecastValueByStrategyV7(values, strategy);
    if (!Number.isFinite(pred)) pred = values.gbdt;
    let rawPredBeforeCalibration = Math.max(0, pred + bias);
    const cal = applyCalibrationLV85(rawPredBeforeCalibration, station, t);
    pred = cal.value;
    row.p = pred;
    forecast.push({
      step:s, time:fmtTime(t), station:row.station, forecast_p_mw:formatNum(pred,3), temp:Number.isFinite(tempUse)?formatNum(tempUse,1):'', rain:rainDefault, holiday:row.holiday,
      model_used:strategy.label || LV7_METHOD_LABELS[strategy.method] || strategy.method || 'LV7',
      strategy_method:strategy.method || '',
      gbdt:formatNum(values.gbdt,3), similar_day:formatNum(values.similar_day,3), same_hour_last_week:formatNum(values.same_hour_last_week,3), trend:formatNum(values.trend,3), bias:formatNum(bias,3), calibration_lv85:cal.applied?1:0, calibration_mw:formatNum(cal.delta,3), calibration_source:cal.source || '', forecast_before_calibration_mw:formatNum(rawPredBeforeCalibration,3)
    });
    lastTime = t;
  }
  return {forecast, series, interval, model, station, strategy, weights:strategy.weights || null};
}

function forecastNext() {
  if (!state.model) throw new Error('Chưa có model.');
  const selected = $('stationSelect')?.value || '__ALL__';
  const station = selected === '__ALL__' && state.model.modelsByStation ? Object.keys(state.model.modelsByStation)[0] : selected;
  const out = forecastForStation(station);
  state.forecastRows = out.forecast;
  applyThresholdsToForecast(state.forecastRows);
  renderTable(state.forecastRows, ['step','time','station','forecast_p_mw','model_used','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong'], 1000);
  const steps = out.forecast.length;
  const actualHist = out.series.slice(-Math.min(steps*2, 240)).map((r,i,arr)=> i < arr.length-steps ? r.p : NaN);
  const predHist = out.series.slice(-Math.min(steps*2, 240)).map((r,i,arr)=> i >= arr.length-steps ? r.p : NaN);
  const times = out.series.slice(-Math.min(steps*2, 240)).map(r=>r.time);
  drawSeries(times, actualHist, predHist, 'Lịch sử gần nhất', 'Dự báo LV7');
  updateForecastMetrics(state.forecastRows);
  renderForecastExplain(out);
  $('exportForecastBtn').disabled = false;
  log(`Đã dự báo LV7 cho ${out.station}: ${out.forecast.length} bước, mô hình ${out.strategy?.label || out.strategy?.method || 'auto'}, mỗi bước ${out.interval} phút.`);
}

function forecastAllStations() {
  if (!state.model) throw new Error('Chưa có model.');
  const stations = state.model.modelsByStation ? Object.keys(state.model.modelsByStation) : [...new Set(state.rows.map(r=>r.station||'ALL'))];
  let all=[];
  for (const st of stations) {
    try { all = all.concat(forecastForStation(st).forecast); }
    catch(e) { log('LV7 bỏ qua dự báo ' + st + ': ' + e.message); }
  }
  state.forecastRows = all;
  applyThresholdsToForecast(state.forecastRows);
  renderTable(all, ['step','time','station','forecast_p_mw','model_used','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong'], 3000);
  updateForecastMetrics(all);
  $('exportForecastBtn').disabled = false;
  log(`LV7 đã dự báo tất cả chỉ danh: ${all.length} dòng.`);
}

function renderForecastExplain(out) {
  const box = $('forecastExplainBox'); if (!box || !out?.forecast?.length) return;
  const first = out.forecast[0];
  const strategy = out.strategy || {};
  const w = strategy.weights;
  const wText = w ? `<span class="pill">w GBDT ${formatNum(w.gbdt,2)}</span><span class="pill">w Similar ${formatNum(w.similar_day,2)}</span><span class="pill">w Tuần ${formatNum(w.same_hour_last_week,2)}</span><span class="pill">w Trend ${formatNum(w.trend,2)}</span>` : '';
  box.innerHTML = `<b>Giải thích LV7:</b> <span class="pill ok">Mô hình dùng: ${escapeHtml(first.model_used || '')}</span>${wText}<span class="pill">GBDT ${first.gbdt}</span><span class="pill">Similar ${first.similar_day}</span><span class="pill">Tuần trước ${first.same_hour_last_week}</span><span class="pill">Xu hướng ${first.trend}</span><span class="pill">Bù sai số ${first.bias}</span><span class="pill ok">Dự báo cuối ${first.forecast_p_mw} MW</span>`;
}

function renderModelInfo() {
  if (!state.model) { $('modelInfo').innerHTML = ''; return; }
  const m = state.model;
  if (m.modelsByStation) {
    const n = Object.keys(m.modelsByStation).length;
    const ns = Object.keys(m.strategyByStation || {}).length;
    $('modelInfo').innerHTML = `<span class="pill modeBadge">Model bundle ${escapeHtml(m.appVersion || 'LV7')}</span><span class="pill ok">Xuất từ: ${escapeHtml(m.exportSource || 'chưa có chữ ký')}</span><span class="pill">${n} model trạm/lộ</span><span class="pill">${ns} chiến lược tự chọn</span><span class="pill">Created: ${escapeHtml(m.createdAt||'')}</span><span class="pill">Exported: ${escapeHtml(m.exportedAt||'')}</span>`;
    $('forecastBtn').disabled = false; if ($('forecastAllBtn')) $('forecastAllBtn').disabled = false;
    return;
  }
  const st = m.station || $('stationSelect')?.value || 'ALL';
  const sg = m.strategyByStation?.[st] || m.v7Strategy;
  $('modelInfo').innerHTML = `
    <span class="pill modeBadge">Model ${escapeHtml(m.appVersion || 'LV7')}</span>
    <span class="pill ok">Xuất từ: ${escapeHtml(m.exportSource || 'chưa có chữ ký')}</span>
    <span class="pill">Loại: ${escapeHtml(m.type || '')}</span>
    <span class="pill">Số cây: ${m.trees?.length || 0}</span>
    <span class="pill">Interval: ${m.intervalMinutes} phút</span>
    <span class="pill">Station: ${escapeHtml(m.station || 'ALL')}</span>
    <span class="pill">MAPE GBDT: ${formatNum(m.metrics?.validation?.mape,2)}%</span>
    ${sg ? `<span class="pill ok">Chiến lược: ${escapeHtml(sg.label || sg.method)}</span>` : ''}`;
}

function exportModel() {
  if (!state.model) return;
  const payload = {
    ...state.model,
    scadaModelPackage: 'SCADA_LOAD_FORECAST_OPERATIONAL_MODEL',
    appVersion: 'LV7.6',
    exportSource: 'MUC_8_EXPORT_MODEL',
    exportSection: '8',
    allowedImportSection: '10',
    modelKind: 'OPERATIONAL_FORECAST_MODEL',
    exportedAt: new Date().toISOString(),
    importRules: {
      onlyImportInSection10: true,
      requiredExportSource: 'MUC_8_EXPORT_MODEL',
      note: 'File này được xuất từ Mục 8 và là file duy nhất được phép nạp ở Mục 10.'
    },
    lv7Config: collectFullConfig(),
    lv7State:{strategyByStation: state.lv7.strategyByStation || {}, savedAt:new Date().toISOString()}
  };
  saveTextFile('model_gbdt_lv7_6_operational.json', JSON.stringify(payload, null, 2), 'application/json');
  log('Đã xuất model_gbdt_lv7_6_operational.json từ Mục 8. Đây là file hợp lệ để nạp ở Mục 10 trong mạng SCADA.');
}

function exportForecast() {
  if (!state.forecastRows.length) return;
  const rows = state.forecastRows.map(r => {
    const d = parseDesignation(r.station);
    return {...r, don_vi:d.unit||'', tram:d.substation||'', lo:d.feeder||'', noi_vong:d.hasRing?`${d.ringUnit}/${d.ringSubstation}/${d.ringFeeder}`:''};
  });
  saveTextFile('forecast_lv7_6.csv', toCSV(rows, ['step','time','don_vi','tram','lo','noi_vong','station','forecast_p_mw','model_used','strategy_method','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong']), 'text/csv;charset=utf-8');
  log('Đã xuất forecast_lv7_6.csv theo cấu trúc LV7 có mô hình được chọn.');
}

function exportV7ModelReport() {
  const rows = state.lv7.modelReports || [];
  if (!rows.length) { log('Chưa có báo cáo so sánh mô hình LV7.'); return; }
  saveTextFile('model_compare_lv7_6.csv', toCSV(formatMetricRowsV7(rows), ['station','method','ten_mo_hinh','n','MAE','MAPE_pct','RMSE','xep_hang','khuyen_nghi','w_gbdt','w_similar','w_week','w_trend']), 'text/csv;charset=utf-8');
  log(`Đã xuất model_compare_lv7_6.csv gồm ${rows.length} dòng.`);
}

function exportV7HourlyReport() {
  const rows = state.lv7.hourlyReports || [];
  if (!rows.length) { log('Chưa có báo cáo sai số theo giờ LV7.'); return; }
  const out = rows.map(r => ({station:r.station, method:r.method, ten_mo_hinh:r.ten_mo_hinh, hour:r.hour, n:r.n, MAE:formatNum(r.mae,3), MAPE_pct:formatNum(r.mape,2), RMSE:formatNum(r.rmse,3)}));
  saveTextFile('hourly_error_lv7_6.csv', toCSV(out, ['station','method','ten_mo_hinh','hour','n','MAE','MAPE_pct','RMSE']), 'text/csv;charset=utf-8');
  log(`Đã xuất hourly_error_lv7_6.csv gồm ${rows.length} dòng.`);
}


function weightsFromReportV72(row) {
  if (!row) return null;
  if (row.method === 'auto_blend' || row.method === 'hybrid_current') {
    const w = {
      gbdt: Number(row.w_gbdt),
      similar_day: Number(row.w_similar),
      same_hour_last_week: Number(row.w_week),
      trend: Number(row.w_trend)
    };
    const valid = Object.values(w).some(v => Number.isFinite(v) && v > 0);
    return valid ? w : null;
  }
  return null;
}

function makeStrategyV72(station, method, label, validation=null, weights=null, intervalMinutes=null) {
  return {
    version:'LV7.6',
    selectedAt:new Date().toISOString(),
    appliedFromSection:'9',
    station,
    method,
    label: label || LV7_METHOD_LABELS[method] || method,
    validation: validation || null,
    weights: weights || null,
    intervalMinutes: intervalMinutes || detectIntervalMinutes(getStationRowsForV7(station)) || null
  };
}

function setApplyStatusV72(message, ok=true) {
  const box = $('lv7ApplyStatus');
  if (!box) return;
  box.innerHTML = `<span class="pill ${ok?'ok':'warn'}">${escapeHtml(message)}</span><span class="pill">Bước tiếp theo: Mục 8 → Xuất model vận hành cho mục 10</span>`;
}

function applyStrategyToModelV72(station, strategy) {
  ensureStrategyContainerV7();
  const st = station || $('stationSelect')?.value || '__ALL__';
  state.model.strategyByStation[st] = strategy;
  state.lv7.strategyByStation[st] = strategy;
  renderModelInfo();
  setApplyStatusV72(`Đã áp dụng chiến lược ${strategy.label || strategy.method} cho ${st}`);
  log(`LV7.6 đã áp dụng chiến lược cho ${st}: ${strategy.label || strategy.method}. Hãy xuất model vận hành ở Mục 8 để dùng ở Mục 10.`);
  return strategy;
}

function applySelectedStrategyV72() {
  const st = $('stationSelect')?.value || '__ALL__';
  const method = $('forecastStrategy')?.value || 'auto';
  if (method === 'auto') {
    return selectBestModelV7(state.lv7.lastResult || evaluateModelsV7());
  }
  // Nếu người dùng chọn thủ công, vẫn lấy validation của dòng tương ứng nếu đã có bảng so sánh.
  const rows = (state.lv7.modelReports || []).filter(r => r.station === st);
  const row = rows.find(r => r.method === method);
  const strategy = makeStrategyV72(
    st,
    method,
    LV7_METHOD_LABELS[method] || method,
    row ? {mae:row.mae, mape:row.mape, rmse:row.rmse, n:row.n} : null,
    weightsFromReportV72(row),
    row?.interval_minutes
  );
  return applyStrategyToModelV72(st, strategy);
}

function applyRecommendationsFromReportsV72() {
  const reports = state.lv7.modelReports || [];
  if (!reports.length) throw new Error('Chưa có bảng so sánh. Hãy bấm “So sánh mô hình” hoặc “Huấn luyện + áp dụng tốt nhất tất cả” trước.');
  ensureStrategyContainerV7();
  const byStation = {};
  for (const r of reports) {
    if (!r.station) continue;
    const score = Number.isFinite(r.mape) ? r.mape : Infinity;
    if (!byStation[r.station] || score < (Number.isFinite(byStation[r.station].mape) ? byStation[r.station].mape : Infinity)) byStation[r.station] = r;
  }
  const applied = [];
  for (const [st, r] of Object.entries(byStation)) {
    const strategy = makeStrategyV72(
      st,
      r.method,
      r.ten_mo_hinh || LV7_METHOD_LABELS[r.method] || r.method,
      {mae:r.mae, mape:r.mape, rmse:r.rmse, n:r.n},
      weightsFromReportV72(r),
      r.interval_minutes
    );
    state.model.strategyByStation[st] = strategy;
    state.lv7.strategyByStation[st] = strategy;
    applied.push(st);
  }
  renderModelInfo();
  setApplyStatusV72(`Đã áp dụng gợi ý tốt nhất cho ${applied.length} chỉ danh`);
  log(`LV7.6 đã áp dụng gợi ý tốt nhất từ bảng so sánh cho ${applied.length} chỉ danh. Hãy xuất model vận hành ở Mục 8 để dùng ở Mục 10.`);
  return applied;
}

function lv7BindEvents() {
  $('evaluateModelsBtn')?.addEventListener('click', () => { try { evaluateModelsV7(); } catch(e) { log('Lỗi so sánh mô hình LV7: ' + e.message); } });
  $('selectBestModelBtn')?.addEventListener('click', () => { try { selectBestModelV7(); } catch(e) { log('Lỗi áp dụng gợi ý LV7.6: ' + e.message); } });
  $('applySelectedStrategyBtn')?.addEventListener('click', () => { try { applySelectedStrategyV72(); } catch(e) { log('Lỗi áp dụng chiến lược đang chọn LV7.6: ' + e.message); } });
  $('applyAllReportedStrategiesBtn')?.addEventListener('click', () => { try { applyRecommendationsFromReportsV72(); } catch(e) { log('Lỗi áp dụng gợi ý từ bảng LV7.6: ' + e.message); } });
  $('trainSelectBestAllBtn')?.addEventListener('click', () => { trainSelectBestAllStationsV7().catch(e => log('Lỗi huấn luyện + áp dụng tốt nhất LV7.6: ' + e.message)); });
  $('exportModelReportBtn')?.addEventListener('click', exportV7ModelReport);
  $('exportHourlyReportBtn')?.addEventListener('click', exportV7HourlyReport);
  $('forecastStrategy')?.addEventListener('change', saveSettingsToLocal);
  $('forecastStrategyScada')?.addEventListener('change', saveSettingsToLocal);
}

setTimeout(() => {
  try {
    lv7BindEvents();
    state.lv5.version = 'LV7.6';
    if ($('versionInfo')) $('versionInfo').innerHTML = '<span class="pill modeBadge">LV7.6</span><span class="pill ok">Có hướng dẫn luồng công việc</span>';
    log('Sẵn sàng LV7.6: có hướng dẫn luồng công việc trong HTML, so sánh mô hình, áp dụng chiến lược vào model, rồi xuất model vận hành ở Mục 8 để nạp Mục 10.');
  } catch(e) { log('Lỗi khởi tạo LV7: ' + e.message); }
}, 0);
// ====================== END LV7 EXTENSIONS ======================


// ======================== LV7.6 DATA QUALITY RAM FOCUS OVERRIDE ========================
// Mục tiêu:
// - Mục 5 kiểm tra chất lượng trên toàn bộ dữ liệu đang ở RAM.
// - Sau khi kiểm tra, Mục 4 tự chuyển sang chế độ "chỉ hiện dòng lỗi" để người vận hành theo dõi/sửa.
// - Mục 6 dùng chính phạm vi lỗi đang hiển thị ở Mục 4 nếu người dùng chưa chọn dòng cụ thể.
// - Không còn phụ thuộc việc dòng lỗi có đang nằm trên trang hiện tại hay không.

function ensureEditorFocusStateLV74() {
  if (!state.editor) state.editor = {};
  if (!state.editor.focus) state.editor.focus = {mode:'none', label:'', indices:[]};
  if (!Array.isArray(state.editor.focus.indices)) state.editor.focus.indices = [];
  return state.editor.focus;
}

function getEditorFocusSetLV74() {
  const focus = ensureEditorFocusStateLV74();
  if (!focus || focus.mode === 'none') return null;
  const arr = (focus.indices || []).filter(i => Number.isInteger(i) && i >= 0 && i < state.rawRows.length);
  return arr.length ? new Set(arr) : new Set();
}

function setEditorFocusRowsLV74(indices, label='dòng cần xử lý', mode='custom', options={}) {
  const uniq = [...new Set((indices || []).filter(i => Number.isInteger(i) && i >= 0 && i < state.rawRows.length))].sort((a,b)=>a-b);
  state.editor.focus = {mode, label, indices: uniq};
  state.editor.page = 1;
  if ($('editorPage')) $('editorPage').value = 1;
  if (options.select) {
    state.editor.selected.clear();
    uniq.forEach(i => state.editor.selected.add(i));
  }
  renderEditorTable();
  const msg = uniq.length ? `Mục 4 đang chỉ hiển thị ${uniq.length} ${label}.` : `Không có ${label} để hiển thị tại Mục 4.`;
  if (options.log !== false) log(msg);
  return uniq;
}

function clearEditorFocusLV74(options={}) {
  state.editor.focus = {mode:'none', label:'', indices:[]};
  state.editor.page = 1;
  if ($('editorPage')) $('editorPage').value = 1;
  renderEditorTable();
  if (options.log !== false) log('Đã bỏ chế độ chỉ hiện dòng lỗi/sự kiện. Mục 4 trở lại bộ lọc thường.');
}

function editorFocusLabelHtmlLV74() {
  const focus = ensureEditorFocusStateLV74();
  if (!focus || focus.mode === 'none') return '';
  const n = (focus.indices || []).filter(i => i >= 0 && i < state.rawRows.length).length;
  const cls = n ? 'warn' : 'bad';
  return `<span class="pill ${cls}">Đang chỉ hiện ${n} ${escapeHtml(focus.label || 'dòng')}</span>`;
}

function updateEditorFocusNoticeLV74() {
  const box = $('editorFocusNotice');
  if (!box) return;
  const focus = ensureEditorFocusStateLV74();
  if (!focus || focus.mode === 'none') {
    box.innerHTML = '<span class="pill ok">Bảng đang hiển thị theo bộ lọc thường</span><span class="pill">Mục 5/6 có thể ép bảng chỉ hiện dòng lỗi để xử lý</span>';
    return;
  }
  const n = (focus.indices || []).filter(i => i >= 0 && i < state.rawRows.length).length;
  box.innerHTML = `<span class="pill warn">Đang khóa bảng theo ${escapeHtml(focus.label || 'dòng cần xử lý')}: ${n} dòng</span><span class="pill">Mục 6 sẽ ưu tiên xử lý đúng các dòng này nếu chưa chọn dòng khác</span>`;
}

function getEditorFilteredIndices() {
  const focusSet = getEditorFocusSetLV74();
  if (focusSet) {
    const q = norm(state.editor.query);
    const arr = [...focusSet].sort((a,b)=>a-b).filter(i => {
      const raw = state.rawRows[i];
      if (!raw) return false;
      if (!q) return true;
      const joined = norm(state.headers.map(h => raw[h]).join(' '));
      return joined.includes(q);
    });
    return arr;
  }
  const q = norm(state.editor.query);
  const filter = state.editor.filter;
  const indices = [];
  for (let i=0; i<state.rawRows.length; i++) {
    const raw = state.rawRows[i];
    if (q) {
      const joined = norm(state.headers.map(h => raw[h]).join(' '));
      if (!joined.includes(q)) continue;
    }
    if (!editorDateFilterPass(raw)) continue;
    if (filter === 'invalid' && validateRawRow(raw).length === 0) continue;
    if (filter === 'abnormal' && !rawRowIsAbnormal(raw)) continue;
    indices.push(i);
  }
  return indices;
}

function renderEditorStatus(extra='') {
  const total = state.rawRows.length;
  const invalid = state.rawRows.reduce((n, r) => n + (validateRawRow(r).length ? 1 : 0), 0);
  const abnormal = state.rawRows.reduce((n, r) => n + (rawRowIsAbnormal(r) ? 1 : 0), 0);
  const selected = state.editor.selected.size;
  const dirty = state.editor.dirty;
  const status = [];
  status.push(`<span class="pill">${total} dòng thô</span>`);
  status.push(`<span class="pill ${invalid?'bad':'ok'}">${invalid} dòng lỗi</span>`);
  status.push(`<span class="pill ${abnormal?'warn':''}">${abnormal} dòng bất thường</span>`);
  const focusPill = editorFocusLabelHtmlLV74();
  if (focusPill) status.push(focusPill);
  if (selected) status.push(`<span class="pill warn">đã chọn ${selected}</span>`);
  status.push(`<span class="pill ${dirty?'warn':'ok'}">${dirty?'có thay đổi chưa áp dụng':'đã áp dụng'}</span>`);
  if (extra) status.push(`<span class="pill">${escapeHtml(extra)}</span>`);
  const el = $('editorStatus');
  if (el) el.innerHTML = status.join('');
  updateEditorFocusNoticeLV74();
}

function resetDateFilter() {
  state.editor.dateMode = 'all'; state.editor.dateSingle = ''; state.editor.dateMulti = ''; state.editor.dateFrom = ''; state.editor.dateTo = '';
  if ($('dateFilterMode')) $('dateFilterMode').value = 'all';
  if ($('dateFilterSingle')) $('dateFilterSingle').value = '';
  if ($('dateFilterMulti')) $('dateFilterMulti').value = '';
  if ($('dateFilterFrom')) $('dateFilterFrom').value = '';
  if ($('dateFilterTo')) $('dateFilterTo').value = '';
  state.editor.page = 1; if ($('editorPage')) $('editorPage').value = 1;
  clearEditorFocusLV74({log:false});
  log('Đã bỏ bộ lọc ngày và chế độ chỉ hiện dòng lỗi/sự kiện.');
}

function focusEditorOnQualityIssuesLV74(select=false) {
  const rows = [...new Set((state.qualityIssues || []).map(x => x.rowIndex).filter(Number.isInteger))];
  return setEditorFocusRowsLV74(rows, 'dòng lỗi chất lượng từ Mục 5', 'quality', {select});
}

function focusEditorOnMissingTimeIssuesLV74(select=false) {
  const rows = [...new Set((state.qualityIssues || []).filter(x => String(x.type||'').includes('Mất mốc')).map(x => x.rowIndex).filter(Number.isInteger))];
  return setEditorFocusRowsLV74(rows, 'dòng liền sau khoảng mất mốc từ Mục 5', 'missing_time', {select});
}

function focusEditorOnOperationEventsLV74(select=false) {
  const rows = [...new Set((state.operationEvents || []).map(x => x.rowIndex).filter(Number.isInteger))];
  return setEditorFocusRowsLV74(rows, 'dòng sự kiện vận hành từ Mục 7', 'operation', {select});
}

function runQualityCheck() {
  const issues = runQualityCheckBaseLV73();
  const rows = [...new Set((issues || []).map(x => x.rowIndex).filter(Number.isInteger))];
  setEditorFocusRowsLV74(rows, 'dòng lỗi chất lượng từ Mục 5', 'quality', {select:false, log:false});
  const sum = $('qualitySummary');
  if (sum) sum.innerHTML += rows.length
    ? `<span class="pill warn">Mục 4 đã tự lọc còn ${rows.length} dòng lỗi</span>`
    : '<span class="pill ok">Mục 4 không cần lọc dòng lỗi</span>';
  log(`LV7.6: Kiểm tra chất lượng hoàn tất trên dữ liệu trong RAM; Mục 4 đã chuyển sang chỉ hiện ${rows.length} dòng lỗi.`);
  return issues;
}

function selectQualityRows() {
  const rows = focusEditorOnQualityIssuesLV74(true);
  log(`Đã chọn ${rows.length} dòng lỗi và ép Mục 4 chỉ hiển thị các dòng này.`);
}

function markQualityAbnormal() {
  if (!state.headers.length) return;
  const abnormalCol = ensureMappedColumn('abnormal', 'Bất thường');
  addProcessColumns();
  let n=0;
  const target = getEditorFocusSetLV74() || new Set((state.qualityIssues || []).map(x => x.rowIndex).filter(Number.isInteger));
  for (const idx of target) {
    if (Number.isInteger(idx) && state.rawRows[idx]) {
      const issueTypes = (state.qualityIssues || []).filter(x => x.rowIndex === idx).map(x => x.type).join('; ');
      state.rawRows[idx][abnormalCol] = '1';
      state.rawRows[idx]['ghi_chu_xu_ly'] = `${state.rawRows[idx]['ghi_chu_xu_ly'] || ''} ${issueTypes || 'LV7.6 danh dau bat thuong tu Muc 5'}`.trim();
      n++;
    }
  }
  markEditorDirty(true);
  renderEditorTable();
  log(`Đã đánh dấu bất thường ${n} dòng lỗi đang hiển thị/đang có trong báo cáo chất lượng.`);
}

function getInterpolationTargetIndexSet(options={}) {
  const selected = [...(state.editor?.selected || new Set())].filter(i => i >= 0 && i < state.rawRows.length);
  if (selected.length) return new Set(selected);
  if (options.selectedOnly) return new Set();
  const focusSet = getEditorFocusSetLV74();
  if (focusSet) return focusSet;
  return new Set(getEditorFilteredIndices());
}

function interpolationFocusLimitSetLV74() {
  const focus = ensureEditorFocusStateLV74();
  if (!focus || focus.mode === 'none') return null;
  const arr = (focus.indices || []).filter(i => Number.isInteger(i) && i >= 0 && i < state.rawRows.length);
  return arr.length ? new Set(arr) : new Set();
}

function interpolateMissingTimestamps() {
  if (!state.headers.length) throw new Error('Chưa có dữ liệu.');
  const m = readMap();
  if (!m.time || !m.p) throw new Error('Cần ánh xạ cột thời gian và P.');
  const stationFilter = $('stationSelect')?.value || '__ALL__';
  const scope = $('interpScope')?.value || 'current';
  const method = $('interpMethod')?.value || 'linear';
  const maxGap = Math.max(1, Math.floor(parseNumber($('interpMaxGap')?.value) || 12));
  const syntheticCol = 'du_lieu_noi_suy', noteCol='ghi_chu_xu_ly';
  addProcessColumns();

  // Mục 6 ưu tiên dùng dòng đã chọn. Nếu chưa chọn, dùng phạm vi lỗi đang khóa ở Mục 4. Nếu không có khóa, dùng bộ lọc thường.
  const fixedBadP = fillInvalidP({silent:true});
  const focusLimit = interpolationFocusLimitSetLV74();

  const groups = groupRowsByStationRaw();
  let added=0, skipped=0, skippedByDate=0, skippedByFocus=0;
  for (const [station, arr] of groups) {
    if (scope === 'current' && stationFilter !== '__ALL__' && station !== stationFilter) continue;
    const expected = $('expectedInterval')?.value !== 'auto' ? parseNumber($('expectedInterval')?.value) : detectIntervalMinutes(arr.map(x => ({time:x.time,p:x.p})));
    const interval = expected || 60;
    for (let i=1; i<arr.length; i++) {
      const prev = arr[i-1], next = arr[i];
      if (focusLimit && !(focusLimit.has(prev.rawIndex) || focusLimit.has(next.rawIndex))) { continue; }
      const gap = (next.time - prev.time)/60000;
      const miss = Math.round(gap/interval) - 1;
      if (miss <= 0) continue;
      if (miss > maxGap) { skipped += miss; continue; }
      for (let k=1; k<=miss; k++) {
        const t = new Date(prev.time.getTime() + k*interval*60000);
        if (!editorDateFilterPassTime(t) && !focusLimit) { skippedByDate++; continue; }
        const row = {};
        state.headers.forEach(h => row[h] = '');
        row[m.time] = fmtTime(t);
        row[m.p] = formatNum(inferPForMissing(method, arr, i-1, i, t, interval), 3);
        if (m.station) row[m.station] = station;
        if (m.temp) {
          const a = parseNumber(prev.raw[m.temp]), b = parseNumber(next.raw[m.temp]);
          if (Number.isFinite(a) && Number.isFinite(b)) row[m.temp] = formatNum(a + (b-a)*(k/(miss+1)), 1);
          else if (Number.isFinite(a)) row[m.temp] = formatNum(a,1);
        }
        if (m.rain) row[m.rain] = '0';
        if (m.holiday) row[m.holiday] = holidayByRules(t) ? '1' : '0';
        if (m.abnormal) row[m.abnormal] = '0';
        if (m.outage) row[m.outage] = '0';
        if (m.transfer) row[m.transfer] = '0';
        row[syntheticCol] = '1';
        row[noteCol] = `LV7.6 noi suy moc thieu ${method}${focusLimit ? ' theo pham vi dong loi Muc 5' : ''}`;
        state.rawRows.push(row); added++;
      }
    }
  }
  state.rawRows.sort((a,b) => {
    const ta = parseTime(a[m.time]) || 0, tb = parseTime(b[m.time]) || 0;
    if (ta - tb) return ta - tb;
    const sa = m.station ? String(a[m.station] || '') : '';
    const sb = m.station ? String(b[m.station] || '') : '';
    return sa.localeCompare(sb, 'vi');
  });
  // Sau khi sort, chỉ số dòng thay đổi nên bỏ khóa tập lỗi cũ để tránh trỏ sai dòng.
  if (added) state.editor.focus = {mode:'none', label:'', indices:[]};
  normalizeRows(); applyDataInfo(); renderEditorTable(); previewData(); markEditorDirty(true);
  const parts = [
    `<span class="pill ok">Đã nội suy ${fixedBadP} giá trị P trống/lỗi/P=0/P thấp</span>`,
    `<span class="pill ok">Đã bổ sung ${added} mốc thời gian thiếu</span>`,
    `<span class="pill ${skipped?'warn':''}">Bỏ qua ${skipped} mốc do gap quá lớn</span>`
  ];
  if (focusLimit) parts.push(`<span class="pill warn">Đã giới hạn bổ sung mốc theo ${focusLimit.size} dòng lỗi đang khóa ở Mục 4</span>`);
  if (skippedByDate) parts.push(`<span class="pill warn">Bỏ qua ${skippedByDate} mốc ngoài bộ lọc ngày</span>`);
  if (added) parts.push('<span class="pill">Đã bỏ khóa dòng lỗi vì chỉ số dòng thay đổi sau khi thêm mốc</span>');
  $('interpolationInfo').innerHTML = parts.join('');
  log(`LV7.6 nội suy/bổ sung từ RAM: sửa ${fixedBadP} P, thêm ${added} mốc, bỏ qua ${skipped}.`);
}

function selectVisibleEditorRows() {
  const indices = getCurrentPageEditorIndices();
  for (const idx of indices) state.editor.selected.add(idx);
  renderEditorTable();
  log(`Đã chọn ${indices.length} dòng đang hiển thị${ensureEditorFocusStateLV74().mode !== 'none' ? ' trong chế độ khóa dòng lỗi' : ''}.`);
}

function selectFilteredEditorRows() {
  const indices = getEditorFilteredIndices();
  for (const idx of indices) state.editor.selected.add(idx);
  renderEditorTable();
  log(`Đã chọn ${indices.length} dòng theo phạm vi đang hiển thị/lọc tại Mục 4.`);
}

function clearEditorSelection() {
  state.editor.selected.clear();
  renderEditorTable();
  log('Đã bỏ chọn tất cả dòng. Chế độ chỉ hiện dòng lỗi vẫn giữ nguyên nếu đang bật.');
}

function lv74BindEvents() {
  $('clearEditorFocusBtn')?.addEventListener('click', () => clearEditorFocusLV74());
  $('showQualityInEditorBtn')?.addEventListener('click', () => focusEditorOnQualityIssuesLV74(false));
  $('showMissingInEditorBtn')?.addEventListener('click', () => focusEditorOnMissingTimeIssuesLV74(false));
  $('selectQualityRowsBtn')?.addEventListener('dblclick', () => focusEditorOnQualityIssuesLV74(true));
}

setTimeout(() => {
  try {
    ensureEditorFocusStateLV74();
    lv74BindEvents();
    if ($('versionInfo')) $('versionInfo').innerHTML = '<span class="pill modeBadge">LV7.6</span><span class="pill ok">Mục 5/6 xử lý theo dữ liệu RAM và ép bảng Mục 4 chỉ hiện dòng lỗi</span>';
    updateEditorFocusNoticeLV74();
    log('Sẵn sàng LV7.6: Mục 5 kiểm tra dữ liệu trong RAM và tự ép Mục 4 chỉ hiện dòng lỗi; Mục 6 ưu tiên xử lý dòng đang chọn hoặc phạm vi lỗi đang hiển thị.');
  } catch(e) { log('Lỗi khởi tạo LV7.6: ' + e.message); }
}, 0);
// ====================== END LV7.6 EXTENSIONS ======================

// ====================== LV7.6 EXTENSIONS ======================
// LV7.6: đồng bộ chế độ rà soát lỗi giữa Mục 4/5/6/7,
// hiển thị cả dòng liền trước và liền sau mốc thiếu, và thêm thu gọn/mở từng khu vực.
function getMissingTimeNeighborRowsLV75() {
  const issues = (state.qualityIssues || []).filter(x => String(x.type || '').includes('Mất mốc'));
  const out = new Set();
  if (!issues.length) return [];
  let groups;
  try { groups = groupRowsByStationRaw(); } catch(_) { groups = new Map(); }
  for (const it of issues) {
    const nextIdx = Number.isInteger(it.rowIndex) ? it.rowIndex : null;
    if (nextIdx != null && nextIdx >= 0 && nextIdx < state.rawRows.length) out.add(nextIdx);
    const t = it.time ? parseTime(it.time) : null;
    const station = String(it.station || '').trim();
    let prevIdx = null;
    if (t && groups && groups.size) {
      const arr = groups.get(station) || [...groups.values()].find(a => a.some(r => r.rawIndex === nextIdx)) || [];
      let best = null;
      for (const r of arr) {
        if (nextIdx != null && r.rawIndex === nextIdx) continue;
        if (r.time && r.time < t && (!best || r.time > best.time)) best = r;
      }
      if (best) prevIdx = best.rawIndex;
    }
    // Dự phòng: nếu không tìm được theo thời gian, lấy dòng raw liền trước có cùng chỉ danh.
    if (prevIdx == null && nextIdx != null) {
      const m = readMap();
      const nextRaw = state.rawRows[nextIdx];
      const stNext = m.station ? String(nextRaw?.[m.station] ?? '').trim() : '';
      for (let j = nextIdx - 1; j >= 0; j--) {
        const raw = state.rawRows[j];
        const st = m.station ? String(raw?.[m.station] ?? '').trim() : '';
        if (!m.station || st === stNext) { prevIdx = j; break; }
      }
    }
    if (prevIdx != null && prevIdx >= 0 && prevIdx < state.rawRows.length) out.add(prevIdx);
  }
  return [...out].sort((a,b)=>a-b);
}

function focusEditorOnMissingTimeIssuesLV75(select=false) {
  const rows = getMissingTimeNeighborRowsLV75();
  return setEditorFocusRowsLV74(rows, 'dòng liền trước và liền sau mốc thiếu từ Mục 5', 'missing_time_neighbors', {select});
}

// Ghi đè hàm LV7.4 để nút cũ dùng logic mới.
focusEditorOnMissingTimeIssuesLV74 = focusEditorOnMissingTimeIssuesLV75;

function focusEditorOnOperationEventsLV75(select=false) {
  const rows = [...new Set((state.operationEvents || []).map(x => x.rowIndex).filter(Number.isInteger))];
  return setEditorFocusRowsLV74(rows, 'dòng sự kiện vận hành từ Mục 7', 'operation', {select});
}

function renderOperationFocusHintLV75(count) {
  const sum = $('operationSummary');
  if (!sum) return;
  const extra = count
    ? `<span class="pill warn">Mục 4 đang chỉ hiện ${count} dòng sự kiện</span>`
    : '<span class="pill ok">Mục 4 không có dòng sự kiện cần khóa</span>';
  if (!sum.innerHTML.includes('Mục 4 đang chỉ hiện') && !sum.innerHTML.includes('Mục 4 không có dòng sự kiện')) {
    sum.innerHTML += extra;
  }
}

const analyzeOperationEventsBaseLV75 = analyzeOperationEvents;
analyzeOperationEvents = function() {
  const events = analyzeOperationEventsBaseLV75();
  const rows = focusEditorOnOperationEventsLV75(false);
  renderOperationFocusHintLV75(rows.length);
  log(`LV7.6: Phân tích Mục 7 xong; Mục 4 đã chuyển sang chế độ rà soát ${rows.length} dòng sự kiện vận hành.`);
  return events;
};

selectOperationRows = function() {
  const rows = focusEditorOnOperationEventsLV75(true);
  log(`Đã chọn ${rows.length} dòng sự kiện và ép Mục 4 chỉ hiển thị các dòng này.`);
};

function setupSectionTogglesLV75() {
  const cards = [...document.querySelectorAll('main .card')];
  cards.forEach((card, idx) => {
    const h2 = card.querySelector(':scope > h2');
    if (!h2 || h2.dataset.lv75ToggleReady === '1') return;
    const title = h2.textContent.trim();
    h2.textContent = '';
    const span = document.createElement('span');
    span.className = 'sectionTitleText';
    span.textContent = title;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'sectionToggleBtn secondary';
    const key = 'scada_lv75_card_collapsed_' + idx + '_' + title.slice(0, 24);
    const apply = () => { btn.textContent = card.classList.contains('section-collapsed') ? 'Mở' : 'Thu gọn'; };
    if (localStorage.getItem(key) === '1') card.classList.add('section-collapsed');
    apply();
    btn.addEventListener('click', ev => {
      ev.preventDefault(); ev.stopPropagation();
      card.classList.toggle('section-collapsed');
      localStorage.setItem(key, card.classList.contains('section-collapsed') ? '1' : '0');
      apply();
    });
    h2.appendChild(span);
    h2.appendChild(btn);
    h2.dataset.lv75ToggleReady = '1';
  });
}

function injectGlobalCollapseBarLV75() {
  if (document.getElementById('lv75CollapseBar')) return;
  const section = document.querySelector('main > section.grid');
  if (!section) return;
  const bar = document.createElement('div');
  bar.id = 'lv75CollapseBar';
  bar.className = 'card span12';
  bar.innerHTML = `<h2><span class="sectionTitleText">0.2) Điều khiển thu gọn giao diện</span></h2>
    <div class="row">
      <button class="secondary" id="collapseAllSectionsBtn">Thu gọn tất cả khu vực</button>
      <button class="secondary" id="expandAllSectionsBtn">Mở tất cả khu vực</button>
      <button class="secondary" id="focusEditorSectionBtn">Chỉ mở Mục 4/5/6/7 để rà soát lỗi</button>
    </div>
    <div class="compactNote">Mỗi khu vực đều có nút <b>Thu gọn/Mở</b> ở tiêu đề. Dùng khi cần nhường không gian cho bảng hiệu chỉnh Mục 4 hoặc các bảng báo cáo lỗi/sự kiện.</div>`;
  const first = section.firstElementChild;
  if (first && first.nextElementSibling) section.insertBefore(bar, first.nextElementSibling);
  else section.prepend(bar);
  const setCollapsedByPredicate = pred => {
    [...document.querySelectorAll('main .card')].forEach(card => {
      if (card.id === 'lv75CollapseBar') return;
      const title = card.querySelector(':scope > h2 .sectionTitleText, :scope > h2')?.textContent || '';
      card.classList.toggle('section-collapsed', !!pred(title));
      const btn = card.querySelector(':scope > h2 .sectionToggleBtn');
      if (btn) btn.textContent = card.classList.contains('section-collapsed') ? 'Mở' : 'Thu gọn';
    });
  };
  setTimeout(() => {
    $('collapseAllSectionsBtn')?.addEventListener('click', () => setCollapsedByPredicate(() => true));
    $('expandAllSectionsBtn')?.addEventListener('click', () => setCollapsedByPredicate(() => false));
    $('focusEditorSectionBtn')?.addEventListener('click', () => setCollapsedByPredicate(title => !/^(4\)|5\)|6\)|7\)|0\.2\))/.test(title.trim())));
  }, 0);
}

function lv75BindEvents() {
  $('showMissingInEditorBtn')?.addEventListener('click', () => focusEditorOnMissingTimeIssuesLV75(false));
  $('showOperationInEditorBtn')?.addEventListener('click', () => focusEditorOnOperationEventsLV75(false));
  $('selectOperationRowsBtn')?.addEventListener('click', () => focusEditorOnOperationEventsLV75(true));
}

setTimeout(() => {
  try {
    setupSectionTogglesLV75();
    injectGlobalCollapseBarLV75();
    setupSectionTogglesLV75();
    lv75BindEvents();
    if ($('versionInfo')) $('versionInfo').innerHTML = '<span class="pill modeBadge">LV7.6</span><span class="pill ok">Rà soát lỗi Mục 4/5/6/7 + trước/sau mốc thiếu + thu gọn từng khu vực</span>';
    updateEditorFocusNoticeLV74();
    log('Sẵn sàng LV7.6: Mục 5 hiển thị cả dòng liền trước và liền sau mốc thiếu; Mục 7 cũng áp dụng chế độ rà soát tại Mục 4; các khu vực có nút Thu gọn/Mở.');
  } catch(e) { log('Lỗi khởi tạo LV7.6: ' + e.message); }
}, 0);
// ====================== END LV7.6 EXTENSIONS ======================

// ====================== LV7.6 WORKFLOW + OPERATION RESTORE EXTENSIONS ======================
// LV7.6:
// - Khi thao tác Mục 7, tự thu gọn Mục 5 và Mục 6, chỉ mở Mục 4 và Mục 7.
// - Khôi phục P nền theo ngữ cảnh vận hành:
//   + Sự cố/cắt điện: phục hồi P nền cho dòng nguồn, không chỉnh lộ khác.
//   + Chuyển tải: phục hồi P nền cho dòng nguồn và trừ phần tương ứng ở lộ nhận tải nếu tìm được.
// - Làm rõ luồng B: được dùng dữ liệu mới nếu cùng cấu trúc/cùng quy tắc chỉ danh và thuộc phạm vi model đã huấn luyện.

function getSectionTitleLV76(card) {
  return card?.querySelector(':scope > h2 .sectionTitleText, :scope > h2')?.textContent?.trim() || '';
}

function setSectionCollapsedLV76(card, collapsed) {
  if (!card) return;
  card.classList.toggle('section-collapsed', !!collapsed);
  const btn = card.querySelector(':scope > h2 .sectionToggleBtn');
  if (btn) btn.textContent = card.classList.contains('section-collapsed') ? 'Mở' : 'Thu gọn';
}

function focusOperationWorkspaceLV76() {
  const cards = [...document.querySelectorAll('main .card')];
  for (const card of cards) {
    const title = getSectionTitleLV76(card);
    // Khi vào Mục 7, thu gọn Mục 5/6 theo yêu cầu để tối ưu không gian.
    if (/^5\)/.test(title) || /^6\)/.test(title)) setSectionCollapsedLV76(card, true);
    // Mở Mục 4 và Mục 7 để người vận hành theo dõi bảng hiệu chỉnh + báo cáo vận hành.
    if (/^4\)/.test(title) || /^7\)/.test(title) || /^0\.2\)/.test(title)) setSectionCollapsedLV76(card, false);
  }
  log('LV7.6: Đã thu gọn Mục 5/6 và mở Mục 4/7 để thao tác phân tích vận hành.');
}

function buildGroupsFromNormalizedRowsLV76() {
  if (!state.rows.length) normalizeRows();
  const groups = new Map();
  for (const r of state.rows) {
    if (!groups.has(r.station)) groups.set(r.station, []);
    groups.get(r.station).push(r);
  }
  for (const arr of groups.values()) arr.sort((a,b)=>a.time-b.time);
  return groups;
}

function findNormalizedRowByRawIndexLV76(rawIndex) {
  if (!state.rows.length) normalizeRows();
  return state.rows.find(r => r.idx === rawIndex) || null;
}

function nextCleanItemInGroupLV76(groups, station, time, minP) {
  const arr = groups.get(station) || [];
  for (const it of arr) {
    if (it.time <= time) continue;
    const bad = !Number.isFinite(it.p) || isLowPValue(it.p, minP) || it.outage || it.transfer || it.abnormal || it.excludeTrain;
    if (!bad) return it;
  }
  return null;
}

function estimateCleanPForEventRowLV76(row, groups, minP) {
  const prev = previousCleanItemInGroup(groups, row.station, row.time, minP) || previousItemInGroup(groups, row.station, row.time);
  const next = nextCleanItemInGroupLV76(groups, row.station, row.time, minP);
  if (prev && next && Number.isFinite(prev.p) && Number.isFinite(next.p) && next.time > prev.time) {
    const ratio = (row.time - prev.time) / (next.time - prev.time);
    return prev.p + (next.p - prev.p) * ratio;
  }
  if (prev && Number.isFinite(prev.p)) return prev.p;
  if (next && Number.isFinite(next.p)) return next.p;
  const pg = parseNumber(row.raw?.p_goc);
  if (Number.isFinite(pg) && !isLowPValue(pg, minP)) return pg;
  return NaN;
}

function ensureProcessAuditColumnsLV76() {
  addProcessColumns();
  for (const name of ['p_truoc_khoi_phuc','p_sau_khoi_phuc','p_bi_dieu_chinh_do_chuyen_tai','chi_danh_nguon_chuyen_tai']) {
    if (!state.headers.includes(name)) state.headers.push(name);
    for (const r of state.rawRows) if (!(name in r)) r[name] = '';
  }
}

function restoreOriginalPForSelectedLV76() {
  const m = readMap();
  if (!m.p || !m.time) throw new Error('Cần ánh xạ cột Thời gian và Công suất P.');
  const selected = [...(state.editor?.selected || new Set())].filter(i => i >= 0 && i < state.rawRows.length).sort((a,b)=>a-b);
  if (!selected.length) { log('Chưa chọn dòng để khôi phục P nền.'); return; }

  normalizeRows();
  const minP = getInvalidPThreshold();
  const tolPct = Math.max(1, parseNumber($('transferTolerancePercent')?.value) || 30);
  const windowMode = $('operationTimeWindow')?.value || 'same';
  const groups = buildGroupsFromNormalizedRowsLV76();
  ensureProcessAuditColumnsLV76();

  let restored = 0, receiverAdjusted = 0, skipped = 0;
  const messages = [];

  for (const idx of selected) {
    const raw = state.rawRows[idx];
    const row = findNormalizedRowByRawIndexLV76(idx);
    if (!raw || !row) { skipped++; continue; }

    const oldP = parseNumber(raw[m.p]);
    const targetP = estimateCleanPForEventRowLV76(row, groups, minP);
    if (!Number.isFinite(targetP)) {
      skipped++;
      messages.push(`Dòng ${idx+1}: không đủ mốc sạch trước/sau để khôi phục P nền.`);
      continue;
    }

    const storedOriginalP = parseNumber(raw.p_goc);
    let deltaAdd = targetP - (Number.isFinite(oldP) ? oldP : 0);
    // Trường hợp dòng nguồn đã được nội suy trước đó nhưng lộ nhận chưa bị trừ tải:
    // nếu p_goc thấp/0 và hiện P đã là P nền, vẫn cần dùng phần đã phục hồi để cân bằng lộ nhận.
    let deltaForReceiver = deltaAdd;
    if (Math.abs(deltaAdd) < 1e-9 && row.transfer && Number.isFinite(storedOriginalP) && isLowPValue(storedOriginalP, minP) && Number.isFinite(oldP)) {
      deltaForReceiver = oldP - storedOriginalP;
    }
    if (Math.abs(deltaAdd) < 1e-9 && !(row.transfer && deltaForReceiver > 0)) {
      skipped++;
      messages.push(`Dòng ${idx+1}: P hiện tại đã gần bằng P nền ước tính.`);
      continue;
    }

    if (!raw.p_goc) raw.p_goc = Number.isFinite(oldP) ? String(oldP) : String(raw[m.p] ?? '');
    raw.p_truoc_khoi_phuc = Number.isFinite(oldP) ? formatNum(oldP,3) : String(raw[m.p] ?? '');
    raw[m.p] = formatNum(targetP, 3);
    raw.p_sau_khoi_phuc = formatNum(targetP, 3);
    raw.du_lieu_noi_suy = '1';
    raw.bo_khoi_huan_luyen = '1';

    let note = `LV7.6 khoi phuc P nen dong nguon tu ${raw.p_truoc_khoi_phuc} len ${formatNum(targetP,3)}`;

    // Nếu đây là chuyển tải, khôi phục P cho lộ nguồn phải cắt bớt phần tương ứng tại lộ nhận.
    if (row.transfer && deltaForReceiver > 0) {
      const intervalMs = inferIntervalMsForStation(groups, row.station);
      const windowMs = windowMode === 'near' ? intervalMs * 1.1 : 0;
      const candidates = findTimeWindowCandidates(state.rows, row.time, row.station, windowMs);
      const receiver = chooseTransferReceiver(candidates, groups, row, deltaForReceiver, minP, tolPct);
      if (receiver && Number.isInteger(receiver.idx)) {
        const recvRaw = state.rawRows[receiver.idx];
        const recvOldP = parseNumber(recvRaw?.[m.p]);
        if (recvRaw && Number.isFinite(recvOldP)) {
          if (!recvRaw.p_goc) recvRaw.p_goc = String(recvOldP);
          recvRaw.p_truoc_khoi_phuc = formatNum(recvOldP,3);
          const recvNewP = Math.max(0, recvOldP - deltaForReceiver);
          recvRaw[m.p] = formatNum(recvNewP,3);
          recvRaw.p_sau_khoi_phuc = formatNum(recvNewP,3);
          recvRaw.p_bi_dieu_chinh_do_chuyen_tai = formatNum(deltaForReceiver,3);
          recvRaw.chi_danh_nguon_chuyen_tai = row.stationRaw || row.station;
          recvRaw.du_lieu_noi_suy = '1';
          recvRaw.bo_khoi_huan_luyen = '1';
          recvRaw.ghi_chu_xu_ly = `${recvRaw.ghi_chu_xu_ly || ''} LV7.6 cat bot ${formatNum(deltaForReceiver,3)} MW do khoi phuc nguon chuyen tai dong ${idx+1}`.trim();
          receiverAdjusted++;
          note += `; da cat bot ${formatNum(deltaForReceiver,3)} MW tai lo nhan ${receiver.stationRaw || receiver.station}`;
        }
      } else {
        note += '; chua tim duoc lo nhan de cat bot P, can kiem tra thu cong';
      }
    } else if (row.outage) {
      note += '; su co/cat dien: chi bo sung P nen cho dong nguon, khong dieu chinh lo khac';
    }

    raw.ghi_chu_xu_ly = `${raw.ghi_chu_xu_ly || ''} ${note}`.trim();
    restored++;
  }

  normalizeRows();
  renderEditorTable(); previewData(); markEditorDirty(true);
  const msg = `LV7.6 đã khôi phục P nền cho ${restored}/${selected.length} dòng; đã cắt bớt P ở ${receiverAdjusted} dòng lộ nhận; bỏ qua ${skipped}.`;
  log(msg + (messages.length ? '\n' + messages.slice(0,8).join('\n') : ''));
  if ($('operationSummary')) {
    $('operationSummary').innerHTML += `<span class="pill ok">${msg}</span>`;
  }
}

// Ghi đè nút cũ bằng logic LV7.6.
restoreOriginalPForSelected = restoreOriginalPForSelectedLV76;

// Bọc phân tích Mục 7 để tự thu gọn Mục 5/6 và mở Mục 4/7.
const analyzeOperationEventsBeforeLV76 = analyzeOperationEvents;
analyzeOperationEvents = function() {
  focusOperationWorkspaceLV76();
  const res = analyzeOperationEventsBeforeLV76();
  focusOperationWorkspaceLV76();
  return res;
};

function updateWorkflowLabelsLV76() {
  if ($('versionInfo')) $('versionInfo').innerHTML = '<span class="pill modeBadge">LV7.6</span><span class="pill ok">Mục 7 tự thu gọn Mục 5/6 + khôi phục P nền có cân bằng chuyển tải</span>';
  const btn = $('restoreOriginalPBtn');
  if (btn) btn.textContent = 'Khôi phục P nền cho dòng chọn';
  const barNote = document.querySelector('#lv75CollapseBar .compactNote');
  if (barNote && !barNote.innerHTML.includes('LV7.6')) {
    barNote.innerHTML += '<br><b>LV7.6:</b> Khi chạy Mục 7, ứng dụng tự thu gọn Mục 5/6 và mở Mục 4/7 để tối ưu không gian.';
  }
}

setTimeout(() => {
  try {
    updateWorkflowLabelsLV76();
    log('Sẵn sàng LV7.6: Mục 7 tự thu gọn Mục 5/6; khôi phục P nền cho dòng chọn có xét chuyển tải/cắt điện; luồng B dùng được dữ liệu mới cùng cấu trúc và cùng phạm vi model.');
  } catch(e) { log('Lỗi khởi tạo LV7.6: ' + e.message); }
}, 0);
// ====================== END LV7.6 EXTENSIONS ======================

// ====================== LV8 EXTENSIONS: multi-level forecast for dispatch operation ======================
// LV8 keeps the LV7.6 operational workflow and adds forecast horizons, Pmax/time, energy and shift summaries.
state.lv8 = state.lv8 || {version:'LV8', summaryRows:[], lastHorizonMode:'current_steps'};

function updateWorkflowLabelsLV8() {
  if ($('versionInfo')) $('versionInfo').innerHTML = '<span class="pill modeBadge">LV8</span><span class="pill ok">Dự báo đa cấp: 15/30/60 phút, ngày tới, Pmax, MWh, ca vận hành</span>';
  const title = document.querySelector('h1');
  if (title) title.textContent = 'SCADA Load Forecast Offline PWA LV8';
}

function lv8GetIntervalFromForecastRows(rows) {
  const times = (rows || []).map(r => parseTime(r.time)).filter(d => d instanceof Date && !isNaN(d));
  times.sort((a,b)=>a-b);
  const diffs=[];
  for (let i=1;i<times.length;i++) {
    const d=(times[i]-times[i-1])/60000;
    if (d>0 && d<1440) diffs.push(d);
  }
  if (!diffs.length) {
    const st = $('stationSelect')?.value || '__ALL__';
    const model = resolveModelForStation(st) || state.model;
    return model?.intervalMinutes || detectIntervalMinutes(state.rows || []) || 60;
  }
  diffs.sort((a,b)=>a-b);
  return diffs[Math.floor(diffs.length/2)] || 60;
}

function lv8HorizonToSteps(intervalMinutes) {
  const mode = $('lv8HorizonMode')?.value || 'current_steps';
  state.lv8.lastHorizonMode = mode;
  if (mode === 'current_steps') return Math.max(1, Math.floor(parseNumber($('forecastSteps')?.value) || 24));
  const minMap = {next_15m:15, next_30m:30, next_60m:60, next_day:1440, next_2_days:2880, next_7_days:10080};
  const minutes = minMap[mode] || 1440;
  return Math.max(1, Math.ceil(minutes / Math.max(1, intervalMinutes || 60)));
}

function lv8ShiftName(date, mode) {
  const h = date.getHours();
  if (mode === 'dispatch_3shift') {
    if (h >= 6 && h < 14) return 'Ca 1 06-14';
    if (h >= 14 && h < 22) return 'Ca 2 14-22';
    return 'Ca 3 22-06';
  }
  if (h < 6) return 'Đêm 00-06';
  if (h < 12) return 'Sáng 06-12';
  if (h < 18) return 'Chiều 12-18';
  return 'Tối 18-24';
}

function lv8DateKey(date) {
  return date instanceof Date && !isNaN(date) ? date.toISOString().slice(0,10) : '';
}

function lv8BuildSummary(rows) {
  const interval = lv8GetIntervalFromForecastRows(rows);
  const shiftMode = $('lv8ShiftMode')?.value || 'standard';
  const groups = new Map();
  for (const r of rows || []) {
    const t = parseTime(r.time);
    const p = parseNumber(r.forecast_p_mw);
    if (!(t instanceof Date) || isNaN(t) || !Number.isFinite(p)) continue;
    const station = r.station || 'ALL';
    const key = station + '|' + lv8DateKey(t);
    if (!groups.has(key)) groups.set(key, {station, date:lv8DateKey(t), items:[]});
    groups.get(key).items.push({...r, _t:t, _p:p, _shift:lv8ShiftName(t, shiftMode)});
  }
  const out=[];
  const shiftNames = shiftMode === 'dispatch_3shift' ? ['Ca 1 06-14','Ca 2 14-22','Ca 3 22-06'] : ['Đêm 00-06','Sáng 06-12','Chiều 12-18','Tối 18-24'];
  for (const g of groups.values()) {
    g.items.sort((a,b)=>a._t-b._t);
    const ps = g.items.map(x=>x._p).filter(Number.isFinite);
    if (!ps.length) continue;
    const maxP = Math.max(...ps), minP = Math.min(...ps), avgP = mean(ps);
    const maxItem = g.items.find(x=>x._p === maxP) || g.items[0];
    const warnCount = g.items.filter(x => String(x.trang_thai_nguong || '').includes('CẢNH') || String(x.trang_thai_nguong || '').includes('NGUY')).length;
    const dangerCount = g.items.filter(x => String(x.trang_thai_nguong || '').includes('NGUY')).length;
    const d = parseDesignation(g.station);
    const row = {
      ngay:g.date,
      station:g.station,
      don_vi:d.unit||'', tram:d.substation||'', lo:d.feeder||'', noi_vong:d.hasRing?`${d.ringUnit}/${d.ringSubstation}/${d.ringFeeder}`:'',
      so_moc:g.items.length,
      buoc_phut:formatNum(interval,0),
      pmax_mw:formatNum(maxP,3),
      gio_pmax:fmtTime(maxItem._t),
      pmin_mw:formatNum(minP,3),
      ptb_mw:formatNum(avgP,3),
      san_luong_mwh:formatNum(ps.reduce((a,b)=>a+b,0) * interval / 60,3),
      canh_bao_moc:warnCount,
      nguy_hiem_moc:dangerCount,
      ket_luan: dangerCount ? 'Có mốc nguy hiểm' : (warnCount ? 'Có mốc cảnh báo' : 'Bình thường')
    };
    for (const name of shiftNames) {
      const si = g.items.filter(x => x._shift === name);
      const vals = si.map(x=>x._p).filter(Number.isFinite);
      const prefix = name.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/[^a-z0-9]+/g,'_').replace(/^_|_$/g,'');
      row[`${prefix}_ptb_mw`] = vals.length ? formatNum(mean(vals),3) : '';
      row[`${prefix}_pmax_mw`] = vals.length ? formatNum(Math.max(...vals),3) : '';
      row[`${prefix}_mwh`] = vals.length ? formatNum(vals.reduce((a,b)=>a+b,0) * interval / 60,3) : '';
    }
    out.push(row);
  }
  out.sort((a,b)=>String(a.station).localeCompare(String(b.station),'vi') || String(a.ngay).localeCompare(String(b.ngay)));
  state.lv8.summaryRows = out;
  return out;
}

function lv8SummaryHeaders() {
  const shiftMode = $('lv8ShiftMode')?.value || 'standard';
  const base = ['ngay','don_vi','tram','lo','noi_vong','station','so_moc','buoc_phut','pmax_mw','gio_pmax','pmin_mw','ptb_mw','san_luong_mwh','canh_bao_moc','nguy_hiem_moc','ket_luan'];
  const shifts = shiftMode === 'dispatch_3shift' ? ['ca_1_06_14','ca_2_14_22','ca_3_22_06'] : ['dem_00_06','sang_06_12','chieu_12_18','toi_18_24'];
  for (const s of shifts) base.push(`${s}_ptb_mw`, `${s}_pmax_mw`, `${s}_mwh`);
  return base;
}

function lv8RenderSummary(rows) {
  const box = $('lv8SummaryBox');
  const tbl = $('lv8SummaryTable');
  const rowsIn = rows || state.lv8.summaryRows || [];
  if (box) {
    if (!rowsIn.length) box.innerHTML = '<span class="pill warn">Chưa có báo cáo đa cấp LV8</span>';
    else {
      const pmax = Math.max(...rowsIn.map(r=>parseNumber(r.pmax_mw)).filter(Number.isFinite));
      const maxRow = rowsIn.find(r => parseNumber(r.pmax_mw) === pmax) || rowsIn[0];
      const energy = rowsIn.map(r=>parseNumber(r.san_luong_mwh)).filter(Number.isFinite).reduce((a,b)=>a+b,0);
      const warn = rowsIn.map(r=>parseNumber(r.canh_bao_moc)).filter(Number.isFinite).reduce((a,b)=>a+b,0);
      const danger = rowsIn.map(r=>parseNumber(r.nguy_hiem_moc)).filter(Number.isFinite).reduce((a,b)=>a+b,0);
      box.innerHTML = `<span class="pill ok">LV8: ${rowsIn.length} dòng tổng hợp</span><span class="pill">Pmax ${formatNum(pmax,3)} MW</span><span class="pill">Giờ Pmax ${escapeHtml(maxRow.gio_pmax || '')}</span><span class="pill">Sản lượng ${formatNum(energy,3)} MWh</span><span class="pill ${danger?'bad':(warn?'warn':'ok')}">Cảnh báo/Nguy hiểm: ${warn}/${danger}</span>`;
    }
  }
  if (tbl) renderTableInBox(tbl, rowsIn, lv8SummaryHeaders(), 1000);
}

function lv8CreateSummary() {
  if (!state.forecastRows || !state.forecastRows.length) throw new Error('Chưa có dữ liệu dự báo. Hãy bấm dự báo ở Mục 10 trước.');
  const rows = lv8BuildSummary(state.forecastRows);
  lv8RenderSummary(rows);
  if ($('lv8ExportSummaryBtn')) $('lv8ExportSummaryBtn').disabled = !rows.length;
  log(`LV8 đã tạo báo cáo đa cấp: ${rows.length} dòng tổng hợp.`);
}

function lv8QuickForecast() {
  if (!state.model) throw new Error('Chưa có model vận hành.');
  const selected = $('stationSelect')?.value || '__ALL__';
  const model = resolveModelForStation(selected === '__ALL__' && state.model.modelsByStation ? Object.keys(state.model.modelsByStation)[0] : selected) || state.model;
  const interval = model?.intervalMinutes || detectIntervalMinutes(state.rows || []) || 60;
  const steps = lv8HorizonToSteps(interval);
  if ($('forecastSteps')) $('forecastSteps').value = steps;
  const station = selected === '__ALL__' && state.model.modelsByStation ? Object.keys(state.model.modelsByStation)[0] : selected;
  const out = forecastForStation(station, steps);
  state.forecastRows = out.forecast;
  applyThresholdsToForecast(state.forecastRows);
  renderTable(state.forecastRows, ['step','time','station','forecast_p_mw','model_used','strategy_method','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong'], 3000);
  const actualHist = out.series.slice(-Math.min(steps*2, 360)).map((r,i,arr)=> i < arr.length-steps ? r.p : NaN);
  const predHist = out.series.slice(-Math.min(steps*2, 360)).map((r,i,arr)=> i >= arr.length-steps ? r.p : NaN);
  const times = out.series.slice(-Math.min(steps*2, 360)).map(r=>r.time);
  drawSeries(times, actualHist, predHist, 'Lịch sử gần nhất', 'Dự báo LV8');
  updateForecastMetrics(state.forecastRows);
  renderForecastExplain(out);
  if ($('exportForecastBtn')) $('exportForecastBtn').disabled = false;
  if ($('lv8SummaryBtn')) $('lv8SummaryBtn').disabled = false;
  lv8CreateSummary();
  log(`LV8 đã dự báo nhanh ${state.lv8.lastHorizonMode}: ${out.forecast.length} bước, mỗi bước ${out.interval} phút.`);
}

// Override forecast functions to enable LV8 summary controls and update labels.
const forecastNextLV76 = forecastNext;
forecastNext = function() {
  forecastNextLV76();
  if ($('lv8SummaryBtn')) $('lv8SummaryBtn').disabled = !(state.forecastRows || []).length;
  if ((state.forecastRows || []).length) {
    try { lv8CreateSummary(); } catch(e) { log('LV8 chưa tạo được báo cáo đa cấp: ' + e.message); }
  }
};

const forecastAllStationsLV76 = forecastAllStations;
forecastAllStations = function() {
  forecastAllStationsLV76();
  if ($('lv8SummaryBtn')) $('lv8SummaryBtn').disabled = !(state.forecastRows || []).length;
  if ((state.forecastRows || []).length) {
    try { lv8CreateSummary(); } catch(e) { log('LV8 chưa tạo được báo cáo đa cấp: ' + e.message); }
  }
};

function exportForecast() {
  if (!state.forecastRows.length) return;
  const rows = state.forecastRows.map(r => {
    const d = parseDesignation(r.station);
    return {...r, don_vi:d.unit||'', tram:d.substation||'', lo:d.feeder||'', noi_vong:d.hasRing?`${d.ringUnit}/${d.ringSubstation}/${d.ringFeeder}`:''};
  });
  saveTextFile('forecast_lv8.csv', toCSV(rows, ['step','time','don_vi','tram','lo','noi_vong','station','forecast_p_mw','model_used','strategy_method','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong']), 'text/csv;charset=utf-8');
  log('Đã xuất forecast_lv8.csv theo cấu trúc LV8.');
}

function exportModel() {
  if (!state.model) return;
  const payload = {
    ...state.model,
    scadaModelPackage: 'SCADA_LOAD_FORECAST_OPERATIONAL_MODEL',
    appVersion: 'LV8.2',
    exportSource: 'MUC_8_EXPORT_MODEL',
    exportSection: '8',
    allowedImportSection: '10',
    modelKind: 'OPERATIONAL_FORECAST_MODEL',
    exportedAt: new Date().toISOString(),
    importRules: {
      onlyImportInSection10: true,
      requiredExportSource: 'MUC_8_EXPORT_MODEL',
      note: 'File này được xuất từ Mục 8 và được phép nạp ở Mục 10. LV8 dùng được cho dữ liệu mới cùng cấu trúc/chỉ danh/phạm vi đã huấn luyện.'
    },
    lv8Config: collectFullConfig(),
    lv8State:{strategyByStation: state.lv7?.strategyByStation || state.model?.strategyByStation || {}, savedAt:new Date().toISOString(), multiLevelForecast:true}
  };
  saveTextFile('model_gbdt_lv8_5_operational.json', JSON.stringify(payload, null, 2), 'application/json');
  log('Đã xuất model_gbdt_lv8_5_operational.json từ Mục 8. Đây là file hợp lệ để nạp ở Mục 10 trong mạng SCADA.');
}

function exportLV8Summary() {
  const rows = state.lv8.summaryRows || [];
  if (!rows.length) { log('Chưa có báo cáo đa cấp LV8 để xuất.'); return; }
  saveTextFile('forecast_summary_lv8.csv', toCSV(rows, lv8SummaryHeaders()), 'text/csv;charset=utf-8');
  log(`Đã xuất forecast_summary_lv8.csv gồm ${rows.length} dòng.`);
}

function lv8BindButtons() {
  const quick = $('lv8QuickForecastBtn');
  const sum = $('lv8SummaryBtn');
  const exp = $('lv8ExportSummaryBtn');
  if (quick && !quick.dataset.boundLv8) {
    quick.dataset.boundLv8 = '1'; quick.disabled = false;
    quick.addEventListener('click', () => { try { lv8QuickForecast(); } catch(e) { log('Lỗi dự báo nhanh LV8: ' + e.message); } });
  }
  if (sum && !sum.dataset.boundLv8) {
    sum.dataset.boundLv8 = '1';
    sum.addEventListener('click', () => { try { lv8CreateSummary(); } catch(e) { log('Lỗi tạo báo cáo LV8: ' + e.message); } });
  }
  if (exp && !exp.dataset.boundLv8) {
    exp.dataset.boundLv8 = '1';
    exp.addEventListener('click', () => { try { exportLV8Summary(); } catch(e) { log('Lỗi xuất báo cáo LV8: ' + e.message); } });
  }
}

// Rename LV7.6 report exporters in UI/log only; core functions remain compatible.
const exportV7ModelReportBeforeLV8 = exportV7ModelReport;
exportV7ModelReport = function() {
  const rows = state.lv7.modelReports || [];
  if (!rows.length) { log('Chưa có báo cáo so sánh mô hình.'); return; }
  saveTextFile('model_compare_lv8.csv', toCSV(formatMetricRowsV7(rows), ['station','method','ten_mo_hinh','n','MAE','MAPE_pct','RMSE','xep_hang','khuyen_nghi','w_gbdt','w_similar','w_week','w_trend']), 'text/csv;charset=utf-8');
  log(`Đã xuất model_compare_lv8.csv gồm ${rows.length} dòng.`);
};

const exportV7HourlyReportBeforeLV8 = exportV7HourlyReport;
exportV7HourlyReport = function() {
  const rows = state.lv7.hourlyReports || [];
  if (!rows.length) { log('Chưa có báo cáo sai số theo giờ.'); return; }
  const out = rows.map(r => ({station:r.station, method:r.method, ten_mo_hinh:r.ten_mo_hinh, hour:r.hour, n:r.n, MAE:formatNum(r.mae,3), MAPE_pct:formatNum(r.mape,2), RMSE:formatNum(r.rmse,3)}));
  saveTextFile('hourly_error_lv8.csv', toCSV(out, ['station','method','ten_mo_hinh','hour','n','MAE','MAPE_pct','RMSE']), 'text/csv;charset=utf-8');
  log(`Đã xuất hourly_error_lv8.csv gồm ${rows.length} dòng.`);
};

setTimeout(() => {
  try {
    updateWorkflowLabelsLV8();
    lv8BindButtons();
    log('Sẵn sàng LV8: bổ sung dự báo đa cấp 15/30/60 phút, ngày tới, Pmax, giờ Pmax, MWh và phụ tải theo ca.');
  } catch(e) { log('Lỗi khởi tạo LV8: ' + e.message); }
}, 0);
// ====================== END LV8 EXTENSIONS ======================

// ====================== LV8.1 CLARITY + REPORT FLOW FIX ======================
// LV8.1 clarifies horizon semantics and makes the multi-level report button usable:
// if there is no forecast yet, the report button automatically runs the selected LV8.1 forecast first.
state.lv8 = state.lv8 || {};
state.lv8.version = 'LV8.1';

function lv81HorizonText(mode, interval, steps) {
  const map = {
    current_steps: `Theo ${steps} bước đang nhập ở Mục 10`,
    next_15m: '15 phút tới sau mốc cuối dữ liệu',
    next_30m: '30 phút tới sau mốc cuối dữ liệu',
    next_60m: '1 giờ tới sau mốc cuối dữ liệu',
    next_day: '24 giờ tới sau mốc cuối dữ liệu',
    next_2_days: '48 giờ tới sau mốc cuối dữ liệu',
    next_7_days: '7 ngày / 168 giờ tới sau mốc cuối dữ liệu'
  };
  const raw = map[mode] || map.next_day;
  const realMinutes = Math.max(1, steps) * Math.max(1, interval || 60);
  const q = (interval && interval > 0) ? `Chu kỳ dữ liệu/model: ${interval} phút; số bước thực tế: ${steps}; khoảng dự báo thực tế: ${realMinutes} phút.` : `Số bước thực tế: ${steps}.`;
  return `${raw}. ${q}`;
}

function lv81StationListForForecast() {
  const selected = $('stationSelect')?.value || '__ALL__';
  if (selected !== '__ALL__') return [selected];
  if (state.model?.modelsByStation) return Object.keys(state.model.modelsByStation);
  const s = [...new Set((state.rows || []).map(r => r.station || 'ALL'))].filter(Boolean);
  return s.length ? s : ['__ALL__'];
}

function lv81LastTimeForStation(station) {
  const rows = (state.rows || []).filter(r => station === '__ALL__' || station === 'ALL' || r.station === station);
  const times = rows.map(r => r.time).filter(t => t instanceof Date && !isNaN(t));
  if (!times.length) return null;
  times.sort((a,b)=>a-b);
  return times[times.length - 1];
}

function lv81UpdateHorizonExplain(stepsOverride=null) {
  const box = $('lv8HorizonExplainBox');
  if (!box) return;
  const selected = $('stationSelect')?.value || '__ALL__';
  const mode = $('lv8HorizonMode')?.value || 'current_steps';
  const stations = lv81StationListForForecast();
  let interval = 60;
  const firstStation = stations[0] || selected;
  const model = resolveModelForStation(firstStation) || state.model;
  interval = model?.intervalMinutes || detectIntervalMinutes(state.rows || []) || 60;
  const steps = stepsOverride || lv8HorizonToSteps(interval);
  const lastTimes = stations.map(st => ({st, t: lv81LastTimeForStation(st)})).filter(x => x.t instanceof Date && !isNaN(x.t));
  let baseText = '';
  if (selected === '__ALL__') {
    const minT = lastTimes.length ? new Date(Math.min(...lastTimes.map(x=>x.t.getTime()))) : null;
    const maxT = lastTimes.length ? new Date(Math.max(...lastTimes.map(x=>x.t.getTime()))) : null;
    baseText = lastTimes.length ? `Dự báo tất cả chỉ danh: mỗi chỉ danh bắt đầu sau mốc cuối riêng. Mốc cuối sớm nhất: ${fmtTime(minT)}; muộn nhất: ${fmtTime(maxT)}.` : 'Chưa xác định được mốc cuối dữ liệu.';
  } else {
    const t = lastTimes[0]?.t;
    baseText = t ? `Dự báo cho ${selected}: mốc dữ liệu cuối là ${fmtTime(t)}, mốc dự báo đầu tiên là ${fmtTime(new Date(t.getTime() + interval*60000))}.` : `Chưa xác định được mốc cuối dữ liệu cho ${selected}.`;
  }
  box.innerHTML = `<span class="pill ok">LV8.1</span><span class="pill">${escapeHtml(lv81HorizonText(mode, interval, steps))}</span><span class="pill">${escapeHtml(baseText)}</span><span class="pill warn">“24 giờ tới” là 24 giờ sau mốc cuối dữ liệu, không bắt buộc là ngày lịch kế tiếp.</span>`;
}

function lv81ForecastAllSelectedStations(steps) {
  if (!state.model) throw new Error('Chưa có model vận hành. Hãy nạp model ở Mục 10.');
  const stations = lv81StationListForForecast();
  let all = [];
  const skipped = [];
  for (const st of stations) {
    try {
      const out = forecastForStation(st, steps);
      all = all.concat(out.forecast);
    } catch(e) {
      skipped.push(`${st}: ${e.message}`);
    }
  }
  if (!all.length) throw new Error('Không tạo được dòng dự báo nào. ' + skipped.join('; '));
  state.forecastRows = all;
  applyThresholdsToForecast(state.forecastRows);
  renderTable(all, ['step','time','station','forecast_p_mw','model_used','strategy_method','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong'], 5000);
  updateForecastMetrics(all);
  if ($('exportForecastBtn')) $('exportForecastBtn').disabled = false;
  if ($('lv8SummaryBtn')) $('lv8SummaryBtn').disabled = false;
  if (skipped.length) log('LV8.1 bỏ qua một số chỉ danh: ' + skipped.join(' | '));
  return {forecast: all, skipped};
}

function lv8QuickForecast(options={}) {
  if (!state.model) throw new Error('Chưa có model vận hành. Hãy nạp model ở Mục 10.');
  const selected = $('stationSelect')?.value || '__ALL__';
  const modelStation = selected === '__ALL__' ? (lv81StationListForForecast()[0] || '__ALL__') : selected;
  const model = resolveModelForStation(modelStation) || state.model;
  const interval = model?.intervalMinutes || detectIntervalMinutes(state.rows || []) || 60;
  const steps = lv8HorizonToSteps(interval);
  if ($('forecastSteps')) $('forecastSteps').value = steps;
  lv81UpdateHorizonExplain(steps);
  if (selected === '__ALL__') {
    lv81ForecastAllSelectedStations(steps);
    if (!options.skipSummary) lv8CreateSummary(false);
    log(`LV8.1 đã dự báo nhanh tất cả chỉ danh theo chế độ ${state.lv8.lastHorizonMode}: ${state.forecastRows.length} dòng, ${steps} bước/chỉ danh, chu kỳ ${interval} phút.`);
    return {forecast: state.forecastRows, interval, station:'__ALL__'};
  }
  const out = forecastForStation(selected, steps);
  state.forecastRows = out.forecast;
  applyThresholdsToForecast(state.forecastRows);
  renderTable(state.forecastRows, ['step','time','station','forecast_p_mw','model_used','strategy_method','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong'], 3000);
  const actualHist = out.series.slice(-Math.min(steps*2, 360)).map((r,i,arr)=> i < arr.length-steps ? r.p : NaN);
  const predHist = out.series.slice(-Math.min(steps*2, 360)).map((r,i,arr)=> i >= arr.length-steps ? r.p : NaN);
  const times = out.series.slice(-Math.min(steps*2, 360)).map(r=>r.time);
  drawSeries(times, actualHist, predHist, 'Lịch sử gần nhất', 'Dự báo LV8.1');
  updateForecastMetrics(state.forecastRows);
  renderForecastExplain(out);
  if ($('exportForecastBtn')) $('exportForecastBtn').disabled = false;
  if ($('lv8SummaryBtn')) $('lv8SummaryBtn').disabled = false;
  if (!options.skipSummary) lv8CreateSummary(false);
  log(`LV8.1 đã dự báo nhanh ${state.lv8.lastHorizonMode} cho ${selected}: ${out.forecast.length} bước, mỗi bước ${out.interval} phút.`);
  return out;
}

function lv8CreateSummary(allowAutoForecast=true) {
  if ((!state.forecastRows || !state.forecastRows.length) && allowAutoForecast) {
    log('LV8.1: chưa có kết quả forecast, tự chạy dự báo theo kiểu đang chọn trước khi tạo báo cáo đa cấp.');
    lv8QuickForecast({skipSummary:true});
  }
  if (!state.forecastRows || !state.forecastRows.length) throw new Error('Chưa có dữ liệu dự báo. Hãy nạp model ở Mục 10 và bấm dự báo LV8.1 trước.');
  const rows = lv8BuildSummary(state.forecastRows);
  lv8RenderSummary(rows);
  if ($('lv8ExportSummaryBtn')) $('lv8ExportSummaryBtn').disabled = !rows.length;
  lv81UpdateHorizonExplain();
  log(`LV8.1 đã tạo báo cáo đa cấp: ${rows.length} dòng tổng hợp từ ${state.forecastRows.length} dòng forecast.`);
}

function exportForecast() {
  if (!state.forecastRows.length) return;
  const rows = state.forecastRows.map(r => {
    const d = parseDesignation(r.station);
    return {...r, don_vi:d.unit||'', tram:d.substation||'', lo:d.feeder||'', noi_vong:d.hasRing?`${d.ringUnit}/${d.ringSubstation}/${d.ringFeeder}`:''};
  });
  saveTextFile('forecast_lv8_1.csv', toCSV(rows, ['step','time','don_vi','tram','lo','noi_vong','station','forecast_p_mw','model_used','strategy_method','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong']), 'text/csv;charset=utf-8');
  log('Đã xuất forecast_lv8_1.csv theo cấu trúc LV8.1.');
}

function exportLV8Summary() {
  const rows = state.lv8.summaryRows || [];
  if (!rows.length) { log('Chưa có báo cáo đa cấp LV8.1 để xuất.'); return; }
  saveTextFile('forecast_summary_lv8_1.csv', toCSV(rows, lv8SummaryHeaders()), 'text/csv;charset=utf-8');
  log(`Đã xuất forecast_summary_lv8_1.csv gồm ${rows.length} dòng.`);
}

function updateWorkflowLabelsLV8() {
  if ($('versionInfo')) $('versionInfo').innerHTML = '<span class="pill modeBadge">LV8.1</span><span class="pill ok">Dự báo đa cấp rõ mốc: từ mốc dữ liệu cuối, tự tạo báo cáo nếu chưa có forecast</span>';
  const title = document.querySelector('h1');
  if (title) title.textContent = 'SCADA Load Forecast Offline PWA LV8.1';
}

setTimeout(() => {
  try {
    updateWorkflowLabelsLV8();
    const h = $('lv8HorizonMode');
    if (h && !h.dataset.boundLv81) {
      h.dataset.boundLv81 = '1';
      h.addEventListener('change', () => { try { lv81UpdateHorizonExplain(); } catch(e) {} });
    }
    const shift = $('lv8ShiftMode');
    if (shift && !shift.dataset.boundLv81) {
      shift.dataset.boundLv81 = '1';
      shift.addEventListener('change', () => { try { if ((state.forecastRows||[]).length) lv8CreateSummary(false); } catch(e) { log('LV8.1 chưa cập nhật được ca vận hành: ' + e.message); } });
    }
    if ($('lv8SummaryBtn')) $('lv8SummaryBtn').disabled = false;
    if ($('lv8QuickForecastBtn')) $('lv8QuickForecastBtn').disabled = false;
    lv81UpdateHorizonExplain();
    log('Sẵn sàng LV8.1: đã làm rõ mốc dự báo và sửa nút Tạo báo cáo đa cấp để tự dự báo khi cần.');
  } catch(e) { log('Lỗi khởi tạo LV8.1: ' + e.message); }
}, 0);
// ====================== END LV8.1 CLARITY + REPORT FLOW FIX ======================

// ====================== LV8.2 TARGET-DATE FORECAST + APPEND FORECAST TO RAM ======================
// LV8.2 adds two operational modes requested by dispatch:
// 1) Forecast for the current date or a selected calendar date/time, independent from the next timestamp in sample data.
// 2) Insert forecasted points back into RAM data so the operator can continue forecasting from those predicted points.
state.lv8 = state.lv8 || {};
state.lv8.version = 'LV8.2';
state.lv8.appendedForecastRows = state.lv8.appendedForecastRows || [];

function lv82Pad2(n){ return String(n).padStart(2,'0'); }
function lv82ToLocalDatetimeValue(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  return `${d.getFullYear()}-${lv82Pad2(d.getMonth()+1)}-${lv82Pad2(d.getDate())}T${lv82Pad2(d.getHours())}:${lv82Pad2(d.getMinutes())}`;
}
function lv82ParseDatetimeLocal(v) {
  if (!v) return null;
  const d = new Date(v);
  return d instanceof Date && !isNaN(d) ? d : null;
}
function lv82RoundUpToInterval(d, interval) {
  const t = new Date(d);
  t.setSeconds(0,0);
  const min = t.getHours()*60 + t.getMinutes();
  const step = Math.max(1, Math.floor(interval || 60));
  const rounded = Math.ceil(min / step) * step;
  t.setHours(0,0,0,0);
  t.setMinutes(rounded);
  return t;
}

function lv82HorizonToSteps(intervalMinutes) {
  const mode = $('lv8HorizonMode')?.value || 'current_steps';
  state.lv8.lastHorizonMode = mode;
  if (mode === 'current_steps') return Math.max(1, Math.floor(parseNumber($('forecastSteps')?.value) || 24));
  const minMap = {next_15m:15, next_30m:30, next_60m:60, next_day:1440, next_2_days:2880, next_7_days:10080};
  const minutes = minMap[mode] || 1440;
  return Math.max(1, Math.ceil(minutes / Math.max(1, intervalMinutes || 60)));
}
// Override LV8/LV8.1 step mapping with the same behavior but LV8.2 name.
function lv8HorizonToSteps(intervalMinutes) { return lv82HorizonToSteps(intervalMinutes); }

function lv82HistoricalEstimateAt(rows, targetTime, tempUse) {
  const clean = (rows || []).filter(r => Number.isFinite(r.p) && !r.abnormal && !r.outage && !r.transfer && !r.excludeTrain);
  if (!clean.length) return NaN;
  const targetHour = targetTime.getHours() + targetTime.getMinutes()/60;
  const targetDow = targetTime.getDay();
  const targetHol = holidayByRules(targetTime) ? 1 : 0;
  const lastMs = Math.max(...clean.map(r => r.time.getTime()));
  const cand = [];
  for (const r of clean) {
    if (!(r.time instanceof Date) || isNaN(r.time)) continue;
    const hour = r.time.getHours() + r.time.getMinutes()/60;
    let score = 1;
    score += Math.max(0, 5 - Math.abs(hour - targetHour) * 3);
    if (r.time.getDay() === targetDow) score += 4;
    if ((r.holiday||0) === targetHol) score += 2;
    if (Number.isFinite(tempUse) && Number.isFinite(r.temp)) score += Math.max(0, 3 - Math.abs(r.temp - tempUse)/2);
    const ageDays = Math.max(0, (lastMs - r.time.getTime()) / 86400000);
    score += Math.max(0, 2 - ageDays / 60);
    cand.push({p:r.p, score});
  }
  cand.sort((a,b)=>b.score-a.score);
  const top = cand.slice(0, Math.min(16, cand.length));
  const sw = top.reduce((a,b)=>a+b.score,0);
  return sw ? top.reduce((a,b)=>a+b.p*b.score,0)/sw : mean(clean.map(r=>r.p));
}

function lv82BuildSyntheticContext(sourceRows, station, startTime, interval, nPerDay, tempUse, rainDefault) {
  const contextDays = 8; // enough to provide lag_week for computeFeatureVector
  const count = Math.max(nPerDay * contextDays, nPerDay + 8);
  const start = new Date(startTime.getTime() - count * interval * 60000);
  const series = [];
  for (let i=0; i<count; i++) {
    const t = new Date(start.getTime() + i * interval * 60000);
    const p = lv82HistoricalEstimateAt(sourceRows, t, tempUse);
    series.push({
      time:t,
      p:Number.isFinite(p) ? p : mean(sourceRows.map(r=>r.p)),
      station: station === '__ALL__' ? 'ALL' : station,
      temp:Number.isFinite(tempUse) ? tempUse : NaN,
      rain:Number.isFinite(rainDefault) ? rainDefault : 0,
      holiday:holidayByRules(t)?1:0,
      abnormal:0, outage:0, transfer:0, excludeTrain:1,
      du_lieu_ngu_canh_du_bao:1
    });
  }
  return series;
}

function lv82ResolveTargetStartForStation(station, interval) {
  const mode = $('lv82StartMode')?.value || 'after_latest';
  if (mode === 'after_latest') return null; // use normal recursive forecast after the last available timestamp.
  if (mode === 'today_00') {
    const d = new Date(); d.setHours(0,0,0,0); return d;
  }
  if (mode === 'now_next') return lv82RoundUpToInterval(new Date(), interval);
  const custom = lv82ParseDatetimeLocal($('lv82StartDateTime')?.value || '');
  if (custom) return custom;
  const fallback = new Date(); fallback.setHours(0,0,0,0); return fallback;
}

function forecastForStationTargetLV82(station, startTime, stepsOverride=null) {
  const model = resolveModelForStation(station);
  if (!model) throw new Error('Chưa có model phù hợp cho ' + station);
  const sourceRows = (state.rows || []).filter(r => station === '__ALL__' || station === 'ALL' || r.station === station).sort((a,b)=>a.time-b.time);
  if (sourceRows.length < 5) throw new Error('Chưa có đủ dữ liệu nền cho ' + station);
  const tempDefault = parseNumber($('futureTemp')?.value);
  const rainDefault = parseNumber($('futureRain')?.value) || 0;
  const interval = model.intervalMinutes || detectIntervalMinutes(sourceRows) || 60;
  const nPerDay = model.nPerDay || Math.round(1440/interval);
  const lastTemp = [...sourceRows].reverse().find(r => Number.isFinite(r.temp))?.temp;
  const tempUse = Number.isFinite(tempDefault) ? tempDefault : lastTemp;
  const series = lv82BuildSyntheticContext(sourceRows, station, startTime, interval, nPerDay, tempUse, rainDefault);
  const steps = Math.max(1, Math.floor(stepsOverride || parseNumber($('forecastSteps')?.value) || 24));
  const strategy = getForecastStrategyV7(station);
  const bias = estimateRecentBias(model, sourceRows, Math.max(0, Math.floor(parseNumber($('biasWindow')?.value) || 0)));
  const forecast=[];
  for (let s=1; s<=steps; s++) {
    const t = new Date(startTime.getTime() + (s-1)*interval*60000);
    const row = {time:t, p:NaN, station: station === '__ALL__' ? 'ALL' : station, temp:tempUse, rain:rainDefault, holiday:holidayByRules(t)?1:0, abnormal:0, outage:0, transfer:0, excludeTrain:1};
    series.push(row);
    const idx = series.length - 1;
    const x = computeFeatureVector(series, idx, nPerDay, {temp:tempUse, rain:rainDefault});
    const values = {
      gbdt: Math.max(0, predictModel(model, x)),
      similar_day: similarDayPrediction(series, idx, nPerDay),
      same_hour_last_week: lastWeekPrediction(series, idx, nPerDay),
      trend: trendPrediction(series, idx)
    };
    let pred = forecastValueByStrategyV7(values, strategy);
    if (!Number.isFinite(pred)) pred = values.gbdt;
    let rawPredBeforeCalibration = Math.max(0, pred + bias);
    const cal = applyCalibrationLV85(rawPredBeforeCalibration, station, t);
    pred = cal.value;
    row.p = pred;
    forecast.push({
      step:s, time:fmtTime(t), station:row.station, forecast_p_mw:formatNum(pred,3), temp:Number.isFinite(tempUse)?formatNum(tempUse,1):'', rain:rainDefault, holiday:row.holiday,
      model_used:strategy.label || LV7_METHOD_LABELS[strategy.method] || strategy.method || 'LV8.2',
      strategy_method:strategy.method || '',
      gbdt:formatNum(values.gbdt,3), similar_day:formatNum(values.similar_day,3), same_hour_last_week:formatNum(values.same_hour_last_week,3), trend:formatNum(values.trend,3), bias:formatNum(bias,3), calibration_lv85:cal.applied?1:0, calibration_mw:formatNum(cal.delta,3), calibration_source:cal.source || '', forecast_before_calibration_mw:formatNum(rawPredBeforeCalibration,3),
      lv82_start_mode:$('lv82StartMode')?.value || 'after_latest', lv82_forecast_source:'target_calendar'
    });
  }
  return {forecast, series, interval, model, station, strategy, weights:strategy.weights || null, targetStart:startTime};
}

function lv82StationListForForecast() {
  if (typeof lv81StationListForForecast === 'function') return lv81StationListForForecast();
  const selected = $('stationSelect')?.value || '__ALL__';
  if (selected !== '__ALL__') return [selected];
  if (state.model?.modelsByStation) return Object.keys(state.model.modelsByStation);
  return [...new Set((state.rows || []).map(r => r.station || 'ALL'))].filter(Boolean);
}

function lv82ForecastAllSelectedStations(steps) {
  if (!state.model) throw new Error('Chưa có model vận hành. Hãy nạp model ở Mục 10.');
  const stations = lv82StationListForForecast();
  let all = [];
  const skipped = [];
  for (const st of stations) {
    try {
      const sourceRows = (state.rows || []).filter(r => st === '__ALL__' || st === 'ALL' || r.station === st);
      const model = resolveModelForStation(st) || state.model;
      const interval = model?.intervalMinutes || detectIntervalMinutes(sourceRows) || 60;
      const targetStart = lv82ResolveTargetStartForStation(st, interval);
      const out = targetStart ? forecastForStationTargetLV82(st, targetStart, steps) : forecastForStation(st, steps);
      all = all.concat(out.forecast);
    } catch(e) { skipped.push(`${st}: ${e.message}`); }
  }
  if (!all.length) throw new Error('Không tạo được dòng dự báo nào. ' + skipped.join('; '));
  state.forecastRows = all;
  applyThresholdsToForecast(state.forecastRows);
  renderTable(all, ['step','time','station','forecast_p_mw','model_used','strategy_method','lv82_start_mode','lv82_forecast_source','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong'], 5000);
  updateForecastMetrics(all);
  if ($('exportForecastBtn')) $('exportForecastBtn').disabled = false;
  if ($('lv8SummaryBtn')) $('lv8SummaryBtn').disabled = false;
  if ($('lv82AppendForecastBtn')) $('lv82AppendForecastBtn').disabled = false;
  if (skipped.length) log('LV8.2 bỏ qua một số chỉ danh: ' + skipped.join(' | '));
  return {forecast: all, skipped};
}

function lv82HorizonText(mode, interval, steps) {
  const map = {
    current_steps: `Theo ${steps} bước đang nhập ở Mục 10`,
    next_15m: '15 phút / ít nhất 1 chu kỳ dữ liệu',
    next_30m: '30 phút / ít nhất 1 chu kỳ dữ liệu',
    next_60m: '1 giờ',
    next_day: '24 giờ / 1 ngày',
    next_2_days: '48 giờ / 2 ngày',
    next_7_days: '7 ngày / 168 giờ'
  };
  const realMinutes = Math.max(1, steps) * Math.max(1, interval || 60);
  return `${map[mode] || map.next_day}. Chu kỳ dữ liệu/model: ${interval || 60} phút; số bước thực tế: ${steps}; khoảng dự báo thực tế: ${realMinutes} phút.`;
}

function lv82UpdateHorizonExplain(stepsOverride=null) {
  const box = $('lv8HorizonExplainBox'); if (!box) return;
  const selected = $('stationSelect')?.value || '__ALL__';
  const stations = lv82StationListForForecast();
  const firstStation = stations[0] || selected;
  const model = resolveModelForStation(firstStation) || state.model;
  const interval = model?.intervalMinutes || detectIntervalMinutes(state.rows || []) || 60;
  const steps = stepsOverride || lv82HorizonToSteps(interval);
  const startMode = $('lv82StartMode')?.value || 'after_latest';
  let startText = '';
  if (startMode === 'after_latest') {
    const lastTimes = stations.map(st => ({st, t: (typeof lv81LastTimeForStation === 'function' ? lv81LastTimeForStation(st) : null)})).filter(x => x.t instanceof Date && !isNaN(x.t));
    if (selected === '__ALL__') {
      const minT = lastTimes.length ? new Date(Math.min(...lastTimes.map(x=>x.t.getTime()))) : null;
      const maxT = lastTimes.length ? new Date(Math.max(...lastTimes.map(x=>x.t.getTime()))) : null;
      startText = lastTimes.length ? `Chế độ sau mốc cuối: mỗi chỉ danh bắt đầu sau mốc cuối riêng. Mốc cuối sớm nhất ${fmtTime(minT)}, muộn nhất ${fmtTime(maxT)}.` : 'Chưa xác định được mốc cuối dữ liệu.';
    } else {
      const t = lastTimes[0]?.t;
      startText = t ? `Chế độ sau mốc cuối: ${selected} bắt đầu dự báo từ ${fmtTime(new Date(t.getTime()+interval*60000))}.` : `Chưa xác định được mốc cuối dữ liệu cho ${selected}.`;
    }
  } else {
    const target = lv82ResolveTargetStartForStation(firstStation, interval);
    const modeLabel = startMode === 'today_00' ? 'ngày hiện tại 00:00' : (startMode === 'now_next' ? 'mốc hiện tại/kế tiếp theo chu kỳ' : 'ngày/giờ tự chọn');
    startText = `Chế độ ngày/giờ đích: bắt đầu từ ${fmtTime(target)} (${modeLabel}). Không phụ thuộc mốc cuối của dữ liệu mẫu.`;
  }
  box.innerHTML = `<span class="pill ok">LV8.5</span><span class="pill">${escapeHtml(lv82HorizonText($('lv8HorizonMode')?.value || 'current_steps', interval, steps))}</span><span class="pill">${escapeHtml(startText)}</span><span class="pill warn">Forecast chèn vào RAM sẽ được đánh dấu dữ liệu dự báo và bỏ khỏi huấn luyện.</span>`;
}
// Override LV8.1 explain hook too.
function lv81UpdateHorizonExplain(stepsOverride=null) { return lv82UpdateHorizonExplain(stepsOverride); }

function lv8QuickForecast(options={}) {
  if (!state.model) throw new Error('Chưa có model vận hành. Hãy nạp model ở Mục 10.');
  const selected = $('stationSelect')?.value || '__ALL__';
  const modelStation = selected === '__ALL__' ? (lv82StationListForForecast()[0] || '__ALL__') : selected;
  const model = resolveModelForStation(modelStation) || state.model;
  const interval = model?.intervalMinutes || detectIntervalMinutes(state.rows || []) || 60;
  const steps = lv82HorizonToSteps(interval);
  if ($('forecastSteps')) $('forecastSteps').value = steps;
  lv82UpdateHorizonExplain(steps);
  if (selected === '__ALL__') {
    lv82ForecastAllSelectedStations(steps);
    if (!options.skipSummary) lv8CreateSummary(false);
    log(`LV8.5 đã dự báo ${state.forecastRows.length} dòng cho tất cả chỉ danh, chế độ ${$('lv82StartMode')?.value || 'after_latest'}, ${steps} bước/chỉ danh.`);
    return {forecast: state.forecastRows, interval, station:'__ALL__'};
  }
  const targetStart = lv82ResolveTargetStartForStation(selected, interval);
  const out = targetStart ? forecastForStationTargetLV82(selected, targetStart, steps) : forecastForStation(selected, steps);
  state.forecastRows = out.forecast;
  applyThresholdsToForecast(state.forecastRows);
  renderTable(state.forecastRows, ['step','time','station','forecast_p_mw','model_used','strategy_method','lv82_start_mode','lv82_forecast_source','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong'], 3000);
  const actualHist = out.series.slice(-Math.min(steps*2, 360)).map((r,i,arr)=> i < arr.length-steps ? r.p : NaN);
  const predHist = out.series.slice(-Math.min(steps*2, 360)).map((r,i,arr)=> i >= arr.length-steps ? r.p : NaN);
  const times = out.series.slice(-Math.min(steps*2, 360)).map(r=>r.time);
  drawSeries(times, actualHist, predHist, 'Bối cảnh nền', 'Dự báo LV8.2');
  updateForecastMetrics(state.forecastRows);
  renderForecastExplain(out);
  if ($('exportForecastBtn')) $('exportForecastBtn').disabled = false;
  if ($('lv8SummaryBtn')) $('lv8SummaryBtn').disabled = false;
  if ($('lv82AppendForecastBtn')) $('lv82AppendForecastBtn').disabled = false;
  if (!options.skipSummary) lv8CreateSummary(false);
  log(`LV8.5 đã dự báo ${out.forecast.length} bước cho ${selected}. Điểm bắt đầu: ${targetStart ? fmtTime(targetStart) : 'sau mốc cuối dữ liệu'}.`);
  return out;
}

function lv8CreateSummary(allowAutoForecast=true) {
  if ((!state.forecastRows || !state.forecastRows.length) && allowAutoForecast) {
    log('LV8.5: chưa có kết quả forecast, tự chạy dự báo theo kiểu đang chọn trước khi tạo báo cáo đa cấp.');
    lv8QuickForecast({skipSummary:true});
  }
  if (!state.forecastRows || !state.forecastRows.length) throw new Error('Chưa có dữ liệu dự báo. Hãy nạp model ở Mục 10 và bấm dự báo LV8.2 trước.');
  const rows = lv8BuildSummary(state.forecastRows);
  lv8RenderSummary(rows);
  if ($('lv8ExportSummaryBtn')) $('lv8ExportSummaryBtn').disabled = !rows.length;
  lv82UpdateHorizonExplain();
  log(`LV8.5 đã tạo báo cáo đa cấp: ${rows.length} dòng tổng hợp từ ${state.forecastRows.length} dòng forecast.`);
}

function exportForecast() {
  if (!state.forecastRows.length) return;
  const rows = state.forecastRows.map(r => {
    const d = parseDesignation(r.station);
    return {...r, don_vi:d.unit||'', tram:d.substation||'', lo:d.feeder||'', noi_vong:d.hasRing?`${d.ringUnit}/${d.ringSubstation}/${d.ringFeeder}`:''};
  });
  saveTextFile('forecast_lv8_5.csv', toCSV(rows, ['step','time','don_vi','tram','lo','noi_vong','station','forecast_p_mw','model_used','strategy_method','lv82_start_mode','lv82_forecast_source','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','calibration_lv85','calibration_mw','calibration_source','forecast_before_calibration_mw','nguong_canh_bao','trang_thai_nguong']), 'text/csv;charset=utf-8');
  log('Đã xuất forecast_lv8_5.csv theo cấu trúc LV8.2.');
}

function exportLV8Summary() {
  const rows = state.lv8.summaryRows || [];
  if (!rows.length) { log('Chưa có báo cáo đa cấp LV8.5 để xuất.'); return; }
  saveTextFile('forecast_summary_lv8_5.csv', toCSV(rows, lv8SummaryHeaders()), 'text/csv;charset=utf-8');
  log(`Đã xuất forecast_summary_lv8_5.csv gồm ${rows.length} dòng.`);
}

function lv82EnsureRawColumn(header) {
  if (!state.headers.includes(header)) {
    state.headers.push(header);
    state.rawRows.forEach(r => { if (!(header in r)) r[header] = ''; });
    fillColumnSelects(state.headers);
    refreshQuickCustomColumns?.();
  }
  return header;
}

function lv82AppendForecastToRam() {
  if (!state.forecastRows || !state.forecastRows.length) throw new Error('Chưa có dòng dự báo để cập nhật vào dữ liệu.');
  const timeCol = ensureMappedColumn('time');
  const pCol = ensureMappedColumn('p');
  const stationCol = ensureMappedColumn('station');
  const tempCol = state.colMap?.temp || ensureMappedColumn('temp');
  const rainCol = state.colMap?.rain || ensureMappedColumn('rain');
  const holidayCol = state.colMap?.holiday || ensureMappedColumn('holiday');
  const abnormalCol = state.colMap?.abnormal || ensureMappedColumn('abnormal');
  const outageCol = state.colMap?.outage || ensureMappedColumn('outage');
  const transferCol = state.colMap?.transfer || ensureMappedColumn('transfer');
  const forecastFlagCol = lv82EnsureRawColumn('du_lieu_du_bao');
  const forecastSourceCol = lv82EnsureRawColumn('nguon_du_bao');
  const modelCol = lv82EnsureRawColumn('model_used');
  const strategyCol = lv82EnsureRawColumn('strategy_method');
  const excludeCol = lv82EnsureRawColumn('bo_khoi_huan_luyen');
  const noteCol = lv82EnsureRawColumn('ghi_chu_xu_ly');
  const mode = $('lv82AppendMode')?.value || 'overwrite';
  const keyOfRaw = r => {
    const t = parseTime(r[timeCol]);
    const st = String(r[stationCol] ?? '').trim() || 'ALL';
    return (t instanceof Date && !isNaN(t) ? fmtTime(t) : String(r[timeCol]||'')) + '|' + st;
  };
  const existing = new Map();
  state.rawRows.forEach((r,i)=> existing.set(keyOfRaw(r), i));
  let added=0, overwritten=0, skipped=0;
  for (const fr of state.forecastRows) {
    const st = fr.station || 'ALL';
    const key = String(fr.time || '') + '|' + st;
    if (existing.has(key) && mode === 'skip_existing') { skipped++; continue; }
    const row = {};
    for (const h of state.headers) row[h] = '';
    row[timeCol] = fr.time || '';
    row[pCol] = fr.forecast_p_mw ?? '';
    row[stationCol] = st;
    row[tempCol] = fr.temp ?? '';
    row[rainCol] = fr.rain ?? '0';
    row[holidayCol] = fr.holiday ?? (holidayByRules(parseTime(fr.time)) ? '1' : '0');
    row[abnormalCol] = '0'; row[outageCol] = '0'; row[transferCol] = '0';
    row[forecastFlagCol] = '1';
    row[forecastSourceCol] = 'LV8.2 forecast inserted to RAM';
    row[modelCol] = fr.model_used || '';
    row[strategyCol] = fr.strategy_method || '';
    row[excludeCol] = '1';
    row[noteCol] = 'LV8.5: mốc dự báo được chèn vào RAM để dự báo tiếp; không dùng để huấn luyện model thật';
    if (existing.has(key) && mode === 'overwrite') {
      const idx = existing.get(key);
      state.rawRows[idx] = {...state.rawRows[idx], ...row};
      overwritten++;
    } else {
      state.rawRows.push(row);
      existing.set(key, state.rawRows.length-1);
      added++;
    }
  }
  normalizeRows();
  applyDataInfo();
  renderEditorTable();
  previewData();
  markEditorDirty(true);
  state.lv8.appendedForecastRows.push({at:new Date().toISOString(), added, overwritten, skipped, mode});
  log(`LV8.5 đã cập nhật forecast vào dữ liệu RAM: thêm ${added}, ghi đè ${overwritten}, bỏ qua ${skipped}. Các mốc này được đánh dấu bo_khoi_huan_luyen=1 để không huấn luyện.`);
  return {added, overwritten, skipped, mode};
}

function updateWorkflowLabelsLV8() {
  if ($('versionInfo')) $('versionInfo').innerHTML = '<span class="pill modeBadge">LV8.5</span><span class="pill ok">Dự báo theo ngày/giờ đích + cập nhật forecast vào RAM để dự báo tiếp</span>';
  const title = document.querySelector('h1');
  if (title) title.textContent = 'SCADA Load Forecast Offline PWA LV8.5';
}

setTimeout(() => {
  try {
    updateWorkflowLabelsLV8();
    const now = new Date();
    const input = $('lv82StartDateTime');
    if (input && !input.value) input.value = lv82ToLocalDatetimeValue(now);
    ['lv8HorizonMode','lv82StartMode','lv82StartDateTime'].forEach(id => {
      const el = $(id);
      if (el && !el.dataset.boundLv82) {
        el.dataset.boundLv82 = '1';
        el.addEventListener(id === 'lv82StartDateTime' ? 'input' : 'change', () => { try { lv82UpdateHorizonExplain(); } catch(e) {} });
      }
    });
    // Quick forecast / summary / export buttons are already bound by LV8 and call the latest LV8.2 functions.
    const append = $('lv82AppendForecastBtn');
    if (append && !append.dataset.boundLv82) {
      append.dataset.boundLv82 = '1';
      append.addEventListener('click', () => { try { lv82AppendForecastToRam(); } catch(e) { log('Lỗi cập nhật forecast vào RAM LV8.5: '+e.message); } });
    }
    if ($('lv8SummaryBtn')) $('lv8SummaryBtn').disabled = false;
    if ($('lv8QuickForecastBtn')) $('lv8QuickForecastBtn').disabled = false;
    lv82UpdateHorizonExplain();
    log('Sẵn sàng LV8.5: có thể dự báo theo ngày hiện tại/ngày giờ tùy chọn, chèn forecast vào RAM và hiển thị thời gian hoàn thành thao tác.');
  } catch(e) { log('Lỗi khởi tạo LV8.5: ' + e.message); }
}, 0);
// ====================== END LV8.2 TARGET-DATE FORECAST + APPEND FORECAST TO RAM ======================

// ====================== LV8.5 ACTION STATUS / RUNTIME NOTIFICATION ======================
(function(){
  function lv83TimeText(d){
    try { return (d || new Date()).toLocaleString('vi-VN', {hour12:false}); } catch(e) { return String(d || new Date()); }
  }
  function lv83Duration(ms){
    ms = Math.max(0, Math.round(ms || 0));
    if (ms < 1000) return ms + ' ms';
    const s = ms / 1000;
    if (s < 60) return s.toFixed(1) + ' giây';
    const m = Math.floor(s / 60);
    const r = Math.round(s % 60);
    return m + ' phút ' + r + ' giây';
  }
  function lv83Status(kind, title, detail){
    const box = $('lv83ActionStatus') || $('lv8SummaryBox') || $('lv8HorizonExplainBox');
    if (!box) return;
    const cls = kind === 'ok' ? 'ok' : (kind === 'err' ? 'bad' : (kind === 'warn' ? 'warn' : ''));
    const label = kind === 'ok' ? 'HOÀN THÀNH' : (kind === 'err' ? 'LỖI' : (kind === 'warn' ? 'CHÚ Ý' : 'ĐANG THỰC HIỆN'));
    box.innerHTML = `<span class="pill ${cls}">${label}</span><span class="pill">${escapeHtml(title || '')}</span><span>${escapeHtml(detail || '')}</span>`;
  }
  function lv83Disable(ids, disabled){
    ids.forEach(id => { const el = $(id); if (el) el.disabled = disabled; });
  }
  function lv83Run(label, fn, after){
    const start = new Date();
    lv83Status('run', label, `Bắt đầu: ${lv83TimeText(start)}. Vui lòng chờ...`);
    lv83Disable(['lv8QuickForecastBtn','lv8SummaryBtn','lv82AppendForecastBtn','lv8ExportSummaryBtn'], true);
    setTimeout(() => {
      try {
        const result = fn ? fn() : null;
        const end = new Date();
        let detail = after ? after(result, start, end) : '';
        if (!detail) detail = `Xong lúc ${lv83TimeText(end)}. Thời gian thực hiện: ${lv83Duration(end - start)}.`;
        lv83Status('ok', label, detail);
      } catch(e) {
        const end = new Date();
        lv83Status('err', label, `Lỗi lúc ${lv83TimeText(end)} sau ${lv83Duration(end - start)}: ${e.message}`);
        log(`Lỗi ${label}: ` + e.message);
      } finally {
        lv83Disable(['lv8QuickForecastBtn','lv8SummaryBtn','lv82AppendForecastBtn','lv8ExportSummaryBtn'], false);
        const exp = $('lv8ExportSummaryBtn');
        if (exp) exp.disabled = !(state.lv8?.summaryRows || []).length;
        const append = $('lv82AppendForecastBtn');
        if (append) append.disabled = !(state.forecastRows || []).length;
      }
    }, 60);
  }
  function lv83CloneButton(id, handler){
    const old = $(id);
    if (!old || old.dataset.boundLv83 === '1') return;
    const neu = old.cloneNode(true);
    neu.dataset.boundLv83 = '1';
    neu.disabled = old.disabled;
    old.parentNode.replaceChild(neu, old);
    neu.addEventListener('click', handler);
  }
  function lv83InstallHandlers(){
    try {
      updateWorkflowLabelsLV8 = function(){
        if ($('versionInfo')) $('versionInfo').innerHTML = '<span class="pill modeBadge">LV8.5</span><span class="pill ok">Có thông báo trạng thái, thời điểm xong và thời gian thực hiện cho báo cáo/RAM</span>';
        const title = document.querySelector('h1');
        if (title) title.textContent = 'SCADA Load Forecast Offline PWA LV8.5';
      };
      updateWorkflowLabelsLV8();
    } catch(e) {}
    lv83CloneButton('lv8QuickForecastBtn', () => lv83Run('Dự báo LV8.5', () => lv8QuickForecast(), (result,start,end) => {
      const rows = (state.forecastRows || []).length;
      const st = $('stationSelect')?.value || '__ALL__';
      return `Xong lúc ${lv83TimeText(end)}. Thời gian thực hiện: ${lv83Duration(end-start)}. Đã tạo ${rows} dòng forecast cho ${st}.`;
    }));
    lv83CloneButton('lv8SummaryBtn', () => lv83Run('Tạo báo cáo đa cấp LV8.5', () => lv8CreateSummary(true), (result,start,end) => {
      const sRows = (state.lv8?.summaryRows || []).length;
      const fRows = (state.forecastRows || []).length;
      return `Xong lúc ${lv83TimeText(end)}. Thời gian thực hiện: ${lv83Duration(end-start)}. Báo cáo có ${sRows} dòng tổng hợp từ ${fRows} dòng forecast.`;
    }));
    lv83CloneButton('lv82AppendForecastBtn', () => lv83Run('Cập nhật forecast vào dữ liệu RAM', () => lv82AppendForecastToRam(), (result,start,end) => {
      result = result || (state.lv8?.appendedForecastRows || []).slice(-1)[0] || {};
      return `Xong lúc ${lv83TimeText(end)}. Thời gian thực hiện: ${lv83Duration(end-start)}. Thêm ${result.added || 0}, ghi đè ${result.overwritten || 0}, bỏ qua ${result.skipped || 0}. Tổng dòng RAM hiện có: ${(state.rawRows || []).length}.`;
    }));
    lv83CloneButton('lv8ExportSummaryBtn', () => lv83Run('Xuất forecast_summary_lv8_5.csv', () => exportLV8Summary(), (result,start,end) => {
      const sRows = (state.lv8?.summaryRows || []).length;
      return `Xong lúc ${lv83TimeText(end)}. Thời gian thực hiện: ${lv83Duration(end-start)}. Đã gọi xuất báo cáo ${sRows} dòng.`;
    }));
    ['lv8QuickForecastBtn','lv8SummaryBtn'].forEach(id => { const el=$(id); if (el) el.disabled=false; });
    const append = $('lv82AppendForecastBtn'); if (append) append.disabled = !(state.forecastRows || []).length;
    const exp = $('lv8ExportSummaryBtn'); if (exp) exp.disabled = !(state.lv8?.summaryRows || []).length;
    lv83Status('ok', 'Sẵn sàng LV8.5', 'Khi bấm dự báo, tạo báo cáo hoặc cập nhật RAM, trạng thái và thời gian hoàn thành sẽ hiển thị tại đây.');
    log('Sẵn sàng LV8.5: đã bổ sung thông báo thời gian thực hiện cho dự báo/báo cáo/cập nhật RAM.');
  }
  setTimeout(lv83InstallHandlers, 150);
})();
// ====================== END LV8.5 ACTION STATUS / RUNTIME NOTIFICATION ======================

// ====================== LV8.5 FORECAST QUALITY EVALUATION ======================
(function(){
  function lv84NowText(d=new Date()) { return d.toLocaleString('vi-VN'); }
  function lv84Duration(ms) { return ms < 1000 ? `${ms} ms` : `${(ms/1000).toFixed(2)} giây`; }
  function lv84Status(kind, title, detail) {
    const box = $('lv84EvalStatus'); if (!box) return;
    const cls = kind === 'ok' ? 'ok' : (kind === 'err' ? 'bad' : (kind === 'warn' ? 'warn' : ''));
    const label = kind === 'ok' ? 'HOÀN THÀNH' : (kind === 'err' ? 'LỖI' : (kind === 'warn' ? 'CHÚ Ý' : 'ĐANG THỰC HIỆN'));
    box.innerHTML = `<span class="pill ${cls}">${label}</span><span class="pill">${escapeHtml(title || '')}</span><span>${escapeHtml(detail || '')}</span>`;
  }
  function lv84DisplayName(name) { return headerDisplayName ? headerDisplayName(name) : name; }
  function lv84NormalizeHeader(h) { return norm(String(h || '')).replace(/\s+/g,'_'); }
  function lv84FindCol(headers, groups) {
    const hs = (headers || []).map(h => ({raw:h, n: lv84NormalizeHeader(h)}));
    for (const g of groups) {
      for (const item of hs) {
        if (g.exact && g.exact.includes(item.n)) return item.raw;
      }
      for (const item of hs) {
        if (g.includes && g.includes.some(x => item.n.includes(x))) return item.raw;
      }
    }
    return '';
  }
  function lv84PickColumns(headers, kind) {
    const time = lv84FindCol(headers, [
      {exact:['thoi_gian','time','timestamp','datetime','date_time','ngay_gio']},
      {includes:['thoi_gian','timestamp','ngay_gio','datetime']}
    ]);
    const station = lv84FindCol(headers, [
      {exact:['lv6_chi_danh_chuan','lv6_chi_danh_du_bao','chi_danh','station','tram_lo','tram_lo_khu_vuc','doi_tuong','object','feeder','lo','tram']},
      {includes:['chi_danh','station','tram_lo','doi_tuong','feeder']}
    ]);
    const actualP = lv84FindCol(headers, [
      {exact:['p_thuc_te','actual_p','actual_p_mw','p_actual','thuc_te_p','p_mw','p','cong_suat_p','cong_suat','mw']},
      {includes:['p_thuc_te','actual','cong_suat_p','cong_suat','p_mw']}
    ]);
    const forecastP = lv84FindCol(headers, [
      {exact:['forecast_p_mw','p_du_bao','du_bao_p','p_forecast','forecast','forecast_p','yhat','p_du_bao_mw','du_bao_mw']},
      {includes:['forecast_p','forecast','du_bao','du_bao_p','yhat']},
      {exact:['p_mw','p','cong_suat_p','cong_suat','mw']}
    ]);
    return {time, station, value: kind === 'forecast' ? forecastP : actualP};
  }
  function lv84KeyTime(v) {
    const d = parseTime(v);
    return d ? fmtTime(d) : '';
  }
  function lv84KeyStation(v) {
    return String(v ?? '').trim() || '__ALL__';
  }
  async function lv84ReadFile(file) {
    if (!file) throw new Error('Chưa chọn đủ file dữ liệu.');
    if (isSpreadsheetFile(file.name)) {
      if (!window.XLSX || !XLSX.read || !XLSX.utils) throw new Error('Thiếu thư viện SheetJS/XLSX để đọc Excel.');
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, {type:'array', cellDates:false});
      const sheet = wb.SheetNames && wb.SheetNames[0];
      if (!sheet) throw new Error('File Excel không có sheet.');
      const parsed = workbookToParsed(wb, sheet);
      parsed.fileName = file.name;
      parsed.sheetName = sheet;
      return parsed;
    }
    const text = await file.text();
    const parsed = parseDataFileText(text, file.name);
    parsed.fileName = file.name;
    return parsed;
  }
  function lv84RowsFromParsed(parsed, kind) {
    const headers = parsed.headers || [];
    const cols = lv84PickColumns(headers, kind);
    if (!cols.time) throw new Error(`${kind === 'forecast' ? 'Dữ liệu dự báo' : 'Dữ liệu thực tế'} thiếu cột thời gian.`);
    if (!cols.value) throw new Error(`${kind === 'forecast' ? 'Dữ liệu dự báo' : 'Dữ liệu thực tế'} thiếu cột giá trị ${kind === 'forecast' ? 'P dự báo' : 'P thực tế'}.`);
    const rows = (parsed.data || []).map((r, idx) => {
      const time = lv84KeyTime(r[cols.time]);
      const p = parseNumber(r[cols.value]);
      const station = cols.station ? lv84KeyStation(r[cols.station]) : '__ALL__';
      return {idx, time, station, p, raw:r};
    }).filter(r => r.time && Number.isFinite(r.p));
    return {rows, cols, headers};
  }
  function lv84RowsFromRamActual() {
    if (!state.rows.length) normalizeRows();
    return {rows: state.rows.map((r, idx) => ({idx, time: fmtTime(r.time), station: lv84KeyStation(r.station), p: r.p, raw: r.raw})), cols:{time:'RAM', station:'station', value:'P'}, headers: state.headers};
  }
  function lv84RowsFromRamForecast() {
    const rows = (state.forecastRows || []).map((r, idx) => ({idx, time: lv84KeyTime(r.time), station: lv84KeyStation(r.station), p: parseNumber(r.forecast_p_mw), raw: r})).filter(r => r.time && Number.isFinite(r.p));
    return {rows, cols:{time:'time', station:'station', value:'forecast_p_mw'}, headers: rows.length ? Object.keys(rows[0].raw || {}) : []};
  }
  function lv84BuildMap(rows) {
    const map = new Map();
    for (const r of rows) {
      const key = `${r.time}||${r.station || '__ALL__'}`;
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(r);
      // Also keep a fallback all-station key if source has no station
      if (r.station === '__ALL__') {
        const k2 = `${r.time}||`;
        if (!map.has(k2)) map.set(k2, []);
        map.get(k2).push(r);
      }
    }
    return map;
  }
  function lv84Evaluate(actualSet, forecastSet, opts={}) {
    const actualRows = actualSet.rows || [];
    const forecastRows = forecastSet.rows || [];
    const aMap = lv84BuildMap(actualRows);
    const rows = [];
    const mapeLimit = parseNumber($('lv84MapeLimit')?.value) || 5;
    const maeWarn = parseNumber($('lv84MaeWarnMw')?.value) || 2;
    for (const f of forecastRows) {
      const key = `${f.time}||${f.station || '__ALL__'}`;
      let list = aMap.get(key);
      if (!list && (f.station && f.station !== '__ALL__')) list = aMap.get(`${f.time}||__ALL__`) || aMap.get(`${f.time}||`);
      if (!list || !list.length) {
        rows.push({time:f.time, station:f.station, actual_p_mw:'', forecast_p_mw:+f.p.toFixed(6), error_mw:'', abs_error_mw:'', error_pct:'', mape_pct:'', trang_thai:'Không có dữ liệu thực tế khớp', model_used:f.raw?.model_used || '', strategy_method:f.raw?.strategy_method || ''});
        continue;
      }
      const a = list[0];
      const err = f.p - a.p;
      const abs = Math.abs(err);
      const pct = a.p !== 0 ? err / a.p * 100 : NaN;
      const mape = Number.isFinite(pct) ? Math.abs(pct) : NaN;
      let status = 'Đạt';
      if (!Number.isFinite(mape)) status = 'Không tính % do P thực tế = 0';
      else if (mape > mapeLimit || abs > maeWarn) status = 'Cần kiểm tra';
      rows.push({
        time:f.time,
        station:f.station,
        actual_p_mw:+a.p.toFixed(6),
        forecast_p_mw:+f.p.toFixed(6),
        error_mw:+err.toFixed(6),
        abs_error_mw:+abs.toFixed(6),
        error_pct:Number.isFinite(pct)? +pct.toFixed(3) : '',
        mape_pct:Number.isFinite(mape)? +mape.toFixed(3) : '',
        trang_thai:status,
        model_used:f.raw?.model_used || '',
        strategy_method:f.raw?.strategy_method || '',
        actual_source_row:a.idx,
        forecast_source_row:f.idx
      });
    }
    return rows;
  }
  function lv84Agg(rows, groupFn) {
    const groups = new Map();
    for (const r of rows) {
      if (!Number.isFinite(parseNumber(r.abs_error_mw))) continue;
      const key = groupFn(r);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(r);
    }
    const out = [];
    for (const [key, arr] of groups.entries()) {
      const n = arr.length;
      const mae = mean(arr.map(r => parseNumber(r.abs_error_mw)));
      const mape = mean(arr.map(r => parseNumber(r.mape_pct)).filter(Number.isFinite));
      const rmse = Math.sqrt(mean(arr.map(r => Math.pow(parseNumber(r.error_mw),2)).filter(Number.isFinite)));
      const bias = mean(arr.map(r => parseNumber(r.error_mw)).filter(Number.isFinite));
      let maxRow = arr[0];
      for (const r of arr) if (parseNumber(r.abs_error_mw) > parseNumber(maxRow.abs_error_mw)) maxRow = r;
      out.push({group:key, n, MAE_MW:+mae.toFixed(4), MAPE_pct:Number.isFinite(mape)? +mape.toFixed(3) : '', RMSE_MW:+rmse.toFixed(4), BIAS_MW:Number.isFinite(bias)? +bias.toFixed(4):'', max_abs_error_mw:+parseNumber(maxRow.abs_error_mw).toFixed(4), max_error_time:maxRow.time, ket_luan:(Number.isFinite(mape) && mape <= (parseNumber($('lv84MapeLimit')?.value)||5)) ? 'Đạt ngưỡng MAPE' : 'Cần kiểm tra'});
    }
    return out.sort((a,b) => String(a.group).localeCompare(String(b.group),'vi'));
  }
  function lv84Summarize(rows) {
    const matched = rows.filter(r => Number.isFinite(parseNumber(r.abs_error_mw)));
    const missing = rows.length - matched.length;
    const byStation = lv84Agg(rows, r => r.station || '__ALL__').map(r => ({level:'station', ...r}));
    const byDate = lv84Agg(rows, r => String(r.time).slice(0,10)).map(r => ({level:'date', ...r}));
    const total = lv84Agg(rows, r => '__TOTAL__').map(r => ({level:'total', ...r}));
    return {total, byStation, byDate, all:[...total, ...byStation, ...byDate], matched, missing, detailCount:rows.length};
  }
  function lv84RenderEvaluation() {
    const detail = state.lv84?.detailRows || [];
    const summary = state.lv84?.summaryRows || [];
    const matched = state.lv84?.summary?.matched || 0;
    const missing = state.lv84?.summary?.missing || 0;
    const statusBox = $('lv84EvalSummaryBox');
    if (statusBox) {
      const total = summary.find(r => r.level === 'total') || {};
      statusBox.innerHTML = `<span class="pill ok">Kết quả LV8.5</span><span class="pill">Mốc khớp: ${matched}</span><span class="pill">Không khớp: ${missing}</span><span class="pill">MAE: ${escapeHtml(total.MAE_MW ?? '-')} MW</span><span class="pill">MAPE: ${escapeHtml(total.MAPE_pct ?? '-')}%</span><span class="pill">RMSE: ${escapeHtml(total.RMSE_MW ?? '-')} MW</span>`;
    }
    const tbl = $('lv84EvalTable');
    if (tbl) renderTableInBox(tbl, detail, ['time','station','actual_p_mw','forecast_p_mw','error_mw','abs_error_mw','error_pct','mape_pct','trang_thai','model_used','strategy_method'], 1000);
    const b1=$('lv84ExportDetailBtn'), b2=$('lv84ExportSummaryBtn');
    if (b1) b1.disabled = !detail.length;
    if (b2) b2.disabled = !summary.length;
    lv84RenderHistory();
  }
  function lv84SaveHistory(sourceActual, sourceForecast, summary) {
    const total = (summary.all || []).find(r => r.level === 'total') || {};
    const rec = {
      id: 'EV' + Date.now(),
      time: new Date().toISOString(),
      actual_file: sourceActual || '',
      forecast_file: sourceForecast || '',
      n_matched: summary.matched || 0,
      n_missing: summary.missing || 0,
      MAE_MW: total.MAE_MW ?? '',
      MAPE_pct: total.MAPE_pct ?? '',
      RMSE_MW: total.RMSE_MW ?? '',
      max_abs_error_mw: total.max_abs_error_mw ?? '',
      max_error_time: total.max_error_time ?? '',
      ket_luan: total.ket_luan ?? ''
    };
    const arr = lv84LoadHistory();
    arr.unshift(rec);
    localStorage.setItem('scada_load_forecast_lv84_eval_history', JSON.stringify(arr.slice(0, 200)));
  }
  function lv84LoadHistory() {
    try { return JSON.parse(localStorage.getItem('scada_load_forecast_lv84_eval_history') || '[]'); }
    catch(e) { return []; }
  }
  function lv84RenderHistory() {
    const box = $('lv84HistoryBox'); if (!box) return;
    const hist = lv84LoadHistory();
    if (!hist.length) { box.innerHTML = '<table><tbody><tr><td>Chưa có lịch sử đánh giá.</td></tr></tbody></table>'; return; }
    renderTableInBox(box, hist, ['time','actual_file','forecast_file','n_matched','n_missing','MAE_MW','MAPE_pct','RMSE_MW','max_abs_error_mw','max_error_time','ket_luan'], 50);
  }
  async function lv84EvaluateFromFiles() {
    const start = new Date();
    lv84Status('run', 'Đánh giá từ 2 file', `Bắt đầu: ${lv84NowText(start)}. Đang đọc file thực tế và file dự báo...`);
    const fActual = $('lv84ActualFile')?.files?.[0];
    const fForecast = $('lv84ForecastFile')?.files?.[0];
    if (!fActual || !fForecast) throw new Error('Cần chọn đủ 2 file: dữ liệu thực tế và dữ liệu dự báo.');
    const [actualParsed, forecastParsed] = await Promise.all([lv84ReadFile(fActual), lv84ReadFile(fForecast)]);
    const actualSet = lv84RowsFromParsed(actualParsed, 'actual');
    const forecastSet = lv84RowsFromParsed(forecastParsed, 'forecast');
    const detail = lv84Evaluate(actualSet, forecastSet);
    const summary = lv84Summarize(detail);
    state.lv84 = {detailRows:detail, summaryRows:summary.all, summary:{matched:summary.matched.length, missing:summary.missing, detailCount:summary.detailCount}, actualFile:fActual.name, forecastFile:fForecast.name, evaluatedAt:new Date().toISOString(), actualCols:actualSet.cols, forecastCols:forecastSet.cols};
    lv84SaveHistory(fActual.name, fForecast.name, {...summary, matched:summary.matched.length});
    lv84RenderEvaluation();
    const end = new Date();
    lv84Status('ok', 'Đánh giá từ 2 file', `Xong lúc ${lv84NowText(end)}. Thời gian thực hiện: ${lv84Duration(end-start)}. Khớp ${summary.matched.length}/${detail.length} mốc, không khớp ${summary.missing}.`);
    log(`LV8.5 đã đánh giá sai số từ 2 file: khớp ${summary.matched.length}/${detail.length} mốc.`);
  }
  function lv84EvaluateFromRam() {
    const start = new Date();
    lv84Status('run', 'Đánh giá RAM + forecast', `Bắt đầu: ${lv84NowText(start)}. Đang ghép dữ liệu RAM và forecast hiện có...`);
    if (!state.rows.length) normalizeRows();
    if (!(state.forecastRows || []).length) throw new Error('Chưa có forecastRows trong RAM. Hãy dự báo trước hoặc nạp file dự báo.');
    const actualSet = lv84RowsFromRamActual();
    const forecastSet = lv84RowsFromRamForecast();
    const detail = lv84Evaluate(actualSet, forecastSet);
    const summary = lv84Summarize(detail);
    state.lv84 = {detailRows:detail, summaryRows:summary.all, summary:{matched:summary.matched.length, missing:summary.missing, detailCount:summary.detailCount}, actualFile:'RAM hiện có', forecastFile:'forecastRows hiện có', evaluatedAt:new Date().toISOString()};
    lv84SaveHistory('RAM hiện có', 'forecastRows hiện có', {...summary, matched:summary.matched.length});
    lv84RenderEvaluation();
    const end = new Date();
    lv84Status('ok', 'Đánh giá RAM + forecast', `Xong lúc ${lv84NowText(end)}. Thời gian thực hiện: ${lv84Duration(end-start)}. Khớp ${summary.matched.length}/${detail.length} mốc, không khớp ${summary.missing}.`);
    log(`LV8.5 đã đánh giá sai số từ RAM + forecast: khớp ${summary.matched.length}/${detail.length} mốc.`);
  }
  function lv84ExportDetail() {
    const rows = state.lv84?.detailRows || [];
    if (!rows.length) { log('Chưa có chi tiết sai số LV8.5 để xuất.'); return; }
    saveTextFile('forecast_error_detail_lv8_5.csv', exportCSVContent(rows, ['time','station','actual_p_mw','forecast_p_mw','error_mw','abs_error_mw','error_pct','mape_pct','trang_thai','model_used','strategy_method','actual_source_row','forecast_source_row']), 'text/csv;charset=utf-8');
    log(`Đã xuất forecast_error_detail_lv8_5.csv gồm ${rows.length} dòng.`);
  }
  function lv84ExportSummary() {
    const rows = state.lv84?.summaryRows || [];
    if (!rows.length) { log('Chưa có tổng hợp sai số LV8.5 để xuất.'); return; }
    saveTextFile('forecast_error_summary_lv8_5.csv', exportCSVContent(rows, ['level','group','n','MAE_MW','MAPE_pct','RMSE_MW','BIAS_MW','max_abs_error_mw','max_error_time','ket_luan']), 'text/csv;charset=utf-8');
    log(`Đã xuất forecast_error_summary_lv8_5.csv gồm ${rows.length} dòng.`);
  }
  function lv84ExportHistory() {
    const rows = lv84LoadHistory();
    if (!rows.length) { log('Chưa có lịch sử đánh giá LV8.5 để xuất.'); return; }
    saveTextFile('forecast_evaluation_history_lv8_5.csv', exportCSVContent(rows, ['time','actual_file','forecast_file','n_matched','n_missing','MAE_MW','MAPE_pct','RMSE_MW','max_abs_error_mw','max_error_time','ket_luan']), 'text/csv;charset=utf-8');
    log(`Đã xuất forecast_evaluation_history_lv8_5.csv gồm ${rows.length} dòng.`);
  }
  function lv84ClearHistory() {
    localStorage.removeItem('scada_load_forecast_lv84_eval_history');
    lv84RenderHistory();
    log('Đã xóa lịch sử đánh giá LV8.5 trong trình duyệt.');
  }
  function lv84Bind() {
    try {
      if ($('versionInfo')) $('versionInfo').innerHTML = '<span class="pill modeBadge">LV8.5</span><span class="pill ok">Đánh giá sai số dự báo + lưu lịch sử</span>';
      const title = document.querySelector('h1'); if (title) title.textContent = 'SCADA Load Forecast Offline PWA LV8.5';
      document.title = 'SCADA Load Forecast Offline PWA LV8.5';
    } catch(e) {}
    const bind = (id, fn) => { const el=$(id); if (el && el.dataset.boundLv84 !== '1') { el.dataset.boundLv84='1'; el.addEventListener('click', async () => { try { await fn(); } catch(e) { lv84Status('err', id, e.message); log('Lỗi LV8.5: ' + e.message); } }); } };
    bind('lv84EvalFilesBtn', lv84EvaluateFromFiles);
    bind('lv84EvalRamBtn', () => lv84EvaluateFromRam());
    bind('lv84ExportDetailBtn', () => lv84ExportDetail());
    bind('lv84ExportSummaryBtn', () => lv84ExportSummary());
    bind('lv84ExportHistoryBtn', () => lv84ExportHistory());
    bind('lv84ClearHistoryBtn', () => lv84ClearHistory());
    lv84RenderHistory();
    lv84Status('ok', 'Sẵn sàng LV8.5', 'Có thể nạp 2 file thực tế/dự báo để tính sai số từng thời điểm và lưu lịch sử đánh giá.');
    log('Sẵn sàng LV8.5: đánh giá chất lượng dự báo từ 2 file hoặc từ RAM + forecast hiện có.');
  }
  setTimeout(lv84Bind, 250);
})();
// ====================== END LV8.5 FORECAST QUALITY EVALUATION ======================


// ====================== LV8.5 MODEL CALIBRATION FROM EVALUATION ======================
(function(){
  function lv85NowText(d=new Date()) { return d.toLocaleString('vi-VN'); }
  function lv85Duration(ms) { return ms < 1000 ? `${ms} ms` : `${(ms/1000).toFixed(2)} giây`; }
  function lv85Status(kind, title, detail) {
    const box = $('lv85CalibrationStatus'); if (!box) return;
    const cls = kind === 'ok' ? 'ok' : (kind === 'err' ? 'bad' : (kind === 'warn' ? 'warn' : ''));
    const label = kind === 'ok' ? 'HOÀN THÀNH' : (kind === 'err' ? 'LỖI' : (kind === 'warn' ? 'CHÚ Ý' : 'ĐANG THỰC HIỆN'));
    box.innerHTML = `<span class="pill ${cls}">${label}</span><span class="pill">${escapeHtml(title || '')}</span><span>${escapeHtml(detail || '')}</span>`;
  }
  function lv85HourFromTime(t) {
    const d = parseTime(t);
    return d ? d.getHours() : '';
  }
  function lv85DateTypeFromTime(t) {
    const d = parseTime(t);
    if (!d) return 'khong_ro';
    if (holidayByRules(d)) return 'ngay_nghi_le';
    const day = d.getDay();
    if (day === 0) return 'chu_nhat';
    if (day === 6) return 'thu_bay';
    return 'ngay_thuong';
  }
  function lv85GroupStats(rows, level, groupFn) {
    const groups = new Map();
    for (const r of rows) {
      const e = parseNumber(r.error_mw);
      const abs = parseNumber(r.abs_error_mw);
      const mape = parseNumber(r.mape_pct);
      const fp = parseNumber(r.forecast_p_mw);
      const ap = parseNumber(r.actual_p_mw);
      if (!Number.isFinite(e) || !Number.isFinite(abs)) continue;
      const key = groupFn(r);
      if (!key) continue;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push({...r, _error:e, _abs:abs, _mape:mape, _forecast:fp, _actual:ap});
    }
    const out = [];
    for (const [key, arr] of groups.entries()) {
      const n = arr.length;
      const biasForecastMinusActual = mean(arr.map(r => r._error).filter(Number.isFinite));
      const correctionMw = Number.isFinite(biasForecastMinusActual) ? -biasForecastMinusActual : 0;
      const ratioArr = arr.map(r => (Number.isFinite(r._forecast) && r._forecast !== 0 && Number.isFinite(r._actual)) ? ((r._actual / r._forecast - 1) * 100) : NaN).filter(Number.isFinite);
      const correctionPct = ratioArr.length ? mean(ratioArr) : 0;
      const mae = mean(arr.map(r => r._abs).filter(Number.isFinite));
      const mape = mean(arr.map(r => r._mape).filter(Number.isFinite));
      const rmse = Math.sqrt(mean(arr.map(r => Math.pow(r._error,2)).filter(Number.isFinite)));
      const sampleTimes = arr.slice(0, 3).map(r => r.time).join(' | ');
      out.push({
        level, group:key, n,
        correction_mw:+correctionMw.toFixed(6),
        correction_pct:+correctionPct.toFixed(6),
        BIAS_forecast_minus_actual_MW:+biasForecastMinusActual.toFixed(6),
        MAE_MW:Number.isFinite(mae)? +mae.toFixed(6) : '',
        MAPE_pct:Number.isFinite(mape)? +mape.toFixed(3) : '',
        RMSE_MW:Number.isFinite(rmse)? +rmse.toFixed(6) : '',
        sample_times:sampleTimes
      });
    }
    return out.sort((a,b) => String(a.group).localeCompare(String(b.group),'vi'));
  }
  function lv85BuildCalibration() {
    const start = new Date();
    lv85Status('run', 'Tạo bảng hiệu chỉnh', `Bắt đầu: ${lv85NowText(start)}. Đang đọc kết quả đánh giá LV8.5...`);
    const detail = state.lv84?.detailRows || [];
    if (!detail.length) throw new Error('Chưa có kết quả đánh giá ở Mục 10.2. Hãy đánh giá từ 2 file hoặc RAM + forecast trước.');
    const valid = detail.filter(r => Number.isFinite(parseNumber(r.error_mw)) && Number.isFinite(parseNumber(r.actual_p_mw)) && Number.isFinite(parseNumber(r.forecast_p_mw)));
    if (!valid.length) throw new Error('Kết quả đánh giá không có mốc nào khớp được dữ liệu thực tế và dự báo.');
    const minSamples = Math.max(1, Math.floor(parseNumber($('lv85MinSamples')?.value) || 3));
    const rows = [];
    rows.push(...lv85GroupStats(valid, 'total', () => '__TOTAL__'));
    rows.push(...lv85GroupStats(valid, 'station', r => r.station || '__ALL__'));
    rows.push(...lv85GroupStats(valid, 'hour', r => String(lv85HourFromTime(r.time)).padStart(2,'0')));
    rows.push(...lv85GroupStats(valid, 'station_hour', r => `${r.station || '__ALL__'}||${String(lv85HourFromTime(r.time)).padStart(2,'0')}`));
    rows.push(...lv85GroupStats(valid, 'day_type', r => lv85DateTypeFromTime(r.time)));
    rows.push(...lv85GroupStats(valid, 'station_day_type', r => `${r.station || '__ALL__'}||${lv85DateTypeFromTime(r.time)}`));
    const usable = rows.map(r => ({...r, usable: r.n >= minSamples ? 1 : 0, min_samples:minSamples}));
    const calibration = {
      version:'LV8.5',
      createdAt:new Date().toISOString(),
      source:'MUC_10_2_EVALUATION',
      actualFile: state.lv84?.actualFile || '',
      forecastFile: state.lv84?.forecastFile || '',
      applyScope:$('lv85ApplyScope')?.value || 'station_hour_station_hour_total',
      correctionMode:$('lv85CorrectionMode')?.value || 'mw',
      enabled:($('lv85EnableMode')?.value || 'enabled') === 'enabled',
      minSamples,
      maxPct:Math.max(1, parseNumber($('lv85MaxPct')?.value) || 15),
      rows:usable,
      note:'Hiệu chỉnh tạo từ sai số thực tế: error_mw = forecast - actual, correction_mw = -BIAS. Không dùng dữ liệu dự báo để huấn luyện model thật.'
    };
    state.lv85 = state.lv85 || {};
    state.lv85.calibration = calibration;
    state.lv85.calibrationRows = usable;
    state.lv85.applied = false;
    lv85RenderCalibration();
    const end = new Date();
    lv85Status('ok', 'Tạo bảng hiệu chỉnh', `Xong lúc ${lv85NowText(end)}. Thời gian: ${lv85Duration(end-start)}. Tạo ${usable.length} dòng hiệu chỉnh từ ${valid.length} mốc sai số hợp lệ.`);
    log(`LV8.5 đã tạo bảng hiệu chỉnh ${usable.length} dòng từ ${valid.length} mốc đánh giá.`);
  }
  function lv85RenderCalibration() {
    const rows = state.lv85?.calibrationRows || state.model?.calibrationLV85?.rows || [];
    const cal = state.lv85?.calibration || state.model?.calibrationLV85 || null;
    const summaryBox = $('lv85CalibrationSummaryBox');
    if (summaryBox) {
      if (!rows.length) summaryBox.innerHTML = '<span class="pill">Chưa có bảng hiệu chỉnh</span>';
      else {
        const usable = rows.filter(r => parseNumber(r.usable) === 1).length;
        const total = rows.find(r => r.level === 'total');
        summaryBox.innerHTML = `<span class="pill ok">Bảng hiệu chỉnh LV8.5</span><span class="pill">Tổng dòng: ${rows.length}</span><span class="pill">Dòng đủ mẫu: ${usable}</span><span class="pill">Chế độ: ${escapeHtml(cal?.correctionMode || '')}</span><span class="pill">Phạm vi: ${escapeHtml(cal?.applyScope || '')}</span><span class="pill">Bù chung: ${escapeHtml(total?.correction_mw ?? '-')} MW / ${escapeHtml(total?.correction_pct ?? '-')}%</span><span class="pill ${cal?.enabled ? 'ok':'warn'}">${cal?.enabled ? 'Đang bật áp dụng' : 'Chưa bật áp dụng'}</span>`;
      }
    }
    const table = $('lv85CalibrationTable');
    if (table) renderTableInBox(table, rows, ['usable','level','group','n','correction_mw','correction_pct','BIAS_forecast_minus_actual_MW','MAE_MW','MAPE_pct','RMSE_MW','sample_times'], 1000);
    const has = rows.length > 0;
    ['lv85ApplyCalibrationBtn','lv85ExportCalibrationBtn','lv85ExportReportBtn'].forEach(id => { const el=$(id); if (el) el.disabled = !has; });
    const mBtn = $('lv85ExportModelBtn'); if (mBtn) mBtn.disabled = !(state.model && (state.lv85?.calibration || state.model?.calibrationLV85));
  }
  function lv85ApplyCalibrationToModel() {
    const start = new Date();
    lv85Status('run', 'Áp dụng hiệu chỉnh', `Bắt đầu: ${lv85NowText(start)}. Đang ghi hiệu chỉnh vào model vận hành...`);
    if (!state.model) throw new Error('Chưa có model vận hành đang nạp. Hãy nạp model ở Mục 10 hoặc huấn luyện ở Mục 8 trước.');
    const cal = state.lv85?.calibration;
    if (!cal || !(cal.rows || []).length) throw new Error('Chưa có bảng hiệu chỉnh. Hãy bấm tạo bảng hiệu chỉnh trước.');
    cal.enabled = ($('lv85EnableMode')?.value || 'enabled') === 'enabled';
    cal.applyScope = $('lv85ApplyScope')?.value || cal.applyScope;
    cal.correctionMode = $('lv85CorrectionMode')?.value || cal.correctionMode;
    cal.minSamples = Math.max(1, Math.floor(parseNumber($('lv85MinSamples')?.value) || cal.minSamples || 3));
    cal.maxPct = Math.max(1, parseNumber($('lv85MaxPct')?.value) || cal.maxPct || 15);
    cal.appliedAt = new Date().toISOString();
    state.model.calibrationLV85 = JSON.parse(JSON.stringify(cal));
    state.model.appVersion = 'LV8.5';
    state.lv85.applied = true;
    renderModelInfo();
    lv85RenderCalibration();
    const end = new Date();
    lv85Status('ok', 'Áp dụng hiệu chỉnh', `Xong lúc ${lv85NowText(end)}. Thời gian: ${lv85Duration(end-start)}. Model hiện tại sẽ áp dụng hiệu chỉnh LV8.5 khi dự báo.`);
    log('LV8.5 đã áp dụng hiệu chỉnh vào model đang nạp. Có thể dự báo lại hoặc xuất model_gbdt_lv8_5_calibrated.json.');
  }
  function lv85ExportCalibration() {
    const cal = state.lv85?.calibration || state.model?.calibrationLV85;
    if (!cal) { log('Chưa có calibration LV8.5 để xuất.'); return; }
    saveTextFile('calibration_lv8_5.json', JSON.stringify(cal, null, 2), 'application/json');
    log('Đã xuất calibration_lv8_5.json.');
  }
  function lv85ExportReport() {
    const rows = state.lv85?.calibrationRows || state.model?.calibrationLV85?.rows || [];
    if (!rows.length) { log('Chưa có calibration report LV8.5 để xuất.'); return; }
    saveTextFile('calibration_report_lv8_5.csv', exportCSVContent(rows, ['usable','level','group','n','correction_mw','correction_pct','BIAS_forecast_minus_actual_MW','MAE_MW','MAPE_pct','RMSE_MW','min_samples','sample_times']), 'text/csv;charset=utf-8');
    log(`Đã xuất calibration_report_lv8_5.csv gồm ${rows.length} dòng.`);
  }
  function lv85ExportCalibratedModel() {
    if (!state.model) { log('Chưa có model để xuất.'); return; }
    const cal = state.lv85?.calibration || state.model?.calibrationLV85;
    if (!cal) { log('Chưa có hiệu chỉnh LV8.5 để xuất cùng model.'); return; }
    const payload = {
      ...state.model,
      calibrationLV85:cal,
      scadaModelPackage:'SCADA_LOAD_FORECAST_OPERATIONAL_MODEL',
      appVersion:'LV8.5',
      exportSource:'MUC_8_EXPORT_MODEL',
      exportSection:'8+10.3',
      allowedImportSection:'10',
      modelKind:'OPERATIONAL_FORECAST_MODEL',
      exportedAt:new Date().toISOString(),
      importRules:{onlyImportInSection10:true, requiredExportSource:'MUC_8_EXPORT_MODEL', note:'Model đã hiệu chỉnh LV8.5 từ kết quả đánh giá Mục 10.2/10.3. Dùng được cho dữ liệu mới cùng cấu trúc/chỉ danh/phạm vi đã huấn luyện.'},
      lv85State:{calibrationApplied:true, calibrationCreatedAt:cal.createdAt || '', savedAt:new Date().toISOString()}
    };
    saveTextFile('model_gbdt_lv8_5_calibrated.json', JSON.stringify(payload, null, 2), 'application/json');
    log('Đã xuất model_gbdt_lv8_5_calibrated.json để nạp ở Mục 10 trong mạng SCADA.');
  }
  function lv85ClearCalibration() {
    state.lv85 = state.lv85 || {};
    state.lv85.calibration = null;
    state.lv85.calibrationRows = [];
    state.lv85.applied = false;
    if (state.model && state.model.calibrationLV85) delete state.model.calibrationLV85;
    lv85RenderCalibration();
    lv85Status('warn', 'Đã tắt hiệu chỉnh', 'Model hiện tại sẽ dự báo theo chiến lược/model gốc, không bù sai số LV8.5.');
    log('Đã tắt/xóa hiệu chỉnh LV8.5 khỏi model đang nạp.');
  }
  function lv85Bind() {
    try {
      if ($('versionInfo')) $('versionInfo').innerHTML = '<span class="pill modeBadge">LV8.5</span><span class="pill ok">Đánh giá sai số + hiệu chỉnh model từ thực tế</span>';
      const title = document.querySelector('h1'); if (title) title.textContent = 'SCADA Load Forecast Offline PWA LV8.5';
      document.title = 'SCADA Load Forecast Offline PWA LV8.5';
      const b = (id, fn) => { const el=$(id); if (el && el.dataset.boundLv85 !== '1') { el.dataset.boundLv85='1'; el.addEventListener('click', () => { try { fn(); } catch(e) { lv85Status('err', id, e.message); log('Lỗi LV8.5: ' + e.message); } }); } };
      b('lv85BuildCalibrationBtn', lv85BuildCalibration);
      b('lv85ApplyCalibrationBtn', lv85ApplyCalibrationToModel);
      b('lv85ExportCalibrationBtn', lv85ExportCalibration);
      b('lv85ExportModelBtn', lv85ExportCalibratedModel);
      b('lv85ExportReportBtn', lv85ExportReport);
      b('lv85ClearCalibrationBtn', lv85ClearCalibration);
      ['lv85ApplyScope','lv85CorrectionMode','lv85MinSamples','lv85MaxPct','lv85EnableMode'].forEach(id => { const el=$(id); if (el && el.dataset.boundChangeLv85 !== '1') { el.dataset.boundChangeLv85='1'; el.addEventListener('change', () => { try { if (state.lv85?.calibration) { state.lv85.calibration.applyScope=$('lv85ApplyScope')?.value || state.lv85.calibration.applyScope; state.lv85.calibration.correctionMode=$('lv85CorrectionMode')?.value || state.lv85.calibration.correctionMode; state.lv85.calibration.enabled=($('lv85EnableMode')?.value || 'enabled') === 'enabled'; state.lv85.calibration.minSamples=Math.max(1, Math.floor(parseNumber($('lv85MinSamples')?.value) || 3)); state.lv85.calibration.maxPct=Math.max(1, parseNumber($('lv85MaxPct')?.value) || 15); lv85RenderCalibration(); } } catch(e) {} }); } });
      if (state.model?.calibrationLV85) { state.lv85 = state.lv85 || {}; state.lv85.calibration = state.model.calibrationLV85; state.lv85.calibrationRows = state.model.calibrationLV85.rows || []; }
      lv85RenderCalibration();
      lv85Status('ok', 'Sẵn sàng LV8.5', 'Sau khi đánh giá ở Mục 10.2, có thể tạo hiệu chỉnh, áp dụng vào model, xuất calibration/model hiệu chỉnh.');
      log('Sẵn sàng LV8.5: hiệu chỉnh mô hình từ kết quả đánh giá sai số thực tế.');
    } catch(e) { log('Lỗi khởi tạo LV8.5: ' + e.message); }
  }
  setTimeout(lv85Bind, 700);
})();

function lv85PickCalibrationRow(cal, station, time) {
  if (!cal || !cal.enabled || !(cal.rows || []).length) return null;
  const hour = String((time instanceof Date ? time : parseTime(time))?.getHours?.() ?? '').padStart(2,'0');
  const rows = (cal.rows || []).filter(r => parseNumber(r.usable) === 1);
  const by = (level, group) => rows.find(r => r.level === level && String(r.group) === String(group));
  const scope = cal.applyScope || 'station_hour_station_hour_total';
  const stationKey = station || '__ALL__';
  const candidates = [];
  if (scope === 'station_hour_station_hour_total') candidates.push(['station_hour', `${stationKey}||${hour}`], ['station', stationKey], ['hour', hour], ['total','__TOTAL__']);
  else if (scope === 'station_hour_station_total') candidates.push(['station_hour', `${stationKey}||${hour}`], ['station', stationKey], ['total','__TOTAL__']);
  else if (scope === 'station_only') candidates.push(['station', stationKey], ['total','__TOTAL__']);
  else if (scope === 'hour_only') candidates.push(['hour', hour], ['total','__TOTAL__']);
  else candidates.push(['total','__TOTAL__']);
  for (const [lvl, grp] of candidates) { const r = by(lvl, grp); if (r) return r; }
  return null;
}

function applyCalibrationLV85(pred, station, time) {
  const cal = state.model?.calibrationLV85 || state.lv85?.calibration;
  if (!cal || !cal.enabled) return {value:pred, applied:false, delta:0, source:''};
  const row = lv85PickCalibrationRow(cal, station, time);
  if (!row) return {value:pred, applied:false, delta:0, source:'khong_co_nhom_du_mau'};
  const mode = cal.correctionMode || 'mw';
  const maxPct = Math.max(1, parseNumber(cal.maxPct) || 15);
  let delta = 0;
  if (mode === 'percent') {
    const pct = Math.max(-maxPct, Math.min(maxPct, parseNumber(row.correction_pct) || 0));
    delta = pred * pct / 100;
  } else if (mode === 'combined') {
    const mw = parseNumber(row.correction_mw) || 0;
    const pct = Math.max(-maxPct, Math.min(maxPct, parseNumber(row.correction_pct) || 0));
    const pctDelta = pred * pct / 100;
    delta = Number.isFinite(mw) && Number.isFinite(pctDelta) ? (0.7 * mw + 0.3 * pctDelta) : (mw || pctDelta || 0);
  } else {
    delta = parseNumber(row.correction_mw) || 0;
  }
  const pctCapMw = Math.abs(pred) * maxPct / 100;
  if (Number.isFinite(pctCapMw) && pctCapMw > 0 && Math.abs(delta) > pctCapMw) delta = Math.sign(delta) * pctCapMw;
  const value = Math.max(0, pred + delta);
  return {value, applied:true, delta:value - pred, source:`${row.level}:${row.group}`};
}


// Show calibration status in the existing model info panel.
const renderModelInfoBaseLV85 = renderModelInfo;
renderModelInfo = function() {
  renderModelInfoBaseLV85();
  try {
    const box = $('modelInfo');
    const cal = state.model?.calibrationLV85;
    if (box && cal) {
      const usable = (cal.rows || []).filter(r => parseNumber(r.usable) === 1).length;
      box.innerHTML += `<span class="pill ok">Calibration LV8.5: ${cal.enabled ? 'đang bật' : 'đang tắt'}</span><span class="pill">${usable} nhóm đủ mẫu</span><span class="pill">${escapeHtml(cal.correctionMode || '')}</span>`;
    }
  } catch(e) {}
};

// Override exportModel for LV8.5 so Mục 8 exports model containing calibration when applied.
function exportModel() {
  if (!state.model) return;
  const payload = {
    ...state.model,
    scadaModelPackage: 'SCADA_LOAD_FORECAST_OPERATIONAL_MODEL',
    appVersion: 'LV8.5',
    exportSource: 'MUC_8_EXPORT_MODEL',
    exportSection: '8',
    allowedImportSection: '10',
    modelKind: 'OPERATIONAL_FORECAST_MODEL',
    exportedAt: new Date().toISOString(),
    importRules: {
      onlyImportInSection10: true,
      requiredExportSource: 'MUC_8_EXPORT_MODEL',
      note: 'File này được xuất từ Mục 8 và được phép nạp ở Mục 10. Nếu có calibrationLV85, model sẽ bù sai số theo kết quả đánh giá thực tế.'
    },
    lv85Config: collectFullConfig(),
    lv85State:{strategyByStation: state.lv7?.strategyByStation || state.model?.strategyByStation || {}, calibrationApplied: !!state.model.calibrationLV85, savedAt:new Date().toISOString(), multiLevelForecast:true}
  };
  saveTextFile('model_gbdt_lv8_5_operational.json', JSON.stringify(payload, null, 2), 'application/json');
  log('Đã xuất model_gbdt_lv8_5_operational.json từ Mục 8. Nếu đã áp dụng calibration LV8.5, file này chứa hiệu chỉnh để nạp ở Mục 10.');
}
// ====================== END LV8.5 MODEL CALIBRATION FROM EVALUATION ======================

// ====================== LV8.9 FULL WIDTH LAYOUT EXTENSIONS ======================
// LV8.9:
// - Mặc định bỏ giới hạn max-width 1480px để các khung chức năng tràn hết màn hình.
// - Bổ sung nút đổi bố cục: Rộng / Siêu rộng / Cố định 1480px.
// - Tối ưu bảng: bảng dùng min-width 100%, khung bảng không bị bó ngang.

function applyLayoutModeLV86(mode) {
  const html = document.documentElement;
  html.classList.remove('layout-wide','layout-ultra','layout-comfort');
  const m = mode || localStorage.getItem('scada_lv86_layout_mode') || 'wide';
  html.classList.add('layout-' + m);
  localStorage.setItem('scada_lv86_layout_mode', m);
  const info = document.getElementById('lv86LayoutInfo');
  if (info) {
    const label = m === 'ultra' ? 'Siêu rộng: sát mép màn hình, khung gọn hơn' : (m === 'comfort' ? 'Cố định 1480px: giống bố cục cũ' : 'Rộng: tràn toàn bộ chiều ngang màn hình');
    info.innerHTML = `<span class="pill ok">Bố cục LV8.9</span><span class="pill">${escapeHtml(label)}</span>`;
  }
  try { window.dispatchEvent(new Event('resize')); } catch(e) {}
}

function injectLayoutControlsLV86() {
  let bar = document.getElementById('lv75CollapseBar');
  if (!bar) return;
  if (document.getElementById('lv86LayoutControls')) return;
  const panel = document.createElement('div');
  panel.id = 'lv86LayoutControls';
  panel.className = 'toolPanel';
  panel.innerHTML = `
    <div class="row">
      <button class="secondary" id="lv86WideBtn">Bố cục rộng toàn màn hình</button>
      <button class="secondary" id="lv86UltraBtn">Bố cục siêu rộng / gọn khung</button>
      <button class="secondary" id="lv86ComfortBtn">Bố cục cố định 1480px</button>
    </div>
    <div id="lv86LayoutInfo" class="compactNote" style="margin-top:8px"></div>
    <div class="compactNote">LV8.9 mặc định dùng bố cục rộng để các khung chức năng và bảng dữ liệu tận dụng toàn bộ chiều ngang màn hình. Chọn “cố định 1480px” nếu muốn quay lại kiểu hiển thị cũ.</div>`;
  bar.appendChild(panel);
  document.getElementById('lv86WideBtn')?.addEventListener('click', () => applyLayoutModeLV86('wide'));
  document.getElementById('lv86UltraBtn')?.addEventListener('click', () => applyLayoutModeLV86('ultra'));
  document.getElementById('lv86ComfortBtn')?.addEventListener('click', () => applyLayoutModeLV86('comfort'));
  applyLayoutModeLV86(localStorage.getItem('scada_lv86_layout_mode') || 'wide');
}

setTimeout(() => {
  try {
    applyLayoutModeLV86(localStorage.getItem('scada_lv86_layout_mode') || 'wide');
    injectLayoutControlsLV86();
    const title = document.querySelector('h1');
    if (title) title.textContent = 'SCADA Load Forecast Offline PWA LV8.9';
    document.title = 'SCADA Load Forecast Offline PWA LV8.9';
    if ($('versionInfo')) $('versionInfo').innerHTML = '<span class="pill modeBadge">LV8.9</span><span class="pill ok">Bố cục rộng toàn màn hình + bảng tự giãn ngang</span>';
    log('Sẵn sàng LV8.9: bố cục mặc định tràn hết màn hình; có thể đổi Rộng / Siêu rộng / Cố định 1480px tại Mục 0.2.');
  } catch(e) { log('Lỗi khởi tạo LV8.9 layout: ' + e.message); }
}, 50);
// ====================== END LV8.9 FULL WIDTH LAYOUT EXTENSIONS ======================

// ====================== LV8.9 SAFE DASHBOARD + OPERATION LOG ======================
// LV8.9 tiếp tục từ lõi ổn định LV8.6/LV8.8.3 Safe.
// Nguyên tắc: chỉ quan sát và ghi nhận, không khóa nút, không can thiệp luồng Mục 5/6/7/8/10.
(function(){
  const VERSION = 'LV8.9';
  const LOG_KEY = 'scada_load_forecast_lv89_operation_log';

  function lv89EnsureState(){
    state.lv89 = state.lv89 || {operationLog: [], compareRows: []};
    if (!Array.isArray(state.lv89.operationLog)) state.lv89.operationLog = lv89LoadLog();
    if (!Array.isArray(state.lv89.compareRows)) state.lv89.compareRows = [];
    return state.lv89;
  }
  function lv89Now(){ return new Date().toLocaleString('vi-VN'); }
  function lv89Iso(){ return new Date().toISOString(); }
  function lv89Num(v){ return Number.isFinite(parseNumber(v)) ? parseNumber(v) : NaN; }
  function lv89SafeText(v){ return String(v ?? ''); }
  function lv89LoadLog(){
    try { return JSON.parse(localStorage.getItem(LOG_KEY) || '[]'); } catch(e) { return []; }
  }
  function lv89SaveLog(){
    try { localStorage.setItem(LOG_KEY, JSON.stringify((state.lv89?.operationLog || []).slice(-1000))); } catch(e) {}
  }
  function lv89AddLog(action, detail='', level='info'){
    try {
      const st = lv89EnsureState();
      const row = {
        time: lv89Now(),
        iso: lv89Iso(),
        level,
        action: lv89SafeText(action).slice(0,160),
        detail: lv89SafeText(detail).slice(0,400),
        raw_rows: state.rawRows?.length || 0,
        ram_rows: state.rows?.length || 0,
        forecast_rows: state.forecastRows?.length || 0,
        quality_issues: state.qualityIssues?.length || 0,
        operation_events: state.operationEvents?.length || 0,
        model_loaded: state.model ? 1 : 0,
        calibration: state.model?.calibrationLV85 ? 1 : 0
      };
      st.operationLog.push(row);
      if (st.operationLog.length > 1000) st.operationLog = st.operationLog.slice(-1000);
      lv89SaveLog();
      lv89RenderLog();
      lv89RenderDashboard();
    } catch(e) {}
  }

  function lv89SetText(id, value){ const el=$(id); if (el) el.textContent = value; }
  function lv89DateRange(){
    const times = (state.rows || []).map(r => r.time).filter(t => t instanceof Date && !isNaN(t)).sort((a,b)=>a-b);
    if (!times.length) return '-';
    return `${fmtTime(times[0]).slice(0,10)} → ${fmtTime(times[times.length-1]).slice(0,10)}`;
  }
  function lv89StationCount(){
    const set = new Set((state.rows || []).map(r => r.station).filter(Boolean));
    return set.size || 0;
  }
  function lv89RecentMape(){
    try {
      const total = (state.lv84?.summaryRows || []).find(r => r.level === 'total') || (state.lv84?.summaryRows || [])[0];
      if (total && total.MAPE_pct !== '' && total.MAPE_pct != null) return `${total.MAPE_pct}%`;
      const hist = JSON.parse(localStorage.getItem('scada_load_forecast_lv84_eval_history') || '[]');
      const h = hist && hist[0];
      const ht = h?.total?.[0] || h?.summary?.total?.[0];
      if (ht && ht.MAPE_pct !== '' && ht.MAPE_pct != null) return `${ht.MAPE_pct}%`;
    } catch(e) {}
    return '-';
  }
  function lv89ForecastAlerts(){
    const rows = state.forecastRows || [];
    const warn = rows.filter(r => /CẢNH|NGUY|canh|nguy/i.test(String(r.trang_thai_nguong || ''))).length;
    return warn;
  }
  function lv89DashboardSnapshot(){
    const modelLabel = state.model ? `${state.model.appVersion || state.model.modelVersion || state.model.version || 'model'} / ${state.model.station || 'multi'}` : 'Chưa nạp';
    const cal = state.model?.calibrationLV85;
    return {
      appVersion: VERSION,
      sourceFileName: state.sourceFileName || '',
      rawRows: state.rawRows?.length || 0,
      ramRows: state.rows?.length || 0,
      dataRange: lv89DateRange(),
      stationCount: lv89StationCount(),
      modelLoaded: !!state.model,
      modelLabel,
      calibration: cal ? (cal.enabled ? 'Đang bật' : 'Có nhưng đang tắt') : 'Chưa có',
      forecastRows: state.forecastRows?.length || 0,
      recentMape: lv89RecentMape(),
      qualityIssues: state.qualityIssues?.length || 0,
      operationEvents: state.operationEvents?.length || 0,
      thresholdAlerts: lv89ForecastAlerts(),
      logRows: state.lv89?.operationLog?.length || 0,
      updatedAt: new Date().toISOString()
    };
  }
  function lv89RenderDashboard(){
    try {
      const d = lv89DashboardSnapshot();
      lv89SetText('lv89RowsMetric', `${d.ramRows}/${d.rawRows}`);
      lv89SetText('lv89RangeMetric', d.dataRange);
      lv89SetText('lv89StationMetric', String(d.stationCount));
      lv89SetText('lv89ModelMetric', d.modelLabel);
      lv89SetText('lv89CalibrationMetric', d.calibration);
      lv89SetText('lv89ForecastMetric', String(d.forecastRows));
      lv89SetText('lv89MapeMetric', d.recentMape);
      lv89SetText('lv89QualityMetric', `${d.qualityIssues} lỗi / ${d.thresholdAlerts} cảnh báo tải`);
      const box = $('lv89DashboardStatus');
      if (box) {
        const readyData = d.ramRows > 0;
        const readyModel = d.modelLoaded;
        const readyForecast = d.forecastRows > 0;
        box.innerHTML = `
          <span class="pill ${readyData?'ok':'warn'}">Dữ liệu: ${readyData ? 'đã có' : 'chưa có'}</span>
          <span class="pill ${readyModel?'ok':'warn'}">Model: ${readyModel ? 'đã nạp' : 'chưa nạp'}</span>
          <span class="pill ${readyForecast?'ok':'warn'}">Forecast: ${readyForecast ? d.forecastRows + ' dòng' : 'chưa có'}</span>
          <span class="pill ${d.qualityIssues?'warn':'ok'}">Lỗi dữ liệu: ${d.qualityIssues}</span>
          <span class="pill ${d.operationEvents?'warn':'ok'}">Sự kiện vận hành: ${d.operationEvents}</span>
          <span class="pill">Cập nhật: ${new Date().toLocaleTimeString('vi-VN')}</span>
          <div class="compactNote" style="margin-top:6px">LV8.9 chỉ quan sát trạng thái, không khóa/mở nút tự động. Nếu thiếu dữ liệu hoặc model, thực hiện theo hướng dẫn Mục 0.1.</div>`;
      }
    } catch(e) {}
  }
  function lv89RenderLog(){
    const box = $('lv89OperationLogBox');
    if (!box) return;
    const rows = (state.lv89?.operationLog || lv89LoadLog()).slice(-80).reverse();
    if (!rows.length) {
      box.innerHTML = '<table><tbody><tr><td>Chưa có nhật ký thao tác LV8.9.</td></tr></tbody></table>';
      return;
    }
    const headers = ['time','level','action','detail','ram_rows','forecast_rows','quality_issues','operation_events','model_loaded','calibration'];
    renderTableInBox(box, rows, headers, 80);
  }
  function lv89ExportLog(){
    const rows = state.lv89?.operationLog || lv89LoadLog();
    if (!rows.length) { log('LV8.9: chưa có nhật ký thao tác để xuất.'); return; }
    saveTextFile('operation_log_lv8_9.csv', toCSV(rows, ['time','iso','level','action','detail','raw_rows','ram_rows','forecast_rows','quality_issues','operation_events','model_loaded','calibration']), 'text/csv;charset=utf-8');
    log(`Đã xuất operation_log_lv8_9.csv gồm ${rows.length} dòng.`);
  }
  function lv89ClearLog(){
    lv89EnsureState().operationLog = [];
    try { localStorage.removeItem(LOG_KEY); } catch(e) {}
    lv89RenderLog();
    lv89RenderDashboard();
    log('Đã xóa nhật ký thao tác LV8.9.');
  }
  function lv89ExportState(){
    const payload = {dashboard: lv89DashboardSnapshot(), operationLogTail:(state.lv89?.operationLog || []).slice(-50)};
    saveTextFile('dashboard_lv8_9.json', JSON.stringify(payload, null, 2), 'application/json');
    log('Đã xuất dashboard_lv8_9.json.');
  }

  function lv89KeyTimeValue(v){
    const t = (v instanceof Date) ? v : parseTime(v);
    return t && !isNaN(t) ? fmtTime(t) : String(v || '').slice(0,16).replace('T',' ');
  }
  function lv89BuildActualMap(){
    if (!state.rows.length && state.rawRows.length) { try { normalizeRows(); } catch(e) {} }
    const map = new Map();
    for (const r of state.rows || []) {
      if (!r.time || !(r.time instanceof Date) || isNaN(r.time) || !Number.isFinite(r.p)) continue;
      const st = String(r.station || '__ALL__');
      const key = `${fmtTime(r.time)}||${st}`;
      if (!map.has(key)) map.set(key, r);
      if (st !== '__ALL__' && !map.has(`${fmtTime(r.time)}||__ALL__`)) map.set(`${fmtTime(r.time)}||__ALL__`, r);
    }
    return map;
  }
  function lv89Metrics(rows, field){
    const arr = rows.map(r => ({err: parseNumber(r[field]), actual: parseNumber(r.actual_p_mw)})).filter(x => Number.isFinite(x.err));
    if (!arr.length) return {n:0, mae:'', mape:'', rmse:'', bias:''};
    const mae = mean(arr.map(x => Math.abs(x.err)));
    const rmse = Math.sqrt(mean(arr.map(x => x.err*x.err)));
    const bias = mean(arr.map(x => x.err));
    const pct = arr.map(x => Number.isFinite(x.actual) && x.actual !== 0 ? Math.abs(x.err / x.actual * 100) : NaN).filter(Number.isFinite);
    const mape = pct.length ? mean(pct) : NaN;
    return {n:arr.length, mae:+mae.toFixed(4), mape:Number.isFinite(mape)? +mape.toFixed(3) : '', rmse:+rmse.toFixed(4), bias:+bias.toFixed(4)};
  }
  function lv89CompareCalibration(){
    if (!state.forecastRows?.length) { throw new Error('Chưa có forecast hiện có để so sánh. Hãy dự báo ở Mục 10.1 trước.'); }
    const actualMap = lv89BuildActualMap();
    const detail = [];
    for (const f of state.forecastRows) {
      const time = lv89KeyTimeValue(f.time);
      const station = String(f.station || '__ALL__');
      const actual = actualMap.get(`${time}||${station}`) || actualMap.get(`${time}||__ALL__`) || actualMap.get(`${time}||`);
      if (!actual) continue;
      const after = parseNumber(f.forecast_p_mw);
      const beforeRaw = parseNumber(f.forecast_before_calibration_mw);
      const before = Number.isFinite(beforeRaw) ? beforeRaw : after - (parseNumber(f.calibration_mw) || 0);
      if (!Number.isFinite(after) || !Number.isFinite(before) || !Number.isFinite(actual.p)) continue;
      const errBefore = before - actual.p;
      const errAfter = after - actual.p;
      detail.push({
        time, station,
        actual_p_mw:+actual.p.toFixed(6),
        forecast_before_calibration_mw:+before.toFixed(6),
        forecast_after_calibration_mw:+after.toFixed(6),
        error_before_mw:+errBefore.toFixed(6),
        error_after_mw:+errAfter.toFixed(6),
        abs_error_before_mw:+Math.abs(errBefore).toFixed(6),
        abs_error_after_mw:+Math.abs(errAfter).toFixed(6),
        improvement_mw:+(Math.abs(errBefore)-Math.abs(errAfter)).toFixed(6),
        improvement_pct: Math.abs(errBefore) > 1e-9 ? +((Math.abs(errBefore)-Math.abs(errAfter))/Math.abs(errBefore)*100).toFixed(2) : '',
        calibration_mw: parseNumber(f.calibration_mw) || 0,
        calibration_source: f.calibration_source || '',
        ket_luan: Math.abs(errAfter) <= Math.abs(errBefore) ? 'Tốt hơn hoặc không xấu hơn' : 'Xấu hơn, cần kiểm tra hiệu chỉnh'
      });
    }
    const beforeM = lv89Metrics(detail, 'error_before_mw');
    const afterM = lv89Metrics(detail, 'error_after_mw');
    const byStationMap = new Map();
    for (const r of detail) {
      if (!byStationMap.has(r.station)) byStationMap.set(r.station, []);
      byStationMap.get(r.station).push(r);
    }
    const rows = [];
    rows.push({level:'total', group:'__TOTAL__', n:beforeM.n, MAE_before:beforeM.mae, MAE_after:afterM.mae, MAPE_before:beforeM.mape, MAPE_after:afterM.mape, RMSE_before:beforeM.rmse, RMSE_after:afterM.rmse, BIAS_before:beforeM.bias, BIAS_after:afterM.bias, cai_thien_MAE_pct: beforeM.mae ? +((beforeM.mae-afterM.mae)/beforeM.mae*100).toFixed(2) : '', ket_luan: afterM.mae <= beforeM.mae ? 'Nên dùng hiệu chỉnh' : 'Hiệu chỉnh làm xấu, cần kiểm tra'});
    for (const [st, arr] of byStationMap) {
      const b = lv89Metrics(arr, 'error_before_mw'), a = lv89Metrics(arr, 'error_after_mw');
      rows.push({level:'station', group:st, n:b.n, MAE_before:b.mae, MAE_after:a.mae, MAPE_before:b.mape, MAPE_after:a.mape, RMSE_before:b.rmse, RMSE_after:a.rmse, BIAS_before:b.bias, BIAS_after:a.bias, cai_thien_MAE_pct: b.mae ? +((b.mae-a.mae)/b.mae*100).toFixed(2) : '', ket_luan: a.mae <= b.mae ? 'Tốt hơn' : 'Xấu hơn'});
    }
    state.lv89.compareRows = rows;
    state.lv89.compareDetail = detail;
    const box = $('lv89CompareTable');
    if (box) renderTableInBox(box, rows, ['level','group','n','MAE_before','MAE_after','MAPE_before','MAPE_after','RMSE_before','RMSE_after','BIAS_before','BIAS_after','cai_thien_MAE_pct','ket_luan'], 500);
    const status = $('lv89CompareStatus');
    if (status) {
      const total = rows[0];
      status.innerHTML = `<span class="pill ${total && String(total.ket_luan).includes('Nên')?'ok':'warn'}">Đã so sánh ${detail.length} mốc</span><span class="pill">MAE trước: ${total?.MAE_before ?? '-'}</span><span class="pill">MAE sau: ${total?.MAE_after ?? '-'}</span><span class="pill">Cải thiện MAE: ${total?.cai_thien_MAE_pct ?? '-'}%</span><span class="pill">${escapeHtml(total?.ket_luan || '')}</span>`;
    }
    const btn = $('lv89ExportCompareBtn'); if (btn) btn.disabled = !rows.length;
    lv89AddLog('So sánh trước/sau hiệu chỉnh LV8.9', `Khớp ${detail.length} mốc; ${rows[0]?.ket_luan || ''}`, 'info');
    return rows;
  }
  function lv89ExportCompare(){
    const rows = state.lv89?.compareRows || [];
    if (!rows.length) { log('LV8.9: chưa có dữ liệu so sánh trước/sau hiệu chỉnh để xuất.'); return; }
    saveTextFile('calibration_compare_lv8_9.csv', toCSV(rows, ['level','group','n','MAE_before','MAE_after','MAPE_before','MAPE_after','RMSE_before','RMSE_after','BIAS_before','BIAS_after','cai_thien_MAE_pct','ket_luan']), 'text/csv;charset=utf-8');
    if (state.lv89?.compareDetail?.length) saveTextFile('calibration_compare_detail_lv8_9.csv', toCSV(state.lv89.compareDetail, ['time','station','actual_p_mw','forecast_before_calibration_mw','forecast_after_calibration_mw','error_before_mw','error_after_mw','abs_error_before_mw','abs_error_after_mw','improvement_mw','improvement_pct','calibration_mw','calibration_source','ket_luan']), 'text/csv;charset=utf-8');
    log(`Đã xuất calibration_compare_lv8_9.csv gồm ${rows.length} dòng.`);
  }

  // Export model LV8.9: giống LV8.5, chỉ cập nhật phiên bản và không thêm Workflow Guard.
  const exportModelBaseLV89 = exportModel;
  exportModel = function(){
    if (!state.model) return;
    const payload = {
      ...state.model,
      scadaModelPackage: 'SCADA_LOAD_FORECAST_OPERATIONAL_MODEL',
      appVersion: VERSION,
      exportSource: 'MUC_8_EXPORT_MODEL',
      exportSection: '8',
      allowedImportSection: '10',
      modelKind: 'OPERATIONAL_FORECAST_MODEL',
      exportedAt: new Date().toISOString(),
      importRules: {
        onlyImportInSection10: true,
        requiredExportSource: 'MUC_8_EXPORT_MODEL',
        note: 'File này được xuất từ Mục 8 và được phép nạp ở Mục 10. LV8.9 giữ lõi dự báo ổn định, có thể chứa calibrationLV85 nếu đã áp dụng hiệu chỉnh.'
      },
      lv89Config: collectFullConfig(),
      lv89State:{strategyByStation: state.lv7?.strategyByStation || state.model?.strategyByStation || {}, calibrationApplied: !!state.model.calibrationLV85, savedAt:new Date().toISOString(), dashboardOnly:true, noWorkflowGuard:true}
    };
    saveTextFile('model_gbdt_lv8_9_operational.json', JSON.stringify(payload, null, 2), 'application/json');
    log('Đã xuất model_gbdt_lv8_9_operational.json từ Mục 8. LV8.9 không dùng Workflow Guard khóa nút.');
  };

  function lv89Bind(){
    try {
      lv89EnsureState();
      state.lv89.operationLog = lv89LoadLog();
      $('lv89RefreshDashboardBtn')?.addEventListener('click', () => { lv89RenderDashboard(); lv89AddLog('Cập nhật Dashboard LV8.9', 'Người dùng bấm cập nhật thủ công.'); });
      $('lv89ExportStateBtn')?.addEventListener('click', () => { lv89ExportState(); lv89AddLog('Xuất trạng thái Dashboard LV8.9', 'dashboard_lv8_9.json'); });
      $('lv89ExportLogBtn')?.addEventListener('click', lv89ExportLog);
      $('lv89ClearLogBtn')?.addEventListener('click', lv89ClearLog);
      $('lv89CompareCalibrationBtn')?.addEventListener('click', () => { try { lv89CompareCalibration(); } catch(e) { const st=$('lv89CompareStatus'); if (st) st.innerHTML = `<span class="pill bad">Lỗi so sánh</span><span class="pill">${escapeHtml(e.message)}</span>`; log('Lỗi LV8.9 so sánh hiệu chỉnh: ' + e.message); } });
      $('lv89ExportCompareBtn')?.addEventListener('click', lv89ExportCompare);
      document.addEventListener('click', (ev) => {
        const btn = ev.target && ev.target.closest ? ev.target.closest('button') : null;
        if (!btn || !btn.id || btn.dataset.lv89NoAutoLog === '1') return;
        const label = (btn.textContent || btn.id).replace(/\s+/g,' ').trim();
        // Ghi nhận sau một nhịp để số dòng/trạng thái phản ánh kết quả thao tác.
        setTimeout(() => lv89AddLog('Bấm nút: ' + label, 'id=' + btn.id, 'click'), 150);
      }, true);
      setInterval(lv89RenderDashboard, 2500);
      setInterval(lv89RenderLog, 5000);
      lv89RenderDashboard();
      lv89RenderLog();
      const title = document.querySelector('h1'); if (title) title.textContent = 'SCADA Load Forecast Offline PWA LV8.9';
      document.title = 'SCADA Load Forecast Offline PWA LV8.9';
      if ($('versionInfo')) $('versionInfo').innerHTML = '<span class="pill modeBadge">LV8.9</span><span class="pill ok">Dashboard + nhật ký thao tác, không khóa nút</span>';
      lv89AddLog('Khởi động LV8.9', 'Dashboard/Operation Log chỉ quan sát; không dùng Workflow Guard.', 'init');
      log('Sẵn sàng LV8.9: Dashboard và Nhật ký thao tác chỉ theo dõi, không khóa nút và không can thiệp luồng chức năng.');
    } catch(e) { log('Lỗi khởi tạo LV8.9: ' + e.message); }
  }
  setTimeout(lv89Bind, 1800);
})();
// ====================== END LV8.9 SAFE DASHBOARD + OPERATION LOG ======================
