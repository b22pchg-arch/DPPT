'use strict';

const state = {
  headers: [], rawRows: [], rows: [], colMap: {}, model: null, trainingResult: null, forecastRows: [], sampleLoaded: false,
  sourceFileName: '', delimiter: ',', workbook: null, currentSheet: '', editor: {page: 1, pageSize: 100, query: '', filter: 'all', selected: new Set(), dirty: false, dateMode: 'all', dateSingle: '', dateMulti: '', dateFrom: '', dateTo: ''}
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

function runQualityCheck() {
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
    pred = Math.max(0, pred + bias);
    row.p = pred;
    forecast.push({step:s, time:fmtTime(t), station:row.station, forecast_p_mw:formatNum(pred,3), temp:Number.isFinite(tempUse)?formatNum(tempUse,1):'', rain:rainDefault, holiday:row.holiday, gbdt:formatNum(gbdt,3), similar_day:formatNum(similar,3), same_hour_last_week:formatNum(week,3), trend:formatNum(trend,3), bias:formatNum(bias,3)});
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
  renderTable(state.forecastRows, ['step','time','station','forecast_p_mw','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','nguong_canh_bao','trang_thai_nguong'], 1000);
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
  for (const st of stations) {
    try { all = all.concat(forecastForStation(st).forecast); }
    catch(e) { log('Bỏ qua ' + st + ': ' + e.message); }
  }
  state.forecastRows = all;
  applyThresholdsToForecast(state.forecastRows);
  renderTable(all, ['step','time','station','forecast_p_mw','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','nguong_canh_bao','trang_thai_nguong'], 2000);
  updateForecastMetrics(all);
  $('exportForecastBtn').disabled = false;
  log(`Đã dự báo tất cả trạm/lộ: ${all.length} dòng.`);
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
  const headers = ['step','time','station','forecast_p_mw','temp','rain','holiday','gbdt','similar_day','same_hour_last_week','trend','bias','nguong_canh_bao','trang_thai_nguong'];
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

$('modelFile').addEventListener('change', e => {
  const file = e.target.files[0]; if (!file) return;
  loadTextFile(file, text => {
    try {
      const m = JSON.parse(text);
      if ((!m.trees || !m.featureNames) && !m.modelsByStation) throw new Error('Không đúng cấu trúc model_gbdt.json.');
      state.model = m;
      updateMetrics(m.metrics?.validation || {}, m.intervalMinutes);
      renderModelInfo();
      $('forecastBtn').disabled = false;
      if ($('forecastAllBtn')) $('forecastAllBtn').disabled = false;
      $('exportModelBtn').disabled = false;
      log(`Đã nạp model: ${file.name}, ${m.modelsByStation ? Object.keys(m.modelsByStation).length + ' model trạm/lộ' : m.trees.length + ' cây'}.`);
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
log('Sẵn sàng LV5.4. Bước 1: nạp file Excel .xlsx/.xlsm hoặc CSV/TXT/TSV/JSON. Có thể lọc theo ngày, chọn nhiều dòng, điền nhanh nhiệt độ/cờ vận hành rồi huấn luyện hoặc nạp model.');


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
  log('Đã nạp bản sửa nội suy LV5.4: P=0/P thấp được xử lý theo lựa chọn cờ vận hành, có lưu p_goc.');
} catch(_) {}
