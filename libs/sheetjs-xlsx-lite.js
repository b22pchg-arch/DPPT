/*
  SheetJS-compatible XLSX Lite Reader for SCADA Load Forecast PWA LV3.
  Offline fallback with a small subset of the SheetJS browser API:
    XLSX.read(arrayBuffer, {type:'array'})
    XLSX.utils.sheet_to_json(worksheet, {defval:'', header:1})
  Supported by this fallback: .xlsx/.xlsm basic worksheets with shared strings, inline strings, numbers and booleans.
  Legacy binary .xls requires the official SheetJS xlsx.full.min.js library.
*/
(function(global){
  'use strict';
  if (global.XLSX && global.XLSX.read && global.XLSX.utils && global.XLSX.utils.sheet_to_json) return;

  const TD = new TextDecoder('utf-8');
  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  function u16(dv, off){ return dv.getUint16(off, true); }
  function u32(dv, off){ return dv.getUint32(off, true); }
  function decodeXml(u8){ return TD.decode(u8).replace(/^\uFEFF/, ''); }
  function parseXml(text){
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    const err = doc.querySelector ? doc.querySelector('parsererror') : doc.getElementsByTagName('parsererror')[0];
    if (err) throw new Error('Không đọc được XML trong file Excel.');
    return doc;
  }
  function attr(el, name){ return el.getAttribute(name) || el.getAttributeNS(NS_REL, name.replace(/^r:/,'')) || ''; }
  function childText(el, tag){ const n = el.getElementsByTagName(tag)[0]; return n ? n.textContent : ''; }
  function colIndex(ref){
    const m = String(ref || '').match(/^[A-Z]+/i);
    if (!m) return 0;
    let n = 0;
    for (const ch of m[0].toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  function uniqueHeaders(headers){
    const seen = new Map();
    return headers.map((h, i) => {
      let name = String(h == null || h === '' ? `cot_${i+1}` : h).trim() || `cot_${i+1}`;
      const n = seen.get(name) || 0;
      seen.set(name, n + 1);
      return n ? `${name}_${n+1}` : name;
    });
  }

  function findEndOfCentralDirectory(u8){
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const min = Math.max(0, u8.length - 0xFFFF - 22);
    for (let i = u8.length - 22; i >= min; i--) {
      if (u32(dv, i) === 0x06054b50) return i;
    }
    throw new Error('Không tìm thấy cấu trúc ZIP của file .xlsx.');
  }

  function readZip(input){
    const u8 = input instanceof Uint8Array ? input : new Uint8Array(input);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const eocd = findEndOfCentralDirectory(u8);
    const total = u16(dv, eocd + 10);
    const cdOffset = u32(dv, eocd + 16);
    const files = new Map();
    let p = cdOffset;
    for (let i = 0; i < total; i++) {
      if (u32(dv, p) !== 0x02014b50) throw new Error('Central directory ZIP không hợp lệ.');
      const method = u16(dv, p + 10);
      const compSize = u32(dv, p + 20);
      const uncompSize = u32(dv, p + 24);
      const nameLen = u16(dv, p + 28);
      const extraLen = u16(dv, p + 30);
      const commentLen = u16(dv, p + 32);
      const localOffset = u32(dv, p + 42);
      const name = TD.decode(u8.slice(p + 46, p + 46 + nameLen)).replace(/\\/g, '/');
      files.set(name, {method, compSize, uncompSize, localOffset, name});
      p += 46 + nameLen + extraLen + commentLen;
    }
    function get(name){
      name = String(name).replace(/^\//, '').replace(/\\/g, '/');
      const f = files.get(name);
      if (!f) return null;
      if (u32(dv, f.localOffset) !== 0x04034b50) throw new Error('Local header ZIP không hợp lệ: ' + name);
      const nl = u16(dv, f.localOffset + 26);
      const el = u16(dv, f.localOffset + 28);
      const start = f.localOffset + 30 + nl + el;
      const comp = u8.slice(start, start + f.compSize);
      if (f.method === 0) return comp;
      if (f.method === 8) {
        if (!global.pako || !global.pako.inflateRaw) throw new Error('Thiếu pako.min.js để giải nén file .xlsx.');
        return global.pako.inflateRaw(comp);
      }
      throw new Error(`File ${name} dùng phương thức nén ZIP chưa hỗ trợ: ${f.method}`);
    }
    return {files, get};
  }

  function resolvePath(baseDir, target){
    target = String(target || '').replace(/\\/g, '/');
    if (target.startsWith('/')) return target.slice(1);
    const parts = (baseDir + '/' + target).split('/');
    const out = [];
    for (const part of parts) {
      if (!part || part === '.') continue;
      if (part === '..') out.pop(); else out.push(part);
    }
    return out.join('/');
  }

  function parseSharedStrings(zip){
    const u8 = zip.get('xl/sharedStrings.xml');
    if (!u8) return [];
    const doc = parseXml(decodeXml(u8));
    return Array.from(doc.getElementsByTagName('si')).map(si => {
      const ts = Array.from(si.getElementsByTagName('t')).map(t => t.textContent || '');
      return ts.join('');
    });
  }

  function parseWorkbookRels(zip){
    const u8 = zip.get('xl/_rels/workbook.xml.rels');
    const map = new Map();
    if (!u8) return map;
    const doc = parseXml(decodeXml(u8));
    Array.from(doc.getElementsByTagName('Relationship')).forEach(r => {
      map.set(r.getAttribute('Id'), r.getAttribute('Target'));
    });
    return map;
  }

  function parseSheets(zip){
    const wbXml = zip.get('xl/workbook.xml');
    if (!wbXml) throw new Error('Không tìm thấy xl/workbook.xml trong file Excel.');
    const wb = parseXml(decodeXml(wbXml));
    const rels = parseWorkbookRels(zip);
    return Array.from(wb.getElementsByTagName('sheet')).map((s, i) => {
      const name = s.getAttribute('name') || `Sheet${i+1}`;
      const rid = s.getAttribute('r:id') || s.getAttributeNS(NS_REL, 'id') || '';
      const target = rels.get(rid) || `worksheets/sheet${i+1}.xml`;
      return {name, path: resolvePath('xl', target)};
    });
  }

  function parseCellValue(c, sharedStrings){
    const t = c.getAttribute('t') || '';
    if (t === 'inlineStr') {
      const is = c.getElementsByTagName('is')[0];
      if (!is) return '';
      return Array.from(is.getElementsByTagName('t')).map(n => n.textContent || '').join('');
    }
    const v = childText(c, 'v');
    if (t === 's') return sharedStrings[Number(v)] ?? '';
    if (t === 'str') return v;
    if (t === 'b') return v === '1' ? 1 : 0;
    if (v === '') return '';
    const n = Number(v);
    return Number.isFinite(n) ? n : v;
  }

  function parseWorksheet(zip, path, sharedStrings){
    const u8 = zip.get(path);
    if (!u8) throw new Error('Không tìm thấy worksheet: ' + path);
    const doc = parseXml(decodeXml(u8));
    const rows = [];
    Array.from(doc.getElementsByTagName('row')).forEach(rowEl => {
      let arr = [];
      Array.from(rowEl.getElementsByTagName('c')).forEach(c => {
        const idx = colIndex(c.getAttribute('r') || 'A1');
        arr[idx] = parseCellValue(c, sharedStrings);
      });
      while (arr.length && (arr[arr.length-1] == null || arr[arr.length-1] === '')) arr.pop();
      rows.push(arr.map(v => v == null ? '' : v));
    });
    return {__rows: rows};
  }

  function read(data, opts){
    const zip = readZip(data);
    const sharedStrings = parseSharedStrings(zip);
    const sheetMeta = parseSheets(zip);
    const out = {SheetNames: [], Sheets: {}, Workbook: {Sheets: sheetMeta.map(s => ({name:s.name}))}, __lite: true};
    sheetMeta.forEach(s => {
      out.SheetNames.push(s.name);
      out.Sheets[s.name] = parseWorksheet(zip, s.path, sharedStrings);
    });
    return out;
  }

  function sheet_to_json(ws, opts){
    opts = opts || {};
    const defval = Object.prototype.hasOwnProperty.call(opts, 'defval') ? opts.defval : '';
    const rows = (ws && ws.__rows ? ws.__rows : []).map(r => r.slice());
    if (opts.header === 1) return rows.map(r => r.map(v => v == null ? defval : v));
    if (!rows.length) return [];
    const headers = uniqueHeaders(rows[0].map(v => v == null || v === '' ? '' : String(v)));
    return rows.slice(1).filter(r => r.some(v => v != null && String(v).trim() !== '')).map(r => {
      const o = {};
      headers.forEach((h, i) => o[h] = r[i] == null || r[i] === '' ? defval : r[i]);
      return o;
    });
  }

  global.XLSX = {version: 'lite-xlsx-reader-0.1', read, utils: {sheet_to_json}};
})(window);
