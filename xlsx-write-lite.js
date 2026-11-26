/**
 * xlsx-write-lite —— 纯前端生成 .xlsx（无依赖）
 *
 * 原理：xlsx 本质是个 ZIP，里面放几份 XML。
 * ZIP 允许「不压缩」（store），所以不需要任何压缩库，手写头即可。
 *
 * 用法：
 *   const bytes = XlsxWrite.build([
 *     { name: '缺口', rows: [['物料编码','名称'], ['10417112','连接器']] }
 *   ]);
 *   // bytes 是 Uint8Array，配合 Blob 下载
 */
(function (root) {
  'use strict';

  const enc = new TextEncoder();

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[i] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function zipStore(files) {
    const parts = [];
    const central = [];
    let offset = 0;

    const now = new Date();
    const dTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() / 2)) & 0xFFFF;
    const dDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xFFFF;

    for (const f of files) {
      const nameB = enc.encode(f.name);
      const data = f.data;
      const crc = crc32(data);

      const lh = new Uint8Array(30 + nameB.length);
      const v = new DataView(lh.buffer);
      v.setUint32(0, 0x04034b50, true);
      v.setUint16(4, 20, true);          // version needed
      v.setUint16(6, 0x0800, true);      // UTF-8 文件名
      v.setUint16(8, 0, true);           // store
      v.setUint16(10, dTime, true);
      v.setUint16(12, dDate, true);
      v.setUint32(14, crc, true);
      v.setUint32(18, data.length, true);
      v.setUint32(22, data.length, true);
      v.setUint16(26, nameB.length, true);
      v.setUint16(28, 0, true);
      lh.set(nameB, 30);
      parts.push(lh, data);

      const ch = new Uint8Array(46 + nameB.length);
      const cv = new DataView(ch.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, dTime, true);
      cv.setUint16(14, dDate, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, data.length, true);
      cv.setUint32(24, data.length, true);
      cv.setUint16(28, nameB.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      ch.set(nameB, 46);
      central.push(ch);

      offset += lh.length + data.length;
    }

    const cdSize = central.reduce((a, b) => a + b.length, 0);
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, files.length, true);
    ev.setUint16(10, files.length, true);
    ev.setUint32(12, cdSize, true);
    ev.setUint32(16, offset, true);
    ev.setUint16(20, 0, true);

    const all = parts.concat(central, [eocd]);
    const total = all.reduce((a, b) => a + b.length, 0);
    const out = new Uint8Array(total);
    let p = 0;
    for (const a of all) { out.set(a, p); p += a.length; }
    return out;
  }

  function esc(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;')
      // 去掉 XML 不允许的控制字符
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  }

  function colName(i) {
    let s = '';
    i += 1;
    while (i > 0) {
      const m = (i - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      i = Math.floor((i - 1) / 26);
    }
    return s;
  }

  function sheetXml(rows) {
    const out = [];
    out.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    out.push('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
    out.push('<sheetData>');
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] || [];
      out.push(`<row r="${r + 1}">`);
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        if (v === null || v === undefined || v === '') continue;
        const ref = colName(c) + (r + 1);
        if (typeof v === 'number' && isFinite(v)) {
          out.push(`<c r="${ref}"><v>${v}</v></c>`);
        } else {
          out.push(`<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`);
        }
      }
      out.push('</row>');
    }
    out.push('</sheetData></worksheet>');
    return out.join('');
  }

  function build(sheets) {
    const list = (sheets || []).filter(s => s && s.rows);
    if (!list.length) throw new Error('没有可写的工作表');

    const files = [];
    const push = (name, text) => files.push({ name, data: enc.encode(text) });

    const ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'];
    const wbSheets = [];
    const wbRels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'];

    list.forEach((s, i) => {
      const n = i + 1;
      const safe = String(s.name || `Sheet${n}`).replace(/[\\/?*\[\]:]/g, '_').slice(0, 31);
      ct.push(`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
      push(`xl/worksheets/sheet${n}.xml`, sheetXml(s.rows));
      wbSheets.push(`<sheet name="${esc(safe)}" sheetId="${n}" r:id="rId${n}"/>`);
      wbRels.push(`<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/>`);
    });

    ct.push('</Types>');
    wbRels.push('</Relationships>');

    push('[Content_Types].xml', ct.join(''));
    push('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
      + '</Relationships>');
    push('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
      + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      + '<sheets>' + wbSheets.join('') + '</sheets></workbook>');
    push('xl/_rels/workbook.xml.rels', wbRels.join(''));

    return zipStore(files);
  }

  const api = { build, zipStore, crc32, colName };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.XlsxWrite = api;
})(typeof window !== 'undefined' ? window : globalThis);
