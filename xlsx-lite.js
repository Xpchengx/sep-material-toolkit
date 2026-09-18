/* ============================================================================
   xlsx-lite.js —— 零依赖读取 xlsx / csv
   ----------------------------------------------------------------------------
   为什么不用 SheetJS：
     页面要求单文件、可离线。引入第三方库要么内联几百 KB，要么走 CDN（断网就废）。
     xlsx 本质是个 ZIP 包，浏览器有原生 DecompressionStream 可以解 deflate，
     所以自己解析完全可行。

   实现范围（够用就好，不做通用库）：
     · ZIP：中央目录 + 本地头，支持 stored(0) / deflate(8)
     · xlsx：workbook.xml 取表名、rels 取路径、sharedStrings、sheet XML
     · 单元格：共享字符串 / 内联字符串 / 公式缓存值 / 数字 / 布尔 / 日期(序列号)
     · csv：带引号转义、BOM、CRLF

   已知不做：
     · ZIP64（单文件 > 4GB）—— 台账不会这么大
     · 复杂格式、图表、批注 —— 与取数无关
     · 合并单元格的填充复制（按左上角取值，其余为空，下游按需处理）

   纯逻辑，不依赖 DOM，因此可以在 Node 里跑测试。
   ========================================================================== */
(function (root) {
  'use strict';

  // ========================================================== ZIP

  const SIG_EOCD = 0x06054b50;
  const SIG_CEN = 0x02014b50;
  const SIG_LOC = 0x04034b50;

  /** 用浏览器/Node 原生能力解 deflate-raw */
  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('当前环境不支持 DecompressionStream，无法解压 xlsx（需要 Chrome 80+ / Node 18+）');
    }
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  /** 找到中央目录结束记录 */
  function findEOCD(dv, len) {
    // EOCD 最小 22 字节，注释最长 65535
    const start = Math.max(0, len - 22 - 65535);
    for (let i = len - 22; i >= start; i--) {
      if (dv.getUint32(i, true) === SIG_EOCD) return i;
    }
    return -1;
  }

  /**
   * 解开 ZIP，返回 Map<文件名, Uint8Array>
   * 只解压需要的条目（懒解压由调用方决定，这里为简单起见全解，xlsx 内部文件都很小）
   */
  async function readZip(buffer) {
    const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const len = bytes.byteLength;

    const eocd = findEOCD(dv, len);
    if (eocd < 0) throw new Error('不是有效的 xlsx（找不到 ZIP 结构）');

    const count = dv.getUint16(eocd + 10, true);
    let off = dv.getUint32(eocd + 16, true);
    if (off === 0xffffffff) throw new Error('不支持 ZIP64 格式的文件');

    const files = new Map();
    const dec = new TextDecoder('utf-8');

    for (let n = 0; n < count; n++) {
      if (dv.getUint32(off, true) !== SIG_CEN) break;
      const method = dv.getUint16(off + 10, true);
      const compSize = dv.getUint32(off + 20, true);
      const nameLen = dv.getUint16(off + 28, true);
      const extraLen = dv.getUint16(off + 30, true);
      const cmtLen = dv.getUint16(off + 32, true);
      const locOff = dv.getUint32(off + 42, true);
      const name = dec.decode(bytes.subarray(off + 46, off + 46 + nameLen));

      // 本地头里的 extra 长度可能与中央目录不同，必须重新读
      if (dv.getUint32(locOff, true) === SIG_LOC) {
        const lNameLen = dv.getUint16(locOff + 26, true);
        const lExtraLen = dv.getUint16(locOff + 28, true);
        const dataStart = locOff + 30 + lNameLen + lExtraLen;
        const raw = bytes.subarray(dataStart, dataStart + compSize);
        if (raw.byteLength > 0) {
          try {
            files.set(name, method === 0 ? raw : await inflateRaw(raw));
          } catch (e) {
            // 单个条目解压失败不阻断整体
            files.set(name, null);
          }
        }
      }
      off += 46 + nameLen + extraLen + cmtLen;
    }
    return files;
  }

  // ========================================================== XML 小工具

  const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

  function unescapeXml(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#x?[0-9a-fA-F]+|amp|lt|gt|quot|apos);/g, (m, g) => {
      if (g[0] === '#') {
        const code = g[1] === 'x' || g[1] === 'X'
          ? parseInt(g.slice(2), 16) : parseInt(g.slice(1), 10);
        return isNaN(code) ? m : String.fromCodePoint(code);
      }
      return ENT[g] !== undefined ? ENT[g] : m;
    });
  }

  function attr(tag, name) {
    const m = tag.match(new RegExp('\\b' + name + '\\s*=\\s*"([^"]*)"'));
    return m ? unescapeXml(m[1]) : null;
  }

  /** 列引用 A1 / AB12 -> {col, row}  (0-based) */
  function cellRef(ref) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref || '');
    if (!m) return null;
    let col = 0;
    for (let i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64);
    return { col: col - 1, row: parseInt(m[2], 10) - 1 };
  }

  // ========================================================== xlsx

  const BUILTIN_DATE_FMT = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

  /**
   * 解析 styles.xml，返回 { dateStyleIdx: Set<number>, base1904: bool }
   * cellXfs 的序号即单元格 s 属性所指的索引
   */
  function parseStyles(xml) {
    const dateStyleIdx = new Set();
    if (!xml) return dateStyleIdx;
    const custom = new Set();
    const nfRe = /<numFmt\b[^>]*\/?>/g;
    let m;
    while ((m = nfRe.exec(xml))) {
      const id = parseInt(attr(m[0], 'numFmtId'), 10);
      const code = attr(m[0], 'formatCode') || '';
      // 去掉引号内的字面量再判断，避免 "年" 之类误判
      const bare = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
      if (/[ymdhs]/i.test(bare) && /[ymd]/i.test(bare)) custom.add(id);
    }
    const xfBlock = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
    if (xfBlock) {
      const xfRe = /<xf\b[^>]*\/?>/g;
      let i = 0;
      while ((m = xfRe.exec(xfBlock[1]))) {
        const id = parseInt(attr(m[0], 'numFmtId') || '0', 10);
        if (BUILTIN_DATE_FMT.has(id) || custom.has(id)) dateStyleIdx.add(i);
        i++;
      }
    }
    return dateStyleIdx;
  }

  /** xlsx 日期序列号 -> Date（含 1900 闰年 bug 的修正） */
  function serialToDate(serial, base1904) {
    const epoch = base1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    return new Date(epoch + Math.round(serial * 86400000));
  }

  function parseSharedStrings(xml) {
    const out = [];
    if (!xml) return out;
    const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g;
    let m;
    while ((m = siRe.exec(xml))) {
      const body = m[1] || '';
      // 富文本会有多个 <t>，拼起来
      let text = '';
      const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>|<t\b[^>]*\/>/g;
      let t;
      while ((t = tRe.exec(body))) text += unescapeXml(t[1] || '');
      out.push(text);
    }
    return out;
  }

  /** 解析一个 sheet XML -> 二维数组（稀疏处补 null，按需扩列） */
  function parseSheet(xml, shared, dateStyleIdx, base1904, maxRows) {
    const rows = [];
    let maxCol = 0;
    const rowRe = /<row\b([^>]*)>([\s\S]*?)<\/row>|<row\b([^>]*)\/>/g;
    let rm;
    while ((rm = rowRe.exec(xml))) {
      if (maxRows && rows.length >= maxRows) break;
      const attrs = rm[1] || rm[3] || '';
      const body = rm[2] || '';
      const rIdx = parseInt(attr(attrs, 'r') || String(rows.length + 1), 10) - 1;
      const cells = [];

      const cRe = /<c\b([^>]*)\/>|<c\b([^>]*)>([\s\S]*?)<\/c>/g;
      let cm;
      while ((cm = cRe.exec(body))) {
        const cAttrs = cm[1] || cm[2] || '';
        const cBody = cm[3] || '';
        const ref = cellRef(attr(cAttrs, 'r'));
        const ci = ref ? ref.col : cells.length;
        const type = attr(cAttrs, 't') || 'n';
        const styleIdx = parseInt(attr(cAttrs, 's') || '-1', 10);

        let val = null;
        if (type === 'inlineStr') {
          const is = /<is\b[^>]*>([\s\S]*?)<\/is>/.exec(cBody);
          let text = '';
          if (is) {
            const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>|<t\b[^>]*\/>/g;
            let t;
            while ((t = tRe.exec(is[1]))) text += unescapeXml(t[1] || '');
          }
          val = text;
        } else {
          const v = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(cBody);
          const raw = v ? unescapeXml(v[1]) : null;
          if (raw === null || raw === '') {
            val = null;
          } else if (type === 's') {
            val = shared[parseInt(raw, 10)] !== undefined ? shared[parseInt(raw, 10)] : '';
          } else if (type === 'str' || type === 'e') {
            val = raw;                       // 公式结果是字符串 / 错误值
          } else if (type === 'b') {
            val = raw === '1';
          } else {
            const num = Number(raw);
            if (isFinite(num) && styleIdx >= 0 && dateStyleIdx.has(styleIdx) && num > 0) {
              val = serialToDate(num, base1904);   // 日期
            } else {
              val = isFinite(num) ? num : raw;
            }
          }
        }
        cells[ci] = val;
        if (ci + 1 > maxCol) maxCol = ci + 1;
      }
      rows[rIdx] = rows[rIdx] || cells;
      if (rIdx + 1 > rows.length) rows.length = rIdx + 1;
      for (let i = 0; i < cells.length; i++) {
        if (rows[rIdx][i] === undefined) rows[rIdx][i] = null;
      }
    }
    // 补齐短行 + 去空洞
    const out = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r) { out.push(new Array(maxCol).fill(null)); continue; }
      const fixed = new Array(maxCol);
      for (let j = 0; j < maxCol; j++) fixed[j] = r[j] === undefined ? null : r[j];
      out.push(fixed);
    }
    return out;
  }

  /** 读 xlsx -> { sheets: [{name, rows}], warnings: [] } */
  async function readXlsx(buffer, opts) {
    opts = opts || {};
    const files = await readZip(buffer);
    const dec = new TextDecoder('utf-8');
    const warnings = [];

    const get = (n) => {
      const b = files.get(n);
      return b ? dec.decode(b) : null;
    };

    // 表名 -> 路径
    const wbXml = get('xl/workbook.xml');
    if (!wbXml) throw new Error('不是有效的 xlsx（缺少 xl/workbook.xml）');
    const relsXml = get('xl/_rels/workbook.xml.rels') || '';
    const rels = {};
    let m;
    const relRe = /<Relationship\b[^>]*\/?>/g;
    while ((m = relRe.exec(relsXml))) {
      const id = attr(m[0], 'Id');
      let target = attr(m[0], 'Target') || '';
      if (target && !target.startsWith('/')) {
        target = target.replace(/^\.\//, '');
        target = 'xl/' + target;
      } else if (target.startsWith('/')) {
        target = target.slice(1);
      }
      if (id) rels[id] = target;
    }

    const sheets = [];
    const shRe = /<sheet\b[^>]*\/?>/g;
    while ((m = shRe.exec(wbXml))) {
      const name = attr(m[0], 'name') || '';
      const rid = attr(m[0], 'r:id') || attr(m[0], 'id');
      sheets.push({ name, path: rels[rid] || null });
    }

    // 1904 基准
    const base1904 = /date1904\s*=\s*"(1|true)"/.test(wbXml);
    const dateStyleIdx = parseStyles(get('xl/styles.xml'));
    const shared = parseSharedStrings(get('xl/sharedStrings.xml'));
    const maxRows = opts.maxRows || 200000;

    const out = [];
    for (const s of sheets) {
      if (!s.path || !files.has(s.path)) {
        out.push({ name: s.name, rows: [], missing: true });
        warnings.push(`工作表「${s.name}」找不到数据文件（${s.path || '路径未解析'}）`);
        continue;
      }
      const xml = dec.decode(files.get(s.path));
      const rows = parseSheet(xml, shared, dateStyleIdx, base1904, maxRows);
      if (rows.length >= maxRows) warnings.push(`工作表「${s.name}」超过 ${maxRows} 行，只读了前 ${maxRows} 行`);
      out.push({ name: s.name, rows });
    }

    return { sheets: out, warnings };
  }

  // ========================================================== csv

  /** 解析 CSV 文本 -> 二维数组（支持引号转义、BOM、CRLF） */
  function parseCsv(text) {
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    const rows = [];
    let row = [], field = '', inQ = false, i = 0;
    while (i < text.length) {
      const c = text[i];
      if (inQ) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQ = false; i++; continue;
        }
        field += c; i++; continue;
      }
      if (c === '"') { inQ = true; i++; continue; }
      if (c === ',') { row.push(field); field = ''; i++; continue; }
      if (c === '\r') { i++; continue; }
      if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
      field += c; i++;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  function readCsv(text) {
    const rows = parseCsv(text);
    // 数值化：纯数字的单元格转成 number
    const conv = rows.map(r => r.map(v => {
      const s = String(v).trim();
      if (s === '') return null;
      if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
      return s;
    }));
    return { sheets: [{ name: 'CSV', rows: conv }], warnings: [] };
  }

  // ========================================================== 统一入口

  /**
   * 读一个文件（File / Blob / ArrayBuffer / 字符串）
   * -> { sheets, warnings, kind }
   */
  async function readTable(input, opts) {
    opts = opts || {};
    const name = (input && input.name) || opts.name || '';

    // 字符串：按 csv 处理
    if (typeof input === 'string') {
      const t = input.trim();
      if (t.startsWith('{') || t.startsWith('[')) {
        return { sheets: [], kind: 'json', json: JSON.parse(t), warnings: [] };
      }
      return Object.assign(readCsv(input), { kind: 'csv' });
    }

    let buf;
    if (input instanceof ArrayBuffer || ArrayBuffer.isView(input)) {
      buf = input;
    } else if (typeof input.arrayBuffer === 'function') {
      buf = await input.arrayBuffer();
    } else {
      throw new Error('不支持的数据类型');
    }

    const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    // 魔数判断，别只看扩展名（集成商经常改错后缀）
    const isZip = bytes[0] === 0x50 && bytes[1] === 0x4B;
    const head = new TextDecoder('utf-8').decode(bytes.subarray(0, 512));

    if (isZip) {
      return Object.assign(await readXlsx(bytes, opts), { kind: 'xlsx' });
    }
    if (/^\s*[{[]/.test(head)) {
      const txt = new TextDecoder('utf-8').decode(bytes);
      return { sheets: [], kind: 'json', json: JSON.parse(txt), warnings: [] };
    }
    // .xls（老的 BIFF 二进制格式）单独提示
    if (bytes[0] === 0xD0 && bytes[1] === 0xCF) {
      throw new Error('这是老的 .xls 格式（Excel 97-2003），浏览器读不了。请用 Excel 另存为 .xlsx');
    }
    return Object.assign(readCsv(new TextDecoder('utf-8').decode(bytes)), { kind: 'csv' });
  }

  // ========================================================== 表头识别

  /**
   * 在前若干行里找表头行：返回 {rowIndex, headers}
   * 判据：该行含有 "型号"（台账的关键列），且非空单元格数 >= 3
   */
  function findHeaderRow(rows, opts) {
    opts = opts || {};
    const keys = opts.keys || ['型号', '物料', '规格'];
    const limit = Math.min(rows.length, opts.limit || 12);
    for (let i = 0; i < limit; i++) {
      const r = rows[i] || [];
      const texts = r.map(v => (v === null || v === undefined) ? '' : String(v).trim());
      const filled = texts.filter(Boolean).length;
      if (filled < 3) continue;
      if (keys.some(k => texts.some(t => t.includes(k)))) {
        return { rowIndex: i, headers: texts };
      }
    }
    return null;
  }

  /** 按关键字找列，返回索引或 -1 */
  function findColumn(headers, keywords) {
    for (let i = 0; i < headers.length; i++) {
      for (const k of keywords) {
        if (headers[i] && headers[i].includes(k)) return i;
      }
    }
    return -1;
  }

  const api = {
    readZip, readXlsx, readCsv, readTable, parseCsv,
    findHeaderRow, findColumn, cellRef, serialToDate, unescapeXml,
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.XlsxLite = api;
})(typeof window !== 'undefined' ? window : globalThis);
