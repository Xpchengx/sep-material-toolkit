/* ============================================================================
   wb-merge.js —— 把拖入的原始表格合并成工作台数据
   ----------------------------------------------------------------------------
   输入：若干「数据源」
     { id, fileName, supplier, type, sheets:[{name,rows}] }
     type: 库存 | 入库 | 出库 | 退料

   输出：与 Python 端 build_workbench_data.py 一致的 workbench 数据结构
     { snapshotDate, trend, safetyLevel, codes, materials:[[...]] }

   ⚠️ 口径必须与 Python 端严格一致，否则两条路径的结果对不上：
     · 近 N 日 = [快照日-(N-1), 快照日]（含首尾）
     · lastOut = 距最后出库的天数；从未出库记 999
     · warehouse = 持有该物料的集成商，按持有量取前 2 个，用 / 连接，空则「—」
     · 库存按 标准型号 + 单位 聚合（单位不同不合并）

   纯逻辑，不依赖 DOM，可在 Node 里测试。
   ========================================================================== */
(function (root) {
  'use strict';

  const XL = root.XlsxLite ||
    (typeof require !== 'undefined' ? require('./xlsx-lite.js') : null);

  // 兜底分类规则（Python 端会通过 meta.categories 覆盖）
  const DEFAULT_CATS = [
    ['光模块', ['光模块', '彩光模块', 'CWDM', '模块-', '1.25G模块', '25GE']],
    ['主设备', ['RRU', 'BBU', 'AAU', '基带板', '主控', 'INCR', '皮站', '直放站',
                'RHUB', 'RHBU', 'pRRU', '远端单元', '轿厢', '载波功放', '交转直',
                '近端', '远端', '电源模块']],
    ['天线', ['天线', '吸顶', '板状', '对数周期', '射灯', '天馈']],
    ['线缆', ['馈线', '电缆', '光缆', '光电', '电源线', '混合缆', '地线', '接地线',
              '钢管', 'PVC管', '跳纤', '尾纤', '槽道']],
    ['光配', ['分纤箱', '分线箱', '野战光缆', '终端盒']],
    ['接头', ['接头', 'N-JJ', 'N-KK', '直角头', '转接']],
    ['无源器件', ['功分', '耦合', 'dB', '合路器', '合分波', '电桥', '负载', '衰减']],
    ['辅料', ['挡风板', '标签', '辅料', '胶带', '扎带']],
  ];

  const DAY = 86400000;

  // ======================================================== 日期

  /** Excel 序列号 -> Date（含 1900 闰年 bug 修正） */
  function serialDate(n) {
    return new Date(Date.UTC(1899, 11, 30) + Math.round(n) * DAY);
  }

  /** 判断一个数字是否像 Excel 日期序列号（1954~2118 年区间） */
  function looksSerial(n) {
    return typeof n === 'number' && isFinite(n) && n > 20000 && n < 80000;
  }

  /**
   * 解析各种日期写法 -> Date（UTC 零点）或 null
   * 支持：Date / Excel 序列号 / 2026-09-08 / 2026.9.8 / 9月8日 / 9-8
   * 中文与「9-8」没有年份，用 fallbackYear 补
   */
  function parseDate(v, fallbackYear) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === 'number') return looksSerial(v) ? serialDate(v) : null;
    const s = String(v).trim();
    let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    m = s.match(/^(\d{1,2})\s*月\s*(\d{1,2})\s*[号日]/);
    if (m && fallbackYear) return new Date(Date.UTC(fallbackYear, +m[1] - 1, +m[2]));
    m = s.match(/^(\d{1,2})[-/](\d{1,2})$/);
    if (m && fallbackYear) return new Date(Date.UTC(fallbackYear, +m[1] - 1, +m[2]));
    return null;
  }

  /** 取值里是否自带绝对年份（Date / 序列号 / 带年 ISO）。自带就往里取。 */
  function absDate(v) {
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === 'number') return looksSerial(v) ? serialDate(v) : null;
    if (typeof v === 'string') {
      const m = v.trim().match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
      if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    }
    return null;
  }

  /** 无年份写法取出月日：9月8号 / 9-8 */
  function monthDayOf(v) {
    if (typeof v !== 'string') return null;
    const s = v.trim();
    let m = s.match(/^(\d{1,2})\s*月\s*(\d{1,2})\s*[号日]/);
    if (m) return { m: +m[1], d: +m[2] };
    m = s.match(/^(\d{1,2})[-/](\d{1,2})$/);
    if (m) return { m: +m[1], d: +m[2] };
    return null;
  }

  /**
   * 解析一整列日期。
   *
   * 三个层次，按可靠性递进：
   *   ① 自带年份的值（Date / Excel 序列号 / 2026-09-08）—— 直接用
   *   ② 无年份写法（9月8号）—— 需要补年
   *   ③ 补年的依据：**按行序检测月份回绕**
   *
   * ⚠️ ③ 是这里最容易做错的地方。台账里常见「序列号 + 无年份中文日期」混排，
   *    而中文那段往往**跨年**（8月…12月、1月…9月）。
   *    无论统一补哪一年都会错 —— 必须按行序：月份从 ≥10 月跳到 ≤3 月时年份 +1。
   *    （此坑在本项目的某家入库明细上真实存在，统一补年会导致近7日入库少算一笔。）
   */
  function parseDateColumn(rows, col, fallbackYear) {
    const abs = rows.map(r => (r ? absDate(r[col]) : null));

    let maxY = 0;
    for (const a of abs) if (a) maxY = Math.max(maxY, a.getUTCFullYear());
    const fallback = maxY || fallbackYear || new Date().getUTCFullYear();

    // 起始锚点：第一个无年份值**之前最近的**绝对日期。台账按时间排，
    // 紧邻的前一条绝对日期最接近真实年份。
    let seedYear = fallback, seedMonth = null;
    const firstMD = rows.findIndex((r, i) => r && !abs[i] && monthDayOf(r[col]));
    if (firstMD >= 0) {
      for (let i = firstMD - 1; i >= 0; i--) {
        if (abs[i]) { seedYear = abs[i].getUTCFullYear(); seedMonth = abs[i].getUTCMonth() + 1; break; }
      }
    }

    const out = new Array(rows.length).fill(null);
    let curYear = seedYear;
    let prevMonth = seedMonth;

    for (let i = 0; i < rows.length; i++) {
      if (!rows[i]) continue;
      if (abs[i]) {
        out[i] = abs[i];
        continue;                    // ⚠️ 绝对日期不更新 curYear/prevMonth：
                                     //    否则夹在中间会把无年份序列的年份重置回去
      }
      const md = monthDayOf(rows[i][col]);
      if (!md) continue;
      // 台账按时间排：月份变小就意味着跨年
      if (prevMonth !== null && md.m < prevMonth) curYear += 1;
      out[i] = new Date(Date.UTC(curYear, md.m - 1, md.d));
      prevMonth = md.m;
    }
    return out;
  }

  /**
   * 从文件名里读数据日期。
   * 「XX库存(9月8日).xlsx」这种命名本身就标明了快照时点，比从流水里取最大值可靠
   * —— 流水里常混入异常日期（比如误录成半年后的日期），会带偏整个快照。
   */
  function dateFromFileName(name, refYear) {
    const s = String(name || '');
    const m = s.match(/(20\d{2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    const c = s.match(/(\d{1,2})\s*月\s*(\d{1,2})\s*[号日]/);
    if (c && refYear) return new Date(Date.UTC(refYear, +c[1] - 1, +c[2]));
    const p = s.match(/(20\d{2})(\d{2})(\d{2})/);
    if (p) return new Date(Date.UTC(+p[1], +p[2] - 1, +p[3]));
    return null;
  }

  function daysBetween(a, b) {
    return Math.round((a.getTime() - b.getTime()) / DAY);
  }

  function iso(d) {
    return d ? d.toISOString().slice(0, 10) : null;
  }

  // ======================================================== 数值 / 文本

  function num(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    const s = String(v).replace(/[,，\s]/g, '').replace(/^约/, '');
    const n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }

  function text(v) {
    return v === null || v === undefined ? '' : String(v).trim();
  }

  function classify(name, cats) {
    for (const [cat, kws] of cats) {
      for (const k of kws) if (name.includes(k)) return cat;
    }
    return '其他';
  }

  // ======================================================== 表结构识别

  /**
   * 在一张工作表里找表头，并判断它是什么表。
   * 返回 { role:'stock'|'flow', rowIndex, cols:{...} } 或 null
   */
  function detectSheet(sheet) {
    const rows = sheet.rows || [];
    const h = XL.findHeaderRow(rows, { keys: ['型号', '物料', '规格'], limit: 12 });
    if (!h) return null;
    const H = h.headers;

    const c = {
      spec: XL.findColumn(H, ['型号', '规格']),
      unit: XL.findColumn(H, ['单位']),
      name: XL.findColumn(H, ['名称', '物料名称']),
      stock: XL.findColumn(H, ['库存', '结存']),
      inQty: XL.findColumn(H, ['材料入库', '入库数量', '入库']),
      outQty: XL.findColumn(H, ['材料出库', '出库数量', '出库']),
      date: XL.findColumn(H, ['日期', '时间']),
      qty: XL.findColumn(H, ['数量', '求和项:数量']),
      safe: XL.findColumn(H, ['安全库存', '预警线']),
      src: XL.findColumn(H, ['材料来源', '来源']),
      project: XL.findColumn(H, ['项目']),
      person: XL.findColumn(H, ['领料人', '退料人员', '领料']),
      site: XL.findColumn(H, ['材料去向', '站点', '去向']),
    };

    if (c.spec < 0) return null;
    if (c.date >= 0 && c.qty >= 0) return { role: 'flow', rowIndex: h.rowIndex, cols: c };
    if (c.stock >= 0) return { role: 'stock', rowIndex: h.rowIndex, cols: c };
    return null;
  }

/**
 * 挑出需要的表。
 * ⚠️ 一份 xlsx 里常常同时有「入库明细」和「出库明细」两张流水表，必须都取。
 * 流水方向按**工作表名**判断（比按文件类型准）。
 */
function pickSheets(sheets) {
  const stocks = [], flows = [];
  for (const s of sheets || []) {
    const d = detectSheet(s);
    if (!d) continue;
    if (d.role === 'stock') {
      stocks.push({ sheet: s, d });
    } else if (d.role === 'flow') {
      let dir = null;
      if (/出库|发料|领料/.test(s.name)) dir = 'out';
      else if (/入库|到货/.test(s.name)) dir = 'in';
      flows.push({ sheet: s, d, dir });
    }
  }
  // 库存表优先选名字像"汇总/库存"的
  stocks.sort((a, b) => (/汇总|库存/.test(b.sheet.name) ? 1 : 0) -
                        (/汇总|库存/.test(a.sheet.name) ? 1 : 0));
  return { stocks, flows };
}

  // ======================================================== 主流程

  /**
   * @param sources 数据源数组
   * @param opts    { mapping, mappingRaw, master, safety, categories, days, snapshotDate }
   * @returns { data, report }
   */
  function mergeSources(sources, opts) {
    opts = opts || {};
    const days = opts.days || 7;
    const cats = opts.categories && opts.categories.length ? opts.categories : DEFAULT_CATS;
    const bySup = opts.mapping || {};        // {集成商: {原始型号: 标准型号}}
    const byRaw = opts.mappingRaw || {};     // {原始型号: 标准型号}（仅唯一映射）
    const master = opts.master || {};        // {标准型号: {code,name,unit}}
    const safety = opts.safety || {};
    const hasMaster = Object.keys(master).length > 0;
    const report = {
      sources: [], unresolved: {}, dropped: 0, warnings: [],
      counts: {}, snapshotDate: null, hasMaster,
    };

    // ---- 归一：(集成商, 原始型号) -> 标准型号，与 Python 端 to_std 一致 ----
    function toStd(sup, raw) {
      const s = text(raw);
      if (!s) return null;
      const perSup = bySup[sup];
      if (perSup && perSup[s]) return perSup[s];
      if (byRaw[s]) return byRaw[s];
      report.unresolved[s] = (report.unresolved[s] || 0) + 1;
      return null;
    }

    // ---- 第一遍：读出每份数据源里的记录 ----
    const stockRows = [];   // {std, qty, supplier}
    const flowRows = [];    // {std, date, qty, dir}
    let dateFallbackYear = null;

    for (const src of sources || []) {
      const picked = pickSheets(src.sheets);
      const sup = src.supplier || '';
      const rec = { id: src.id, fileName: src.fileName, supplier: sup,
                    type: src.type, sheets: (src.sheets || []).length,
                    stockRows: 0, flowRows: 0, dropped: 0, dataDate: src.dataDate || null };
      report.sources.push(rec);
      if (!picked.stocks.length && !picked.flows.length) {
        report.warnings.push(`${src.fileName}：没找到可识别的表（需含「型号」列，且含「库存」或「日期+数量」）`);
        continue;
      }

      for (const pick of picked.stocks) {
        const { sheet, d } = pick;
        const rows = sheet.rows;
        for (let i = d.rowIndex + 1; i < rows.length; i++) {
          const r = rows[i];
          if (!r) continue;
          const rawSpec = text(r[d.cols.spec]);
          if (!rawSpec) continue;
          if (/^(合计|小计|总计|汇总)/.test(rawSpec)) continue;
          const std = toStd(sup, rawSpec);
          if (!std) { rec.dropped++; report.dropped++; continue; }
          stockRows.push({ std, qty: num(r[d.cols.stock]), supplier: sup });
          rec.stockRows++;
        }
      }

      for (const pick of picked.flows) {
        const { sheet, d } = pick;
        const rows = sheet.rows;
        const dir = pick.dir || (src.type === '入库' ? 'in' : 'out');
        const dates = parseDateColumn(rows, d.cols.date, dateFallbackYear);
        for (let i = d.rowIndex + 1; i < rows.length; i++) {
          const r = rows[i];
          if (!r) continue;
          const rawSpec = text(r[d.cols.spec]);
          if (!rawSpec) continue;
          if (/^(合计|小计|总计|汇总)/.test(rawSpec)) continue;
          const dt = dates[i] || null;
          if (!dt) continue;                       // Python 端也是「没日期就跳过」
          if (!dateFallbackYear) dateFallbackYear = dt.getUTCFullYear();
          const std = toStd(sup, rawSpec);
          if (!std) { rec.dropped++; report.dropped++; continue; }
          flowRows.push({ std, date: dt, qty: num(r[d.cols.qty]), dir });
          rec.flowRows++;
        }
        const valid = dates.filter(Boolean);
        if (valid.length) {
          const mx = new Date(Math.max.apply(null, valid.map(x => x.getTime())));
          if (!rec.dataDate || mx > new Date(rec.dataDate)) rec.dataDate = mx;
          // 年份参考取「见过的最大年份」，不是第一条的年份——后者常被历史数据带偏
          const y = mx.getUTCFullYear();
          dateFallbackYear = Math.max(dateFallbackYear || 0, y);
        }
      }

      if (!rec.dataDate && src.dataDate) rec.dataDate = src.dataDate;
    }

    // ---- 快照日 ----
    // 优先级：用户指定 > 各表文件名里的日期 > 流水里的最大日期
    // 文件名优先，是因为「XX库存(9月8日)」本身就标明了快照时点；
    // 而流水里常混入异常日期（误录成半年后的日期），拿它当快照会整个带偏。
    let snapshot = null;
    if (opts.snapshotDate) {
      snapshot = new Date(opts.snapshotDate + 'T00:00:00Z');
    } else {
      const fromName = [];
      for (const s of sources) {
        const d = dateFromFileName(s.fileName, dateFallbackYear);
        if (d) fromName.push(d);
      }
      let cands = fromName;
      if (!cands.length) {
        cands = [];
        for (const s of sources) if (s.dataDate) cands.push(new Date(s.dataDate));
        for (const f of flowRows) if (f.date) cands.push(f.date);
      }
      if (cands.length) snapshot = new Date(Math.max.apply(null, cands.map(d => d.getTime())));
    }
    if (!snapshot || isNaN(snapshot.getTime())) snapshot = new Date();
    snapshot = new Date(Date.UTC(snapshot.getUTCFullYear(), snapshot.getUTCMonth(), snapshot.getUTCDate()));
    report.snapshotDate = iso(snapshot);
    const winFrom = new Date(snapshot.getTime() - (days - 1) * DAY);

    // ---- 聚合（按标准型号，不带单位——与 Python 一致）----
    const stockByStd = new Map();
    const supByStd = new Map();       // std -> Map<supplier, 正库存量>
    const in7 = new Map(), out7 = new Map(), lastOut = new Map();

    for (const r of stockRows) {
      stockByStd.set(r.std, (stockByStd.get(r.std) || 0) + r.qty);
      if (r.qty > 0) {
        let m = supByStd.get(r.std);
        if (!m) { m = new Map(); supByStd.set(r.std, m); }
        m.set(r.supplier, (m.get(r.supplier) || 0) + r.qty);
      }
    }

    for (const f of flowRows) {
      if (f.date >= winFrom && f.date <= snapshot) {
        const m = f.dir === 'in' ? in7 : out7;
        m.set(f.std, (m.get(f.std) || 0) + f.qty);
      }
      if (f.dir === 'out' && f.date <= snapshot) {
        const cur = lastOut.get(f.std);
        if (!cur || f.date > cur) lastOut.set(f.std, f.date);
      }
    }

    // ---- 输出：只遍历 master，再用「有库存或有流动」过滤 ----
    const seen = new Set();
    const rows = [];
    function emit(std) {
      if (seen.has(std)) return;
      seen.add(std);
      const meta = master[std] || {};
      const stock = Math.round((stockByStd.get(std) || 0) * 100) / 100;
      const i7 = Math.round((in7.get(std) || 0) * 100) / 100;
      const o7 = Math.round((out7.get(std) || 0) * 100) / 100;
      if (!(stock > 0 || i7 || o7)) return;        // 与 Python 的过滤一致
      const lo = lastOut.get(std);
      const age = lo ? daysBetween(snapshot, lo) : 9999;
      const supMap = supByStd.get(std);
      const wh = supMap
        ? Array.from(supMap.entries()).sort((a, b) => b[1] - a[1]).slice(0, 2)
            .map(x => x[0]).join('/')
        : '';
      const name = meta.name || std;
      rows.push([
        name, std, meta.unit || '',
        stock,
        Math.round((safety[std] || 0) * 100) / 100,
        i7, o7,
        age < 9999 ? age : 999,
        wh || '—',
        classify(name, cats),
      ]);
    }

    if (hasMaster) {
      for (const std of Object.keys(master)) emit(std);
    } else {
      // 没拿到物料主数据时退化为「遇到什么算什么」
      for (const std of stockByStd.keys()) emit(std);
      for (const std of in7.keys()) emit(std);
      for (const std of out7.keys()) emit(std);
    }

    rows.sort((a, b) => (b[3] - a[3]) || a[1].localeCompare(b[1], 'zh'));

    const materials = rows;
    const codes = materials.map(m => (master[m[1]] || {}).code || '');

    // ---- 趋势：本地只有单点快照，画成水平线 ----
    const total = materials.reduce((s, m) => s + m[3], 0);
    const trend = [[report.snapshotDate, Math.round(total * 100) / 100]];
    const safetyLevel = Math.round(materials.reduce((s, m) => s + m[4], 0) * 100) / 100;

    report.counts = {
      sources: report.sources.length,
      stockRows: stockRows.length,
      flowRows: flowRows.length,
      materials: materials.length,
      dropped: report.dropped,
      unresolved: Object.keys(report.unresolved).length,
      noSafe: materials.filter(m => !m[4]).length,
    };

    return {
      data: { snapshotDate: report.snapshotDate, trend, safetyLevel, codes, materials },
      report,
    };
  }

  const api = { mergeSources, parseDate, parseDateColumn, dateFromFileName, detectSheet,
                pickSheets, classify, num, text, serialDate, iso, looksSerial,
                absDate, monthDayOf, DEFAULT_CATS };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.WbMerge = api;
})(typeof window !== 'undefined' ? window : globalThis);
