/**
 * wb-transfer —— 物资调拨测算（设计量表 + 站点现有量表 → 缺口与调拨建议）
 *
 * 口径（与 Python 版一致，不可改）：
 *   可调出余量 = 现网可用量 − 该站点设计量      >0 才可调出，<0 必须调入
 *   分配：缺口封顶 → 项目级备料优先 → 站点余量按余量从大到小 → 中转站最后
 *
 * 依赖：xlsx-lite.js（读）、xlsx-write-lite.js（导出）
 */
(function (root) {
  'use strict';

  const CODE_RE = /^\d{6,10}$/;
  const DISTRICTS = ['开发区', '城区', '夷陵', '枝江', '当阳', '秭归', '长阳', '远安',
    '宜都', '五峰', '兴山', '点军', '猇亭', '西陵', '伍家岗'];

  // ------------------------------------------------------------ 基础工具
  const txt = v => (v === null || v === undefined ? '' : String(v).trim());
  const num = v => {
    if (v === null || v === undefined || v === '') return null;
    const n = Number(v);
    return isFinite(n) ? n : null;
  };

  function parseQty(v) {
    const s = txt(v);
    if (!s) return null;
    const m = s.match(/^([0-9]+(?:\.[0-9]+)?)/);
    return m ? Number(m[1]) : null;
  }

  function normSite(s) {
    let x = txt(s).replace(/\s+/g, '');
    // 去掉城市前缀：仅当站名不是以区县开头时，才取「首个区县出现处」为起点。
    // 不写死城市名，换个地市照样能用。
    if (!DISTRICTS.some(d => x.startsWith(d))) {
      for (const d of DISTRICTS) {
        const i = x.indexOf(d);
        if (i > 0 && i <= 4) { x = x.slice(i); break; }
      }
    }
    x = x.replace(/[-－]H5S$/, '');
    x = x.replace(/[（(]物资[）)]$/, '');
    x = x.replace(/设备$/, '');
    x = x.replace(/[-－]H5S$/, '');
    return x;
  }

  function distOf(s) {
    const n = normSite(s);
    for (const d of DISTRICTS) if (n.startsWith(d)) return d;
    return '';
  }

  function isBucket(s) {
    const x = txt(s);
    return x.includes('物资') && (x.includes('工程') || x.endsWith('（物资）'));
  }

  function findHeader(rows, must, anyOf) {
    const lim = Math.min(rows.length, 15);
    for (let i = 0; i < lim; i++) {
      const r = rows[i] || [];
      const joined = r.map(txt).join('|');
      const okMust = (must || []).every(m => joined.includes(m));
      const okAny = !anyOf || anyOf.length === 0 || anyOf.some(a => joined.includes(a));
      if (okMust && okAny) return i;
    }
    return -1;
  }

  function colOf(header, names, exclude) {
    for (let j = 0; j < (header || []).length; j++) {
      const h = txt(header[j]);
      if (!h) continue;
      if (exclude && exclude.some(e => h.includes(e))) continue;
      if (names.some(n => h.includes(n))) return j;
    }
    return -1;
  }

  // ------------------------------------------------------------ 读三张表
  function readGap(rows) {
    const hi = findHeader(rows, ['设计'], ['差异', '支出']);
    if (hi < 0) throw new Error('到货表：找不到含「设计量 + 差异/支出」的表头行');
    const h = rows[hi];
    const cCode = colOf(h, ['物料编码', '编码']);
    const cName = colOf(h, ['物料名称', '物料描述', '名称']);
    const cUnit = colOf(h, ['单位']);
    const cDes = colOf(h, ['设计量', '设计']);
    const cUse = colOf(h, ['支出', '已用']);
    const cDiff = colOf(h, ['差异', '余额']);
    if (cCode < 0 || cDiff < 0) throw new Error('到货表：缺「物料编码」或「差异」列');
    const out = {};
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const code = txt(r[cCode]);
      if (!CODE_RE.test(code)) continue;
      const diff = num(r[cDiff]);
      if (diff === null || diff >= 0) continue;
      out[code] = {
        name: cName >= 0 ? txt(r[cName]) : '',
        unit: cUnit >= 0 ? txt(r[cUnit]) : '',
        design: (cDes >= 0 ? num(r[cDes]) : 0) || 0,
        spend: (cUse >= 0 ? num(r[cUse]) : 0) || 0,
        gap: -diff,
      };
    }
    return out;
  }

  function readDesign(rows) {
    const hi = findHeader(rows, ['站点名称']);
    if (hi < 0) throw new Error('设计量表：找不到含「站点名称」的表头行');
    const header = rows[hi];
    const codeRow = hi >= 1 ? rows[hi - 1] : [];
    const cStage = colOf(header, ['工期', '阶段']);
    const cSite = colOf(header, ['站点名称', '站点']);
    if (cStage < 0 || cSite < 0) throw new Error('设计量表：缺「工期」或「站点名称」列');

    const col2code = [];
    for (let j = 0; j < codeRow.length; j++) {
      const v = txt(codeRow[j]);
      if (CODE_RE.test(v)) col2code.push({ col: j, code: v });
    }
    if (!col2code.length) throw new Error('设计量表：第 2 行没找到物料编码');

    const design = new Map();       // stage\u0001site -> {code: qty}
    const stageSites = new Map();   // stage -> Set(site)
    const stageRows = new Map();

    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const stage = txt(r[cStage]);
      const site = txt(r[cSite]);
      if (!stage || !site) continue;
      const nsite = normSite(site);
      if (!stageSites.has(stage)) stageSites.set(stage, new Set());
      stageSites.get(stage).add(nsite);
      stageRows.set(stage, (stageRows.get(stage) || 0) + 1);
      const row = {};
      for (const { col, code } of col2code) {
        const q = parseQty(r[col]);
        if (q) row[code] = q;
      }
      design.set(stage + '\u0001' + nsite, row);
    }
    return { design, stageSites, stageRows, col2code };
  }

  function readStock(rows) {
    const hi = findHeader(rows, ['可用量']);
    if (hi < 0) throw new Error('现有量表：找不到「可用量」列');
    const h = rows[hi];
    const cCode = colOf(h, ['物料/组合物料编码', '物料编码'], ['扩展']);
    const cAv = colOf(h, ['可用量']);
    const cSite = colOf(h, ['任务名称', '站点']);
    const cTask = colOf(h, ['任务编码']);
    if (cCode < 0 || cAv < 0 || cSite < 0) throw new Error('现有量表：缺「物料编码/可用量/任务名称」列');

    const agg = new Map();
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const code = txt(r[cCode]);
      const site = txt(r[cSite]);
      if (!code || !site) continue;
      const k = site + '\u0001' + code;
      const cur = agg.get(k) || { site, code, avail: 0, task: '' };
      cur.avail += num(r[cAv]) || 0;
      if (!cur.task && cTask >= 0) cur.task = txt(r[cTask]);
      agg.set(k, cur);
    }
    return [...agg.values()];
  }

  function detectStage(stockRows, stageSites) {
    const sites = new Set(stockRows.map(r => normSite(r.site)));
    if (!sites.size) return { stage: null, ratio: 0 };
    let best = null, bestR = 0;
    for (const [stage, ss] of stageSites) {
      let hit = 0;
      for (const s of sites) if (ss.has(s)) hit++;
      const r = hit / sites.size;
      if (r > bestR) { best = stage; bestR = r; }
    }
    return { stage: best, ratio: bestR };
  }

  // ------------------------------------------------------------ 核心测算
  function compute(input) {
    const { gapRows, designRows, stocks, target, hub } = input;
    const targets = target.split(/[,，]/).map(s => s.trim()).filter(Boolean);
    const hubs = hub ? hub.split(/[,，]/).map(s => s.trim()).filter(Boolean) : [];

    const gap = readGap(gapRows);
    const D = readDesign(designRows);
    const codeSet = new Set(D.col2code.map(x => x.code));
    const codes = Object.keys(gap).filter(c => codeSet.has(c));

    // 每份现有量表 → 识别工期（rows 传原始行或已解析行都行）
    const fileInfo = [];
    const stock = [];
    for (const st of stocks) {
      const parsed = (st.rows && Array.isArray(st.rows[0])) ? readStock(st.rows) : st.rows;
      const det = detectStage(parsed, D.stageSites);
      const ok = det.ratio >= 0.3;
      fileInfo.push({
        file: st.name, stage: ok ? det.stage : '未识别', ratio: det.ratio,
        sites: new Set(parsed.map(r => normSite(r.site))).size, rows: parsed.length,
      });
      for (const r of parsed) stock.push(Object.assign({}, r, { stage: ok ? det.stage : '' }));
    }

    // 聚合
    const av = new Map();       // stage\u0001site\u0001code -> qty
    const taskOf = new Map();
    const rawOf = new Map();    // stage\u0001site -> 原始站名（判断备料桶要用）
    for (const r of stock) {
      const nsite = normSite(r.site);
      const k = r.stage + '\u0001' + nsite + '\u0001' + r.code;
      av.set(k, (av.get(k) || 0) + r.avail);
      if (!taskOf.has(k)) taskOf.set(k, r.task || '');
      const k2 = r.stage + '\u0001' + nsite;
      if (!rawOf.has(k2)) rawOf.set(k2, r.site);
    }

    // 分类
    const hubKeys = new Set(), matchedKeys = new Set(), unmatchedKeys = new Set();
    for (const k of av.keys()) {
      const [stage, nsite] = k.split('\u0001');
      const k2 = stage + '\u0001' + nsite;
      if (hubs.includes(stage)) hubKeys.add(k2);
      else if (D.design.has(k2)) { if (!targets.includes(stage)) matchedKeys.add(k2); }
      else unmatchedKeys.add(k2);
    }

    // 逐站点 可用 vs 设计
    const perSite = [];
    for (const k2 of matchedKeys) {
      const [stage, nsite] = k2.split('\u0001');
      const row = D.design.get(k2) || {};
      for (const code of codes) {
        const q = row[code] || 0;
        const k = k2 + '\u0001' + code;
        const a = av.get(k) || 0;
        if (Math.abs(q) < 1e-9 && Math.abs(a) < 1e-9) continue;
        perSite.push({
          stage, site: rawOf.get(k2) || nsite, task: taskOf.get(k) || '',
          code, avail: a, design: q, left: Math.round((a - q) * 1e4) / 1e4,
        });
      }
    }

    const surplus = new Map();  // code -> [rows]
    for (const p of perSite) {
      if (p.left <= 0.01) continue;
      if (!surplus.has(p.code)) surplus.set(p.code, []);
      surplus.get(p.code).push(Object.assign({}, p));
    }

    // 目标需求
    const targetSet = new Set(targets);
    const demand = new Map();   // code -> [[site, need]]
    for (const [k2, row] of D.design) {
      const [stage, nsite] = k2.split('\u0001');
      if (!targetSet.has(stage)) continue;
      for (const code of codes) {
        if (!row[code]) continue;
        if (!demand.has(code)) demand.set(code, []);
        demand.get(code).push([nsite, row[code]]);
      }
    }

    // 分配
    const plan = [];
    const remain = {};
    const used = new Map();     // code\u0001stage\u0001site -> 调出量

    for (const code of codes.slice().sort((a, b) => gap[b].gap - gap[a].gap)) {
      const total = gap[code].gap;
      const dem = (demand.get(code) || []).slice().sort((a, b) => b[1] - a[1]);

      const donors = [];
      for (const d of (surplus.get(code) || [])) {
        donors.push({ type: '站点余量', pri: 1, stage: d.stage, site: d.site, task: d.task,
          avail: d.avail, design: d.design, left: d.left });
      }
      for (const k2 of [...unmatchedKeys].sort()) {
        if (!isBucket(rawOf.get(k2) || '')) continue;
        const [stage, nsite] = k2.split('\u0001');
        const k = k2 + '\u0001' + code;
        const v = av.get(k) || 0;
        if (v <= 0) continue;
        donors.push({ type: '项目级备料', pri: 0, stage, site: rawOf.get(k2) || nsite,
          task: taskOf.get(k) || '', avail: v, design: 0, left: v });
      }
      for (const k2 of [...hubKeys].sort()) {
        const [stage, nsite] = k2.split('\u0001');
        const k = k2 + '\u0001' + code;
        const v = av.get(k) || 0;
        if (v <= 0) continue;
        donors.push({ type: '中转站', pri: 2, stage, site: rawOf.get(k2) || nsite,
          task: taskOf.get(k) || '', avail: v, design: 0, left: v });
      }
      donors.sort((a, b) => (a.pri - b.pri) || (b.left - a.left));

      const got = new Map();
      let remaining = total;
      for (const d of donors) {
        if (remaining <= 1e-9) break;
        const dd = distOf(d.site);
        const order = dem.slice().sort((a, b) =>
          ((distOf(a[0]) !== dd) - (distOf(b[0]) !== dd)) || (b[1] - a[1]));
        for (const [dsite, dneed] of order) {
          if (remaining <= 1e-9) break;
          const still = dneed - (got.get(dsite) || 0);
          if (still <= 1e-9) continue;
          const take = Math.min(still, d.left, remaining);
          if (take <= 1e-9) continue;
          plan.push({
            code, name: gap[code].name, unit: gap[code].unit, total,
            type: d.type, stage: d.stage, site: d.site, task: d.task,
            avail: d.avail, design: d.design, left: Math.round(d.left * 1e4) / 1e4,
            take: Math.round(take * 1e4) / 1e4, dsite, dneed,
            same: !!dd && dd === distOf(dsite),
          });
          d.left -= take;
          got.set(dsite, (got.get(dsite) || 0) + take);
          remaining -= take;
          const uk = code + '\u0001' + d.stage + '\u0001' + d.site;
          used.set(uk, (used.get(uk) || 0) + take);
        }
      }
      remain[code] = Math.round(remaining * 1e4) / 1e4;
    }

    // 余料
    const leftovers = [];
    for (const code of codes) {
      for (const d of (surplus.get(code) || [])) {
        const uk = code + '\u0001' + d.stage + '\u0001' + d.site;
        const left = Math.round((d.left - (used.get(uk) || 0)) * 1e4) / 1e4;
        if (left > 0.01) leftovers.push({
          code, name: gap[code].name, unit: gap[code].unit,
          stage: d.stage, site: d.site, task: d.task, left,
        });
      }
    }

    // 待确认候选
    const cand = [];
    for (const k2 of [...unmatchedKeys].sort()) {
      if (isBucket(rawOf.get(k2) || '')) continue;
      const [stage, nsite] = k2.split('\u0001');
      for (const code of codes) {
        const k = k2 + '\u0001' + code;
        const v = av.get(k) || 0;
        if (v > 0) cand.push({
          code, name: gap[code].name, unit: gap[code].unit,
          stage, site: rawOf.get(k2) || nsite, task: taskOf.get(k) || '', avail: v,
        });
      }
    }
    const candCover = new Map();
    for (const c of cand) candCover.set(c.code, (candCover.get(c.code) || 0) + c.avail);

    const unresolved = codes.filter(c => remain[c] > 0.01).map(c => ({
      code: c, name: gap[c].name, unit: gap[c].unit,
      total: gap[c].gap,
      solved: Math.round((gap[c].gap - remain[c]) * 1e4) / 1e4,
      left: remain[c],
    }));

    const covered = new Set(fileInfo.filter(f => f.stage !== '未识别').map(f => f.stage));
    const coverage = [...D.stageSites.entries()]
      .map(([stage, ss]) => ({
        stage, sites: ss.size, rows: D.stageRows.get(stage) || 0,
        has: covered.has(stage),
        role: targets.includes(stage) ? '目标' : (hubs.includes(stage) ? '中转站' : '来源'),
      }))
      .sort((a, b) => b.sites - a.sites);

    return {
      gap, codes, plan, unresolved, perSite, leftovers, cand, candCover,
      coverage, fileInfo, targets, hubs,
      stats: {
        codes: codes.length, closed: codes.length - unresolved.length,
        transfers: plan.length,
        sources: new Set(plan.map(p => p.site)).size,
        dests: new Set(plan.map(p => p.dsite)).size,
        leftover: leftovers.length,
        sites: matchedKeys.size, hubSites: hubKeys.size, unmatched: unmatchedKeys.size,
      },
    };
  }

  // ------------------------------------------------------------ 导出 xlsx
  function toSheets(r) {
    const g = r.gap;
    const s1 = [['物料编码', '物料名称', '单位', '设计量', '已支出', '缺口']];
    r.codes.slice().sort((a, b) => g[b].gap - g[a].gap).forEach(c => {
      s1.push([c, g[c].name, g[c].unit, g[c].design, g[c].spend, g[c].gap]);
    });

    const s2 = [['序', '物料编码', '物料名称', '单位', '来源项目', '调出站点', '调出任务编码',
      '调出前可用', '该站设计量', '可调余量', '调出量', '调入站点', '该站设计需求', '类型']];
    r.plan.forEach((p, i) => s2.push([i + 1, p.code, p.name, p.unit, p.stage, p.site, p.task || '—',
      p.avail, p.design, p.left, p.take, p.dsite, p.dneed,
      p.type + (p.same ? '（同区县）' : '')]));

    const s3 = [['来源项目', '站点', '物料编码', '物料名称', '单位', '现网可用', '设计量', '差额', '判定']];
    r.perSite.slice().sort((a, b) => (b.left > 0) - (a.left > 0) || Math.abs(b.left) - Math.abs(a.left))
      .forEach(p => {
        const v = p.left > 0.01 ? '可调出' : (p.left < -0.01 ? '需调入' : '平衡');
        s3.push([p.stage, p.site, p.code, g[p.code].name, g[p.code].unit, p.avail, p.design, p.left, v]);
      });

    const s4 = [['物料编码', '物料名称', '单位', '缺口', '已调', '仍缺', '候选存量', '结论']];
    r.unresolved.forEach(u => {
      const cv = Math.round((r.candCover.get(u.code) || 0) * 1e4) / 1e4;
      s4.push([u.code, u.name, u.unit, u.total, u.solved, u.left, cv,
        cv >= u.left ? `候选 ${cv} 已够，确认站点对应即可补` : `须其他项目调拨（候选仅 ${cv}）`]);
    });

    const s5 = [['物料编码', '物料名称', '单位', '缺口', '已调', '仍缺', '状态', '主要来源']];
    r.codes.slice().sort((a, b) => g[b].gap - g[a].gap).forEach(c => {
      const solved = r.plan.filter(p => p.code === c).reduce((s, p) => s + p.take, 0);
      const left = Math.round((g[c].gap - solved) * 1e4) / 1e4;
      const cv = r.candCover.get(c) || 0;
      const status = left <= 0.01 ? '已闭环'
        : (cv >= left ? '待确认（候选可补）' : (solved > 0 ? '部分缺口' : '无来源'));
      const srcs = [...new Set(r.plan.filter(p => p.code === c).map(p => p.site))];
      s5.push([c, g[c].name, g[c].unit, g[c].gap, Math.round(solved * 1e4) / 1e4, left,
        status, srcs.join('、') || '—']);
    });

    const s6 = [['工期（项目）', '在线站点数', '在线明细行', '已给现有量表', '角色', '说明']];
    r.coverage.forEach(cv => s6.push([cv.stage, cv.sites, cv.rows, cv.has ? '有' : '缺', cv.role,
      cv.has ? '可参与调拨' : '未提供，不能判断是否有余量']));
    s6.push([]);
    s6.push(['本次读入文件：']);
    r.fileInfo.forEach(f => s6.push([
      `${f.file}  →  ${f.stage}（站点名重合 ${(f.ratio * 100).toFixed(0)}%，${f.sites} 站 / ${f.rows} 行）`]));

    const hubName = r.hubs.join('、') || '中转站';
    const s7 = [['物料编码', '物料名称', '单位', '来源项目', '所在站点', '任务编码', '可暂存量', '建议']];
    r.leftovers.slice().sort((a, b) => a.code.localeCompare(b.code) || b.left - a.left)
      .forEach(l => s7.push([l.code, l.name, l.unit, l.stage, l.site, l.task || '—', l.left,
        `调 ${hubName} 暂存`]));

    return [
      { name: '1-缺口', rows: s1 },
      { name: '2-调拨执行单', rows: s2 },
      { name: '3-站点可用vs设计', rows: s3 },
      { name: '4-尚未闭环', rows: s4 },
      { name: '5-按物料汇总', rows: s5 },
      { name: '6-数据源覆盖', rows: s6 },
      { name: '7-余料暂存建议', rows: s7 },
    ];
  }

  // ------------------------------------------------------------ UI
  const esc = s => String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  const fmt = n => {
    if (n === null || n === undefined || n === '') return '';
    const f = Number(n);
    if (!isFinite(f)) return '';
    return Math.abs(f - Math.round(f)) < 1e-6
      ? Math.round(f).toLocaleString('zh-CN')
      : f.toLocaleString('zh-CN', { maximumFractionDigits: 4 });
  };

  const state = { gapRows: null, gapName: '', designRows: null, designName: '', stocks: [], result: null };

  function log(msg, kind) {
    const el = document.getElementById('tfLog');
    if (!el) return;
    const t = new Date().toTimeString().slice(0, 8);
    el.insertAdjacentHTML('beforeend',
      `<div class="${kind === 'err' ? 'tf-err' : ''}">[${t}] ${esc(msg)}</div>`);
    el.scrollTop = el.scrollHeight;
  }

  async function readFile(file) {
    const r = await root.XlsxLite.readTable(file);
    const sh = (r.sheets || [])[0];
    if (!sh) throw new Error('这个文件里没读到工作表');
    return sh.rows;
  }

  function bindSlot(id, multiple, onDone) {
    const box = document.getElementById(id);
    if (!box) return;
    const input = box.querySelector('input[type=file]');
    const nameEl = box.querySelector('.tf-name');
    const pick = () => input.click();
    box.addEventListener('click', pick);
    box.addEventListener('dragover', e => { e.preventDefault(); box.classList.add('tf-over'); });
    box.addEventListener('dragleave', () => box.classList.remove('tf-over'));
    box.addEventListener('drop', async e => {
      e.preventDefault(); box.classList.remove('tf-over');
      await handle([...(e.dataTransfer.files || [])]);
    });
    input.addEventListener('change', async () => { await handle([...input.files]); input.value = ''; });

    async function handle(files) {
      if (!files.length) return;
      const use = multiple ? files : files.slice(0, 1);
      try {
        box.classList.add('tf-busy');
        await onDone(use);
        if (nameEl) nameEl.textContent = multiple
          ? (state.stocks.length ? `${state.stocks.length} 份：` + state.stocks.map(s => s.name).join('、') : '未选择')
          : (multiple ? '' : use[0].name);
      } catch (e) {
        log(e.message || String(e), 'err');
      } finally {
        box.classList.remove('tf-busy');
      }
    }
  }

  function render(r) {
    const st = r.stats;
    const kpi = document.getElementById('tfKpi');
    if (kpi) {
      kpi.innerHTML = [
        ['缺口项数', st.codes], ['已闭环', `${st.closed}/${st.codes}`],
        ['调拨笔数', st.transfers], ['调出点', st.sources], ['调入站点', st.dests],
        ['可暂存余料', st.leftover],
      ].map(([k, v]) => `<div class="kpi"><div class="kpi-k">${k}</div><div class="kpi-v">${v}</div></div>`).join('');
    }
    document.getElementById('tfKpiCard').style.display = '';

    const out = document.getElementById('tfOut');
    const g = r.gap;
    const LIMIT = 300;

    // 执行单
    let h = `<div class="card"><div class="card-h"><h2>调拨执行单</h2>
      <span class="hint">共 ${r.plan.length} 笔；来源须先满足自身设计量</span></div><div class="card-b">
      <table><thead><tr><th>序</th><th>物料编码</th><th>物料名称</th><th>单位</th><th>来源项目</th>
      <th>调出站点</th><th>调出任务编码</th><th class="num">可调余量</th><th class="num">调出量</th>
      <th>调入站点</th><th class="num">该站需求</th><th>类型</th></tr></thead><tbody>`;
    r.plan.slice(0, LIMIT).forEach((p, i) => {
      h += `<tr${p.type === '中转站' ? ' class="tf-hub"' : (p.type === '项目级备料' ? ' class="tf-bucket"' : '')}>
        <td class="num">${i + 1}</td><td>${esc(p.code)}</td><td>${esc(p.name)}</td><td>${esc(p.unit)}</td>
        <td>${esc(p.stage)}</td><td>${esc(p.site)}</td><td>${esc(p.task || '—')}</td>
        <td class="num">${fmt(p.left)}</td><td class="num"><b>${fmt(p.take)}</b></td>
        <td>${esc(p.dsite)}</td><td class="num">${fmt(p.dneed)}</td>
        <td>${esc(p.type)}${p.same ? '（同区县）' : ''}</td></tr>`;
    });
    h += '</tbody></table>';
    if (r.plan.length > LIMIT) h += `<div class="count-note">仅显示前 ${LIMIT} 行，完整清单请导出 xlsx</div>`;
    h += '</div></div>';

    // 尚未闭环
    if (r.unresolved.length) {
      h += `<div class="card"><div class="card-h"><h2>尚未闭环</h2>
        <span class="hint">${r.unresolved.length} 项还差</span></div><div class="card-b">
        <table><thead><tr><th>物料编码</th><th>物料名称</th><th>单位</th>
        <th class="num">缺口</th><th class="num">已调</th><th class="num">仍缺</th><th class="num">候选存量</th></tr></thead><tbody>`;
      r.unresolved.forEach(u => {
        h += `<tr><td>${esc(u.code)}</td><td>${esc(u.name)}</td><td>${esc(u.unit)}</td>
          <td class="num">${fmt(u.total)}</td><td class="num">${fmt(u.solved)}</td>
          <td class="num bad">${fmt(u.left)}</td>
          <td class="num">${fmt(r.candCover.get(u.code) || 0)}</td></tr>`;
      });
      h += '</tbody></table></div></div>';
    }

    // 按物料汇总
    h += `<div class="card"><div class="card-h"><h2>按物料汇总</h2></div><div class="card-b">
      <table><thead><tr><th>物料编码</th><th>物料名称</th><th>单位</th><th class="num">缺口</th>
      <th class="num">已调</th><th class="num">仍缺</th><th>状态</th><th>主要来源</th></tr></thead><tbody>`;
    r.codes.slice().sort((a, b) => g[b].gap - g[a].gap).forEach(c => {
      const solved = r.plan.filter(p => p.code === c).reduce((s, p) => s + p.take, 0);
      const left = Math.round((g[c].gap - solved) * 1e4) / 1e4;
      const cv = r.candCover.get(c) || 0;
      const status = left <= 0.01 ? '<span class="ok">已闭环</span>'
        : (cv >= left ? '<span class="warn">待确认（候选可补）</span>'
          : `<span class="bad">${solved > 0 ? '部分缺口' : '无来源'}</span>`);
      const srcs = [...new Set(r.plan.filter(p => p.code === c).map(p => p.site))];
      h += `<tr><td>${esc(c)}</td><td>${esc(g[c].name)}</td><td>${esc(g[c].unit)}</td>
        <td class="num">${fmt(g[c].gap)}</td><td class="num">${fmt(solved)}</td>
        <td class="num">${left > 0 ? fmt(left) : '0'}</td><td>${status}</td>
        <td>${esc(srcs.slice(0, 2).join('、'))}${srcs.length > 2 ? ` 等 ${srcs.length} 处` : ''}</td></tr>`;
    });
    h += '</tbody></table></div></div>';

    // 余料
    if (r.leftovers.length) {
      const hubName = r.hubs.join('、') || '中转站';
      h += `<div class="card"><div class="card-h"><h2>余料暂存建议</h2>
        <span class="hint">${r.leftovers.length} 条；建议调 ${esc(hubName)} 暂存</span></div><div class="card-b">
        <table><thead><tr><th>物料编码</th><th>物料名称</th><th>单位</th><th>来源项目</th>
        <th>所在站点</th><th class="num">可暂存量</th></tr></thead><tbody>`;
      r.leftovers.slice(0, 100).forEach(l => {
        h += `<tr><td>${esc(l.code)}</td><td>${esc(l.name)}</td><td>${esc(l.unit)}</td>
          <td>${esc(l.stage)}</td><td>${esc(l.site)}</td><td class="num">${fmt(l.left)}</td></tr>`;
      });
      h += '</tbody></table>';
      if (r.leftovers.length > 100) h += '<div class="count-note">仅显示前 100 条，完整清单请导出 xlsx</div>';
      h += '</div></div>';
    }

    // 数据源覆盖
    h += `<div class="card"><div class="card-h"><h2>数据源覆盖</h2>
      <span class="hint">没给现有量表的项目，一律视为不可用</span></div><div class="card-b">
      <table><thead><tr><th>工期（项目）</th><th class="num">在线站点数</th><th class="num">明细行</th>
      <th>现有量表</th><th>角色</th></tr></thead><tbody>`;
    r.coverage.forEach(cv => {
      h += `<tr><td>${esc(cv.stage)}</td><td class="num">${cv.sites}</td><td class="num">${cv.rows}</td>
        <td>${cv.has ? '<span class="ok">有</span>' : '<span class="bad">缺</span>'}</td>
        <td>${esc(cv.role)}</td></tr>`;
    });
    h += '</tbody></table><div class="tf-files">';
    r.fileInfo.forEach(f => {
      h += `<div>${esc(f.file)} → <b>${esc(f.stage)}</b>
        （站点名重合 ${(f.ratio * 100).toFixed(0)}%，${f.sites} 站 / ${f.rows} 行）</div>`;
    });
    h += '</div></div></div>';

    out.innerHTML = h;
    document.getElementById('tfExport').disabled = false;
  }

  function boot() {
    if (!document.getElementById('v-transfer')) return false;

    const savedT = localStorage.getItem('tf_target');
    const savedH = localStorage.getItem('tf_hub');
    if (savedT) document.getElementById('tfTarget').value = savedT;
    if (savedH) document.getElementById('tfHub').value = savedH;

    bindSlot('tfGap', false, async files => {
      state.gapRows = await readFile(files[0]);
      state.gapName = files[0].name;
      log(`到货表：${files[0].name}（${state.gapRows.length} 行）`);
    });
    bindSlot('tfDesign', false, async files => {
      state.designRows = await readFile(files[0]);
      state.designName = files[0].name;
      log(`设计量表：${files[0].name}（${state.designRows.length} 行）`);
    });
    bindSlot('tfStock', true, async files => {
      for (const f of files) {
        const rows = await readFile(f);
        state.stocks = state.stocks.filter(s => s.name !== f.name);
        state.stocks.push({ name: f.name, rows });
        log(`现有量表：${f.name}（${rows.length} 行）`);
      }
    });

    document.getElementById('tfClear').addEventListener('click', () => {
      state.stocks = [];
      state.result = null;
      document.querySelector('#tfStock .tf-name').textContent = '未选择';
      document.getElementById('tfKpiCard').style.display = 'none';
      document.getElementById('tfOut').innerHTML = '';
      document.getElementById('tfExport').disabled = true;
      document.getElementById('tfLog').innerHTML = '';
      log('已清空现有量表');
    });

    document.getElementById('tfRun').addEventListener('click', () => {
      const target = document.getElementById('tfTarget').value.trim();
      const hub = document.getElementById('tfHub').value.trim();
      if (!state.gapRows) return log('还没有选「到货/支出表」', 'err');
      if (!state.designRows) return log('还没有选「设计量表」', 'err');
      if (!state.stocks.length) return log('还没有选「站点现有量表」', 'err');
      if (!target) return log('请填目标工期', 'err');
      localStorage.setItem('tf_target', target);
      localStorage.setItem('tf_hub', hub);
      try {
        const t0 = performance.now();
        const r = compute({ gapRows: state.gapRows, designRows: state.designRows,
          stocks: state.stocks, target, hub });
        state.result = r;
        render(r);
        r.fileInfo.forEach(f => log(`${f.file} → ${f.stage}（重合 ${(f.ratio * 100).toFixed(0)}%）`));
        log(`测算完成：${r.stats.transfers} 笔，闭环 ${r.stats.closed}/${r.stats.codes}，`
          + `用时 ${(performance.now() - t0).toFixed(0)} ms`);
      } catch (e) {
        log(e.message || String(e), 'err');
        console.error(e);
      }
    });

    document.getElementById('tfExport').addEventListener('click', () => {
      const r = state.result;
      if (!r) return;
      const t = document.getElementById('tfTarget').value.trim().replace(/[\\/:*?"<>|]/g, '');
      const bytes = root.XlsxWrite.build(toSheets(r));
      const blob = new Blob([bytes], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `${t || '调拨'}项目调拨执行方案.xlsx`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
      log('已导出 xlsx');
    });

    return true;
  }

  /** 直接跑一轮（跳过文件选择；测试和控制台调试用） */
  function runWith(input) {
    const r = compute(input);
    state.result = r;
    render(r);
    return r;
  }

  const api = { boot, compute, runWith, readGap, readDesign, readStock, detectStage,
    toSheets, normSite, distOf, isBucket };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.WbTransfer = api;
})(typeof window !== 'undefined' ? window : globalThis);
