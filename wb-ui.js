/* ============================================================================
   wb-ui.js —— 拖拽载入 / 数据源登记 / 更新追踪 / 在线同步
   ----------------------------------------------------------------------------
   依赖：xlsx-lite.js（读表格）、wb-merge.js（合并归一）
        以及 index.html 里的 applyData / refreshUI / toast / $ 等全局函数。

   三条数据来源，最终都汇到 applyData()：
     ① 拖入或选择本地表格（xlsx/csv）→ 解析 → 合并 → 应用
     ② 选择生成好的 workbench-data.json / .js → 直接应用
     ③ 填入在线地址 → 拉取 → 按 ① 或 ② 处理

   ⚠️ 数据全程留在本机，不上传。
   ========================================================================== */
(function (root) {
  'use strict';

  const DB = 'material-workbench', STORE_SRC = 'wb-sources', STORE_SET = 'wb-settings';
  const LS_KEY = 'wb-ui-settings';

  // 过期阈值（天）：数据日期距今天数超过 bad 判为「长期未更新」
  const DEF_SETTINGS = { warnDays: 7, badDays: 15, onlineUrl: '', autoLoad: true };

  let settings = Object.assign({}, DEF_SETTINGS);
  let registry = [];        // 本次会话载入的原始表格（内存中，含解析结果）
  let busy = false;

  // ------------------------------------------------------------ IndexedDB

  function idb() {
    return new Promise((res, rej) => {
      const r = indexedDB.open(DB, 2);
      r.onupgradeneeded = () => {
        const db = r.result;
        if (!db.objectStoreNames.contains('handles')) db.createObjectStore('handles');
        if (!db.objectStoreNames.contains(STORE_SRC)) db.createObjectStore(STORE_SRC, { keyPath: 'key' });
        if (!db.objectStoreNames.contains(STORE_SET)) db.createObjectStore(STORE_SET);
      };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }
  async function idbPut(store, val, key) {
    const db = await idb();
    return new Promise((res, rej) => {
      const tx = db.transaction(store, 'readwrite');
      key === undefined ? tx.objectStore(store).put(val) : tx.objectStore(store).put(val, key);
      tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error);
    });
  }
  async function idbAll(store) {
    const db = await idb();
    return new Promise((res, rej) => {
      const tx = db.transaction(store, 'readonly');
      const rq = tx.objectStore(store).getAll();
      rq.onsuccess = () => res(rq.result || []); rq.onerror = () => rej(rq.error);
    });
  }
  async function idbGet(store, key) {
    const db = await idb();
    return new Promise((res, rej) => {
      const tx = db.transaction(store, 'readonly');
      const rq = tx.objectStore(store).get(key);
      rq.onsuccess = () => res(rq.result); rq.onerror = () => rej(rq.error);
    });
  }
  async function idbClear(store) {
    const db = await idb();
    return new Promise((res, rej) => {
      const tx = db.transaction(store, 'readwrite');
      tx.objectStore(store).clear();
      tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error);
    });
  }

  // ------------------------------------------------------------ 设置

  function loadSettings() {
    try {
      const s = JSON.parse(localStorage.getItem(LS_KEY) || '{}');
      settings = Object.assign({}, DEF_SETTINGS, s);
    } catch (e) { settings = Object.assign({}, DEF_SETTINGS); }
    return settings;
  }
  function saveSettings() {
    try { localStorage.setItem(LS_KEY, JSON.stringify(settings)); } catch (e) {}
  }

  // ------------------------------------------------------------ 工具

  function daysSince(d) {
    const iso = toISODate(d);
    if (!iso) return null;
    const t = new Date(iso + 'T00:00:00Z').getTime();
    if (isNaN(t)) return null;
    const now = new Date();
    const today = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
    return Math.round((today - t) / 86400000);
  }

  /** 统一成 YYYY-MM-DD。日期可能来自 Date 对象、ISO 串或各种规范写法 */
  function toISODate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v.toISOString().slice(0, 10);
    const s = String(v).trim();
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    // 兜底：能喂给 Date 的英文日期串
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
  }
  function todayISO() {
    const n = new Date();
    return new Date(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate())).toISOString().slice(0, 10);
  }
  function fmtSize(n) {
    if (n < 1024) return n + ' B';
    if (n < 1048576) return (n / 1024).toFixed(0) + ' KB';
    return (n / 1048576).toFixed(1) + ' MB';
  }

  /** 从文件名猜集成商与类型；猜不出交给用户在下拉里改 */
  function guessMeta(fileName) {
    const s = String(fileName || '');
    let supplier = '', type = '';
    const sups = root.WB_SUPPLIERS
      || (root.WB_META && root.WB_META.suppliers)
      || [];
    // 长名字优先，避免「中兴」误配到「中兴通讯」这类更长的名字上
    for (const x of sups.slice().sort((a, b) => b.length - a.length)) {
      if (x && s.includes(x)) { supplier = x; break; }
    }
    if (/出库|发料|领料/.test(s)) type = '出库';
    else if (/入库|到货/.test(s)) type = '入库';
    else if (/退料|废料/.test(s)) type = '退料';
    else if (/盘点/.test(s)) type = '盘点';
    else if (/库存|汇总/.test(s)) type = '库存';
    return { supplier, type };
  }

  // ------------------------------------------------------------ 载入流程

  /** 解析一批文件（拖入或选择），并入注册表后重算 */
  async function addFiles(files) {
    if (busy) return;
    busy = true; showDropBusy(true);
    const errs = [];
    try {
      for (const f of files) {
        try {
          const t0 = Date.now();
          const res = await root.XlsxLite.readTable(f);
          const meta = guessMeta(f.name);

          if (res.kind === 'json' && res.json && res.json.materials) {
            // 直接是可用的工作台数据
            applyWorkbench(res.json, 'local', f.name);
            continue;
          }
          if (!res.sheets || !res.sheets.length) {
            errs.push(`${f.name}：没读到工作表`);
            continue;
          }
          const rec = {
            key: 'live:' + f.name + ':' + (f.lastModified || Date.now()),
            fileName: f.name,
            supplier: meta.supplier,
            type: meta.type || '库存',
            size: f.size,
            sheetNames: res.sheets.map(s => s.name),
            warnings: res.warnings || [],
            ms: Date.now() - t0,
            sheets: res.sheets,
            live: true,
          };
          const i = registry.findIndex(x => x.fileName === f.name);
          if (i >= 0) registry[i] = rec; else registry.push(rec);
        } catch (e) {
          errs.push(`${f.name}：${e.message}`);
        }
      }
    } finally {
      busy = false; showDropBusy(false);
    }

    if (errs.length) toast('部分文件未载入：' + errs[0] + (errs.length > 1 ? ` 等 ${errs.length} 项` : ''), false);
    await recompute();
  }

  /** 用当前注册表重算工作台数据 */
  async function recompute() {
    if (!registry.length) { await renderSources(); return; }
    const meta = root.WB_META || {};
    let merged;
    try {
      merged = root.WbMerge.mergeSources(
        registry.filter(r => r.live).map(r => ({
          id: r.key, fileName: r.fileName, supplier: r.supplier,
          type: r.type, sheets: r.sheets,
        })),
        {
          mapping: meta.mapping, mappingRaw: meta.mappingRaw, master: meta.master,
          safety: meta.safety, categories: meta.categories, days: 7,
        });
    } catch (e) {
      toast('合并失败：' + e.message, false);
      return;
    }
    const { data, report } = merged;
    applyWorkbench(data, 'dropped', registry.map(r => r.fileName).join('、'));

    // 登记元信息（持久化，用于更新追踪）
    const now = new Date().toISOString();
    for (const s of report.sources) {
      await idbPut(STORE_SRC, {
        key: 'src:' + s.fileName,
        fileName: s.fileName, supplier: s.supplier, type: s.type,
        dataDate: toISODate(s.dataDate),
        loadedAt: now, stockRows: s.stockRows, flowRows: s.flowRows,
        sheets: (registry.find(r => r.fileName === s.fileName) || {}).sheetNames || [],
      });
    }
    window.WB_LAST_REPORT = report;

    const c = report.counts;
    const tip = `已载入 ${report.sources.length} 张表 · 物料 ${c.materials} 种`
      + (c.dropped ? ` · 跳过未匹配 ${c.dropped} 行` : '');
    toast(tip, true);
    await renderSources();
  }

  /** 应用最终数据（复用页面已有的 applyData） */
  function applyWorkbench(obj, origin, name) {
    if (typeof applyData === 'function') applyData(obj, origin, name);
  }

  // ------------------------------------------------------------ 更新追踪

  /** 计算每个数据源的「陈旧度」 */
  async function staleList() {
    const all = await idbAll(STORE_SRC);
    return all.map(s => {
      const age = daysSince(s.dataDate);
      let level = 'ok';
      if (age !== null) {
        if (age >= settings.badDays) level = 'bad';
        else if (age >= settings.warnDays) level = 'warn';
      }
      return Object.assign({}, s, { age, level });
    }).sort((a, b) => (b.age === null ? -1 : b.age) - (a.age === null ? -1 : a.age));
  }

  /** 总览页的提醒横幅 */
  async function renderBanner() {
    const box = document.getElementById('staleBanner');
    if (!box) return;
    const list = await staleList();
    const bad = list.filter(x => x.level === 'bad');
    const warn = list.filter(x => x.level === 'warn');
    if (!bad.length && !warn.length) {
      box.style.display = 'none';
      box.innerHTML = '';
      return;
    }
    const mk = (x) => `<b>${esc(x.fileName)}</b>（${esc(x.supplier)} ${esc(x.type)}，数据日期 ${x.dataDate || '未知'}，${x.age} 天前）`;
    box.style.display = '';
    box.className = 'stale-banner' + (bad.length ? ' bad' : '');
    box.innerHTML = (bad.length
      ? `<div class="sb-t">有 ${bad.length} 张表长期未更新，建议催报</div><div class="sb-l">${bad.slice(0, 4).map(mk).join('<br>')}</div>`
      : '')
      + (warn.length
        ? `<div class="sb-t${bad.length ? ' mt' : ''}">有 ${warn.length} 张表已超过 ${settings.warnDays} 天未更新</div><div class="sb-l">${warn.slice(0, 4).map(mk).join('<br>')}</div>`
        : '');
  }

  function esc(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  // ------------------------------------------------------------ 数据源面板

  async function renderSources() {
    const box = document.getElementById('srcList');
    if (!box) { await renderBanner(); return; }
    const list = await staleList();
    const rep = window.WB_LAST_REPORT;

    const liveRows = registry.map(r => {
      const m = (rep && rep.sources.find(s => s.fileName === r.fileName)) || {};
      return {
        fileName: r.fileName, supplier: r.supplier, type: r.type,
        dataDate: toISODate(m.dataDate),
        stockRows: m.stockRows, flowRows: m.flowRows,
        size: r.size, sheets: r.sheetNames, ms: r.ms, live: true,
      };
    });

    if (!liveRows.length && !list.length) {
      box.innerHTML = `<tr><td colspan="7" class="dim">还没有载入任何表格。把集成商发来的 xlsx 拖到页面上即可。</td></tr>`;
    } else if (liveRows.length) {
      box.innerHTML = liveRows.map(r => {
        const age = daysSince(r.dataDate);
        const lv = age === null ? 'ok' : (age >= settings.badDays ? 'bad' : (age >= settings.warnDays ? 'warn' : 'ok'));
        return `<tr>
          <td>${esc(r.fileName)}</td>
          <td>${esc(r.supplier) || '<span class="dim">未识别</span>'}</td>
          <td>${esc(r.type)}</td>
          <td class="mono">${r.dataDate || '<span class="dim">—</span>'}</td>
          <td>${age === null ? '—' : age + ' 天'}</td>
          <td>${r.stockRows || 0} / ${r.flowRows || 0} 行</td>
          <td>${lvlTag(lv, age)}</td>
        </tr>`;
      }).join('');
    } else {
      box.innerHTML = `<tr><td colspan="7" class="dim">本次会话未载入表格；以下为历史记录（数据仍显示上次载入的结果）。</td></tr>`
        + list.map(r => `<tr class="dim">
            <td>${esc(r.fileName)}</td><td>${esc(r.supplier)}</td><td>${esc(r.type)}</td>
            <td class="mono">${r.dataDate || '—'}</td>
            <td>${r.age === null ? '—' : r.age + ' 天'}</td>
            <td>${r.stockRows || 0} / ${r.flowRows || 0} 行</td>
            <td>${lvlTag(r.level, r.age)}</td>
          </tr>`).join('');
    }

    const hist = document.getElementById('srcHistory');
    if (hist) {
      hist.textContent = list.length ? `共 ${list.length} 条历史记录，最近载入 ${(list[0].loadedAt || '').slice(0, 16).replace('T', ' ')}` : '';
    }
    // 侧栏徽标：提醒条数（长期未更新 + 偏旧）
    const badge = document.getElementById('navSrc');
    if (badge) {
      const n = list.filter(x => x.level === 'bad' || x.level === 'warn').length;
      badge.textContent = n;
      badge.className = 'badge' + (n ? '' : ' neutral');
    }
    await renderBanner();
  }

  function lvlTag(level, age) {
    if (level === 'bad') return `<span class="tag bad">长期未更新</span>`;
    if (level === 'warn') return `<span class="tag warn">偏旧</span>`;
    return `<span class="tag ok">正常</span>`;
  }

  // ------------------------------------------------------------ 在线同步

  /**
   * 在线地址同步。
   * ⚠️ 浏览器只能拉「允许跨域(CORS)的直链」。金山/腾讯文档这类在线表格
   *    是 JS 渲染的页面，且不开放跨域，**拉不到**——这里会明确说明并给替代路径。
   */
  async function syncOnline() {
    const inp = document.getElementById('onlineUrl');
    const url = (inp && inp.value || '').trim();
    if (!url) { toast('请先填入数据地址'); return; }

    if (/docs\.qq\.com|kdocs\.cn|365\.kdocs|doc\.weixin\.qq\.com|feishu|dingtalk/i.test(url)) {
      toast('在线文档链接无法直接抓取，请看下方说明', false);
      const hint = document.getElementById('onlineHint');
      if (hint) hint.style.display = '';
      return;
    }

    settings.onlineUrl = url; saveSettings();
    btnBusy('btnOnline', true, '同步中…');
    try {
      const res = await fetch(url, { cache: 'no-store' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      const buf = await res.arrayBuffer();
      const name = url.split('/').pop().split('?')[0] || 'online';
      const table = await root.XlsxLite.readTable(buf, { name });
      if (table.kind === 'json' && table.json && table.json.materials) {
        applyWorkbench(table.json, 'online', url);
        toast('已从在线地址载入工作台数据', true);
      } else if (table.sheets && table.sheets.length) {
        const meta = guessMeta(name);
        registry = registry.filter(r => r.fileName !== name);
        registry.push({ key: 'online:' + url, fileName: name, supplier: meta.supplier,
                        type: meta.type || '库存', sheets: table.sheets, live: true,
                        sheetNames: table.sheets.map(s => s.name), size: buf.byteLength });
        await recompute();
      } else {
        throw new Error('内容不是可识别的表格');
      }
    } catch (e) {
      toast('拉取失败：' + e.message + '（多半是对方未开放跨域，见下方说明）', false);
      const hint = document.getElementById('onlineHint');
      if (hint) hint.style.display = '';
    } finally {
      btnBusy('btnOnline', false, '⟳ 在线同步');
    }
  }

  function btnBusy(id, on, text) {
    const b = document.getElementById(id);
    if (!b) return;
    b.disabled = on;
    if (text) b.textContent = text;
  }

  // ------------------------------------------------------------ 拖拽

  let dragDepth = 0;
  function showDropBusy(on) {
    const o = document.getElementById('dropOverlay');
    if (o) o.classList.toggle('busy', !!on);
    const t = document.getElementById('dropText');
    if (t && on) t.textContent = '正在解析…';
    else if (t) t.textContent = '松手即可载入';
  }

  function bindDrop() {
    const ov = document.getElementById('dropOverlay');
    const show = (on) => { if (ov) ov.classList.toggle('on', on); if (!on) dragDepth = 0; };

    window.addEventListener('dragenter', e => {
      if (!e.dataTransfer || !Array.from(e.dataTransfer.types || []).includes('Files')) return;
      e.preventDefault(); dragDepth++; show(true);
    });
    window.addEventListener('dragover', e => {
      if (!e.dataTransfer || !Array.from(e.dataTransfer.types || []).includes('Files')) return;
      e.preventDefault(); e.dataTransfer.dropEffect = 'copy';
    });
    window.addEventListener('dragleave', e => {
      dragDepth = Math.max(0, dragDepth - 1);
      if (!dragDepth) show(false);
    });
    window.addEventListener('drop', async e => {
      if (!e.dataTransfer || !e.dataTransfer.files || !e.dataTransfer.files.length) return;
      e.preventDefault(); show(false);
      await addFiles(Array.from(e.dataTransfer.files));
    });
    // 拖到输入框上不要触发页面跳转
    document.addEventListener('dragover', e => {
      if (e.target && e.target.tagName === 'INPUT') e.preventDefault();
    });
  }

  // ------------------------------------------------------------ 设置面板

  function bindSettings() {
    const w = document.getElementById('setWarn'), b = document.getElementById('setBad');
    if (w) { w.value = settings.warnDays; w.addEventListener('change', () => {
      settings.warnDays = Math.max(1, +w.value || 7); saveSettings(); renderSources(); }); }
    if (b) { b.value = settings.badDays; b.addEventListener('change', () => {
      settings.badDays = Math.max(2, +b.value || 15); saveSettings(); renderSources(); }); }

    const clr = document.getElementById('btnClearSrc');
    if (clr) clr.addEventListener('click', async () => {
      await idbClear(STORE_SRC);
      registry = [];
      toast('已清除数据源记录', true);
      renderSources();
    });

    const pick = document.getElementById('btnPickFiles');
    if (pick && document.getElementById('filePick2')) {
      pick.addEventListener('click', () => document.getElementById('filePick2').click());
      document.getElementById('filePick2').addEventListener('change', async e => {
        const fs = Array.from(e.target.files || []);
        if (fs.length) await addFiles(fs);
        e.target.value = '';
      });
    }
    const on = document.getElementById('btnOnline');
    if (on) on.addEventListener('click', syncOnline);
    const ou = document.getElementById('onlineUrl');
    if (ou) { ou.value = settings.onlineUrl || ''; ou.addEventListener('change', () => {
      settings.onlineUrl = ou.value.trim(); saveSettings(); }); }
  }

  // ------------------------------------------------------------ 启动

  async function boot() {
    loadSettings();
    if (root.WB_META) { /* 已由页面注入 */ }
    bindDrop();
    bindSettings();
    // 上次的合并结果直接复用
    try {
      const saved = await idbGet('handles', 'wb:lastData');
      if (saved && saved.data && settings.autoLoad) {
        applyWorkbench(saved.data, 'local', saved.name || '上次载入');
      }
    } catch (e) {}
    await renderSources();
  }

  /** 保存合并结果，供下次打开时直接复用 */
  async function persistLast(data, name) {
    try { await idbPut('handles', { data, name, at: new Date().toISOString() }, 'wb:lastData'); } catch (e) {}
  }

  root.WbUI = { boot, addFiles, recompute, renderSources, staleList, syncOnline,
                get settings() { return settings; }, persistLast, esc, daysSince };
})(typeof window !== 'undefined' ? window : globalThis);
