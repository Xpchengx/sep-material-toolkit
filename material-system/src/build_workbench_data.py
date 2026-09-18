# -*- coding: utf-8 -*-
"""
工作台数据生成器 build_workbench_data.py
==========================================
从材料台账明细（入库/出库/库存汇总/型号映射）计算出工作台需要的数据，
输出 workbench-data.js —— 页面只负责显示，数据由本脚本生成。

数据流：
    金山文档《材料台账工作簿》（人工维护）
        ↓  kdocs 连接器拉取 / 或本地 xlsx 经 extract.py 抽取
    material-system/out/*.csv
        ↓  本脚本
    output/workbench-data.js   ← 页面通过 <script src> 加载

用法：
    python build_workbench_data.py
    python build_workbench_data.py --snapshot 2026-09-08 --days 14
    python build_workbench_data.py --safe-config safe_stock.json

关于安全库存：源头数据里没有这个字段，它是管理参数，必须人工定。
脚本会在首次运行时生成 safe_stock.csv 模板（按近7日出库量给出建议值），
由你修订后再跑一次即可生效。
"""
import sys, os, csv, json, re, math, collections, datetime, argparse

sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DEFAULT_SRC = os.path.join(ROOT, 'out')
DEFAULT_OUT = os.path.join(os.path.dirname(ROOT), 'output', 'workbench-data.js')

# ------------------------------------------------------------ 分类规则
CAT_RULES = [
    ('光模块', ('光模块', '彩光模块', 'CWDM', '模块-', '1.25G模块', '25GE')),
    ('主设备', ('RRU', 'BBU', 'AAU', '基带板', '主控', 'INCR', '皮站', '直放站',
              'RHUB', 'RHBU', 'pRRU', '远端单元', '轿厢', '载波功放', '交转直',
              '近端', '远端', '电源模块', 'AU（基站放大器）')),
    ('天线', ('天线', '吸顶', '板状', '对数周期', '射灯', 'GPS天线', '天馈')),
    ('线缆', ('馈线', '电缆', '光缆', '光电', '电源线', '混合缆', '地线', '接地线',
             '钢管', 'PVC管', '跳纤', '尾纤', '槽道')),
    ('光配', ('分纤箱', '分线箱', '野战光缆', '终端盒')),
    ('接头', ('接头', 'N-JJ', 'N-KK', '直角头', '转接')),
    ('无源器件', ('功分', '耦合', 'dB', '合路器', '合分波', '电桥', '负载', '衰减')),
    ('辅料', ('挡风板', '标签', '辅料', '胶带', '扎带')),
]


def classify(name):
    for cat, kws in CAT_RULES:
        if any(k in name for k in kws):
            return cat
    return '其他'


# ------------------------------------------------------------ 日期解析
CN = re.compile(r'(\d{1,2})\s*月\s*(\d{1,2})\s*[号日]')
ISO = re.compile(r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})')


def parse_dates(rows, keys=('日期',)):
    """
    明细表的日期列是【混合格式】：早期为 ISO（2024-11-10），
    后期为中文文本（8月26号，且没有年份）。
    按行序做年份推断：ISO 行给出锚点年份；中文行沿用锚点，月份回绕时年份 +1。
    """
    out = []
    year = None
    prev_m = None
    for r in rows:
        raw = ''
        for k in keys:
            if r.get(k):
                raw = r[k].strip()
                break
        d = None
        m1 = ISO.match(raw)
        if m1:
            y, mo, dd = map(int, m1.groups())
            try:
                d = datetime.date(y, mo, dd)
                year = y
                prev_m = mo
            except ValueError:
                d = None
        else:
            m2 = CN.search(raw)
            if m2:
                mo, dd = int(m2.group(1)), int(m2.group(2))
                if year is None:
                    out.append(None)
                    continue
                if prev_m is not None and mo < prev_m:
                    year += 1
                prev_m = mo
                try:
                    d = datetime.date(year, mo, dd)
                except ValueError:
                    d = None
        out.append(d)
    return out


def num(v):
    if v is None or v == '':
        return 0.0
    try:
        return float(str(v).replace(',', '').strip())
    except ValueError:
        return 0.0


def load(src, name):
    p = os.path.join(src, name)
    if not os.path.exists(p):
        return []
    with open(p, encoding='utf-8-sig') as f:
        return list(csv.DictReader(f))


# ------------------------------------------------------------ 主流程
def build(src, snapshot, days, safe_path):
    master = load(src, '物料主数据_归一.csv')
    mapping = load(src, '型号映射.csv')
    stock = load(src, '集成商库存汇总.csv')
    inb = load(src, '入库明细.csv')
    oub = load(src, '出库明细.csv')

    if not master:
        raise SystemExit(f'未找到物料主数据_归一.csv（目录 {src}）')

    # (集成商, 原始型号) -> 标准型号
    m2s = {(r['出现于集成商'], r['原始型号']): r['标准型号'] for r in mapping}
    raw2std = {}
    for (sup, raw), std in m2s.items():
        raw2std.setdefault(raw, set()).add(std)

    def to_std(sup, raw):
        v = m2s.get((sup, raw))
        if v:
            return v
        cands = raw2std.get(raw)
        return next(iter(cands)) if cands and len(cands) == 1 else None

    # ---- 库存：按标准型号汇总，并统计涉及集成商 ----
    stock_by_std = collections.defaultdict(float)
    sup_by_std = collections.defaultdict(collections.Counter)
    for r in stock:
        std = to_std(r['集成商'], r['型号'])
        if not std:
            continue
        q = num(r['库存'])
        stock_by_std[std] += q
        if q > 0:
            sup_by_std[std][r['集成商']] += q

    # ---- 明细日期 ----
    din = parse_dates(inb)
    dou = parse_dates(oub)

    if snapshot is None:
        all_d = [d for d in din + dou if d]
        snapshot = max(all_d) if all_d else datetime.date.today()
    snap = snapshot if isinstance(snapshot, datetime.date) else \
        datetime.date(*map(int, snapshot.split('-')))

    win_from = snap - datetime.timedelta(days=days - 1)

    # ---- 近N日 入库/出库 ----
    in7 = collections.defaultdict(float)
    out7 = collections.defaultdict(float)
    last_out = {}
    daily_in = collections.defaultdict(float)
    daily_out = collections.defaultdict(float)
    din_std = collections.defaultdict(lambda: collections.defaultdict(float))
    dout_std = collections.defaultdict(lambda: collections.defaultdict(float))

    for r, d in zip(inb, din):
        if not d:
            continue
        std = to_std(r['集成商'], r['型号'])
        if not std:
            continue
        q = num(r['数量'])
        daily_in[d] += q
        din_std[std][d] += q
        if win_from <= d <= snap:
            in7[std] += q

    for r, d in zip(oub, dou):
        if not d:
            continue
        std = to_std(r['集成商'], r['型号'])
        if not std:
            continue
        q = num(r['数量'])
        daily_out[d] += q
        dout_std[std][d] += q
        if win_from <= d <= snap:
            out7[std] += q
        if d <= snap:
            last_out[std] = max(last_out.get(std, d), d)

    # ---- 安全库存 ----
    # 优先级：① 金山表格「库存总览」表的安全库存列（人工在那维护最自然）
    #         ② 本地 safe_stock.csv
    safe_cfg = {}
    ov = load(src, '库存总览.csv')
    if ov and '安全库存' in ov[0]:
        for r in ov:
            m = (r.get('标准型号') or '').strip()
            v = (r.get('安全库存') or '').strip()
            if m and v:
                try:
                    safe_cfg[m] = float(v)
                except ValueError:
                    pass
    safe_src = f'库存总览.csv（{len(safe_cfg)} 项）' if safe_cfg else None

    if not safe_cfg and safe_path and os.path.exists(safe_path):
        with open(safe_path, encoding='utf-8-sig') as f:
            for r in csv.DictReader(f):
                try:
                    safe_cfg[r['标准型号']] = float(r['安全库存'])
                except (KeyError, ValueError):
                    pass
        safe_src = f'safe_stock.csv（{len(safe_cfg)} 项）'

    # ---- 组装物料 ----
    materials = []
    for m in master:
        std = m['标准型号']
        code = m['标准编码'] or ''
        try:
            mid = int(re.sub(r'\D', '', code) or 0)
        except ValueError:
            mid = 0
        qty = stock_by_std.get(std, 0.0)
        sups = sup_by_std.get(std)
        wh = '/'.join(k for k, _ in sups.most_common(2)) if sups else ''
        lo = last_out.get(std)
        age = (snap - lo).days if lo else 9999
        materials.append({
            'id': mid, 'code': code, 'name': m['标准名称'] or std, 'spec': std,
            'unit': m['单位'] or '', 'stock': round(qty, 2),
            'safe': safe_cfg.get(std, 0.0),
            'in7': round(in7.get(std, 0.0), 2),
            'out7': round(out7.get(std, 0.0), 2),
            'lastOut': age if age < 9999 else 999,
            'neverOut': lo is None,
            'warehouse': wh or '—',
            'category': classify(m['标准名称'] or std),
        })

    # 只保留有库存或有流动的，其余不展示
    materials = [m for m in materials if m['stock'] > 0 or m['in7'] or m['out7']]

    # ---- 趋势：由当前库存倒推 ----
    # 注意：必须只用「入选物料」的流水，否则趋势与库存合计对不上
    included = {m['spec'] for m in materials}
    f_in = collections.defaultdict(float)
    f_out = collections.defaultdict(float)
    for std in included:
        for d, q in din_std.get(std, {}).items():
            f_in[d] += q
        for d, q in dout_std.get(std, {}).items():
            f_out[d] += q

    total_now = sum(m['stock'] for m in materials)
    trend = []
    for i in range(days - 1, -1, -1):
        d = snap - datetime.timedelta(days=i)
        # 库存(t) = 库存(now) - Σ_{x>t} 入库(x) + Σ_{x>t} 出库(x)
        later_in = sum(v for k, v in f_in.items() if k > d)
        later_out = sum(v for k, v in f_out.items() if k > d)
        trend.append([d.strftime('%Y-%m-%d'),
                      round(total_now - later_in + later_out, 0)])

    safety_level = sum(m['safe'] for m in materials)

    return {
        'snapshot': snap, 'materials': materials, 'trend': trend,
        'safetyLevel': safety_level, 'totalNow': total_now,
        'safe_src': safe_src,
        'src_rows': {'master': len(master), 'mapping': len(mapping),
                     'stock': len(stock), 'in': len(inb), 'out': len(oub)},
        # 供 emit_meta 使用，不参与 emit_js / emit_json
        '_meta': {'mapping': mapping, 'master': master, 'safe': safe_cfg},
    }


def emit_meta(data, out_path, over_days=180, cover_ratio=0.3):
    """
    输出 wb-meta.json —— 浏览器端合并原始表格时要用的"规则与字典"。
    这样归一表、分类规则、安全库存、物料主数据只在 Python 端维护一份，
    页面端不重复实现，避免两条路径算出来对不上。

    结构对齐 build() 的口径：
      mapping[集成商][原始型号] = 标准型号      ← 精确匹配优先
      mappingRaw[原始型号]      = 标准型号      ← 仅当该型号在所有集成商下唯一时才给出
      master[标准型号]          = {code,name,unit}   ← 物料列表以此为准
    """
    mi = data.get('_meta') or {}
    mapping_rows = mi.get('mapping') or []
    master = mi.get('master') or []
    safe_cfg = mi.get('safe') or {}

    # 归一表：按集成商分组
    by_sup = {}
    raw2std = {}
    for r in mapping_rows:
        sup = (r.get('出现于集成商') or '').strip()
        raw = (r.get('原始型号') or '').strip()
        std = (r.get('标准型号') or '').strip()
        if not (sup and raw and std):
            continue
        by_sup.setdefault(sup, {}).setdefault(raw, std)
        raw2std.setdefault(raw, set()).add(std)

    # 物料主数据
    master_map = {}
    for r in master:
        std = (r.get('标准型号') or '').strip()
        if not std:
            continue
        master_map[std] = {
            'code': (r.get('标准编码') or '').strip(),
            'name': (r.get('标准名称') or '').strip() or std,
            'unit': (r.get('单位') or '').strip(),
        }

    # 集成商名单：从归一表的「出现于集成商」列拆出来，供页面识别拖入的文件
    sups = sorted({p.strip()
                   for r in mapping_rows
                   for p in (r.get('出现于集成商') or '').replace('|', '/').split('/')
                   if p.strip()})

    payload = {
        'generatedAt': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'snapshotDate': str(data['snapshot']),
        'thresholds': {'overDays': over_days, 'coverRatio': cover_ratio},
        'categories': [[c, list(k)] for c, k in CAT_RULES],
        'mapping': by_sup,
        # 只在唯一映射时给出，与 Python 端 to_std 的兜底逻辑一致
        'mappingRaw': {raw: next(iter(s)) for raw, s in raw2std.items() if len(s) == 1},
        'master': master_map,
        'safety': {k: float(v) for k, v in safe_cfg.items()},
        'suppliers': sups,
        'stats': {
            'mapping': sum(len(v) for v in by_sup.values()),
            'mappingRaw': len(raw2std),
            'master': len(master_map),
            'safetySet': len(safe_cfg),
            'suppliers': len(sups),
        },
    }
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, separators=(',', ':'))

    # 同名 .js 版本：页面用 <script src> 引入。
    # 不能只给 .json —— 本地 file:// 打开时 fetch 会被跨域策略拦掉。
    js_path = os.path.splitext(out_path)[0] + '.js'
    with open(js_path, 'w', encoding='utf-8') as f:
        f.write('/* 由 build_workbench_data.py 生成：归一表 / 分类规则 / 安全库存 / 阈值 */\n')
        f.write(f'/* 生成于 {payload["generatedAt"]} */\n')
        f.write('window.WB_META = ')
        json.dump(payload, f, ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')

    return payload['stats']


def emit_js(data, out_path):
    def esc(s):
        return str(s).replace('\\', '\\\\').replace('"', '\\"')

    L = []
    L.append('/* 由 build_workbench_data.py 自动生成，请勿手工编辑。')
    L.append('   数据源：金山文档《材料台账工作簿》')
    L.append(f"   生成时间：{datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
    L.append(f"   快照日期：{data['snapshot']}   物料 {len(data['materials'])} 种")
    L.append('   重新生成：python src/build_workbench_data.py */')
    L.append('window.RAW_DATA = {')
    L.append(f'  snapshotDate: "{data["snapshot"]}",')
    L.append('  trend: [')
    for d, v in data['trend']:
        L.append(f'    ["{d}", {int(v)}],')
    L.append('  ],')
    L.append(f'  safetyLevel: {int(data["safetyLevel"])},')
    L.append('  materials: [')
    L.append('    // 名称, 规格, 单位, 库存, 安全库存, 近7日入库, 近7日出库, 距最后出库天数, 周转仓, 分类')
    for m in data['materials']:
        L.append(
            f'    ["{esc(m["name"])}","{esc(m["spec"])}","{esc(m["unit"])}",'
            f'{m["stock"]:g},{m["safe"]:g},{m["in7"]:g},{m["out7"]:g},'
            f'{m["lastOut"]},"{esc(m["warehouse"])}","{esc(m["category"])}"],'
        )
    L.append('  ],')
    L.append('  /* 物料编码（与 materials 一一对应） */')
    L.append('  codes: [' + ','.join(f'"{esc(m["code"])}"' for m in data['materials']) + ']')
    L.append('};')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')


def emit_safe_template(data, path, reset=False):
    """
    安全库存是管理参数，源数据没有，必须人工定。
    这里给一个保守的建议值：近 7 日出库 × 0.5（约 3~4 日用量），
    对周转仓这类周转仓偏保守，避免把正常库存误判成短缺。
    """
    if os.path.exists(path) and not reset:
        return False
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['标准型号', '单位', '当前库存', '近7日出库', '建议安全库存', '安全库存'])
        for m in sorted(data['materials'], key=lambda x: -x['out7']):
            sug = int(math.ceil(m['out7'] * 0.5))
            w.writerow([m['spec'], m['unit'], f"{m['stock']:g}", f"{m['out7']:g}",
                        sug, sug])
    return True


def emit_json(data, out_path):
    """
    另出一份纯 JSON。页面用文件选择器读数据时优先按 JSON 解析，比求值 JS 更稳。
    """
    payload = {
        'snapshotDate': str(data['snapshot']),
        'safetyLevel': int(data['safetyLevel']),
        'trend': [[d, int(v)] for d, v in data['trend']],
        'codes': [m['code'] for m in data['materials']],
        'materials': [
            [m['name'], m['spec'], m['unit'], m['stock'], m['safe'],
             m['in7'], m['out7'], m['lastOut'], m['warehouse'], m['category']]
            for m in data['materials']
        ],
    }
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, separators=(',', ':'))
    return len(json.dumps(payload, ensure_ascii=False))


def emit_bundle(data, template_path, out_html):
    """
    生成「单文件实数据版」HTML：把数据内联进页面，替换掉 <script src="workbench-data.js">。
    用途：预览面板 / 发给别人看时，外链脚本可能取不到，单文件版最稳。
    """
    if not os.path.exists(template_path):
        return False
    with open(template_path, encoding='utf-8') as f:
        html = f.read()
    marker = '<script src="workbench-data.js"></script>'
    if marker not in html:
        return False

    js = ['<script>',
          '/* 数据由 build_workbench_data.py 内联生成，源：金山文档《材料台账工作簿》 */',
          f'/* 快照 {data["snapshot"]}　物料 {len(data["materials"])} 种　'
          f'生成 {datetime.datetime.now():%Y-%m-%d %H:%M} */',
          'window.RAW_DATA = {',
          f'  snapshotDate: "{data["snapshot"]}",',
          '  trend: [']
    for d, v in data['trend']:
        js.append(f'    ["{d}", {int(v)}],')
    js.append('  ],')
    js.append(f'  safetyLevel: {int(data["safetyLevel"])},')
    js.append('  materials: [')

    def esc(s):
        return str(s).replace('\\', '\\\\').replace('"', '\\"')

    for m in data['materials']:
        js.append(f'    ["{esc(m["name"])}","{esc(m["spec"])}","{esc(m["unit"])}",'
                  f'{m["stock"]:g},{m["safe"]:g},{m["in7"]:g},{m["out7"]:g},'
                  f'{m["lastOut"]},"{esc(m["warehouse"])}","{esc(m["category"])}"],')
    js.append('  ],')
    js.append('  codes: [' + ','.join(f'"{esc(m["code"])}"' for m in data['materials']) + ']')
    js.append('};')
    js.append('</script>')

    html = html.replace(marker, '\n'.join(js))
    with open(out_html, 'w', encoding='utf-8') as f:
        f.write(html)
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--src', default=DEFAULT_SRC, help='CSV 目录')
    ap.add_argument('--out', default=DEFAULT_OUT, help='输出的 workbench-data.js')
    ap.add_argument('--snapshot', default=None, help='快照日期 YYYY-MM-DD')
    ap.add_argument('--days', type=int, default=14, help='趋势天数')
    ap.add_argument('--safe-config', default=os.path.join(DEFAULT_SRC, 'safe_stock.csv'))
    ap.add_argument('--reset-safe', action='store_true',
                    help='重新生成安全库存模板（会覆盖现有 safe_stock.csv）')
    ap.add_argument('--bundle', default=None,
                    help='另出一份内联数据的单文件 HTML（外链取不到时用）')
    a = ap.parse_args()

    snap = None
    if a.snapshot:
        snap = datetime.date(*map(int, a.snapshot.split('-')))

    data = build(a.src, snap, a.days, a.safe_config)

    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    emit_js(data, a.out)

    # 同名 .json：供页面用文件选择器读取（不依赖求值 JS）
    json_out = os.path.splitext(a.out)[0] + '.json'
    n = emit_json(data, json_out)

    # wb-meta.json：页面端合并原始表格时用的规则与字典
    meta_out = os.path.join(os.path.dirname(os.path.abspath(json_out)), 'wb-meta.json')
    st = emit_meta(data, meta_out)

    created = emit_safe_template(data, a.safe_config, reset=a.reset_safe)

    ms = data['materials']
    n_safe = sum(1 for m in ms if m['safe'] > 0)
    short = [m for m in ms if m['stock'] < m['safe'] or
             (m['out7'] > 0 and m['stock'] < m['out7'] * 0.3)]
    over = [m for m in ms if m['lastOut'] > 180 and m['stock'] > 0]
    ratio = (len(short) + len(over)) / len(ms) if ms else 0

    print(f"快照日期 : {data['snapshot']}")
    print(f"物料品种 : {len(ms)}")
    print(f"库存合计 : {data['totalNow']:,.0f}   （跨单位合计，仅作趋势参考）")
    print(f"安全库存 : {n_safe} / {len(ms)} 项已设定"
          + (f"　来源：{data.get('safe_src')}" if data.get('safe_src') else ''))
    print(f"预警情况 : 短缺 {len(short)} 项、积压 {len(over)} 项，"
          f"占比 {ratio:.0%}（健康度 {1-ratio:.0%}）")
    print(f"趋势末端 : {data['trend'][-1][1]:,.0f}")
    print(f"数据源行数 : {data['src_rows']}")
    print(f"输出 -> {a.out}")
    print(f"        {json_out}  ({n:,} 字符，供页面文件选择器读取)")
    print(f"        {meta_out}  (归一表 {st['mapping']} 条、安全库存 {st['safetySet']} 项，供页面端合并原始表格)")

    if a.bundle:
        # 优先用仓库根的 index.html（PWA 版），回退到 output/ 里的旧模板
        ws = os.path.dirname(ROOT)
        tpl = os.path.join(ws, 'index.html')
        if not os.path.exists(tpl):
            tpl = os.path.join(os.path.dirname(os.path.abspath(a.out)),
                               '仓库物料库存管理智能体-工作台.html')
        ok = emit_bundle(data, tpl, a.bundle)
        print(("内联单文件 -> " + a.bundle) if ok
              else "[!] 未找到模板 HTML，跳过内联版")

    if ratio > 0.30:
        print(f"\n[!] 预警占比 {ratio:.0%} 偏高。通常是安全库存没调准——")
        print(f"    请修订 {a.safe_config} 的「安全库存」列后重跑。")
    else:
        print(f"\n[OK] 预警占比 {ratio:.0%}，在合理区间。")
    if created:
        print(f"\n[!] 已生成安全库存模板 -> {a.safe_config}")
        print("    源数据没有安全库存字段，必须人工定：建议值仅按近7日出库量推算，")
        print("    请结合实际补货周期与供应时效修订该文件的「安全库存」列。")


if __name__ == '__main__':
    main()
