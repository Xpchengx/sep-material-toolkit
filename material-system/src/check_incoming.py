# -*- coding: utf-8 -*-
"""
check_incoming.py —— 原始文件入库校验
========================================
在文件进入流水线**之前**体检，把格式问题挡在源头。
格式问题在源头修，比在脚本里兜底省十倍力气。

用法：
    python check_incoming.py <文件...>                     # 单文件或多个
    python check_incoming.py --dir <目录>                  # 批量
    python check_incoming.py <文件> --type 入库
    python check_incoming.py <文件> --std 标准型号表.csv
    python check_incoming.py --selftest                    # 自检工具本身

集成商名单从外部配置读入（仓库不存真实名单）：
    --suppliers 集成商A,集成商B,集成商C
    或环境变量 LEDGER_SUPPLIERS=集成商A,集成商B,集成商C
    或同目录 suppliers.json  {"suppliers": ["...", "..."]}
未配置时不校验集成商，只提示需人工确认。

退出码：0 = 无错误；1 = 有错误（可据此拦下）
"""
import os, re, sys, csv, json, argparse, datetime

sys.stdout.reconfigure(encoding='utf-8')

# ------------------------------------------------------------ 规范定义
# 命名规范：{集成商}_{类型}_{YYYYMMDD}.xlsx
TYPE_TOKENS = ['入库明细', '出库明细', '退料明细', '库存', '盘点']
NAME_RE = re.compile(
    r'^([^_]{1,12})_(' + '|'.join(TYPE_TOKENS) + r')_(\d{8})\.(xlsx|xls|csv)$')

# 每种数据类型的关键列（用关键字匹配表头，不要求完全一致）
SCHEMA = {
    '入库': {'required': ['日期', '型号', '数量'], 'suggest': ['集成商', '单位', '项目']},
    '出库': {'required': ['日期', '型号', '数量'], 'suggest': ['集成商', '项目', '去向']},
    '退料': {'required': ['日期', '型号', '数量'], 'suggest': ['集成商', '单位', '项目']},
    '库存': {'required': ['型号', '库存'],       'suggest': ['集成商', '名称', '单位']},
    '盘点': {'required': ['型号', '数量'],       'suggest': ['集成商', '单位']},
}

CN_DATE = re.compile(r'(\d{1,2})\s*月\s*(\d{1,2})\s*[号日]')
ISO_DATE = re.compile(r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})')
COMPACT_DATE = re.compile(r'(20\d{2})(\d{2})(\d{2})')
NUM_RE = re.compile(r'^-?[\d,]+(\.\d+)?$')


def load_suppliers(cli_value):
    """集成商名单：命令行 > 环境变量 > 同目录 suppliers.json > 空。"""
    if cli_value:
        return [s.strip() for s in cli_value.split(',') if s.strip()]
    env = os.environ.get('LEDGER_SUPPLIERS')
    if env:
        return [s.strip() for s in env.split(',') if s.strip()]
    cfg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'suppliers.json')
    if os.path.exists(cfg):
        try:
            with open(cfg, encoding='utf-8') as f:
                return [str(s).strip() for s in json.load(f).get('suppliers', [])]
        except Exception:
            pass
    return []


def guess_type(name):
    if '出库' in name or '发料' in name:
        return '出库'
    if '入库' in name or '到货' in name:
        return '入库'
    if '退料' in name or '废料' in name:
        return '退料'
    if '盘点' in name:
        return '盘点'
    if '库存' in name:
        return '库存'
    return None


def find_supplier(name, suppliers):
    """从文件名里认出集成商。名单为空或认不出时返回 None。"""
    for s in suppliers:
        if s and s in name:
            return s
    return None


def read_rows(path, max_rows=200000):
    """返回 (表头列表, 数据行列表, 元信息)。大文件流式读。"""
    meta = {'merged': 0, 'sheets': []}
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv':
        with open(path, encoding='utf-8-sig', newline='') as f:
            rd = list(csv.reader(f))
        meta['sheets'] = ['(csv)']
        return (rd[0] if rd else []), rd[1:max_rows + 1], meta

    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    meta['sheets'] = wb.sheetnames
    ws = wb[wb.sheetnames[0]]
    try:
        if hasattr(ws, 'merged_cells'):
            meta['merged'] = len(list(ws.merged_cells.ranges))
    except Exception:
        pass
    rows, header = [], None
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            header = [('' if c is None else str(c).strip()) for c in row]
            continue
        rows.append(list(row))
        if len(rows) >= max_rows:
            break
    wb.close()
    return (header or []), rows, meta


def find_col(header, keywords):
    for i, h in enumerate(header):
        for k in keywords:
            if k in h:
                return i, h
    return None, None


class Report:
    def __init__(self, path):
        self.path = path
        self.items = []

    def add(self, level, item, msg):
        self.items.append((level, item, msg))

    @property
    def errors(self):
        return [x for x in self.items if x[0] == 'ERR']

    def dump(self):
        icon = {'ERR': '✗', 'WARN': '⚠', 'INFO': 'ℹ', 'OK': '✓'}
        out = [f'检查: {os.path.basename(self.path)}']
        for lv, item, msg in self.items:
            out.append(f'  {icon[lv]} {item:<8} {msg}')
        n_err = len(self.errors)
        n_warn = len([x for x in self.items if x[0] == 'WARN'])
        if n_err:
            out.append(f'\n结论: ✗ 有 {n_err} 项错误、{n_warn} 项警告 —— 不建议直接入库')
        elif n_warn:
            out.append(f'\n结论: ⚠ 无错误，{n_warn} 项警告 —— 可用，但建议先确认')
        else:
            out.append('\n结论: ✓ 通过')
        return '\n'.join(out)


def check(path, dtype=None, std_models=None, suppliers=None):
    suppliers = suppliers or []
    r = Report(path)
    name = os.path.basename(path)

    # ---- 1. 文件名 ----
    m = NAME_RE.match(name)
    if m:
        sup, t, d = m.group(1), m.group(2), m.group(3)
        if suppliers and sup not in suppliers:
            r.add('WARN', '文件名', f'集成商「{sup}」不在配置名单里（{suppliers}）')
        else:
            r.add('OK', '文件名', f'符合规范（集成商={sup}，类型={t}，日期={d}）')
        if dtype is None:
            dtype = t.replace('明细', '')
    else:
        r.add('ERR', '文件名', f'不符合命名规范。建议改为：'
                              f'{suggest_name(name, dtype, suppliers)}')

    if dtype is None:
        dtype = guess_type(name)
    if dtype is None:
        r.add('WARN', '类型', '无法从文件名判断数据类型，请用 --type 指定')
        dtype = '库存'

    # ---- 2. 读文件 ----
    try:
        header, rows, meta = read_rows(path)
    except Exception as e:
        r.add('ERR', '读取', f'打不开：{type(e).__name__}: {e}')
        return r
    r.add('OK', '读取', f'{len(rows)} 行 × {len(header)} 列；工作表 {meta["sheets"][:3]}')

    if len(meta['sheets']) > 1:
        r.add('WARN', '工作表', f'含 {len(meta["sheets"])} 个表，默认只读第一个：'
                                f'{meta["sheets"][0]}')

    # ---- 3. 表头 ----
    spec = SCHEMA[dtype]
    missing = [k for k in spec['required'] if find_col(header, [k])[0] is None]
    if missing:
        r.add('ERR', '表头', f'缺少必需列：{missing}；实际表头 {header[:12]}')
    else:
        r.add('OK', '表头', f'关键列齐全；实际表头 {header[:12]}')
    miss_sug = [k for k in spec['suggest'] if find_col(header, [k])[0] is None]
    if miss_sug:
        r.add('INFO', '表头', f'建议补充列：{miss_sug}')

    # ---- 4. 日期 ----
    di, dname = find_col(header, ['日期', '时间'])
    if di is not None:
        iso = cn = bad = empty = 0
        for row in rows:
            v = row[di] if di < len(row) else None
            s = '' if v is None else str(v).strip()
            if not s:
                empty += 1
                continue
            if isinstance(v, (datetime.date, datetime.datetime)) or ISO_DATE.search(s):
                iso += 1
            elif CN_DATE.search(s):
                cn += 1
            else:
                bad += 1
        if bad:
            r.add('ERR', '日期', f'{bad} 行无法解析（列「{dname}」）')
        if cn:
            r.add('WARN', '日期', f'{cn} 行是中文格式且**无年份**（如「8月6号」），'
                                  f'脚本需靠行序推断年份，容易错。建议统一为 YYYY-MM-DD')
        if empty:
            r.add('WARN', '日期', f'{empty} 行为空')
        if not bad and not cn and (iso or empty):
            r.add('OK', '日期', f'{iso} 行全部为标准日期格式')

    # ---- 5. 数量 ----
    qi, qname = find_col(header, ['数量', '库存', '设计量'])
    if qi is not None:
        badvals, neg = [], 0
        for idx, row in enumerate(rows, start=2):
            v = row[qi] if qi < len(row) else None
            if v is None or str(v).strip() == '':
                continue
            if isinstance(v, (int, float)):
                if v < 0:
                    neg += 1
                continue
            s = str(v).strip()
            m2 = NUM_RE.match(s)
            if not m2:
                if len(badvals) < 5:
                    badvals.append(f'第{idx}行 "{s}"')
            else:
                try:
                    if float(s.replace(',', '')) < 0:
                        neg += 1
                except ValueError:
                    pass
        if badvals:
            r.add('ERR', '数量', f'{len(badvals)}+ 个非数值：{badvals}')
        if neg:
            r.add('WARN', '数量', f'{neg} 行为负数（退料可能正常，请确认）')
        if not badvals and not neg:
            r.add('OK', '数量', f'列「{qname}」全部为数值')

    # ---- 6. 单位 ----
    ui, uname = find_col(header, ['单位'])
    if ui is not None:
        uniq = {}
        for row in rows:
            v = row[ui] if ui < len(row) else None
            s = '' if v is None else str(v).strip()
            if s:
                uniq[s] = uniq.get(s, 0) + 1
        if len(uniq) > 6:
            r.add('WARN', '单位', f'取值多达 {len(uniq)} 种：{list(uniq)[:8]}… 建议限定词表')
        else:
            r.add('OK', '单位', f'取值 {list(uniq.items())[:8]}')

    # ---- 7. 空行 ----
    key_idx = [i for i, _ in filter(
        lambda t: t[0] is not None,
        [find_col(header, ['型号']), find_col(header, ['日期'])])]
    if key_idx:
        blank = sum(1 for row in rows
                    if all((i >= len(row) or row[i] is None or not str(row[i]).strip())
                           for i in key_idx))
        if blank:
            r.add('WARN', '空行', f'{blank} 行关键列为空（可能是分隔行或汇总行）')

    # ---- 8. 型号 ----
    mi, mname = find_col(header, ['型号'])
    if mi is not None and std_models:
        unknown = {}
        for row in rows:
            v = row[mi] if mi < len(row) else None
            s = '' if v is None else str(v).strip()
            if s and s not in std_models:
                unknown[s] = unknown.get(s, 0) + 1
        if unknown:
            r.add('WARN', '型号', f'{len(unknown)} 个型号不在标准型号表中，需人工确认归并：'
                                  f'{list(unknown)[:8]}')
        else:
            r.add('OK', '型号', '全部型号均可在标准型号表中找到')
    elif mi is not None:
        r.add('INFO', '型号', '未提供标准型号表，跳过型号核对（--std 指定）')

    return r


def suggest_name(name, dtype, suppliers=None):
    """
    根据现有文件名给规范化建议。
    日期识别顺序：2026-09-08 → 20260908 → 8月8日（无年份，会标注需核对）。
    集成商从配置名单里认；认不出就标注出来让人工补。
    """
    suppliers = suppliers or []
    sup = find_supplier(name, suppliers)
    note = ''
    if sup is None:
        sup = '集成商X'
        note = '　（集成商没认出来，请手工补）'

    t = dtype or guess_type(name) or '库存'

    m = ISO_DATE.search(name)
    if m:
        d = f'{m.group(1)}{int(m.group(2)):02d}{int(m.group(3)):02d}'
    else:
        mc = COMPACT_DATE.search(name)
        if mc:
            d = f'{mc.group(1)}{mc.group(2)}{mc.group(3)}'
        else:
            cm = CN_DATE.search(name)
            if cm:
                ym = re.search(r'(20\d{2})', name)
                if ym:
                    y = ym.group(1)
                else:
                    y = str(datetime.date.today().year)
                    note += '　（年份是推断的，请核对）'
                d = f'{y}{int(cm.group(1)):02d}{int(cm.group(2)):02d}'
            else:
                d = 'YYYYMMDD'
                note += '　（没识别到日期，请手工填）'

    suffix = '' if t == '库存' else '明细'
    return f'{sup}_{t}{suffix}_{d}.xlsx{note}'


def self_test():
    """
    自检：应能给出合法建议的名字，必须真的合法。
    expect=True 的用例不带 note（即集成商与日期都能认出来）。
    """
    SUP = ['集成商A', '集成商B', '集成商C']
    cases = [
        ('集成商B-材料库存2025-08-08.xlsx', None, True),
        ('材料库存（集成商C）_20250106.xlsx', None, True),
        ('集成商A材料库存8月8日).xlsx', None, True),
        ('随便一个名字.xlsx', None, False),        # 没有集成商也没有日期
    ]
    bad = []
    for c, t, expect in cases:
        s = suggest_name(c, t, SUP)
        ok = bool(NAME_RE.match(s.split('　')[0]))
        if expect and not ok:
            bad.append((c, s))
    return cases, bad, SUP


def load_std(path):
    if not path or not os.path.exists(path):
        return None
    s = set()
    with open(path, encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            for k in ('标准型号', '标准名称'):
                if row.get(k):
                    s.add(row[k].strip())
    return s or None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('paths', nargs='*', help='要检查的文件（可给多个）')
    ap.add_argument('--dir', help='批量检查目录')
    ap.add_argument('--type', choices=list(SCHEMA), help='数据类型')
    ap.add_argument('--std', help='标准型号表 CSV（用于型号核对）')
    ap.add_argument('--suppliers', help='集成商名单，逗号分隔（也可用 LEDGER_SUPPLIERS）')
    ap.add_argument('--selftest', action='store_true', help='自检命名建议逻辑')
    a = ap.parse_args()

    if a.selftest:
        cases, bad, sup = self_test()
        print(f'命名建议自检（示例名单 {sup}）：')
        for c, t, expect in cases:
            s = suggest_name(c, t, sup)
            mark = 'OK ' if (bool(NAME_RE.match(s.split('　')[0])) == expect) else '!! '
            print(f'   {mark}{c}')
            print(f'       -> {s}')
            if not expect:
                print('       （这类输入本就给不出合法建议，需人工补集成商/日期）')
        print()
        print('全部符合预期' if not bad else f'!! 有 {len(bad)} 个「本应合法却不合法」：{bad}')
        return 1 if bad else 0

    suppliers = load_suppliers(a.suppliers)
    std = load_std(a.std)
    if not suppliers:
        print('[提示] 未配置集成商名单，将不校验集成商。'
              '配置方式：--suppliers A,B,C 或环境变量 LEDGER_SUPPLIERS 或同目录 suppliers.json\n')

    targets = []
    if a.dir:
        for f in sorted(os.listdir(a.dir)):
            if f.lower().endswith(('.xlsx', '.xls', '.csv')) and not f.startswith('~$'):
                targets.append(os.path.join(a.dir, f))
    elif a.paths:
        targets = list(a.paths)
    else:
        ap.error('请给出文件路径或 --dir')

    if not targets:
        print('没有找到可检查的文件（已跳过 ~$ 临时文件）')
        return 0

    total_err = 0
    for p in targets:
        rep = check(p, a.type, std, suppliers)
        print(rep.dump())
        print()
        total_err += len(rep.errors)

    print(f'===== 共检查 {len(targets)} 个文件，{total_err} 项错误 =====')
    return 1 if total_err else 0


if __name__ == '__main__':
    sys.exit(main())
