# -*- coding: utf-8 -*-
"""
月报数据引擎 monthly_report_data.py  —  终版
=============================================
从三家集成商库存工作簿抽取「2026年8月」的入库/出库/盘存数据 -> JSON。

三家数据源的坑（已在代码中逐条处理）：
1. 集成商C「出库明细」max_row=104万（格式残留）→ 必须 read_only 流式读。
2. 集成商A日期列是【混合格式】：早期为 Excel 序列号(45597)，
   后期为中文文本('8月6号')，且**不含年份**。
   数据整体按时间顺序排列、末尾到 9 月 2 号 → 视为 2026 年；
   对序列号与中文文本分别解析，中文文本一律按 2026 年处理。
3. 集成商A「汇总」无「2025年结存」列，集成商B有 → 表头按关键字定位。
"""
import os, sys, os, json, datetime, re, collections
sys.stdout.reconfigure(encoding='utf-8')
import openpyxl

EXCEL_EPOCH = datetime.date(1899, 12, 30)
YEAR, MONTH = 2026, 8
# 台账根目录：优先取环境变量 LEDGER_DIR，默认用仓库内 data/集成商
_REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = os.environ.get('LEDGER_DIR') or os.path.join(_REPO, 'data', '集成商')
FILES = {
    '集成商A': os.path.join(BASE, '集成商A', '上海集成商A材料库存(9月1日）.xlsx'),
    '集成商B': os.path.join(BASE, '集成商B', '集成商B-材料库存2026-8-26.xlsx'),
    '集成商C': os.path.join(BASE, '集成商C', '集成商C库存(9月8日).xlsx'),
}
CN_MD = re.compile(r'(\d{1,2})\s*月\s*(\d{1,2})\s*[号日]')


def to_date(v, default_year=YEAR):
    """解析日期：Excel序列号 / datetime / 中文文本('8月6号') / 标准字符串。"""
    if v is None or v == '':
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, (int, float)):
        n = float(v)
        if 20000 < n < 60000:
            return EXCEL_EPOCH + datetime.timedelta(days=int(n))
        return None
    t = str(v).strip()
    m = CN_MD.search(t)                      # 中文「8月6号」→ 按 default_year 补年
    if m:
        try:
            return datetime.date(default_year, int(m.group(1)), int(m.group(2)))
        except ValueError:
            return None
    m = re.match(r'(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})', t)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def num(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', '').strip())
    except ValueError:
        return 0.0


def s(v):
    return '' if v is None else str(v).strip()


def hcol(row, *kw, exclude=()):
    for i, v in enumerate(row):
        t = s(v)
        if t and any(k in t for k in kw) and not any(x in t for x in exclude):
            return i
    return None


def extract(name, path):
    out = {'集成商': name, '文件': path, '异常': []}
    if not os.path.exists(path):
        out['异常'].append('文件不存在')
        return out
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out['工作表'] = list(wb.sheetnames)

    def load_detail(sheet, model_kw, extra):
        ws = wb[sheet]
        it = ws.iter_rows(values_only=True)
        head = []
        for i, row in enumerate(it, 1):
            head.append(row)
            if i >= 3:
                break
        hr = None
        for row in head:
            if hcol(row, *model_kw) is not None and hcol(row, '数量') is not None:
                hr = row
                break
        if hr is None:
            return []
        ci = hcol(hr, *model_kw)
        cq = hcol(hr, '数量')
        cd = hcol(hr, '日期')
        cex = {k: hcol(hr, *v) for k, v in extra.items()}
        recs = []
        blank = 0
        for row in it:
            if ci is None or ci >= len(row):
                continue
            model = s(row[ci])
            if not model:
                blank += 1
                if blank > 200:
                    break
                continue
            blank = 0
            g = lambda c: (row[c] if (c is not None and c < len(row)) else None)
            rec = {'型号': model,
                   '数量': num(g(cq)),
                   '日期': to_date(g(cd))}
            for k, c in cex.items():
                rec[k] = s(g(c))
            recs.append(rec)
        return recs

    for tag, sheet, mkw, extra in (
        ('入库', '入库明细', ('型号',), {'来源': ('材料来源', '来源'), '项目': ('项目',)}),
        ('出库', '出库明细', ('型号',), {'去向': ('材料去向', '去向'),
                                     '项目': ('项目',), '领料人': ('领料人',)}),
    ):
        if sheet not in wb.sheetnames:
            out[f'当月{tag}_条数'] = 0
            continue
        recs = load_detail(sheet, mkw, extra)
        cur = [r for r in recs if r['日期'] and (r['日期'].year, r['日期'].month) == (YEAR, MONTH)]
        out[f'{tag}_明细_总行数'] = len(recs)
        out[f'{tag}_有日期行数'] = len([r for r in recs if r['日期']])
        out[f'当月{tag}_条数'] = len(cur)
        out[f'当月{tag}_型号数'] = len(set(r['型号'] for r in cur))
        out[f'当月{tag}_总量'] = round(sum(r['数量'] for r in cur), 2)
        print(f'   [{name}][{tag}] 明细{len(recs)}行 / 有日期{out[f"{tag}_有日期行数"]} / '
              f'当月{len(cur)}条 / 合计{out[f"当月{tag}_总量"]}', flush=True)

        d = collections.OrderedDict()
        for r in cur:
            d[r['型号']] = d.get(r['型号'], 0.0) + r['数量']
        out[f'{tag}按型号'] = [[k, round(v, 2)] for k, v in sorted(d.items(), key=lambda kv: -kv[1])]

        # 入库：每个型号对应的主要来源（按量取前几个），供「来源」列
        if tag == '入库':
            ms = collections.OrderedDict()
            for r in cur:
                ms.setdefault(r['型号'], collections.OrderedDict())
                k = r.get('来源') or '未标注'
                ms[r['型号']][k] = ms[r['型号']].get(k, 0.0) + r['数量']
            out['入库型号来源'] = {m: [k for k, _ in sorted(v2.items(), key=lambda kv: -kv[1])]
                              for m, v2 in ms.items()}

        if tag == '入库':
            g = collections.OrderedDict()
            for r in cur:
                k = r.get('来源') or '未标注'
                e = g.setdefault(k, {'量': 0.0, '笔': 0})
                e['量'] += r['数量']; e['笔'] += 1
            out['入库来源分布'] = {k: {'量': round(v['量'], 2), '笔': v['笔']} for k, v in
                              sorted(g.items(), key=lambda kv: -kv[1]['量'])}
        else:
            # 出库：项目分布（用于项目维度汇总）
            gp = collections.OrderedDict()
            for r in cur:
                k = r.get('项目') or '未标注'
                e = gp.setdefault(k, {'量': 0.0, '笔': 0})
                e['量'] += r['数量']; e['笔'] += 1
            out['出库项目分布'] = [[k, round(v['量'], 2), v['笔']] for k, v in
                              sorted(gp.items(), key=lambda kv: -kv[1]['量'])]
            # 出库站点：优先「材料去向」；该列为空时（集成商B把站点写在项目列）用项目列兜底
            gs = collections.OrderedDict()
            for r in cur:
                k = (r.get('去向') or '').strip() or (r.get('项目') or '').strip()
                if not k:
                    continue
                e = gs.setdefault(k, {'量': 0.0, '笔': 0})
                e['量'] += r['数量']; e['笔'] += 1
            out['出库站点清单'] = [[k, round(v['量'], 2), v['笔']] for k, v in
                              sorted(gs.items(), key=lambda kv: -kv[1]['量'])]
            # 标注站点来源口径
            out['站点来源'] = '材料去向' if any((r.get('去向') or '').strip() for r in cur) else '项目列（材料去向为空）'

    # ---------------- 汇总表（累计口径 + 库存） ----------------
    if '汇总' in wb.sheetnames:
        ws = wb['汇总']
        it = ws.iter_rows(values_only=True)
        head = []
        for i, row in enumerate(it, 1):
            head.append(row)
            if i >= 4:
                break
        hr = None
        for row in head:
            if hcol(row, '型号') is not None and hcol(row, '库存') is not None:
                hr = row
                break
        if hr is not None:
            cm, cu = hcol(hr, '型号'), hcol(hr, '单位')
            cbeg = hcol(hr, '结存', '期初')
            cin = hcol(hr, '材料入库', '入库', exclude=('出库',))
            cout = hcol(hr, '材料出库', '出库')
            cstk = hcol(hr, '库存')
            rows = []
            for row in it:
                if cm is None or cm >= len(row):
                    continue
                m = s(row[cm])
                if not m:
                    continue
                g = lambda c: num(row[c]) if (c is not None and c < len(row)) else 0.0
                rows.append({'型号': m,
                             '单位': s(row[cu]) if (cu is not None and cu < len(row)) else '',
                             '期初': g(cbeg), '累计入库': g(cin),
                             '累计出库': g(cout), '库存': g(cstk)})
            out['汇总表'] = rows
            out['累计入库_总量'] = round(sum(r['累计入库'] for r in rows), 2)
            out['累计出库_总量'] = round(sum(r['累计出库'] for r in rows), 2)
            out['累计库存_总量'] = round(sum(r['库存'] for r in rows), 2)
            out['累计入库_型号数'] = len([r for r in rows if r['累计入库'] > 0])
            out['有库存_型号数'] = len([r for r in rows if r['库存'] > 0])
            out['库存明细'] = sorted([r for r in rows if r['库存'] > 0], key=lambda x: -x['库存'])

    out['盘存候选表'] = [x for x in wb.sheetnames if '盘' in x]
    wb.close()
    return out


def main():
    global YEAR, MONTH
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('--month', type=int, default=MONTH)
    ap.add_argument('--year', type=int, default=YEAR)
    ap.add_argument('--out', default='out/monthly_data.json')
    a = ap.parse_args()
    YEAR, MONTH = a.year, a.month
    print(f'### 目标区间：{YEAR}年{MONTH}月\n', flush=True)
    res = {'_meta': {'year': YEAR, 'month': MONTH}}
    for name, path in FILES.items():
        print(f'>> {name}', flush=True)
        res[name] = extract(name, path)
        r = res[name]
        print(f"   入库{r.get('当月入库_条数')}笔/{r.get('当月入库_总量')} | "
              f"出库{r.get('当月出库_条数')}笔/{r.get('当月出库_总量')} | "
              f"库存{r.get('累计库存_总量')}/{r.get('有库存_型号数')}种", flush=True)
    os.makedirs('out', exist_ok=True)
    with open(a.out, 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=1, default=str)
    print('\nOK ->', a.out)


if __name__ == '__main__':
    main()
