# -*- coding: utf-8 -*-
"""
月报生成器 build_monthly_report.py
====================================
读 out/monthly_aug_data.json  ->  生成《材料管理周转仓管理月报（N月）》
一家一节，三家同文档；版式对齐 7 月原月报（A4 / 微软雅黑 / 蓝色标题 / Table Grid）。

节结构（与 7 月原月报一致）：
  一、当月入库情况   （型号汇总表 + 主要设备类说明）
  二、累计入库情况
  三、当月出库情况   （型号汇总表 + 主要设备说明 + 站点明细表）
  四、当月盘存情况   （高价值设备 / 线缆类 / 辅料及配件 三张表）
  五、每月物资报废情况
  六、入库清单检查痕迹   [待贴图]
  七、盘存清单           [待贴图]
  八、制度上墙           [待贴图]
"""
import sys, os, json, datetime, argparse
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = '微软雅黑'
C_H1 = RGBColor(0x37, 0x60, 0x92)
C_H2 = RGBColor(0x4F, 0x81, 0xBD)
C_NOTE = RGBColor(0x7F, 0x7F, 0x7F)

# 盘存分类关键字
K_CABLE = ('馈线', '电缆', '光缆', '光电', '电源线', '混合缆', '地线', '钢管', '管', '跳纤', '尾纤', '绳')
K_DEVICE = ('RRU', 'BBU', 'AAU', '基带板', '主控板', '主控传输', '电源模块', 'INCR', '皮站',
            '直放站', 'RHUB', 'RHBU', 'pRRU', '近端', '远端', '轿厢', '载波功放', '挡风板',
            '交转直', 'GPS', '天馈', '合分波')

REPORT_TITLE = '材料管理周转仓管理月报'
NAME_OVERRIDE = {}      # 如 {'集成商A': '集成商A'}


def set_font(run, size=None, bold=None, color=None, name=FONT):
    run.font.name = name
    run.font.size = Pt(size) if size else run.font.size
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rpr.append(rf)
    rf.set(qn('w:ascii'), name)
    rf.set(qn('w:hAnsi'), name)
    rf.set(qn('w:eastAsia'), name)


def para(doc, text='', size=10.5, bold=False, color=None, align=None,
         space_after=4, indent_first=False, style=None):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    if indent_first:
        p.paragraph_format.first_line_indent = Pt(21)
    if text:
        set_font(p.add_run(text), size=size, bold=bold, color=color)
    return p


def h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(8)
    set_font(p.add_run(text), size=14, bold=True, color=C_H1)
    return p


def h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(5)
    set_font(p.add_run(text), size=13, bold=True, color=C_H2)
    return p


def h3(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    set_font(p.add_run(text), size=11, bold=True, color=C_H2)
    return p


def make_table(doc, header, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, htxt in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = ''
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        set_font(p.add_run(str(htxt)), size=9, bold=True)
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = ''
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i != 1 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            set_font(p.add_run('' if v is None else str(v)), size=9)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    return t


def fmt(v):
    """数值格式化：整数不带小数点、不带千分位（与原月报一致）。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return str(int(round(f))) if abs(f - round(f)) < 1e-6 else f'{f:.2f}'


def unit_of(v, model):
    """从库存明细里取单位。"""
    for r in v.get('库存明细', []):
        if r['型号'] == model and r.get('单位'):
            return r['单位']
    for r in v.get('汇总表', []):
        if r['型号'] == model and r.get('单位'):
            return r['单位']
    return ''


def split_stock(v):
    """把库存按 高价值设备 / 线缆 / 辅料配件 分三类。"""
    dev, cab, aux = [], [], []
    for r in v.get('库存明细', []):
        m = r['型号']
        if any(k in m for k in K_DEVICE):
            dev.append(r)
        elif any(k in m for k in K_CABLE):
            cab.append(r)
        else:
            aux.append(r)
    return dev, cab, aux


def stock_word(tag):
    return {'集成商A': '上海集成商A', '集成商B': '集成商B', '集成商C': '集成商C'}.get(tag, tag)


def build_supplier(doc, tag, v, month, year):
    """写一个集成商的完整章节。"""
    sname = NAME_OVERRIDE.get(tag, tag)
    pre = f'{stock_word(tag)}周转仓'
    m = month

    in_num = v.get('当月入库_条数', 0)
    in_mod = v.get('当月入库_型号数', 0)
    in_tot = v.get('当月入库_总量', 0)
    out_num = v.get('当月出库_条数', 0)
    out_mod = v.get('当月出库_型号数', 0)
    out_tot = v.get('当月出库_总量', 0)

    # ---------------------------------------------------- 一、当月入库
    h2(doc, '一、当月入库情况')
    src = v.get('入库来源分布', {})
    if src:
        top = list(src.items())[:3]
        srtxt = '、'.join(f'{k}（{fmt(b["量"])}，{b["笔"]}笔）' for k, b in top)
        more = '等' if len(src) > 3 else ''
        head = (f'{year}年{m}月，{pre}当月入库共计{in_num}笔，涉及{in_mod}种型号物资，'
                f'物资总量{fmt(in_tot)}。入库物资主要来源于{srtxt}{more}。')
    else:
        head = (f'{year}年{m}月，{pre}当月入库共计{in_num}笔，涉及{in_mod}种型号物资，'
                f'物资总量{fmt(in_tot)}。')
    para(doc, head, indent_first=True)
    para(doc, '按型号汇总当月入库明细：')

    inrows = []
    srcof = v.get('入库型号来源', {})
    for i, (model, qty) in enumerate(v.get('入库按型号', []), 1):
        srcs = srcof.get(model, [])
        stray = '、'.join(srcs[:2]) + ('等' if len(srcs) > 2 else '')
        inrows.append([i, model, fmt(qty), unit_of(v, model), stray])
    make_table(doc, ['序号', '型号', '数量', '单位', '来源'], inrows,
               widths=[1.2, 7.6, 2.0, 1.4, 3.8])

    dev_in = [(mo, q) for mo, q in v.get('入库按型号', []) if any(k in mo for k in K_DEVICE)]
    if dev_in:
        para(doc, '其中主要设备类物资入库：' +
             '、'.join(f'{mo} {fmt(q)}{unit_of(v, mo)}' for mo, q in dev_in[:14]) + '等。')

    # ---------------------------------------------------- 二、累计入库
    h2(doc, '二、累计入库情况')
    para(doc, f'截至{year}年{m}月底，{pre}累计入库物资总量{fmt(v.get("累计入库_总量", 0))}，'
              f'涵盖{v.get("累计入库_型号数", 0)}种型号；累计出库{fmt(v.get("累计出库_总量", 0))}，'
              f'当前库存{fmt(v.get("累计库存_总量", 0))}，有库存物资{v.get("有库存_型号数", 0)}种。'
              f'物资覆盖5G八期各阶段、网优更新改造、专网工程等多个项目，'
              f'种类包括室内天线、功分器、耦合器、馈线、接头、转接头、光缆分纤箱、电力电缆、'
              f'尾纤、大功率皮站设备（RRU/BBU/AAU等）、光模块、光缆、辅料等大类。',
         indent_first=True)

    # ---------------------------------------------------- 三、当月出库
    h2(doc, '三、当月出库情况')
    proj = v.get('出库项目分布', [])
    if proj:
        ptxt = '、'.join(f'{k}{fmt(q)}（{n}笔）' for k, q, n in proj[:6])
        head = (f'{year}年{m}月，{pre}当月出库共计{out_num}笔，涉及{out_mod}种型号物资，'
                f'物资总量{fmt(out_tot)}。出库项目分布：{ptxt}。')
    else:
        head = (f'{year}年{m}月，{pre}当月出库共计{out_num}笔，涉及{out_mod}种型号物资，'
                f'物资总量{fmt(out_tot)}。')
    para(doc, head, indent_first=True)
    para(doc, '按型号汇总当月出库明细：')

    outrows = []
    for i, (model, qty) in enumerate(v.get('出库按型号', []), 1):
        outrows.append([i, model, fmt(qty), unit_of(v, model)])
    make_table(doc, ['序号', '型号', '数量', '单位'], outrows,
               widths=[1.2, 9.5, 2.5, 1.8])

    dev_out = [(mo, q) for mo, q in v.get('出库按型号', []) if any(k in mo for k in K_DEVICE)]
    if dev_out:
        para(doc, '主要设备出库：' +
             '、'.join(f'{mo} {fmt(q)}{unit_of(v, mo)}' for mo, q in dev_out[:14]) + '等。')

    sit = v.get('出库站点清单', [])
    if sit:
        para(doc, '出库站点明细：')
        srows = [[i, k, fmt(q), n] for i, (k, q, n) in enumerate(sit, 1)]
        make_table(doc, ['序号', '站点名称', '出库数量', '笔数'], srows,
                   widths=[1.2, 9.5, 2.5, 1.8])
        if v.get('站点来源', '').startswith('项目列'):
            para(doc, '注：该集成商出库台账「材料去向」列留空，站点名称取「项目」列。',
                 size=9, color=C_NOTE)
    else:
        para(doc, '本月出库台账未登记具体站点信息。', size=9, color=C_NOTE)

    # ---------------------------------------------------- 四、当月盘存
    h2(doc, '四、当月盘存情况')
    para(doc, f'截至{year}年{m}月底，{pre}盘存物资总量{fmt(v.get("累计库存_总量", 0))}，'
              f'有库存物资{v.get("有库存_型号数", 0)}种。以下为重点物资库存明细：')
    dev, cab, aux = split_stock(v)
    for title, group in (('（一）高价值设备库存', dev),
                         ('（二）线缆类库存', cab),
                         ('（三）辅料及配件库存', aux)):
        if not group:
            continue
        h3(doc, title)
        rows = [[i, r['型号'], fmt(r['库存']), r.get('单位', '') or unit_of(v, r['型号'])]
                for i, r in enumerate(group, 1)]
        make_table(doc, ['序号', '型号', '库存数量', '单位'], rows,
                   widths=[1.2, 9.5, 2.5, 1.8])

    # ---------------------------------------------------- 五、报废
    h2(doc, '五、每月物资报废情况')
    para(doc, f'{year}年{m}月，{pre}无物资报废记录（报废数据以集成商退废料台账为准）。')

    # ---------------------------------------------------- 六七八 待贴图
    for t in ('六、入库清单检查痕迹', '七、盘存清单', '八、制度上墙'):
        h2(doc, t)
        para(doc, '【待贴图】本节为现场佐证材料，请粘贴对应的检查记录、盘点签字单、制度上墙照片。',
             size=9, color=C_NOTE)


def build(data, month, year, out_path):
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = FONT
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin = sec.bottom_margin = Cm(2.0)

    h1(doc, f'{REPORT_TITLE}（{month}月）')
    para(doc, f'统计区间：{year}年{month}月1日 — {month}月{_last_day(year, month)}日　|　'
              f'编制单位：＜编制单位＞　|　集成商：集成商A、集成商B、集成商C',
         size=9, color=C_NOTE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    for tag in ('集成商A', '集成商B', '集成商C'):
        v = data.get(tag)
        if not v:
            continue
        h1(doc, f'【{tag}】')
        build_supplier(doc, tag, v, month, year)
        doc.add_page_break()

    doc.save(out_path)
    print('OK ->', out_path)


def _last_day(y, m):
    if m == 12:
        return 31
    return (datetime.date(y, m + 1, 1) - datetime.timedelta(days=1)).day


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--month', type=int, default=8)
    ap.add_argument('--year', type=int, default=2026)
    ap.add_argument('--data', default=None)
    ap.add_argument('--out', default=None)
    a = ap.parse_args()
    dfile = a.data or f'out/monthly_data.json'
    with open(dfile, encoding='utf-8') as f:
        data = json.load(f)
    data.pop('_meta', None)
    o = a.out or f'out/材料管理周转仓管理月报（{a.month}月）.docx'
    build(data, a.month, a.year, o)
