# -*- coding: utf-8 -*-
"""按系统真实表的表头，生成物资调拨导入的三份模板。

① 到货/支出表  —— 从《7.3采购与设计量到货情况》取表头（精简为单行）
② 站点设计量表 —— 从《站点设计量_20261126》取三行表头（分组/编码/名称）
③ 站点现有量表 —— 从《站点现有量明细列表》取 35 列表头
"""
import os
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

FONT = '微软雅黑'
F_T = Font(name=FONT, size=15, bold=True, color='1F4E79')
F_H = Font(name=FONT, size=10, bold=True, color='FFFFFF')
F_B = Font(name=FONT, size=10, bold=True)
F_N = Font(name=FONT, size=10)
F_S = Font(name=FONT, size=9, color='595959')
F_R = Font(name=FONT, size=10, bold=True, color='C00000')
FILL_H = PatternFill('solid', fgColor='1F4E79')
FILL_H2 = PatternFill('solid', fgColor='2E75B6')
FILL_MUST = PatternFill('solid', fgColor='C00000')
FILL_Y = PatternFill('solid', fgColor='FFF2CC')
FILL_G = PatternFill('solid', fgColor='E2EFDA')
TH = Side(style='thin', color='BFBFBF')
BD = Border(left=TH, right=TH, top=TH, bottom=TH)
AL_C = Alignment(horizontal='center', vertical='center', wrap_text=True)
AL_L = Alignment(horizontal='left', vertical='center', wrap_text=True)

OUT = r'C:\Users\qazws\WorkBuddy\2026-09-09-09-19-37\output\物资调拨导入模板'
os.makedirs(OUT, exist_ok=True)


def add_howto(wb, title, items, need_cols, opt_cols, rules):
    ws = wb.create_sheet('说明')
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 92

    r = 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c = ws.cell(row=r, column=2, value=title)
    c.font, c.alignment = F_T, AL_L
    ws.row_dimensions[r].height = 34

    r += 2
    for line in items:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2, value=line)
        c.font, c.alignment = F_N, AL_L
        ws.row_dimensions[r].height = 22
        r += 1

    def block(name, rows, header, fill):
        nonlocal r
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2, value=name)
        c.font, c.alignment = F_B, AL_L
        r += 1
        for i, h in enumerate(header):
            cell = ws.cell(row=r, column=2 + i, value=h)
            cell.font, cell.fill, cell.border, cell.alignment = F_H, fill, BD, AL_C
        r += 1
        for row in rows:
            for i, v in enumerate(row):
                cell = ws.cell(row=r, column=2 + i, value=v)
                cell.font, cell.border = F_N, BD
                cell.alignment = AL_L if i == 1 else AL_C
            ws.row_dimensions[r].height = 26
            r += 1

    block('一、必须有的列（表头文字要一字不差）', need_cols, ['列名', '说明'], FILL_MUST)
    if opt_cols:
        block('二、可以有的列（有就更好，没有也能算）', opt_cols, ['列名', '说明'], FILL_H2)

    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c = ws.cell(row=r, column=2, value='三、填写规则')
    c.font, c.alignment = F_B, AL_L
    r += 1
    for i, line in enumerate(rules, 1):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2, value=f'{i}. {line}')
        c.font, c.alignment = F_N, AL_L
        ws.row_dimensions[r].height = 32
        r += 1
    return ws


# ============================================================ ① 到货 / 支出表
P1 = r'C:\Users\qazws\AppData\Local\Temp\7zOCCE2551E\7.3采购与设计量到货情况.xlsx'
wb1 = load_workbook(P1, read_only=True, data_only=True)
raw1 = list(wb1.worksheets[0].iter_rows(min_row=1, max_row=1, values_only=True))[0]
wb1.close()

# 系统原表是多层表头；模板取单层，列名与系统一致
T1 = ['序号', '物料编码', '物料名称', '单位', '设计量', '汇总量', '发货量',
      '项目支出量', '支出与设计差异']
MUST1 = [
    ('物料编码', '8 位数字，和系统物资编码一致；文本格式，不要变成科学计数法'),
    ('物料名称', '从系统复制，不要手打，避免串码'),
    ('单位', '个 / 套 / 条 / 面 / 米 / 皮长公里 等，照系统填'),
    ('设计量', '该项目这个物料的应配总量'),
    ('项目支出量', '已经支出（已领用）的数量'),
    ('支出与设计差异', '＝ 项目支出量 − 设计量；负数就是缺口。也可以只给设计量和支出量，让表算'),
]
OPT1 = [
    ('序号', '仅为好看，脚本不用'),
    ('汇总量', '系统里「汇总量」列，原样带过来'),
    ('发货量', '集成商已发货数量，供参考'),
]

wb = Workbook()
ws = wb.active
ws.title = '数据'
ws.sheet_view.showGridLines = False
for i, h in enumerate(T1, 1):
    c = ws.cell(row=1, column=i, value=h)
    c.font, c.fill, c.border, c.alignment = F_H, FILL_H, BD, AL_C
    ws.column_dimensions[get_column_letter(i)].width = 16 if i not in (2, 3) else 34
ws.row_dimensions[1].height = 26
ws.freeze_panes = 'A2'
# 预留 200 行格式
for r in range(2, 202):
    ws.cell(row=r, column=2).number_format = '@'     # 物料编码按文本
    for i in range(1, len(T1) + 1):
        ws.cell(row=r, column=i).border = BD

add_howto(wb, '① 到货 / 支出表 · 填写说明',
          ['用途：告诉脚本「这个项目缺什么」——缺口 = 支出与设计差异（负数）。',
           '这是项目级的总账，一行一个物料，不区分站点。',
           '可以照系统导出的《XX采购与设计量到货情况》直接整理，列名要与系统一致。'],
          MUST1, OPT1,
          ['一个物料只写一行，不要重复；重复会被当成两笔相加。',
           '物料编码必须是 8 位数字，前后不要有空格；Excel 里若显示成 1.04E+07，把该列设为「文本」再填。',
           '数量列必须是纯数字，不要写「50个」「约50」这类；单位单独放「单位」列。',
           '不要合并单元格、不要多行表头、不要加合计行——脚本只认第 1 行表头。',
           '缺口为 0 或正数（多余）的物料可以不写，脚本只看负数。'])

wb.save(os.path.join(OUT, '①到货与支出表_模板.xlsx'))
print('① 已生成')

# ============================================================ ② 站点设计量表
P2 = r'C:\Users\qazws\Downloads\站点设计量_20261126.xlsx'
wb2 = load_workbook(P2, read_only=True, data_only=True)
rows2 = [list(r) for r in wb2.worksheets[0].iter_rows(min_row=1, max_row=3, values_only=True)]
wb2.close()
g_row, c_row, n_row = rows2[0], rows2[1], rows2[2]
ncol = len(n_row)

wb = Workbook()
ws = wb.active
ws.title = '数据'
ws.sheet_view.showGridLines = False
for j in range(ncol):
    col = j + 1
    for r, val, fill in ((1, g_row[j], FILL_H), (2, c_row[j], FILL_H), (3, n_row[j], FILL_H2)):
        c = ws.cell(row=r, column=col, value=val)
        c.font = F_H if r < 3 else F_B
        c.fill, c.border, c.alignment = fill, BD, AL_C
    if col <= 4:
        ws.column_dimensions[get_column_letter(col)].width = 16
    else:
        ws.column_dimensions[get_column_letter(col)].width = 11
for r in (1, 2):
    ws.row_dimensions[r].height = 22
ws.row_dimensions[3].height = 120
ws.freeze_panes = 'E4'
for r in range(4, 304):
    for col in range(1, 5):
        ws.cell(row=r, column=col).border = BD
        ws.cell(row=r, column=col).alignment = AL_C

add_howto(wb, '② 站点设计量表 · 填写说明',
          ['用途：告诉脚本「每个站点、每种物料该配多少」——这是判断谁缺、谁有余量的基准。',
           '这是宽表：一行一个站点，一列一种物料。表头固定 3 行（分组 / 物料编码 / 物料名称），'
           '从第 4 行起填数据。**这三行不要删、不要改名、不要改顺序。**',
           '建议直接从在线《材料汇总表》整表导出，再删掉不用的行；不要重新排列表头。'],
          [('第 1 列 序号', '可留空'),
           ('第 2 列 工期', '项目阶段名（如「七期三阶段」）。脚本靠它分项目，必须填'),
           ('第 3 列 单位', '实际是集成商/施工单位名，原样保留即可'),
           ('第 4 列 站点名称', '要与《站点现有量明细》里的「任务名称」能对上'),
           ('第 5 列起', '每个物料一列，第 2 行是 8 位物料编码，第 3 行是名称')],
          [],
          ['第 1~3 行是表头，绝对不能动；数据从第 4 行开始。',
           '单元格只填数字。设计量若有附注（如「少268」），写在数字后面的括号里可以，脚本只取前面数字。',
           '同一站点不要拆成两行；同一工期下站点名不要重复。',
           '工期名要保持稳定——脚本按工期名识别项目，改名会导致该项目认不出来。',
           '不要加「合计」行，也不要在表头上面插入标题行。'])

wb.save(os.path.join(OUT, '②站点设计量表_模板.xlsx'))
print('② 已生成，列数', ncol)

# ============================================================ ③ 站点现有量表
P3 = r'C:\Users\qazws\Downloads\站点现有量明细列表 (40).xlsx'
wb3 = load_workbook(P3, read_only=True, data_only=True)
h3 = [c.value for c in next(wb3.worksheets[0].iter_rows(min_row=1, max_row=1))]
wb3.close()

MUST3 = [
    ('物料/组合物料编码', '8 位数字。注意不是「物料扩展编码」，也不是「细化产品编码」'),
    ('物料/组合物料描述', '物料名称，供核对用'),
    ('物料计量单位', '个 / 套 / 条 / 面 / 米 / 皮长公里'),
    ('可用量', '关键列。已经是扣掉占用量之后的数，直接用它'),
    ('任务名称', '站点名。要和设计量表里的「站点名称」对得上'),
    ('项目编码', '如 C25303523AA100C；用于区分不同项目'),
]
OPT3 = [
    ('数量', '账面数量；脚本用「可用量」不用它'),
    ('占用量', '已被订单占用的数量'),
    ('项目名称', '辅助识别项目'),
    ('任务编码', '下单时要用，建议保留'),
    ('库存组织名称、OU名称、子库编码', '可留空，不影响计算'),
    ('供应商名称、订单编号、采购接收日期等', '可留空'),
]

wb = Workbook()
ws = wb.active
ws.title = '数据'
ws.sheet_view.showGridLines = False
for i, h in enumerate(h3, 1):
    c = ws.cell(row=1, column=i, value=h)
    c.font, c.fill, c.border, c.alignment = F_H, FILL_H, BD, AL_C
    ws.column_dimensions[get_column_letter(i)].width = max(12, min(26, len(str(h)) * 2 + 4))
ws.row_dimensions[1].height = 30
ws.freeze_panes = 'A2'
for r in range(2, 302):
    for i in range(1, len(h3) + 1):
        ws.cell(row=r, column=i).border = BD
    ws.cell(row=r, column=9).number_format = '@'   # 物料编码

add_howto(wb, '③ 站点现有量表 · 填写说明',
          ['用途：告诉脚本「每个站点现场有多少可用物资」——这是判断能不能往别处调的依据。',
           '一个项目导一份，文件名带上项目名（如「8.1站点现有量」）；多个项目可以一次都拖进去，'
           '脚本按站点名自动识别各属于哪个工期。',
           '最省事的做法：从系统直接导出《站点现有量明细列表》，一列都别删，原样存 xlsx。'],
          MUST3, OPT3,
          ['一个站点 + 一个物料写一行；同一组合有多行（不同批次）没关系，脚本会自动相加。',
           '「可用量」已扣占用量，别再自己减一遍。',
           '物料编码必须是 8 位数字；用文本格式，避免前导 0 丢失或变成科学计数法。',
           '站点名不要手工改；如果和设计量表对不上，脚本会把它列进「待确认」，不会乱算。',
           '不要合并单元格、不要加合计行、不要留多行表头。'])

wb.save(os.path.join(OUT, '③站点现有量表_模板.xlsx'))
print('③ 已生成，列数', len(h3))

print('\n输出目录:', OUT)
for f in sorted(os.listdir(OUT)):
    print('  ', f, f'{os.path.getsize(os.path.join(OUT, f)):,} 字节')
