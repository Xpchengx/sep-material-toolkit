# -*- coding: utf-8 -*-
"""把抽取出的 CSV 组装成带公式的 .xlsx 工作簿。"""
import os, io, csv, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = r"C:\Users\qazws\WorkBuddy\2026-09-09-09-19-37\material-system\out"
DEST = os.path.join(OUT, "材料台账工作簿.xlsx")

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(bold=True, size=13, color="1F4E79")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def rd(name):
    p = os.path.join(OUT, name + ".csv")
    if not os.path.exists(p):
        return []
    with io.open(p, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    return rows[1:] if rows else []


def conv(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s == "":
        return ""
    if re.fullmatch(r"-?\d+\.0", s):
        return int(float(s))
    if re.fullmatch(r"-?\d+(\.\d+)?", s):
        return float(s)
    return s


wb = Workbook()
wb.remove(wb.active)


def sheet(title, header, rows, widths=None, freeze="A2"):
    ws = wb.create_sheet(title)
    ws.append(header)
    for c in range(1, len(header) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    for r in rows:
        ws.append([conv(x) for x in r])
    for i, h in enumerate(header, start=1):
        ws.column_dimensions[get_column_letter(i)].width = (
            widths[i - 1] if widths and i - 1 < len(widths) else max(10, min(28, len(str(h)) * 2 + 6)))
    ws.freeze_panes = freeze
    ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{max(2, ws.max_row)}"
    return ws


# ---------------- 1 说明 ----------------
ws = wb.create_sheet("说明")
notes = [
    ("材料台账工作簿", ""),
    ("", ""),
    ("一、这个工作簿解决什么", ""),
    ("把原来分散在「采购与设计量与到货量」「集成商框架材料汇总表」和三家集成商各自的库存表里的数据，", ""),
    ("合并成一套统一的台账；库存、发货核对、项目进度全部由公式自动算出，不再手工比对。", ""),
    ("", ""),
    ("二、各表说明", ""),
    ("集成商档案 / 物料主数据 / 项目档案 / 站点档案", "基础档案，先维护这里，其他表都是引用它"),
    ("入库明细 / 出库明细 / 退料明细", "三家集成商的流水，日常录入就在这里（按集成商筛选）"),
    ("站点设计量明细", "每个站点每样物资应发多少（来自集成商C设计明细，其余集成商待补）"),
    ("项目采购到货明细", "按项目×物料看 设计量 / 采购量 / 订单量 / 接收量"),
    ("集成商库存汇总 / 盘点记录", "原表快照，留档备查"),
    ("库存总览", "公式自动算：期初 + 入库 − 出库，按型号汇总三家 → 低于安全库存自动标记"),
    ("站点发货核对", "公式自动算：站点×物资 的设计量 vs 已发货量 → 自动判定 未发货/少发/齐套/超发"),
    ("项目进度", "公式自动算：按项目×物料 对比 设计量 / 采购量 / 到货量"),
    ("", ""),
    ("三、怎么用", ""),
    ("1. 新入库/出库，直接在「入库明细」「出库明细」末尾追加一行，A 列填集成商全称（集成商C/集成商A/集成商B）。", ""),
    ("2. 库存、发货核对、项目进度三张表会自动更新，不需要手工维护。", ""),
    ("3. 要给某物资设预警线，在「库存总览」的 G 列「安全库存」填数字即可。", ""),
    ("", ""),
    ("四、已知数据缺口（需要补录）", ""),
    ("1. 「采购与设计量与到货量」里，20 个项目中只有「八期一阶段」「26年更新改造一阶段」填了设计量，其余为空。", ""),
    ("2. 「站点设计量明细」目前只有集成商C的数据；集成商A、集成商B的站点级设计量没有来源表。", ""),
    ("3. 三家集成商的物资型号命名不统一，跨集成商合并统计时需要先做映射（见「物料主数据」）。", ""),
    ("4. 期初库存：只有集成商B有「2025年结存」，集成商C、集成商A按 0 处理；若实际有期初需在「集成商库存汇总」补。", ""),
    ("5. 库存口径已对账：按「期初+入库−出库」算出的库存与三家自己汇总表的库存，233 项中 230 项完全一致（98.7%）。", ""),
]
for i, (a, b) in enumerate(notes, start=1):
    ws.cell(row=i, column=1, value=a)
    ws.cell(row=i, column=2, value=b)
ws.column_dimensions["A"].width = 62
ws.column_dimensions["B"].width = 70
ws.cell(row=1, column=1).font = TITLE_FONT

# ---------------- 2 基础档案 ----------------
# 注意：单位全称与所在地市属内部信息，此处用占位符，实际使用时替换为真实值
sheet("集成商档案",
      ["集成商", "单位全称", "仓库/项目部", "备注"],
      [["集成商C", "＜集成商全称＞", "＜地市＞", ""],
       ["集成商A", "＜集成商全称＞", "＜地市＞", ""],
       ["集成商B", "＜集成商全称＞", "＜地市＞", ""]],
      [10, 34, 12, 16])

sheet("物料主数据", ["标准编码", "标准名称", "标准型号", "单位", "出现于集成商", "采购表物料编码"], rd("物料主数据"))

sheet("项目档案", ["项目名称", "涉及集成商", "备注"], rd("项目档案"))
sheet("站点档案", ["站点名称", "项目", "集成商", "任务编码"], rd("站点档案"))

# ---------------- 3 流水 ----------------
sheet("入库明细", ["集成商", "日期", "月份", "型号", "数量", "单位", "材料来源", "项目", "备注"], rd("入库明细"))
sheet("出库明细", ["集成商", "月份", "日期", "型号", "数量", "项目", "材料去向", "领料人", "备注", "应发未发", "退废料"], rd("出库明细"))
sheet("退料明细", ["集成商", "类型", "日期", "型号", "数量", "单位", "项目", "站点", "退料人员", "备注"], rd("退料明细"))
sheet("站点设计量明细", ["集成商", "项目", "站点", "型号", "设计数量", "备注"], rd("站点设计量明细"))
sheet("项目采购到货明细",
      ["项目", "物料编码", "物料名称", "单位", "设计量", "采购1", "采购2", "采购汇总", "设计与提采差额", "供应商", "订单号", "订单数量", "接收数量"],
      rd("项目采购到货明细"))
sheet("集成商库存汇总", ["集成商", "名称", "型号", "单位", "期初库存", "入库", "出库", "库存"], rd("集成商库存汇总"))
sheet("盘点记录", ["集成商", "盘点表", "物料名称", "型号", "单位", "账面数量", "实盘数量", "差异", "备注"], rd("盘点记录"))

# ---------------- 4 库存总览（公式） ----------------
stock_rows = rd("集成商库存汇总")
agg = {}
for r in stock_rows:
    tag, name, model, unit = r[0], r[1], r[2], r[3]
    if not model:
        continue
    d = agg.setdefault(model, {"unit": unit, "names": []})
    if unit and not d["unit"]:
        d["unit"] = unit
    if name:
        d["names"].append(name)
# 并集：入库/出库明细里出现过的型号也纳入
for r in rd("入库明细"):
    model, unit = r[3], r[5]
    if model:
        d = agg.setdefault(model, {"unit": unit, "names": []})
        if unit and not d["unit"]:
            d["unit"] = unit
for r in rd("出库明细"):
    model = r[3]
    if model:
        agg.setdefault(model, {"unit": "", "names": []})
models = sorted(agg.items())
ws = wb.create_sheet("库存总览")
hdr = ["型号", "单位", "集成商C库存", "集成商A库存", "集成商B库存", "库存合计", "安全库存", "库存状态"]
ws.append(hdr)
for c in range(1, len(hdr) + 1):
    cell = ws.cell(row=1, column=c)
    cell.fill = HDR_FILL
    cell.font = HDR_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for i, (model, d) in enumerate(models, start=2):
    ws.cell(row=i, column=1, value=model)
    ws.cell(row=i, column=2, value=d["unit"])
    for j, tag in enumerate(["集成商C", "集成商A", "集成商B"], start=3):
        f = (f'=SUMIFS(集成商库存汇总!$E:$E,集成商库存汇总!$A:$A,"{tag}",集成商库存汇总!$C:$C,$A{i})'
             f'+SUMIFS(入库明细!$E:$E,入库明细!$A:$A,"{tag}",入库明细!$D:$D,$A{i})'
             f'-SUMIFS(出库明细!$E:$E,出库明细!$A:$A,"{tag}",出库明细!$D:$D,$A{i})')
        ws.cell(row=i, column=j, value=f)
    ws.cell(row=i, column=6, value=f"=SUM(C{i}:E{i})")
    ws.cell(row=i, column=7, value="")
    ws.cell(row=i, column=8, value=f'=IF($F{i}<=0,"无库存",IF($G{i}="","未设预警线",IF($F{i}<=$G{i},"库存不足","正常")))')
for i, wd in enumerate([30, 8, 11, 11, 12, 11, 10, 12], start=1):
    ws.column_dimensions[get_column_letter(i)].width = wd
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:H{max(2, ws.max_row)}"

# ---------------- 5 站点发货核对（公式） ----------------
des = rd("站点设计量明细")
key = {}
for r in des:
    tag, proj, site, model, qty = r[0], r[1], r[2], r[3], r[4]
    if not site or not model:
        continue
    k = (site, model)
    d = key.setdefault(k, {"proj": proj, "tag": tag, "q": 0.0})
    try:
        d["q"] += float(qty) if qty not in ("", None) else 0
    except Exception:
        pass
    if proj:
        d["proj"] = proj
check = sorted(key.items())
ws = wb.create_sheet("站点发货核对")
hdr = ["站点", "项目", "集成商", "型号", "设计数量", "已发货数量", "差异", "核对结论"]
ws.append(hdr)
for c in range(1, len(hdr) + 1):
    cell = ws.cell(row=1, column=c)
    cell.fill = HDR_FILL
    cell.font = HDR_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for i, ((site, model), d) in enumerate(check, start=2):
    ws.cell(row=i, column=1, value=site)
    ws.cell(row=i, column=2, value=d["proj"])
    ws.cell(row=i, column=3, value=d["tag"])
    ws.cell(row=i, column=4, value=model)
    ws.cell(row=i, column=5, value=round(d["q"], 2))
    ws.cell(row=i, column=6, value=f'=SUMIFS(出库明细!$E:$E,出库明细!$G:$G,$A{i},出库明细!$D:$D,$D{i})')
    ws.cell(row=i, column=7, value=f"=$E{i}-$F{i}")
    ws.cell(row=i, column=8, value=f'=IF($F{i}=0,"未发货",IF($G{i}>0,"少发",IF($G{i}=0,"齐套","超发")))')
for i, wd in enumerate([26, 18, 9, 28, 10, 12, 9, 10], start=1):
    ws.column_dimensions[get_column_letter(i)].width = wd
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:H{max(2, ws.max_row)}"

# ---------------- 6 项目进度（公式） ----------------
pur = rd("项目采购到货明细")
pk = {}
for r in pur:
    proj, name, unit = r[0], r[2], r[3]
    if not proj or not name:
        continue
    k = (proj, name)
    d = pk.setdefault(k, {"unit": unit, "dz": None, "cg": 0.0, "od": 0.0, "rc": 0.0})
    if unit and not d["unit"]:
        d["unit"] = unit
    def f(x):
        try:
            return float(x)
        except Exception:
            return None
    dz = f(r[4])
    if dz is not None:
        d["dz"] = dz
    cg = f(r[7])
    if cg is not None:
        d["cg"] += cg
    od = f(r[11])
    if od is not None:
        d["od"] += od
    rc = f(r[12])
    if rc is not None:
        d["rc"] += rc
prows = sorted(pk.items())
ws = wb.create_sheet("项目进度")
hdr = ["项目", "物料名称", "单位", "设计量", "采购量", "订单量", "接收量(到货)", "到货率", "采购-到货差额"]
ws.append(hdr)
for c in range(1, len(hdr) + 1):
    cell = ws.cell(row=1, column=c)
    cell.fill = HDR_FILL
    cell.font = HDR_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for i, ((proj, name), d) in enumerate(prows, start=2):
    ws.cell(row=i, column=1, value=proj)
    ws.cell(row=i, column=2, value=name)
    ws.cell(row=i, column=3, value=d["unit"])
    ws.cell(row=i, column=4, value=d["dz"] if d["dz"] is not None else "")
    ws.cell(row=i, column=5, value=round(d["cg"], 2) if d["cg"] else "")
    ws.cell(row=i, column=6, value=round(d["od"], 2) if d["od"] else "")
    ws.cell(row=i, column=7, value=round(d["rc"], 2) if d["rc"] else "")
    ws.cell(row=i, column=8, value=f'=IF($F{i}="","",IF($F{i}=0,"",$G{i}/$F{i}))')
    ws.cell(row=i, column=9, value=f'=IF(OR($F{i}="",$G{i}=""),"",$F{i}-$G{i})')
for i in range(2, ws.max_row + 1):
    ws.cell(row=i, column=8).number_format = "0.0%"
for i, wd in enumerate([20, 46, 8, 11, 11, 11, 13, 10, 15], start=1):
    ws.column_dimensions[get_column_letter(i)].width = wd
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:I{max(2, ws.max_row)}"

wb.save(DEST)
print("saved:", DEST)
print("sheets:", wb.sheetnames)
