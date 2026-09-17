# -*- coding: utf-8 -*-
"""最终工作簿：把标准型号接进明细表，库存与发货核对按标准型号汇总。"""
import os, io, csv, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = r"C:\Users\qazws\WorkBuddy\2026-09-09-09-19-37\material-system\out"
DEST = os.path.join(OUT, "材料台账工作簿.xlsx")

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR2_FILL = PatternFill("solid", fgColor="2E75B6")
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


# 型号 -> 标准编码/标准型号
std_map = {}
for r in rd("型号映射"):
    std_map[r[1].strip()] = (r[3], r[4])

wb = Workbook()
wb.remove(wb.active)


def mk(title, header, rows, widths=None, note=None):
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
            widths[i - 1] if widths and i - 1 < len(widths)
            else max(10, min(30, len(str(h)) * 2 + 6)))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{max(2, ws.max_row)}"
    if note:
        ws.cell(row=1, column=len(header) + 1, value=note).font = Font(color="C00000", size=10)
    return ws


# ---------------- 说明 ----------------
ws = wb.create_sheet("说明")
notes = [
    ("材料台账工作簿", ""),
    ("", ""),
    ("一、这套表解决什么", ""),
    ("把原来分散在「采购与设计量与到货量」「集成商框架材料汇总表」和三家集成商各自库存表里的数据，", ""),
    ("合并成一套统一台账；库存、发货核对、项目进度全部由公式自动算出，不再手工比对。", ""),
    ("", ""),
    ("二、型号归一（本轮新增，最关键）", ""),
    ("三家集成商对同一物资的写法完全不同（如「1.25G光模块」「1.25G模块」「光模块-1.25G」是同一物）。", ""),
    ("已建立统一物料主数据：303 种原始写法 → 253 项标准物料，并给每项分配了标准编码（M0001 起）。", ""),
    ("做法分三级：", ""),
    ("  A-原名 / A-机械归一：去空格、全半角统一、分隔符统一后自动命中（197 项）", ""),
    ("  B-规则：同义写法、错别字、交流直流缺省等人工规则（106 项）", ""),
    ("  规格相近候选：相似但规格不同，默认保持独立，供人工复核（123 组，见「规格相近候选」表）", ""),
    ("各明细表末尾都新增了「标准编码」「标准型号」两列，跨集成商汇总靠它。", ""),
    ("", ""),
    ("三、各表说明", ""),
    ("集成商档案 / 物料主数据 / 型号映射 / 项目档案 / 站点档案", "基础档案"),
    ("入库明细 / 出库明细 / 退料明细", "三家流水，日常录入选这里（按集成商筛选）"),
    ("站点设计量明细", "站点×物资应发数量（目前只有集成商C有来源）"),
    ("项目采购到货明细", "按项目×物料看 设计量/采购量/订单量/接收量"),
    ("集成商库存汇总 / 盘点记录", "原表快照，留档"),
    ("库存总览", "公式自动算：期初+入库−出库，按【标准型号】汇总三家"),
    ("站点发货核对", "公式自动算：站点×标准型号 的设计量 vs 已发货量，自动判定 未发货/少发/齐套/超发"),
    ("项目进度", "公式自动算：按项目×物料 对比 设计量/采购量/到货量"),
    ("规格相近候选", "相似但规格不同的型号对，默认不合并；确认要合并的告诉我，我加规则"),
    ("", ""),
    ("四、怎么用", ""),
    ("1. 新入库/出库直接在「入库明细」「出库明细」末尾追加一行，A 列填集成商（集成商C/集成商A/集成商B）。", ""),
    ("2. 型号填三家各自的叫法即可——「型号映射」表认得它，会自动补上标准型号。", ""),
    ("   （新出现的型号需要在该表补一行映射，或告诉我，我给你加规则）", ""),
    ("3. 库存、发货核对、项目进度三张表自动更新。", ""),
    ("4. 给某物资设预警线，在「库存总览」的 H 列「安全库存」填数字。", ""),
    ("", ""),
    ("五、已知数据缺口", ""),
    ("1. 采购表「设计量」列，20 个项目中只有「八期一阶段」「26年更新改造一阶段」填了，其余为空。", ""),
    ("2. 站点设计量只有集成商C的数据；集成商A、集成商B无站点级设计量来源表。", ""),
    ("3. 期初库存只有集成商B有「2025年结存」，集成商C/集成商A按 0 处理。", ""),
    ("4. 库存口径已对账：按「期初+入−出」算出的库存 vs 三家汇总表库存，233 项中 230 项一致（98.7%）。", ""),
]
for i, (a, b) in enumerate(notes, start=1):
    ws.cell(row=i, column=1, value=a)
    ws.cell(row=i, column=2, value=b)
ws.column_dimensions["A"].width = 60
ws.column_dimensions["B"].width = 72
ws.cell(row=1, column=1).font = TITLE_FONT

# ---------------- 基础档案 ----------------
# 注意：单位全称与所在地市属内部信息，此处用占位符，实际使用时替换为真实值
mk("集成商档案", ["集成商", "单位全称", "仓库/项目部", "备注"],
   [["集成商C", "＜集成商全称＞", "＜地市＞", ""],
    ["集成商A", "＜集成商全称＞", "＜地市＞", ""],
    ["集成商B", "＜集成商全称＞", "＜地市＞", ""]], [10, 34, 12, 16])

mk("物料主数据", ["标准编码", "标准名称", "标准型号", "单位", "出现于集成商", "归并写法数", "归并明细"],
   rd("物料主数据_归一"), [10, 34, 34, 8, 16, 11, 60])
mk("型号映射", ["出现于集成商", "原始型号", "单位", "标准编码", "标准型号", "匹配方式", "置信度"],
   rd("型号映射"), [16, 34, 8, 10, 34, 12, 8])
mk("规格相近候选", ["相似度", "标准型号A", "标准型号B", "A出现于", "B出现于"],
   rd("待确认归并"), [9, 36, 36, 16, 16], note="← 默认不合并，需人工确认")
mk("项目档案", ["项目名称", "涉及集成商", "备注"], rd("项目档案"))
mk("站点档案", ["站点名称", "项目", "集成商", "任务编码"], rd("站点档案"))


def with_std(rows, model_idx, code_col_out=True):
    out = []
    for r in rows:
        m = r[model_idx].strip() if model_idx < len(r) else ""
        c, s = std_map.get(m, ("", ""))
        out.append(list(r) + [c, s])
    return out


# ---------------- 流水（加标准编码/标准型号） ----------------
mk("入库明细", ["集成商", "日期", "月份", "型号", "数量", "单位", "材料来源", "项目", "备注", "标准编码", "标准型号"],
   with_std(rd("入库明细"), 3))
mk("出库明细", ["集成商", "月份", "日期", "型号", "数量", "项目", "材料去向", "领料人", "备注", "应发未发", "退废料", "标准编码", "标准型号"],
   with_std(rd("出库明细"), 3))
mk("退料明细", ["集成商", "类型", "日期", "型号", "数量", "单位", "项目", "站点", "退料人员", "备注", "标准编码", "标准型号"],
   with_std(rd("退料明细"), 3))
mk("站点设计量明细", ["集成商", "项目", "站点", "型号", "设计数量", "备注", "标准编码", "标准型号"],
   with_std(rd("站点设计量明细"), 3))

mk("项目采购到货明细",
   ["项目", "物料编码", "物料名称", "单位", "设计量", "采购1", "采购2", "采购汇总", "设计与提采差额", "供应商", "订单号", "订单数量", "接收数量"],
   rd("项目采购到货明细"))
mk("集成商库存汇总", ["集成商", "名称", "型号", "单位", "期初库存", "入库", "出库", "库存", "标准编码", "标准型号"],
   with_std(rd("集成商库存汇总"), 2))
mk("盘点记录", ["集成商", "盘点表", "物料名称", "型号", "单位", "账面数量", "实盘数量", "差异", "备注"], rd("盘点记录"))

# ---------------- 库存总览（按标准型号） ----------------
agg = {}
for r in rd("物料主数据_归一"):
    agg[r[2]] = {"code": r[0], "unit": r[3]}
for r in rd("型号映射"):
    s = r[4].strip()
    if s:
        d = agg.setdefault(s, {"code": r[3], "unit": r[2]})
        if r[2] and not d.get("unit"):
            d["unit"] = r[2]
models = sorted(agg.items())
ws = wb.create_sheet("库存总览")
hdr = ["标准编码", "标准型号", "单位", "集成商C库存", "集成商A库存", "集成商B库存", "库存合计", "安全库存", "库存状态"]
ws.append(hdr)
for c in range(1, len(hdr) + 1):
    cell = ws.cell(row=1, column=c)
    cell.fill = HDR_FILL if c <= 7 else HDR2_FILL
    cell.font = HDR_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for i, (model, d) in enumerate(models, start=2):
    ws.cell(row=i, column=1, value=d["code"])
    ws.cell(row=i, column=2, value=model)
    ws.cell(row=i, column=3, value=d["unit"])
    for j, tag in enumerate(["集成商C", "集成商A", "集成商B"], start=4):
        f = (f'=SUMIFS(集成商库存汇总!$E:$E,集成商库存汇总!$J:$J,$B{i},集成商库存汇总!$A:$A,"{tag}")'
             f'+SUMIFS(入库明细!$E:$E,入库明细!$K:$K,$B{i},入库明细!$A:$A,"{tag}")'
             f'-SUMIFS(出库明细!$E:$E,出库明细!$M:$M,$B{i},出库明细!$A:$A,"{tag}")')
        ws.cell(row=i, column=j, value=f)
    ws.cell(row=i, column=7, value=f"=SUM(D{i}:F{i})")
    ws.cell(row=i, column=9, value=f'=IF($G{i}<=0,"无库存",IF($H{i}="","未设预警线",IF($G{i}<=$H{i},"库存不足","正常")))')
for i, wd in enumerate([10, 36, 8, 11, 11, 12, 11, 10, 12], start=1):
    ws.column_dimensions[get_column_letter(i)].width = wd
ws.freeze_panes = "C2"
ws.auto_filter.ref = f"A1:I{max(2, ws.max_row)}"

# ---------------- 站点发货核对（按标准型号） ----------------
key = {}
for r in rd("站点设计量明细"):
    tag, proj, site, model, qty = r[0], r[1], r[2], r[3], r[4]
    if not site or not model:
        continue
    s = std_map.get(model.strip(), ("", ""))[1] or model.strip()
    k = (site, s)
    d = key.setdefault(k, {"proj": proj, "tag": tag, "q": 0.0})
    try:
        d["q"] += float(qty) if qty not in ("", None) else 0
    except Exception:
        pass
    if proj:
        d["proj"] = proj
check = sorted(key.items())
ws = wb.create_sheet("站点发货核对")
hdr = ["站点", "项目", "集成商", "标准型号", "单位", "设计数量", "已发货数量", "差异", "核对结论"]
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
    ws.cell(row=i, column=5, value=agg.get(model, {}).get("unit", ""))
    ws.cell(row=i, column=6, value=round(d["q"], 2))
    ws.cell(row=i, column=7, value=f'=SUMIFS(出库明细!$E:$E,出库明细!$G:$G,$A{i},出库明细!$M:$M,$D{i})')
    ws.cell(row=i, column=8, value=f"=$F{i}-$G{i}")
    ws.cell(row=i, column=9, value=f'=IF($G{i}=0,"未发货",IF($H{i}>0,"少发",IF($H{i}=0,"齐套","超发")))')
for i, wd in enumerate([26, 20, 9, 34, 8, 10, 12, 9, 10], start=1):
    ws.column_dimensions[get_column_letter(i)].width = wd
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:I{max(2, ws.max_row)}"

# ---------------- 项目进度 ----------------
pur = rd("项目采购到货明细")
pk = {}
for r in pur:
    proj, name, unit = r[0], r[2], r[3]
    if not proj or not name:
        continue
    d = pk.setdefault((proj, name), {"unit": unit, "dz": None, "cg": 0.0, "od": 0.0, "rc": 0.0})
    if unit and not d["unit"]:
        d["unit"] = unit

    def f(x):
        try:
            return float(x)
        except Exception:
            return None
    if f(r[4]) is not None:
        d["dz"] = f(r[4])
    for idx, kk in ((7, "cg"), (11, "od"), (12, "rc")):
        v = f(r[idx]) if idx < len(r) else None
        if v is not None:
            d[kk] += v
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
    ws.cell(row=i, column=8, value=f'=IF(OR($F{i}="",$F{i}=0),"",$G{i}/$F{i})')
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
