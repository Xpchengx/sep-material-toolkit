# -*- coding: utf-8 -*-
"""型号归一：把三家各自的写法映射到统一标准型号。

分级：
  A 机械归一（去空格/全半角/分隔符/大小写）自动命中
  B 人工规则表（同义写法、错别字、交流直流缺省）
  C 相似度候选 -> 输出待确认清单给用户
"""
import os, io, csv, re, difflib
from collections import defaultdict, Counter

OUT = r"C:\Users\qazws\WorkBuddy\2026-09-09-09-19-37\material-system\out"


def rd(name):
    p = os.path.join(OUT, name + ".csv")
    if not os.path.exists(p):
        return []
    with io.open(p, encoding="utf-8-sig", newline="") as f:
        r = list(csv.reader(f))
    return r[1:]


def norm(s):
    """机械归一：全角转半角、去空白、统一分隔符、去分隔符、大写。"""
    s = (s or "").strip()
    s = s.replace("（", "(").replace("）", ")").replace("：", ":").replace("，", ",")
    s = s.replace("、", "/").replace("×", "x").replace("*", "x")
    s = re.sub(r"\s+", "", s)
    s = s.replace("\\", "/").replace("-", "/").replace("_", "/")
    s = s.upper()
    s = s.replace("/", "")
    return s


# ---------- B 级：人工规则表（原始型号 -> 标准型号） ----------
SYN = {
    # 光模块
    "1.25G光模块": "光模块-1.25G", "1.25G模块": "光模块-1.25G", "光模块-1.25G": "光模块-1.25G",
    "10G光模块": "光模块-10G(10km)", "光模块-10GE": "光模块-10G(10km)",
    "光模块-10G（10km）": "光模块-10G(10km)",
    "光模块-10G（1.4km）": "光模块-10G(1.4km)", "4G 光模块(9.8G，1.4Km)": "光模块-9.8G(1.4km)",
    "25G传输光模块": "光模块-25G(10km)", "光模块-25GE": "光模块-25G(10km)",
    "5G 光模块(25G，10Km)": "光模块-25G(10km)",
    "5G 光模块(25G，10Km)皮站专用": "光模块-25G(10km)皮站专用",
    "5G 光模块(25G，2Km)": "光模块-25G(2km)",
    "光模块-拉远1.52G": "光模块-拉远1.52G", "光传输模块单元拉远1.52G": "光模块-拉远1.52G",
    # AAU / RRU
    "AAU-5246": "AAU5246", "AAU5246": "AAU5246", "AAU-5246(交流）": "AAU5246",
    "AAU-5246W": "AAU5246W", "AAU-5249W(交流）": "AAU5249W",
    "5G AAU5246": "AAU5246",
    "RRU-5263-18d": "RRU5263-18d", "5G RRU（5263-18d）": "RRU5263-18d",
    "RRU-5263-18d（交流）": "RRU5263-18d",
    "RRU-5263-de（交流）": "RRU5263-de",
    "RRU-5278-d（交流）": "RRU5278-d", "RRU5278-d": "RRU5278-d",
    "RRU-5269": "RRU5269", "5269": "RRU5269", "RRU-5269-d(交流)": "RRU5269-d",
    "RRU-5263de": "RRU5263-de",
    # 基带板 / 主控
    "基带板-g6C": "基带板-g6c", "基带板-g6ce": "基带板-g6ce", "基带板-g6b": "基带板-g6b",
    "基带板-g6i": "基带板-g6i", "基带板-g6m": "基带板-g6m", "基带板-g5b": "基带板-g5b",
    "5G 基带板": "5G基带板", "5G 主控板": "5G主控板", "4G 主控板": "4G主控板",
    "主控传输单元g6": "主控传输单元-g6", "通用主控传输单元-2电FE/GE": "通用主控传输单元-2电FE/GE",
    "4G TDD基带板": "4G TDD基带板",
    # rHUB
    "rHUB": "rHUB", "rhub": "rHUB", "RHUB5775": "rHUB5775", "功能模块-RHBU5775": "rHUB5775",
    # 交转直
    "交转直电源模块": "交转直电源模块", "集成商C交转直电源模块": "交转直电源模块",
    "集成商C集成商C交转直电源模块": "交转直电源模块",
    # 皮站（型/性 错别字）
    "扩展型皮站主机": "扩展型皮站主机", "扩展性皮站主机": "扩展型皮站主机",
    "扩展型皮站扩展单元": "扩展型皮站扩展单元", "扩展性皮站扩展单元": "扩展型皮站扩展单元",
    "扩展型皮站远端单元（外置天线）": "扩展型皮站远端单元",
    "扩展性皮站远端单元（外接天线）": "扩展型皮站远端单元",
    "扩展性皮站远端单元": "扩展型皮站远端单元",
    # 电力电缆（去「\黑」后缀）
    "通信用电力电缆\\ZA-RVV 2×4": "通信用电力电缆\\ZA-RVV 2×4",
    "通信用电力电缆\\ZA-RVV 2×4\\黑": "通信用电力电缆\\ZA-RVV 2×4",
    "通信用电力电缆\\ZA-RVV 2×6": "通信用电力电缆\\ZA-RVV 2×6",
    "通信用电力电缆\\ZA-RVV 2×6\\黑": "通信用电力电缆\\ZA-RVV 2×6",
    "通信用电力电缆\\ZA-RVV 3×2.5\\黑": "通信用电力电缆\\ZA-RVV 3×2.5",
    "通信用电力电缆\\ZA-RVV 2×10": "通信用电力电缆\\ZA-RVV 2×10",
    # 分纤箱 / 光交 / 终端盒
    "24芯室内壁挂分纤箱FC-PC（楼内）": "24芯室内壁挂分纤箱FC-PC(楼内)",
    "24芯室内壁挂分纤箱FC-PC（楼内": "24芯室内壁挂分纤箱FC-PC(楼内)",
    "144芯壁挂光交": "144芯壁挂光交", "144壁挂光交": "144芯壁挂光交",
    "24芯壁挂光缆终端盒": "24芯光缆终端盒", "终端盒24芯": "24芯光缆终端盒",
    # 光电复合缆
    "光电混合缆": "光电复合缆", "光电复合缆": "光电复合缆", "复合光电缆": "光电复合缆",
    "光电混合缆电缆接头": "光电复合缆接头", "光电混合缆光缆接头": "光电复合缆接头",
    "光电混合缆电缆接头": "光电复合缆接头",
    # 接地线
    "1*16地线": "接地线1×16", "地线16平米": "接地线1×16", "接地线1x16": "接地线1×16",
    "接地线1x6": "接地线1×6",
    # 走线 / 直角头
    "后线抓": "走线爪", "走线爪": "走线爪",
    "直角头": "直角头", "直角头N-JK": "直角头",
    # 天线
    "GPS天线(副)": "GPS天线", "GPS天线": "GPS天线",
    "GPS-北斗天馈包": "GPS/北斗天馈包", "北斗天馈包": "GPS/北斗天馈包",
    "GPS北斗三天线包": "GPS/北斗三天线包",
    "GPS一分四": "GPS一分四", "GPS功分器": "GPS功分器",
    "900-D频段\\窄波束高增益\\双通道天线\\电梯": "电梯双通道天线(900-D窄波束)",
    "900-D全频二口水平面大张角天线": "电梯大张角天线(900-D)",
    "特殊场景天线高楼900-D全频二口垂直面大张角天线": "高楼大张角天线(900-D)",
    # 网线
    "五类网线": "五类网线", "五类网线水晶头": "五类网线水晶头",
    "超六类网线": "超六类网线", "超六类网线水晶头": "超六类网线水晶头",
    # 电梯宝
    "电梯宝主控单元": "电梯宝主控单元", "电梯型主控单元": "电梯宝主控单元",
    "电梯型轿厢单元": "电梯宝轿厢单元", "电梯宝轿厢单元": "电梯宝轿厢单元",
    "电梯宝主控单元（一拖一）": "电梯宝主控单元(一拖一)",
    "电梯宝主控单元（一拖二）": "电梯宝主控单元(一拖二)",
    "电梯宝主控单元（一拖三）": "电梯宝主控单元(一拖三)",
    # 直放站
    "光纤直放站近端机-京信": "光纤直放站近端机", "光纤直放站远端机-京信": "光纤直放站远端机",
    "5G直放站近端": "5G直放站近端", "5G直放站远端": "5G直放站远端",
    "光纤分布系统-扩展单元（室外型）（虹信）": "光纤分布系统-扩展单元",
    "光纤分布系统-接入单元（虹信）": "光纤分布系统-接入单元",
    # 合路器
    "合路器-高功率-GSM&DCS/TD F&TD A/TD E(三路)-N型头": "合路器-三路-N型头",
    "合路器-高功率-GSM/DCS/TD F&TD A&TD E(三路)-N型头": "合路器-三路-N型头",
    "合路器高功率GSM&DCS&TD F&TD A/TD E/TD D（三路）N头": "合路器-三路-N型头",
    "合路器-高功率-GSM&DCS/TD F&TD A/TD D(三路)-N型头": "合路器-三路-N型头",
    "合路器-高功率-GSM/DCS/TD F&TD A&TD E/NR(四路)-N型头": "合路器-四路-N型头",
    "合路器-(四路)-N型头": "合路器-四路-N型头",
    # 尾纤 / 跳纤
    "LC-FC铠装尾纤": "LC-FC铠装尾纤", "LC-LC铠装尾纤": "LC-LC铠装尾纤",
    "野战尾纤": "野战尾纤", "野战光缆": "野战光缆",
    "FC-LC 15米": "FC-LC跳纤15米",
    "LC-LC 10米": "LC-LC跳纤10米", "LC-LC 5米": "LC-LC跳纤5米",
    "LC/LC/652单模单芯跳纤10米": "LC/LC/652单模单芯跳纤10米",
    "LC/LC/652单模单芯跳纤15米": "LC/LC/652单模单芯跳纤15米",
    "LC/PC/652单模单芯跳纤10米": "LC/PC/652单模单芯跳纤10米",
    "LC/PC/652单模单芯跳纤15米": "LC/PC/652单模单芯跳纤15米",
    "LC/PC/652单模单芯跳纤5米": "LC/PC/652单模单芯跳纤5米",
    "SC-SC 3米": "SC-SC跳纤3米",
    # 负载
    "普通 负载(200W)": "普通负载(200W)", "普通 负载(50W)": "普通负载(50W)",
    "负载N型": "负载N型",
    # 托盘 / 挡风板
    "挡风板": "挡风板", "1U挡风板": "1U挡风板",
    # 其他
    "同轴电缆": "同轴电缆", "同轴电缆-国标型号：RG-8U": "同轴电缆(RG-8U)",
    "24芯阻燃光缆": "24芯阻燃光缆", "48芯阻燃光缆": "48芯阻燃光缆",
    "槽道光缆": "槽道光缆",
    "室内小容量机柜": "室内小容量机柜",
    "天线": "天线", "整机辅料包": "整机辅料包", "整机辅料包-BBU5900专用": "整机辅料包(BBU5900专用)",
    "INCR": "iNCR", "INCR5852RN(18002600M)": "iNCR5852RN", "INCR移动辅料包": "iNCR移动辅料包",
    "U": "INCR5852RN",
    "prru5653": "pRRU5653GR", "prru5654": "pRRU5654GR",
    "pRRU(内置天线)": "pRRU(内置天线)",
    "多模射频模块-pRRU5653GH": "pRRU5653GH", "多模射频模块-pRRU5653GR": "pRRU5653GR",
    "多模射频模块-pRRU5654GR": "pRRU5654GR", "多模射频模块-pRRU5631GR": "pRRU5631GR",
    "扩展单元-AU（基站放大器）": "扩展单元-AU(基站放大器)",
    "远端单元-RU（多载波功放）": "远端单元-RU(多载波功放)",
    "光缆组件": "光缆组件", "光缆组件DLC/UPC-单模-10m-2芯": "光缆组件DLC/UPC-10m-2芯",
    "光缆组件DLC/UPC-单模-5m-2芯": "光缆组件DLC/UPC-5m-2芯",
    "光缆组件DLC/UPC-单模-50m-2芯": "光缆组件DLC/UPC-50m-2芯",
    "单模块1芯": "单模块1芯", "光缆段数（楼内）": "光缆段数(楼内)",
    "分布式成套工程标签": "分布式成套工程标签", "配电单元辅料包": "配电单元辅料包",
    "配电盒DCDU": "配电盒DCDU", "DCDU": "配电盒DCDU",
    "1U机框": "1U机框", "3U机框": "3U机框",
    "电源适配器": "电源适配器", "理线器": "理线器", "电话线": "电话线",
    "电话线接头": "电话线接头", "天馈接地夹": "天馈接地夹", "天馈线缆安装辅料": "天馈线缆安装辅料",
    "PVC管25": "PVC管25", "钢管25": "钢管25",
    "开检修孔（包含开孔和恢复等费用）": "开检修孔(含恢复费用)",
    "小站室外安装件": "小站室外安装件", "整机辅料包": "整机辅料包",
}

# 明确「不合并」的保护名单（避免相似度误判把不同物资并掉）
GUARD = set()


def main():
    # 收集三家型号
    pairs = defaultdict(set)
    per = defaultdict(Counter)
    unit_of = {}
    for r in rd("入库明细"):
        if r[3].strip():
            per[r[0]][r[3].strip()] += 1
            if r[5]:
                unit_of.setdefault(r[3].strip(), r[5])
    for r in rd("出库明细"):
        if r[3].strip():
            per[r[0]][r[3].strip()] += 1
    for r in rd("集成商库存汇总"):
        if r[2].strip():
            per[r[0]][r[2].strip()] += 1
            if r[3]:
                unit_of.setdefault(r[2].strip(), r[3])
    for r in rd("站点设计量明细"):
        if r[3].strip():
            per[r[0]][r[3].strip()] += 1

    for tag, c in per.items():
        for m in c:
            pairs[m].add(tag)

    # 归一：SYN 优先，其次机械归一
    std_of = {}
    way = {}
    for m in pairs:
        if m in SYN:
            std_of[m] = SYN[m]
            way[m] = "B-规则" if SYN[m] != m else "A-原名"
        else:
            std_of[m] = m
            way[m] = "A-原名"

    # 先按 SYN 分组，再对未进 SYN 的做机械归一聚类
    groups = defaultdict(set)          # 标准名 -> 原始型号集合
    for m, s in std_of.items():
        groups[s].add(m)

    # 机械归一并组：把 norm 相同但标准名不同的合并（选最短的原名作标准）
    mech = defaultdict(list)
    for s in list(groups):
        mech[norm(s)].append(s)
    merged = {}
    for k, ss in mech.items():
        if len(ss) > 1:
            # 选出现次数最多的原名作标准
            best = max(ss, key=lambda x: sum(per[t].get(o, 0) for o in groups[x] for t in per))
            for s in ss:
                for o in groups[s]:
                    std_of[o] = best
                    if way[o] == "A-原名":
                        way[o] = "A-机械归一"
                merged.setdefault(best, []).extend(groups[s])

    groups = defaultdict(set)
    for m, s in std_of.items():
        groups[s].add(m)

    # 分配标准编码
    stds = sorted(groups)
    code_of = {s: f"M{i:04d}" for i, s in enumerate(stds, start=1)}

    # 输出 型号映射
    rows = []
    for m in sorted(pairs):
        s = std_of[m]
        rows.append(["/".join(sorted(pairs[m])), m, unit_of.get(m, ""),
                     code_of[s], s, way[m],
                     "高" if way[m] in ("A-原名", "A-机械归一") else "中"])
    with io.open(os.path.join(OUT, "型号映射.csv"), "w", encoding="utf-8-sig", newline="") as f:
        cw = csv.writer(f)
        cw.writerow(["出现于集成商", "原始型号", "单位", "标准编码", "标准型号", "匹配方式", "置信度"])
        cw.writerows(rows)

    # 输出 物料主数据
    mrows = []
    for s in stds:
        mem = sorted(groups[s])
        tagset = sorted({t for o in mem for t in pairs[o]})
        unit = ""
        for o in mem:
            if unit_of.get(o):
                unit = unit_of[o]
                break
        mrows.append([code_of[s], s, s, unit, "/".join(tagset),
                      len(mem), " | ".join(mem) if len(mem) > 1 else ""])
    with io.open(os.path.join(OUT, "物料主数据_归一.csv"), "w", encoding="utf-8-sig", newline="") as f:
        cw = csv.writer(f)
        cw.writerow(["标准编码", "标准名称", "标准型号", "单位", "出现于集成商", "归并写法数", "归并明细"])
        cw.writerows(mrows)

    # 待确认：相似度 0.70~0.97 且不属于同一标准，且排除「只差数字」的正常不同规格
    def digitless(s):
        return re.sub(r"\d+", "#", norm(s))

    # 规格维度：任一层取值不同即视为不同物料，不算同义写法
    SPEC_DIMS = [
        (re.compile(r"单极化|双极化"), "极化"),
        (re.compile(r"[二三四]功分|[二三四]路"), "路数"),
        (re.compile(r"UPC|APC|FC|SC|LC|DIN|N/?J|N/?K|RG-?8U|PC"), "接口"),
        (re.compile(r"一拖[一二三]"), "拖动"),
        (re.compile(r"防水|不防水"), "防水"),
        (re.compile(r"楼内|楼外|室外|室内|室外型"), "场景"),
        (re.compile(r"黄绿|红|蓝|黑"), "颜色"),
        (re.compile(r"交流|直流"), "电源"),
        (re.compile(r"皮站|微站|小站"), "站型"),
        (re.compile(r"铠装"), "铠装"),
        (re.compile(r"公头|母头"), "头型"),
        (re.compile(r"^1U|^3U"), "机框"),
        (re.compile(r"全向|定向|大张角|板状|吸顶|对数周期|射灯"), "方向形态"),
        (re.compile(r"近端|远端"), "端型"),
        (re.compile(r"[一二三四五六七八]通道|[单双]通道"), "通道"),
        (re.compile(r"尾纤|跳纤"), "纤型"),
        (re.compile(r"光缆|馈线|电缆|网线"), "线缆类型"),
    ]

    def specclean(s):
        """规格提取用：保留 / 和 - 分隔符，只做大小写与全半角统一。"""
        s = (s or "").strip()
        s = s.replace("（", "(").replace("）", ")")
        s = re.sub(r"\s+", "", s)
        return s.upper()

    def specbits(s):
        t = specclean(s)
        return {name: tuple(sorted(set(rx.findall(t)))) for rx, name in SPEC_DIMS}

    def differ_spec(a, b):
        sa, sb = specbits(a), specbits(b)
        return any(sa[k] != sb[k] for k in sa)

    cand = []
    for i, a in enumerate(stds):
        for b in stds[i + 1:]:
            na, nb = norm(a), norm(b)
            if na == nb:
                continue
            # 只差数字（波长/长度/功率/增益等）→ 是不同规格，不是同义写法
            if digitless(a) == digitless(b):
                continue
            # 规格维度不同 → 本来就该分开
            if differ_spec(a, b):
                continue
            r = difflib.SequenceMatcher(None, na, nb).ratio()
            if 0.62 <= r < 0.97:
                cand.append((round(r, 3), a, b,
                             "/".join(sorted({t for o in groups[a] for t in pairs[o]})),
                             "/".join(sorted({t for o in groups[b] for t in pairs[o]}))))
    cand.sort(reverse=True)
    with io.open(os.path.join(OUT, "待确认归并.csv"), "w", encoding="utf-8-sig", newline="") as f:
        cw = csv.writer(f)
        cw.writerow(["相似度", "标准型号A", "标准型号B", "A出现于", "B出现于"])
        cw.writerows(cand[:120])

    # 报告
    lines = []
    lines.append(f"原始型号写法总数 : {len(pairs)}")
    lines.append(f"归并后标准物料数 : {len(stds)}")
    lines.append(f"合并掉的写法数   : {len(pairs) - len(stds)}")
    byway = Counter(way.values())
    for k, v in byway.items():
        lines.append(f"  {k}: {v}")
    multi = [m for m in mrows if m[5] > 1]
    lines.append(f"由多个写法合并成的标准物料: {len(multi)}")
    lines.append("")
    lines.append("—— 归并了 3 个以上写法的（最值得看） ——")
    for m in sorted(multi, key=lambda x: -x[5])[:30]:
        lines.append(f"  [{m[0]}] {m[1]}  <- {m[5]} 种写法: {m[6][:110]}")
    lines.append("")
    lines.append(f"—— 待确认相似候选: {len(cand)} 组（已写入 待确认归并.csv，前30） ——")
    for r, a, b, ta, tb in cand[:30]:
        lines.append(f"  {r}  「{a}」({ta})  vs  「{b}」({tb})")
    r = "\n".join(lines)
    with io.open(os.path.join(OUT, "normalize_report.txt"), "w", encoding="utf-8") as f:
        f.write(r)
    print(r)


main()
