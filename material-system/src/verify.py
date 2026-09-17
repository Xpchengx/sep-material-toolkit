# -*- coding: utf-8 -*-
"""对账验证：明细算出的库存 vs 三家自己汇总表里的库存。"""
import os, io, csv, re
OUT = r"C:\Users\qazws\WorkBuddy\2026-09-09-09-19-37\material-system\out"


def rd(name):
    p = os.path.join(OUT, name + ".csv")
    with io.open(p, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    return rows[1:]


def f(x):
    try:
        return float(x)
    except Exception:
        return None


# 明细汇总
calc = {}
for r in rd("入库明细"):
    tag, model, q = r[0], r[3], f(r[4])
    if not model or q is None:
        continue
    calc.setdefault((tag, model), [0.0, 0.0])[0] += q
for r in rd("出库明细"):
    tag, model, q = r[0], r[3], f(r[4])
    if not model or q is None:
        continue
    calc.setdefault((tag, model), [0.0, 0.0])[1] += q

# 三家自己的汇总表
own = {}
for r in rd("集成商库存汇总"):
    tag, model, kc = r[0], r[2], f(r[7])
    if not model:
        continue
    own[(tag, model)] = (kc, f(r[4]) or 0.0)

tok = tot = 0
bad = []
for k, (i, o) in calc.items():
    qc = own.get(k, (None, 0.0))[1]
    if k not in own:
        continue
    tot += 1
    mine = round(qc + i - o, 2)
    theirs = own[k][0]
    if theirs is None:
        continue
    if abs(mine - theirs) <= 0.01:
        tok += 1
    else:
        bad.append((k, i, o, mine, theirs))

print(f"可比对 (集成商,型号) 组合: {tot}")
print(f"  完全一致: {tok}")
print(f"  有差异  : {len(bad)}")
print()
print("前 25 条差异（集成商 / 型号 / 累计入库 / 累计出库 / 我算的库存 / 原表库存）：")
for k, i, o, mine, theirs in sorted(bad, key=lambda x: -abs(x[4] - x[3] if x[4] is not None else 0))[:25]:
    print(f"  {k[0]:<4} {str(k[1])[:34]:<36} 入{i:>10.0f} 出{o:>10.0f}  算得{mine:>10.0f}  原表{theirs:>10.0f}")

# 覆羽率
if tot:
    print()
    print(f"一致率: {tok/tot*100:.1f}%")
