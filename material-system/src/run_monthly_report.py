# -*- coding: utf-8 -*-
"""
一键月报 CLI  run_monthly_report.py
====================================
用法：
    python src/run_monthly_report.py --month 8 --year 2026

参数直接改成上个月份即可，例如 9 月月报：
    python src/run_monthly_report.py --month 9

数据源约定（三家库存工作簿路径写死在下方 FILES，换文件时改这里）：
    集成商A   <BASE>/集成商A/上海集成商A材料库存(N月N日）.xlsx
    集成商B <BASE>/集成商B/集成商B-材料库存YYYY-M-D.xlsx
    集成商C   <BASE>/集成商C/集成商C库存(N月N日).xlsx
"""
import sys, os, subprocess, argparse, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PY = sys.executable


def run(script, *args):
    cmd = [PY, os.path.join(HERE, script), *args]
    print('>>>', ' '.join(cmd), flush=True)
    r = subprocess.run(cmd, cwd=ROOT)
    if r.returncode != 0:
        raise SystemExit(f'{script} 失败，退出码 {r.returncode}')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--month', type=int, required=True, help='报告月份，如 8')
    ap.add_argument('--year', type=int, default=datetime.date.today().year)
    ap.add_argument('--skip-extract', action='store_true', help='复用已有 JSON，不重新抽取')
    a = ap.parse_args()

    if not a.skip_extract:
        run('monthly_report_data.py')
    out = f'out/材料管理周转仓管理月报（{a.month}月）.docx'
    run('build_monthly_report.py', '--month', str(a.month), '--year', str(a.year), '--out', out)
    print('\n完成 ->', os.path.join(ROOT, out))
