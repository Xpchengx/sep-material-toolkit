# -*- coding: utf-8 -*-
"""
pick_sources.py —— 自动挑选每个集成商最新的台账文件
=====================================================
解决一个老问题：extract.py 里的源文件路径是写死的，新文件来了不会自动用上，
"点了同步"其实同步的还是旧数据。

本脚本扫描台账目录，按规则挑出**每个集成商最新的一份**，写 sources.json 给
extract.py 读。extract.py 检测到该文件就用它，没有则回落原有的硬编码路径。

挑选规则（优先级从高到低）：
  1. 文件名里的日期（2026-09-08 / 2026.9.8 / 9月8日 / 20260908）
     中文日期没有年份，用「同目录内其他文件的年份」或当前年份补
  2. 文件修改时间
两者不一致时以**文件名日期**为准 —— 台账的文件名才是业务时点，
mtime 常因拷贝/下载而失真。

用法：
    python pick_sources.py                     # 用配置文件
    python pick_sources.py --ledger <目录> --dry-run
    python pick_sources.py --list              # 只列出各集成商的候选与选中项
"""
import os, re, sys, json, argparse, datetime

sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
DEFAULT_CFG = os.path.join(os.path.dirname(REPO), 'wb-sync.config.json')

# 文件名里日期的几种写法
ISO = re.compile(r'(20\d{2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})')
COMPACT = re.compile(r'(20\d{2})(\d{2})(\d{2})')
CN = re.compile(r'(\d{1,2})\s*月\s*(\d{1,2})\s*[号日]')

SKIP_EXT = ('.tmp', '.lnk', '.pdf', '.docx', '.doc', '.zip', '.rar', '.msg')


def name_date(name, mtime=None, fallback_year=None):
    """
    从文件名解析日期；解析不出返回 None。

    文件名没写年份时（如「10月26日」）年份只能推。规则：
      在 [mtime年-1, mtime年, mtime年+1] 里挑**与 mtime 最接近**的那个。
    比"用目录里的最大年份"可靠得多 —— 实测「XX材料库存(10月26日).xlsx」这类文件名
    的 mtime 是 2025-10-27，用目录最大年份会误判成 2026 年，
    结果选中一份比现有数据还旧的表。
    """
    m = ISO.search(name)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None

    m = COMPACT.search(name)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None

    m = CN.search(name)
    if not m:
        return None
    mo, dy = int(m.group(1)), int(m.group(2))

    years = []
    if mtime:
        years = [mtime.year - 1, mtime.year, mtime.year + 1]
    elif fallback_year:
        years = [fallback_year]
    if not years:
        return None

    best = None
    for y in years:
        try:
            d = datetime.date(y, mo, dy)
        except ValueError:
            continue
        gap = abs((d - mtime).days) if mtime else 0
        if best is None or gap < best[0]:
            best = (gap, d)
    return best[1] if best else None


def scan_dir(d, fallback_year=None):
    """列出目录下的候选表格，附上解析出的日期与 mtime。"""
    out = []
    if not os.path.isdir(d):
        return out
    for f in os.listdir(d):
        if f.startswith('~$') or f.startswith('.'):
            continue
        p = os.path.join(d, f)
        if not os.path.isfile(p):
            continue
        if os.path.splitext(f)[1].lower() in SKIP_EXT:
            continue
        if not f.lower().endswith(('.xlsx', '.xls', '.csv')):
            continue
        try:
            mt = datetime.date.fromtimestamp(os.path.getmtime(p))
        except OSError:
            mt = None
        out.append({
            'file': f, 'path': p, 'size': os.path.getsize(p),
            'nameDate': name_date(f, mt, fallback_year), 'mtime': mt,
        })
    return out


def pick(cands):
    """挑最新的一份：先比文件名日期，再比 mtime。"""
    if not cands:
        return None

    def key(x):
        return (x['nameDate'] or datetime.date(1900, 1, 1),
                x['mtime'] or datetime.date(1900, 1, 1))

    return max(cands, key=key)


def guess_fallback_year(cands):
    """同目录里能确定的年份，用来补『9月8日』这种没年份的写法。"""
    ys = [c['nameDate'].year for c in cands if c['nameDate']]
    ys += [c['mtime'].year for c in cands if c['mtime']]
    return max(ys) if ys else datetime.date.today().year


def load_cfg(path):
    if path and os.path.exists(path):
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    return {}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--config', default=DEFAULT_CFG, help='配置文件')
    ap.add_argument('--ledger', help='台账根目录（覆盖配置）')
    ap.add_argument('--out', help='输出的 sources.json 路径')
    ap.add_argument('--dry-run', action='store_true', help='只打印，不写文件')
    ap.add_argument('--list', action='store_true', help='列出各集成商候选与选中项')
    a = ap.parse_args()

    cfg = load_cfg(a.config)
    base = a.ledger or cfg.get('ledgerDir') or os.environ.get('LEDGER_DIR')
    if not base:
        raise SystemExit('未指定台账目录：用 --ledger，或在配置里设 ledgerDir，或设环境变量 LEDGER_DIR')
    if not os.path.isdir(base):
        raise SystemExit(f'台账目录不存在：{base}')

    # 集成商：配置里 "suppliers" 是 {目录名: 角色}，角色 A/B/C 代表三种台账格式。
    # 没有该项时，把 集成商/ 下的每个子目录当成一个集成商，标签用目录名。
    sup_cfg = cfg.get('suppliers') or {}
    if not sup_cfg:
        sub = os.path.join(base, '集成商')
        if os.path.isdir(sub):
            for d in sorted(os.listdir(sub)):
                if os.path.isdir(os.path.join(sub, d)):
                    sup_cfg[d] = ''
    folders = sup_cfg or cfg.get('supplierFolders') or {}
    roles = {}

    result, detail = {}, []
    for folder, role in folders.items():
        d = os.path.join(base, '集成商', folder)
        if not os.path.isdir(d):
            d = os.path.join(base, folder)          # 也允许直接放在根下
        cands = scan_dir(d)
        chosen = pick(cands)
        if chosen:
            result[folder] = chosen['path']
            if role:
                roles[str(role).strip().upper()] = folder
            detail.append((folder, folder, cands, chosen))
        else:
            detail.append((folder, folder, cands, None))

    # 采购/设计量文件
    pf = cfg.get('purchaseFile')
    if pf:
        pp = pf if os.path.isabs(pf) else os.path.join(base, pf)
        if os.path.isfile(pp):
            result['采购'] = pp

    # ---- 输出 ----
    for alias, folder, cands, chosen in detail:
        print(f'■ {alias}  ({folder})  候选 {len(cands)} 个')
        cs = sorted(cands, key=lambda x: (x['nameDate'] or datetime.date(1900, 1, 1)), reverse=True)[:4]
        for c in cs:
            mark = '←选' if chosen and c['path'] == chosen['path'] else '  '
            nd = c['nameDate'].isoformat() if c['nameDate'] else '（文件名无日期）'
            print(f'    {mark} {nd:<12} mtime={c["mtime"]}  {c["file"][:52]}')
        if not chosen:
            print('    !! 该集成商没有找到可用表格')
    print()

    missing = [k for k in folders.keys() if k not in result]
    if missing:
        print(f'[注意] 未找到文件的集成商：{missing}')

    out_path = a.out or cfg.get('sourcesOut') or os.path.join(HERE, 'sources.json')
    payload = {
        'generatedAt': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'ledgerDir': base,
        'sources': result,
        # 角色 -> 名称：extract.py 按角色分支处理不同格式的台账
        'roles': roles,
    }
    if a.dry_run:
        print(f'[dry-run] 未写文件。将写入 {out_path}：')
        print(json.dumps(payload, ensure_ascii=False, indent=1))
    else:
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=1)
        print(f'已写入 {out_path}')
        for k, v in result.items():
            print(f'   {k} -> {os.path.basename(v)}')
        if roles:
            print('   角色: ' + ', '.join(f'{k}={v}' for k, v in sorted(roles.items())))

    return 0 if result else 1


if __name__ == '__main__':
    sys.exit(main())
