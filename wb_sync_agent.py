# -*- coding: utf-8 -*-
"""
wb_sync_agent.py —— 本机同步代理
==================================
让网页上的按钮真的能触发本机脚本。

┌─ 为什么需要它 ────────────────────────────────────────────────┐
│ 浏览器**不允许**网页执行本机程序，这是安全边界，绕不过去。      │
│ 所以中间放一个只监听 127.0.0.1 的小服务：                        │
│   网页点「同步」 → POST 到本服务 → 本服务跑脚本 → 回传进度与结果 │
└──────────────────────────────────────────────────────────────┘

它同时把仓库当静态站点服务，因此：

    http://localhost:8765/          ← 打开就是完整工作台，数据自动最新
                                        （同源，无需任何跨域配置）

跨源用法（GitHub Pages 上的那份页面来调它）也支持，见下方 CORS 说明。

接口
    GET  /                    页面（仓库根）
    GET  /workbench-data.js   映射到 output/workbench-data.js（始终最新）
    GET  /wb-meta.js          映射到 output/wb-meta.js
    GET  /api/status          运行状态与各步骤进度
    POST /api/sync            触发一次同步（已在跑则返回 409）
    GET  /api/data            最新的 workbench-data.json
    GET  /api/meta            最新的 wb-meta.json
    GET  /api/config          当前步骤清单（不含本机路径）

安全
    · 只绑 127.0.0.1 —— 局域网和公网都访问不到
    · 校验 Host 头，挡 DNS rebinding
    · CORS 只放行白名单来源，不是 `*`（避免任意网站读走你的库存数据）
    · 可选 requireToken：跨源请求需带令牌（同源不需要）

启动
    python wb_sync_agent.py
    python wb_sync_agent.py --port 8765 --config wb-sync.config.json
"""
import os, sys, json, time, html, socket, argparse, threading, subprocess, datetime
import urllib.parse
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = HERE
DEFAULT_CFG = os.path.join(REPO, 'wb-sync.config.json')
VERSION = '1.0'

# ------------------------------------------------------------ 默认配置

DEFAULT_STEPS = [
    {'name': '挑最新台账文件', 'cwd': 'material-system', 'cmd': ['python', 'src/pick_sources.py']},
    {'name': '抽取台账', 'cwd': 'material-system', 'cmd': ['python', 'src/extract.py']},
    {'name': '型号归一', 'cwd': 'material-system', 'cmd': ['python', 'src/normalize.py']},
    {'name': '生成工作台数据', 'cwd': 'material-system', 'cmd': ['python', 'src/build_workbench_data.py']},
]

DEFAULT_CFG_DATA = {
    'port': 8765,
    'ledgerDir': '',
    'requireToken': False,
    'token': '',
    'steps': DEFAULT_STEPS,
    'allowOrigins': ['https://xpchengx.github.io', 'http://localhost', 'http://127.0.0.1'],
    'stepTimeoutSec': 900,
}

# 页面会按「相对路径」找这几个文件，但它们在 output/ 下 —— 这里做映射
FILE_MAP = {
    '/workbench-data.js': 'output/workbench-data.js',
    '/workbench-data.json': 'output/workbench-data.json',
    '/wb-meta.js': 'output/wb-meta.js',
    '/wb-meta.json': 'output/wb-meta.json',
}

MIME = {
    '.html': 'text/html; charset=utf-8', '.js': 'application/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8', '.css': 'text/css; charset=utf-8',
    '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon',
    '.webmanifest': 'application/manifest+json; charset=utf-8',
    '.md': 'text/markdown; charset=utf-8', '.txt': 'text/plain; charset=utf-8',
}

# ------------------------------------------------------------ 同步状态（单例）

class SyncState:
    def __init__(self):
        self.lock = threading.Lock()
        self.running = False
        self.started_at = None
        self.finished_at = None
        self.ok = None
        self.steps = []
        self.error = ''

    def snapshot(self):
        with self.lock:
            return {
                'running': self.running,
                'startedAt': self.started_at,
                'finishedAt': self.finished_at,
                'ok': self.ok,
                'error': self.error,
                'steps': [dict(s) for s in self.steps],
            }

STATE = SyncState()
CFG = dict(DEFAULT_CFG_DATA)


# ------------------------------------------------------------ 工具

def now_str():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def load_cfg(path):
    cfg = dict(DEFAULT_CFG_DATA)
    if path and os.path.exists(path):
        try:
            with open(path, encoding='utf-8') as f:
                user = json.load(f)
            cfg.update({k: v for k, v in user.items() if v not in (None, '')})
            if not user.get('steps'):
                cfg['steps'] = DEFAULT_STEPS
        except Exception as e:
            print(f'[警告] 配置文件读不了，用默认配置：{e}')
    return cfg


def resolve_python(cmd):
    """把配置里的 'python' 换成本进程的解释器，避免 PATH 里没有 python。"""
    if cmd and cmd[0] in ('python', 'python3', 'py'):
        return [sys.executable] + list(cmd[1:])
    return cmd


def run_steps():
    """按顺序跑配置里的步骤，逐步更新进度。"""
    with STATE.lock:
        STATE.running = True
        STATE.started_at = now_str()
        STATE.finished_at = None
        STATE.ok = None
        STATE.error = ''
        STATE.steps = [{'name': s.get('name') or f'步骤{i+1}', 'status': 'pending',
                        'ms': None, 'tail': ''} for i, s in enumerate(CFG['steps'])]

    ok_all = True
    for i, step in enumerate(CFG['steps']):
        with STATE.lock:
            if i < len(STATE.steps):
                STATE.steps[i]['status'] = 'running'
        cwd = step.get('cwd') or REPO
        if not os.path.isabs(cwd):
            cwd = os.path.join(REPO, cwd)
        cmd = resolve_python(step.get('cmd') or [])
        env = dict(os.environ)
        if CFG.get('ledgerDir'):
            env['LEDGER_DIR'] = CFG['ledgerDir']

        t0 = time.time()
        try:
            r = subprocess.run(cmd, cwd=cwd, env=env, capture_output=True,
                               text=True, encoding='utf-8', errors='replace',
                               timeout=int(CFG.get('stepTimeoutSec') or 900))
            out = (r.stdout or '') + (('\n' + r.stderr) if r.stderr else '')
            code = r.returncode
        except subprocess.TimeoutExpired:
            out, code = f'超时（>{CFG.get("stepTimeoutSec")}s）', -1
        except Exception as e:
            out, code = f'{type(e).__name__}: {e}', -1

        ms = int((time.time() - t0) * 1000)
        tail = '\n'.join(out.strip().splitlines()[-24:])     # 只留尾部，够看
        with STATE.lock:
            if i < len(STATE.steps):
                STATE.steps[i].update({
                    'status': 'done' if code == 0 else 'failed',
                    'ms': ms, 'tail': tail, 'exitCode': code,
                })
        if code != 0:
            ok_all = False
            with STATE.lock:
                STATE.error = f'步骤「{step.get("name")}」失败（exit={code}）'
            break

    with STATE.lock:
        # 没跑到的步骤标成 skipped
        for s in STATE.steps:
            if s['status'] == 'pending':
                s['status'] = 'skipped'
        STATE.running = False
        STATE.finished_at = now_str()
        STATE.ok = ok_all
    print(f'[{now_str()}] 同步结束：{"成功" if ok_all else "失败"}')


def data_summary():
    """给前端看的「当前数据」摘要。"""
    p = os.path.join(REPO, 'output', 'workbench-data.json')
    info = {'exists': os.path.exists(p)}
    if not info['exists']:
        return info
    try:
        st = os.stat(p)
        info['mtime'] = datetime.datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M')
        info['size'] = st.st_size
        with open(p, encoding='utf-8') as f:
            d = json.load(f)
        info['snapshotDate'] = d.get('snapshotDate')
        info['materials'] = len(d.get('materials') or [])
        info['generatedAt'] = d.get('generatedAt') or (d.get('_meta') or {}).get('generatedAt')
    except Exception as e:
        info['error'] = f'{type(e).__name__}: {e}'
    return info


def origin_allowed(origin):
    if not origin:
        return True                       # 同源 / 直接访问
    for a in CFG.get('allowOrigins') or []:
        if origin == a or origin.startswith(a.rstrip('/') + ':') or origin.startswith(a):
            return True
    return False


# ------------------------------------------------------------ HTTP

class Handler(BaseHTTPRequestHandler):
    server_version = 'wb-sync-agent/' + VERSION
    protocol_version = 'HTTP/1.1'

    # ---------- 基础 ----------
    def log_message(self, fmt, *args):
        pass                              # 静音，避免刷屏；要排查可改成 print

    def _cors(self, origin):
        h = {}
        if origin and origin_allowed(origin):
            h['Access-Control-Allow-Origin'] = origin
            h['Vary'] = 'Origin'
            h['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
            h['Access-Control-Allow-Headers'] = 'Content-Type, X-WB-Token'
            h['Access-Control-Max-Age'] = '600'
            # Chrome 的 Private Network Access：公网页面访问本机需显式同意
            h['Access-Control-Allow-Private-Network'] = 'true'
        return h

    def _send(self, code, body, ctype='application/json; charset=utf-8', extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False)
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionAbortedError):
            pass

    def _host_ok(self):
        """挡 DNS rebinding：只认本机主机名。"""
        host = (self.headers.get('Host') or '').split(':')[0].strip().lower()
        return host in ('localhost', '127.0.0.1', '[::1]', '::1', '')

    def _token_ok(self, origin):
        if not CFG.get('requireToken'):
            return True
        if not origin:
            return True                   # 同源不需要
        want = CFG.get('token') or ''
        got = self.headers.get('X-WB-Token') or ''
        return bool(want) and got == want

    def _origin(self):
        return self.headers.get('Origin') or ''

    # ---------- 方法 ----------
    def do_OPTIONS(self):
        origin = self._origin()
        if not self._host_ok() or not origin_allowed(origin):
            return self._send(403, {'error': '来源不被允许'}, extra=self._cors(''))
        self._send(204, b'', 'text/plain', extra=self._cors(origin))

    def do_GET(self):
        if not self._host_ok():
            return self._send(403, {'error': 'Host 不被允许'})
        path = urllib.parse.urlparse(self.path).path
        origin = self._origin()
        cors = self._cors(origin)

        if path.startswith('/api/'):
            if not origin_allowed(origin):
                return self._send(403, {'error': '来源不被允许'}, extra=cors)
            if not self._token_ok(origin):
                return self._send(401, {'error': '需要令牌（本机代理开启了 requireToken）'}, extra=cors)
            return self._api_get(path, cors)

        return self._static(path, cors)

    def do_POST(self):
        if not self._host_ok():
            return self._send(403, {'error': 'Host 不被允许'})
        path = urllib.parse.urlparse(self.path).path
        origin = self._origin()
        cors = self._cors(origin)
        if not origin_allowed(origin):
            return self._send(403, {'error': '来源不被允许'}, extra=cors)
        if not self._token_ok(origin):
            return self._send(401, {'error': '需要令牌'}, extra=cors)

        if path == '/api/sync':
            with STATE.lock:
                if STATE.running:
                    return self._send(409, {'error': '正在同步中', 'status': STATE.snapshot()}, extra=cors)
            threading.Thread(target=run_steps, daemon=True).start()
            time.sleep(0.15)              # 让前端立刻看到 running=true
            return self._send(200, {'ok': True, 'status': STATE.snapshot()}, extra=cors)

        return self._send(404, {'error': 'no such api'}, extra=cors)

    # ---------- API ----------
    def _api_get(self, path, cors):
        if path == '/api/status':
            return self._send(200, {
                'ok': True, 'version': VERSION, 'pid': os.getpid(),
                'host': '127.0.0.1', 'time': now_str(),
                'ledgerDir': CFG.get('ledgerDir') or '',
                'data': data_summary(),
                'sync': STATE.snapshot(),
            }, extra=cors)

        if path == '/api/config':
            return self._send(200, {
                'ok': True,
                'steps': [{'name': s.get('name')} for s in CFG['steps']],
                'requireToken': bool(CFG.get('requireToken')),
            }, extra=cors)

        if path in ('/api/data', '/api/meta'):
            fn = 'workbench-data.json' if path == '/api/data' else 'wb-meta.json'
            p = os.path.join(REPO, 'output', fn)
            if not os.path.exists(p):
                return self._send(404, {'error': f'{fn} 还不存在，先跑一次同步'}, extra=cors)
            with open(p, 'rb') as f:
                return self._send(200, f.read(), 'application/json; charset=utf-8', extra=cors)

        return self._send(404, {'error': 'no such api'}, extra=cors)

    # ---------- 静态文件 ----------
    def _static(self, path, cors):
        rel = FILE_MAP.get(path)
        if rel is None:
            rel = urllib.parse.unquote(path).lstrip('/') or 'index.html'
        rel = rel.replace('\\', '/')
        target = os.path.normpath(os.path.join(REPO, rel))
        # 防目录穿越
        if not target.startswith(os.path.normpath(REPO)):
            return self._send(403, {'error': '越界'}, extra=cors)
        if os.path.isdir(target):
            target = os.path.join(target, 'index.html')
        if not os.path.exists(target):
            if path in ('/', '/index.html'):
                return self._send(200, _missing_page(), 'text/html; charset=utf-8', extra=cors)
            return self._send(404, {'error': f'找不到 {rel}'}, extra=cors)
        ext = os.path.splitext(target)[1].lower()
        with open(target, 'rb') as f:
            body = f.read()
        self._send(200, body, MIME.get(ext, 'application/octet-stream'), extra=cors)


def _missing_page():
    return ('<!doctype html><meta charset="utf-8"><title>缺少页面</title>'
            '<body style="font-family:system-ui;padding:40px;line-height:1.8">'
            '<h2>没找到 index.html</h2>'
            '<p>同步代理已在运行，但仓库根目录下没有 <code>index.html</code>。</p>'
            '<p>接口是通的：<a href="/api/status">/api/status</a></p></body>')


# ------------------------------------------------------------ 启动

def free_port_warn(port):
    s = socket.socket()
    try:
        s.bind(('127.0.0.1', port))
        return True
    except OSError:
        return False
    finally:
        s.close()


def main():
    global CFG
    ap = argparse.ArgumentParser()
    ap.add_argument('--port', type=int, default=None)
    ap.add_argument('--config', default=DEFAULT_CFG)
    ap.add_argument('--init-config', action='store_true', help='生成配置模板后退出')
    a = ap.parse_args()

    if a.init_config:
        if os.path.exists(DEFAULT_CFG):
            print(f'配置已存在，未覆盖：{DEFAULT_CFG}')
        else:
            with open(DEFAULT_CFG, 'w', encoding='utf-8') as f:
                json.dump(DEFAULT_CFG_DATA, f, ensure_ascii=False, indent=2)
            print(f'已生成配置模板：{DEFAULT_CFG}\n按需改 ledgerDir 后重新启动。')
        return 0

    CFG = load_cfg(a.config)
    port = a.port or int(CFG.get('port') or 8765)

    if not free_port_warn(port):
        print(f'[错误] 端口 {port} 已被占用。可能已经在跑了，换个端口：--port {port+1}')
        return 2

    print('=' * 66)
    print('  工作台同步代理')
    print('=' * 66)
    print(f'  地址    http://localhost:{port}/')
    print(f'  仓库    {REPO}')
    print(f'  台账    {CFG.get("ledgerDir") or "（未设置，脚本会用 LEDGER_DIR 或仓库内 data/）"}')
    print(f'  步骤    ' + ' → '.join(s.get("name", "?") for s in CFG['steps']))
    print()
    print('  直接打开上面的地址即可，页面上的「同步数据」会调用本代理执行脚本。')
    print('  只监听 127.0.0.1，局域网访问不到。按 Ctrl+C 停止。')
    print('=' * 66)

    srv = ThreadingHTTPServer(('127.0.0.1', port), Handler)
    srv.daemon_threads = True
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print('\n已停止。')
    finally:
        srv.server_close()
    return 0


if __name__ == '__main__':
    sys.exit(main())
