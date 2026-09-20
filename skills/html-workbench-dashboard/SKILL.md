---
name: html-workbench-dashboard
description: 制作单文件 HTML 工作台/看板（左侧常驻功能栏 + 右侧卡片流），零依赖离线可用，内置可替换的数据块与一键流程演示。当用户要"网页版""HTML 格式""工作台""看板""dashboard"呈现某个系统或数据时使用。触发词：网页版、html、工作台、看板、dashboard、可视化界面、系统界面。
agent_created: true
---

# 单文件 HTML 工作台

产出：**一个 .html 文件**，双击即开、离线可用、无 CDN 依赖。

## 为什么单文件 + 零依赖

- 用户要拿去截图、演示、发给别人看，多文件或外链 CDN 一断网就废。
- 图表**手写 SVG**，不引 ECharts/Chart.js。折线、柱状、进度条都够用。

## 标准骨架

```html
<div class="app">
  <aside class="sidebar">      <!-- 左侧常驻功能栏：品牌区 + 分组导航 + 页脚 -->
  <div class="main">
    <header class="topbar">     <!-- 标题 + 副标题 + 状态 + 主操作按钮 -->
    <main class="content">
      <section class="view active" id="v-overview">   <!-- 每个视图一个 section -->
```

- 导航用 `data-view` + `switchView()` 切换 `.active` 类，不要多页面。
- 视图分组用「监控 / 数据 / 交付 / 设置」这类标题分隔，比平铺一堆菜单清楚。

## 数据块放在文件底部

```js
/* 接真实数据时只替换本段 */
const RAW_DATA = {
  snapshotDate: "2026-09-15",
  trend: [ ["2026-09-02", 87240], ... ],
  materials: [ /* 数组行，字段顺序注释写清楚 */ ]
};
```
**用数组行而不是对象数组**：更紧凑、便于从表格/CSV 直接生成，
字段含义在上一行注释里写清楚。页面里的「数据与说明」视图再把字段口径列成表。

## 必做的三步校验

### 1. 数据自检（最容易翻车的地方）
**演示数据也要过一遍页面自己的判定规则。** 曾出现第一版数据算出 25 项预警、
健康度 47%，明显失真。交付前用脚本按同样的规则复算一遍：

```python
# 用能处理 JS 转义引号的正则，否则含 \" 的行会被漏掉
STR = r'((?:[^"\\]|\\.)*)'
row_re = re.compile(r'\["' + STR + r'","' + STR + r'","' + STR + r'",(\d+),(\d+),(\d+),(\d+),(\d+),"'
                    + STR + r'","' + STR + r'"\]')
```
检查：① 各统计量与曲线末值是否自洽 ② 预警项数量是否合理（一般 10%~20%，
超过 50% 说明阈值定错了）③ 有没有出现"跨单位求和"这类口径错误。

### 2. JS 语法
```bash
python _extract_js.py           # 抽出 <script> 写 .js
node --check _dash_check.js     # 语法校验
```

### 3. 逻辑复算（最有价值的一步）
写一个独立 node 脚本，从 HTML 里正则抽出 `RAW_DATA` 与判定函数，eval 后复算 KPI，
再和页面应显示的值比对：
```js
const ruleShort = html.match(/const isShort = .*/)[0];
const isShort = eval(ruleShort.replace('const isShort =', ''));
console.log(MATERIALS.filter(isShort).length);
```
比肉眼看页面可靠得多。

## 数据从哪来：页面无法直连数据源

**关键认知**：浏览器里的 HTML **不能**直接连企业系统或在线表格——
没有登录态、跨域会被拦、把 token 写进页面也不安全。
正确做法是**数据在外部生成，再喂给页面**：

```
源系统（金山文档 / 内部系统 / API）
    ↓  拉取（连接器 / 脚本）
本地 CSV / JSON
    ↓  生成脚本做计算与归一
workbench-data.js   （window.RAW_DATA = {...}）
    ↓  <script src="workbench-data.js">
HTML 页面（只负责显示）
```

### 页面要写成"数据可替换"的结构

```html
<script src="workbench-data.js"></script>
<script>
const DEMO_DATA = { /* 内置演示数据，保证单文件也能跑 */ };
const RAW_DATA = window.RAW_DATA || DEMO_DATA;
const IS_REAL_DATA = RAW_DATA !== DEMO_DATA;   // 界面上标出用的是真实还是演示数据
</script>
```
好处：数据文件不存在时自动回落演示数据，不会白屏；界面上能一眼看出数据来源。

### 同时产出「内联单文件版」

外链脚本在**预览面板、转发、只服务单文件的环境**里可能 404。
生成器应再输出一份把数据内联进去的单文件 HTML：

```python
html = html.replace('<script src="workbench-data.js"></script>',
                    '<script>window.RAW_DATA = {...};</script>')
```
给用户两个文件：外链版（日常更新用）+ 内联版（演示/转发用）。

## 生成脚本的三个必做校验

### 1. 自洽性：合计值必须等于趋势末值
趋势通常由"当前库存倒推"，**倒推必须只用入选物料的流水**，
否则趋势与库存合计对不上（踩过一次：差 2.5 万）。

```python
included = {m['spec'] for m in materials}       # 入选物料
f_in  = {d: sum of flows where std in included}  # 只统计入选者的流水
stock(t) = stock_now - Σ_{x>t} 入库 + Σ_{x>t} 出库
```

### 2. 预警比例要落在合理区间
算完立刻自检 `(短缺+积压)/总数`，**超过 30% 就报警**——
通常不是数据问题，而是阈值拍脑袋定错了。

### 3. 阈值要给真实分布做依据，不要拍脑袋
下阈值前先看数据分布。踩过的坑：台账跨两年，用「90 天无出库」判积压，
命中 24 项；改「180 天」后 15 项，显然更站得住。
把阈值提成页面顶部常量，便于用户按管理口径调：

```js
const OVER_DAYS = 180, COVER_RATIO = 0.3;
const isShort = m => m.stock < m.safe || (m.out7 > 0 && m.stock < m.out7 * COVER_RATIO);
const isOver  = m => m.lastOut > OVER_DAYS && m.stock > 0;
```

## 源数据没有的字段，不要假装有

典型：**安全库存**。它是管理参数，台账里通常没有。
处理原则：
1. 先查源表里**是否已预留该列**（本次发现金山表里其实已经有「安全库存」列，只是空着）
   → 有就让用户在那里维护，生成器优先读它
2. 没有则生成一份模板 CSV，给**保守**的建议值，并明确标注"需人工确认"
3. 建议值宁小勿大（本次用 `近7日出库 × 0.5`，用 `×2` 时一半物料被误判短缺）

## 升级为 PWA

工作台加上 PWA 后可以装到桌面、离线打开，体验接近原生应用。四样东西：

```
index.html              页面（入口，放仓库根目录，GitHub Pages 直接吃）
manifest.webmanifest    应用清单
sw.js                   Service Worker（缓存应用外壳）
icons/                  图标
```

### manifest 要点

```json
{
  "start_url": "./index.html",
  "scope": "./",
  "display": "standalone",
  "theme_color": "#2563eb",
  "icons": [
    {"src":"icons/icon.svg","sizes":"any","type":"image/svg+xml"},
    {"src":"icons/icon-192.png","sizes":"192x192","type":"image/png"},
    {"src":"icons/icon-512.png","sizes":"512x512","type":"image/png"},
    {"src":"icons/icon-512-maskable.png","sizes":"512x512","type":"image/png","purpose":"maskable"}
  ]
}
```
`start_url` / `scope` 用**相对路径**，这样放在仓库子路径（`user.github.io/repo/`）下也对。
**必须给 192 和 512 两种尺寸的 PNG**，只有 SVG 有些浏览器不认，装不上。
maskable 版要比普通版多留 14% 内边距，否则被系统裁成圆角时会切到内容。

### sw.js 要点

- 页面导航用**网络优先、断网回落缓存**；其它静态资源缓存优先。
- **数据文件不要缓存**（`if (url.pathname.endsWith('workbench-data.js')) return;`），
  否则用户永远拿到旧数据。
- 改动页面后**记得把版本号加一**，不然访客看到的是旧缓存。
- 只在 `http(s)` 下注册（`file://` 会报错，要判断 `location.protocol`）。

### 图标用 PIL 画（无字体依赖）

不要依赖中文字体渲染（环境里常常没有）。用几何图形：圆角方块 + 三条白色看板条。
渐变用「逐行 `putpixel` 生成 1×N 图再 resize」，比叠两层色块干净——
叠色块会在交界处留下可见接缝。

## 数据不进网络：本地文件同步

如果数据敏感、不能放到网上，但又想要「点一下按钮就刷新数据」，
可以**让用户选一次本机数据文件**，页面记住它：

```js
// 首次：文件选择器 + 把句柄存进 IndexedDB
const [h] = await window.showOpenFilePicker({ types:[...], multiple:false });
await idbSet('dataFile', h);
// 之后：直接重读同一文件，不再弹框
const perm = await h.queryPermission({mode:'read'});
if (perm !== 'granted') await h.requestPermission({mode:'read'});
const f = await h.getFile();
applyData(parse(await f.text()), 'local', f.name);
```

要点：
- 用 `File System Access API`（Chrome/Edge 支持），**回退**到 `<input type="file">`。
- 句柄存 IndexedDB（句柄不能存 localStorage）。
- 数据只进内存，**不上传、不联网**。页面上要显著标出来，用户才敢用。
- 顶栏显示「快照日期 + 同步时间」，这是用户判断数据新旧的唯一依据。

### 数据要能被运行时替换

原始写法 `const RAW_DATA = ...` 是冻结的，换数据无法生效。改成：

```js
let RAW_DATA = window.RAW_DATA || DEMO_DATA;
let MATERIALS = [];
function buildMaterials(){ /* 由 RAW_DATA 重建索引 */ }
function refreshUI(){ /* 重画所有依赖数据的视图 */ }
```

### 页面解析数据文件要能吃两种格式

生成器同时输出 `.js`（`window.RAW_DATA = {...}`，供 `<script src>`）
和 `.json`（纯 JSON，供文件选择器读）。解析时先试 `JSON.parse`，
失败再按 JS 对象字面量求值。

## 托管：为什么不能只发 Gist

**GitHub Gist 不能当网站用**：GitHub 故意把 gist 的 raw 文件以 `text/plain` 返回，
浏览器只显示源码不渲染页面（防止被拿去做钓鱼）。真 PWA 还需要 manifest 与
service worker，Gist 没有目录结构，满足不了。

用 **GitHub Pages**：

```bash
gh api -X POST repos/<user>/<repo>/pages -f source[branch]=main -f source[path]=/
```

放根目录的 `index.html` 会被直接作为首页。部署后务必实测各资源的 **MIME 类型**——
`sw.js` 必须是 `application/javascript`，`manifest` 要是
`application/manifest+json`，错了 PWA 就装不上：

```python
urllib.request.urlopen(url).headers.get('Content-Type')
```

## 让用户把表格直接拖进来

页面自己解析 xlsx 是最省事的数据接法。**不要因此引入 SheetJS**：单文件离线页面背几百 KB
第三方库不划算，走 CDN 又断网就废。xlsx 本质是 ZIP，现代浏览器有原生解压：

```js
// 解 ZIP：EOCD → 中央目录 → 本地头 → DecompressionStream('deflate-raw')
const ds = new DecompressionStream('deflate-raw');
const out = new Uint8Array(await new Response(new Blob([raw]).stream().pipeThrough(ds)).arrayBuffer());
```

需要处理的：`xl/workbook.xml`（表名）、`xl/_rels/workbook.xml.rels`（表路径）、
`xl/sharedStrings.xml`、`xl/styles.xml`（判断哪些单元格是日期）、sheet XML 里的
`t="s"|"str"|"inlineStr"|"b"` 与 `<v>` / `<is><t>`。

**验证方式**：与 `openpyxl` 交叉比对同一批文件的工作表名、行列数、逐格取值。
本项目这样测过 3 个真实文件、全部工作表，前 8 行逐格一致才敢用。

**别只看扩展名判断格式**：集成商经常改错后缀。用魔数 —— `PK\x03\x04` 是 ZIP（xlsx），
`\xD0\xCF\x11\xE0` 是老的 .xls（浏览器读不了，要明确告诉他另存为 xlsx）。

## ⚠️ 两套实现必须共用一份规则

页面能算、Python 也能算时，**规则只能有一个来源**。让 Python 生成规则文件
（归一表、分类关键词、安全库存、阈值）给页面读，别在 JS 里再抄一遍——
抄一遍就等着两条路径算出不同结果。

生成物要出**两种格式**：`.js`（`window.XXX = {...}`，页面用 `<script src>` 引）
和 `.json`。**只给 .json 不行** —— 本地 `file://` 打开时 `fetch` 会被跨域策略拦掉。

## ⚠️ 无年份日期跨年：唯一正确的解法是按行序推断

台账里极常见「Excel 序列号 + 无年份中文日期」混排，而中文那段**往往跨年**：

```
… 8月7号 … 12月24号 │ 1月5号 … 9月2号
      上一年          │     下一年
```

**统一补任何一年都会错**。实测影响：一笔 1.26 万米的入库被漏算/多算。

```js
// 按行序：月份变小即进一年（前提是台账按时间排，实际都成立）
if (prevMonth !== null && md.m < prevMonth) curYear += 1;

// ⚠️ 绝对日期（序列号/带年 ISO）不要更新 curYear/prevMonth
//    否则夹在中间的一条会把年份又拉回去，回绕判断永远不触发
```

起手年份取「第一个无年份值**之前最近的**绝对日期」，比取全列最大年份更准。

## 快照日：文件名优先于流水最大值

`XX库存(9月8日).xlsx` 这种命名本身就标明了快照时点。
若拿「流水里的最大日期」当快照，一旦有人误录了半年后的日期，
整个 7 日窗口就偏了（本项目真出现过：文件是 9月1日的，流水里混进 2026-12-29）。

顺序：用户指定 > **文件名里的日期** > 流水最大值。

## DOM 层怎么测：jsdom + 桩

没有 playwright 时，jsdom 也能做真实的端到端验证：

```js
// 必须注入的替代品
window.indexedDB = makeInMemoryIDB();        // jsdom 不带
window.DecompressionStream = DecompressionStream;  // 用 Node 原生的
window.Response = Response; window.Blob = Blob;
window.structuredClone = structuredClone;
window.fetch = async () => { throw new Error('测试环境不联网'); };
```

然后用真实文件构造 `File` 调 `addFiles()`，走完整链路。
**这能测出静态检查完全看不见的 bug**（本项目一次测出 4 个：日期没转 ISO、
集成商识别不到、来源标签漏了一个分支、连带的天数算出 null）。

> 注意：`let` / `const` 声明的顶层变量**不会**挂到 `window` 上。
> 测试里访问 `window.MATERIALS` 得到 `undefined` 是正常的，要读 DOM 文本验证。

## 让网页按钮触发本机脚本（本机同步代理）

用户常会问：「网页上点一下，让它去跑脚本更新数据」。**这件事网页做不到** ——
浏览器不允许网页执行本机程序，这是硬边界，自定义协议也救不了
（页面拿不到执行结果、做不了进度）。正确做法是在本机放一个**只监听 127.0.0.1 的小服务**：

```
网页点「同步」 → POST /api/sync → 本机服务依次跑脚本 → 轮询 /api/status 看进度 → 取回结果
```

### 最小接口集

| 接口 | 作用 |
|---|---|
| `GET /` | 把项目当静态站点服务 —— 打开 `http://localhost:8765/` 就是**同源**的完整应用，零配置 |
| `GET /api/status` | 运行状态 + 每步进度 + 数据摘要（前端轮询这个） |
| `POST /api/sync` | 触发同步；已在跑返回 409；立刻返回后由前端轮询 |
| `GET /api/data` | 返回最新生成的数据文件 |
| `GET /api/config` | 步骤清单（**不回传本机路径**） |

用 `ThreadingHTTPServer` + 一个后台线程跑步骤即可，不必引框架。
每步用 `subprocess.run(capture_output=True)`，只保留输出**尾部**（20 行左右）回传，
否则失败日志会撑爆响应。

### 顺手解决的两个体验问题

1. **页面按相对路径找的生成文件**（如 `workbench-data.js`）实际在 `output/` 下 →
   服务端做个 `FILE_MAP` 把 `/workbench-data.js` 映射过去，页面不用改。
2. **配置里的 `python` 换成本进程解释器**（`sys.executable`），
   否则用户环境里 PATH 没有 `python` 就起不来。

### 安全（缺一不可，别图省事）

```python
# ① 只绑本机，局域网与公网都访问不到
ThreadingHTTPServer(('127.0.0.1', port), Handler)

# ② 校验 Host 头，挡 DNS rebinding
host = (headers.get('Host') or '').split(':')[0].lower()
if host not in ('localhost', '127.0.0.1', '::1'): 403

# ③ CORS 只放行白名单，绝不要用 '*' —— 否则任意网站都能读走你的数据
if origin in ALLOW_ORIGINS: ACAO = origin

# ④ Chrome 的 Private Network Access：公网页面访问本机必须显式同意，
#    不加这个头，跨源请求会被浏览器直接掐掉（且报错信息很不直观）
'Access-Control-Allow-Private-Network: true'

# ⑤ 静态文件服务防目录穿越
target = os.path.normpath(os.path.join(ROOT, rel))
if not target.startswith(os.path.normpath(ROOT)): 403
```

> `http://localhost` / `http://127.0.0.1` 属于「potentially trustworthy origin」，
> 所以 **HTTPS 页面访问它是允许的**，不会被当混合内容拦掉。

### 前端侧

- 启动时探测一次（短超时 2.5s），失败就显示**启动命令**，不要静默失败
- 同步中**锁定按钮**并显示逐步进度，别让用户以为卡死
- 失败时把**失败那一步的输出尾部**显示出来 —— 用户要的是「哪一步错了」，不是「失败了」
- 生成物（如规则文件）变了要提示**刷新页面**，因为它们是 `<script src>` 一次性加载的
- 顶栏按钮做**降级**：代理不在就退回「拖拽/选文件」，两条路都要能走通

## 常见坑

| 坑 | 处理 |
|---|---|
| 隐藏容器的 `clientWidth` 为 0 | 图表函数兜底默认宽度，并在 `switchView` 里重绘 |
| 跨单位求和（米/个/台混加） | 明确标注"跨单位合计，仅作趋势参考"，别叫"库存总量" |
| 日期是混合格式且无年份 | ISO + 中文"8月6号"混排时，按行序做年份推断（见 docx-form-fill-anonymize 同类问题） |
| 依赖外部字体/图标 | 用系统字体栈 + Unicode 字符（◧ ▽ △ ▤ ◔ ✎ ✉ ⚙），不用字体图标库 |
| 深色主题 | 默认按当前 IDE 主题；本用户长期偏好"护眼暗色（底色非纯黑、正文非纯白）"，需暗色时按此调 |

## 顺手做「一键流程演示」

如果成果本身是一条自动化流程（登录→取数→分析→成报→推送），
加个按钮做 5 步进度动画，完成后渲染结果——**这既是演示，也能直接截图当佐证**。

```js
const FLOW_STEPS = [{k:"login",t:"登录平台"}, ...];
// 每 520ms 走一步，class 在 "" / "run" / "done" 之间切换
```

## 脱敏前移

如果这个界面将来要截图用于报奖材料，**设计时就别放**单位名称、地名、人名：
收件对象写「项目经理（项目名）」而不是具体姓名。这样截图天然合规，不用事后打码。
