---
name: docx-form-fill-anonymize
description: 填写 Word 表单类文档（申报书、业务联系单、会议纪要、汇报模板等）并保留原版式，同时做脱敏校验（去单位名称/个人姓名/联系方式）。当用户给出一个 Word 模板要求"按这份材料填写""照着模板填""不要出现单位名/人名"时使用。触发词：填写申报书、填表、按模板填写、脱敏、去掉单位名称、去人名、不得出现单位名称。
agent_created: true
---

# Word 表单填写 + 脱敏

两件事：**把内容填进模板且版式不变**，以及**确保敏感信息不出现**。

## 一、填写：保留原格式的单元格/段落重建法

**核心思路**：不要新建段落（会丢格式），而是克隆模板段落 + 只换文本。

```python
import copy
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_para_text(p_elem, text):
    """保留首个 run（含其 w:rPr）的格式，只替换文本。"""
    runs = p_elem.findall(qn('w:r'))
    if not runs:
        r = OxmlElement('w:r'); p_elem.append(r); runs = [r]
    first = runs[0]
    for r in runs[1:]:
        p_elem.remove(r)
    for t in first.findall(qn('w:t')):
        first.remove(t)
    for br in first.findall(qn('w:br')):
        first.remove(br)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    first.append(t)
```

单元格重建：

```python
def fill_cell(cell, sections):
    """sections: [{'label': '标题文字', 'body': ['段落1', '段落2', ...]}, ...]"""
    tc = cell._tc
    ps = tc.findall(qn('w:p'))
    label_tmpl = copy.deepcopy(ps[0])           # 标题段（原粗体）
    body_tmpl = copy.deepcopy(ps[1]) if len(ps) > 1 else copy.deepcopy(ps[0])
    # ⚠️ 必须剥掉正文模板的粗体，否则正文跟着标题变粗
    for r in body_tmpl.findall(qn('w:r')):
        rpr = r.find(qn('w:rPr'))
        if rpr is not None:
            for tag in ('w:b', 'w:bCs'):
                for e in rpr.findall(qn(tag)):
                    rpr.remove(e)
    for p in ps:
        tc.remove(p)                            # 只删 w:p，保留 tcPr
    for sec in sections:
        lp = copy.deepcopy(label_tmpl); set_para_text(lp, sec['label']); tc.append(lp)
        for txt in sec['body']:
            bp = copy.deepcopy(body_tmpl); set_para_text(bp, txt); tc.append(bp)
```

要点：
- `tc` 下只能删 `w:p`，**`tcPr` 必须留着**，否则单元格宽度/边框丢失。
- 单元格末尾必须是 `w:p`，重建时至少 append 一个段。
- 一个格子里有多组"标题+正文"（如同时含"推广性描述"和"创新性描述"）时，
  `sections` 传多项即可，第二项的 label 也从原第一个 label 克隆。

### 先探结构再动手
```python
from docx import Document
d = Document(PATH)
for i, p in enumerate(d.paragraphs):
    print(i, repr(p.text), p.style.name, [r.font.name for r in p.runs])
t = d.tables[0]
for ri, row in enumerate(t.rows):
    for ci, c in enumerate(row.cells):
        for pi, p in enumerate(c.paragraphs):
            print(f'R{ri}C{ci} p{pi}', repr(p.text[:80]),
                  p.runs[0].font.bold if p.runs else None)
```
看清"哪个是标题段、哪个是说明段、说明段有几段、标题段是第几段"再写填充逻辑。
多组标题的格子（标题不在 p0 后的固定位置）必须按实际序号处理，不能假设只有一组。

## 二、脱敏：硬约束清单

> ⚠️ **写技能/文档/示例代码时，绝不要用真实地名、单位名、人名。**
> 本技能初版就犯过这个错——把脱敏对照表写成了「＜真实地市＞分公司 → 本单位」，
> 结果技能文档本身成了泄露源，提交到代码仓库时被扫描拦下。
> 示例一律用占位符：＜地市＞＜上级单位＞＜姓名1＞＜部门名＞。

填完**必须跑一遍校验**，用代码扫，不靠肉眼：

```python
d2 = Document(OUT)
alltext = '\n'.join(p.text for p in d2.paragraphs)
for tb in d2.tables:
    for row in tb.rows:
        for c in row.cells:
            alltext += '\n' + c.text
banned = ['<申报单位名>', '<地市名>', '<上级单位名>', '<人名1>', '<人名2>', '<部门名>',
          '邮箱', '电话', '手机', '@', '联系']
hits = [b for b in banned if b in alltext]
```

常见处理：

| 原表述 | 改成 |
|---|---|
| ＜地市＞分公司＜部门＞ | 本单位 / 项目组 / 工程建设一线 |
| ＜上级单位＞公司 | （去掉，或写"本单位"） |
| ＜姓名1＞、＜姓名2＞ | 项目组成员 / 团队成员 |
| 封面"单位名称""单位公章" | **留空**，不要填 |
| 匹配项目经理邮箱 | 匹配对应项目经理（连"邮箱"都规避，防被误读为联系方式） |
| 联系人+手机号 | 整段删除 |

注意：
- 封面上的集团公司名称是模板自身的印制单位，不是申报单位，保留。
- 模板自带的"注意：申报书内不得出现…"提示语可以保留，交付时提醒用户是否删除。
- 技术产品名（如具体云平台、大模型、系统名）**不属于**单位名称/人名，可保留且通常必要。

## 三、交付前必做

1. **另存为新文件**，不要覆盖原模板（用户下次还要用）。
2. 校验字数门槛（申报书类常要求"不少于 500 字"，且 500 是硬线）。
   用 `sum(len(x) for x in section_paragraphs)` 统计，注意含标点。
3. 扫描禁用词，输出报告，`present_files` 交付。

## 四、附表：证明材料汇编（配套申报书）

申报书常需配一份《成果证明材料》。若用户要"汇编"，按下面九项结构做，**每项独立成页**：

| 材料 | 内容 | 数据来源 |
|---|---|---|
| 封面 | 成果名称 / 材料份数 / 编制日期（单位名称、公章**留空**） | — |
| 材料清单 | 9 行索引表（序号/名称/主要内容/份数页数） | — |
| 一、成果基本情况 | 成果名称、研发起止时间、成果类别、技术路线、交付形态、应用范围、解决的问题、主要亮点 | 计划书 |
| 二、研发过程与阶段记录 | 阶段 / 起止时间 / 主要工作 / 阶段成果 / 佐证截图号 | 计划书的阶段划分 |
| 三、功能模块与实现清单 | 序号 / 模块 / 主要功能 / 实现方式 / 佐证截图号 | 计划书的任务清单 |
| 四、系统运行记录 | 序号 / 执行日期 / 触发方式 / 执行结果 / 提取数据量 / 佐证 | **留【待填】** |
| 五、日报推送记录 | 序号 / 推送日期 / 推送对象（**按岗位，不写人名**） / 内容构成 / 佐证 | **留【待填】** |
| 六、交付物清单 | 序号 / 交付物名称 / 形态 / 说明 | 计划书+申报书 |
| 七、效益测算依据 | 效益项 / 测算口径 / 年化估算 / 说明 | 从申报书效益段展开 |
| 八、应用情况说明 | 应用场景、应用范围、当前状态、使用与推广、合规与安全（5 小节） | 申报书 |
| 九、佐证截图 | 5 个截图占位框 | 用户自贴 |
| 附、材料真实性说明 | 声明正文 + 签字/盖章/日期下划线 | — |

关键做法：
- **能推导的填实，需真实记录的留【待填】**。运行日期、推送记录这类绝不能编，
  给表格结构 + `【待填】` 占位 + 一行填写建议（如"建议连续取 5 个执行日覆盖两种触发方式"）。
- 截图占位用一个 **1×1 表格**，内含三行（标题 / `（此处粘贴截图）` / 说明），
  并用 `w:trHeight`（`hRule="atLeast"`）撑高到 2400 twips，给人留出粘贴空间。
- 封面键值表用两列无边框表格（`w:tblBorders` 六条边设 `val="none"`）。
- 佐证编号要和正文表里的「佐证」列对得上（截图①~⑤）。

### 脱敏校验要加白名单
证明材料里常保留一句合规提醒，如
"截图内容同样不得出现申报单位名称、个人姓名与联系方式等信息"——
这句含「联系」，会被禁用词扫描误报。校验时先把这类**规则表述句**从待检文本里剔除：

```python
SAFE = ['不得出现申报单位名称、个人姓名与联系方式等信息']
probe = alltext
for s in SAFE:
    probe = probe.replace(s, '')
hits = [b for b in banned if b in probe]
```

## 五、生成 docx 时的两个必查项

### 1. 别在纯文本 run 里写 markdown 的 `**`
python-docx 写的是纯文本，`**加粗**` 会**原样渲染成星号**。要么去掉标记，要么自己解析：

```python
def rich(p, text, size, name=FS, color=None, base_bold=False):
    """支持 **加粗** 的行内标记。"""
    for i, seg in enumerate(str(text).split('**')):
        if seg == '':
            continue
        setf(p.add_run(seg), name=name, size=size,
             bold=(base_bold or i % 2 == 1), color=color)
```

**所有文案出口统一走 rich()**（标题 / 正文 / 警示 / 表格单元格），
只改其中几处必然漏。交付前全文档扫一遍 `**`，计数必须为 0：

```python
bad = sum('**' in p.text for p in d.paragraphs)
for tb in d.tables:
    for r in tb.rows:
        for c in r.cells: bad += '**' in c.text
assert bad == 0
```

### 2. 清单类文档
用 `□`（U+25A1）做勾选框、`⚠`（U+26A0）做风险标记，Word 里显示正常，
比用图片打勾或符号字体省事。

## 六、写「跑通/检查清单」类文档的要点

当用户要的不是成品文档，而是**指导性清单**（"我该怎么做""怎么验证"）时：

- **每一步都要给「跑通的定义」**——不含糊地说"能连续 5 个工作日无人干预成功"，
  而不是"实现自动登录"。可验证的定义才有用。
- **区分「该做什么」和「拍什么」**：每个环节末尾单列一节「补证：截图X」，
  写明拍什么 / 证据强度 / 脱敏要点。
- **主动指出风险，不要只写步骤**。例如"脚本模拟登录公司系统可能违反信息安全规定，
  先确认合规性再开发"——这类前置风险提示比多写十步操作更有价值。
- **时间倒排**：如果下游需要"连续多日记录"，必须提醒它决定了系统跑通后还要留试运行期。
- **区分内部文档与交付文档的口径**：内部清单可不脱敏（方便联调），
  但其中产出的截图用于交付时必须脱敏——在文档开头用一节讲清，别让用户混用。

## 七、字数与效益表述

- 表单写"请勿过度包装/如发现虚假将退回"时，效益**不要编精确数字**。
  用"按每工作日约X小时估算""实际以运行台账统计为准"这类可辩护表述。
- 效益分维度写全：效率、成本、管理、社会（表单通常要求覆盖经济/社会/生态）。

## 八、工具通道与降级

本地 docx 优先走 `tencent-docs-routing` → `tencent-local-office-edit`（editor_sdk）。
**已知故障**：`open_file` 返回 "open started"、`get_pool_status` 能看到实例，
但 `doc_get_content_region` / `doc_resolve_document_structure` / `doc_get_outline`
报 `DocEditor::...: document is not open`，`present_files` 也注册不出 UUID 实例。
此时按路由规则的降级阶梯改用 python-docx 直接读写，**并明确告知用户已降级**。

### 本机环境备忘（Windows / qazws）
- Bash 工具的 PATH 在本机常损坏（`dirname`/`mkdir`/`cat`/`head` 全部 not found）→ 一律用
  `C:\Users\qazws\.workbuddy\binaries\python\envs\default\Scripts\python.exe` 执行脚本。
- PowerShell 工具在本机可能吞掉 stdout（只回 "Command completed with exit code 0"）→
  **让脚本把结果写进 UTF-8 文本文件，再用 Read 工具读**，不要依赖 stdout。
- 重定向到文件时若出现 `Cannot display content of binary file`，说明被写成了 UTF-16；
  改用 Python `open(path, 'w', encoding='utf-8')` 显式写入。

### 视觉校验的替代方案
`doc_to_image` 在本机**不可用**（报 `COS upload failed: can't open file
'F:\Program Files\Tencent\Marvis\...\mcp\doc\upload_to_cos.py'`，editor_sdk 自带脚本缺失），
无法渲染成图看效果。改用**文档流结构校验**代替：

```python
for child in d.element.body.iterchildren():
    if child.tag == qn('w:p'):
        p = Paragraph(child, d)
        if 'type="page"' in child.xml: print('--- 分页 ---')
        elif p.text.strip(): print(p.text[:80])
    elif child.tag == qn('w:tbl'):
        print(f'《表 {len(Table(child,d).rows)} 行》')
```
确认章节顺序、分页位置、表格行列数是否符合设计即可。
注意：`doc_to_image` 传中文路径报 `file not found`，但换 ASCII 路径**仍然失败**——
根因是脚本缺失而非路径编码，别被这个报错误导。
