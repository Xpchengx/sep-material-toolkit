---
name: publish-repo-to-github
description: 把本地工作区上传到 GitHub 并建立后续更新流程。重点是「只上传代码、绝不外泄业务数据」——含 .gitignore 拦截、公开前脱敏、以及清除 git 历史中脱敏前原文的关键手法。当用户说"上传到 GitHub""传到我的仓库""以后在 GitHub 上更新项目""发布到代码仓库"时使用。触发词：上传 GitHub、推到 GitHub、建仓库、发布到 GitHub、git push、开源这个项目。
agent_created: true
---

# 把工作区上传到 GitHub

**先想清楚一件事：这个工作区里有没有不该外泄的东西？**
绝大多数真实工作区都有。默认答案是「有」，然后逐个排掉。

## 一、先盘点，再动手

```bash
# 列出全部文件（含大小），人工过一遍
```

**必须拦下来的东西**（真实案例里全中过）：

| 类型 | 例子 | 为什么危险 |
|---|---|---|
| 业务台账 | `*.csv` 明细、月报 `.xlsx` | 真实的库存/交易数据 |
| 交付物 | 生成的 `.docx` / `.html` | 里面往往带单位名、人名、金额 |
| **盖章件** | 扫描版合同、确认单、`.zip` | 盖章即法律效力，最敏感 |
| 申报/评审材料 | 申报书、填报内容 | 含单位名称与联系人 |
| 本地记忆 | `.workbuddy/memory/`、`.claude/` | 记忆文件本身就含地名与人名 |
| 一次性脚本 | `diag_*.py`、`probe_*.py` | 常硬编码真实路径与单位名 |

### .gitignore 要点

```
# 数据
material-system/out/
*.csv
*.xlsx
# 交付物
output/*.docx
output/*.html
output/站址确认单-分区县/
# 本地记忆与运行时
.workbuddy/
# 根目录临时脚本
/_*
# 一次性脚本目录
*/scratch/
```

加完 `gitignore` 后**必须复查暂存区**，确认数据类文件数为 0：
```bash
git add -A && git status --short   # 人工看一遍
```

## 二、⚠️ 最容易踩的坑：暂存区清不干净

`git rm -r --cached .` 有时**报 exit=0 但实际没清空**，
结果是：被 ignore 的文件仍留在暂存区，而且 `git check-ignore` 会因
「文件已被跟踪」而返回「未忽略」，让人误以为规则写错了。

**用 `git reset` 代替**：
```bash
git reset && git add -A && git status --short
```

## 三、公开仓库：先脱敏

私有仓库可以宽松，**公开仓库要严**——会被搜索引擎和代码镜像站长期抓取，
删库也清不干净缓存。

脱敏口径不能只看「地名/人名」。真正的问题是**组合指纹**：

| 类别 | 例子（占位符写法） |
|---|---|
| 内部系统名 | 内部平台简称 → 物资管理平台 |
| 内部文档/模块名 | XX材料管理系统 → 材料台账工作簿 |
| 供应商名 | 具体厂商 → 集成商A / 集成商B |
| 雇主名 | 具体运营商 → 运营商 |
| 专有名词 | 内部叫法 → 行业通用叫法 |

> ⚠️ **这张表本身也必须用占位符。** 写真实名字当示例，等于把要脱敏的东西
> 又写回了技能文档 —— 而技能文档要进仓库，扫描会当场拦下。
> **这个错我犯了三次。** 每次都是"举例说明"时顺手用了真名。
> 所以：写脱敏文档时，**边写边用占位符**，不要写完再回头扫。

**单项是行业通用词，组合起来能精确定位到具体单位的具体项目。**

保留原则：**功能必需且行业通用的词要留**（如设备品类 RRU/BBU/光模块——
它们是分类规则的关键字，删了代码就坏；且不具识别性）。
替换脚本按「长串优先」排序，否则短串先替换会破坏长串。

### ⚠️ 脱敏后必须 amend，不能新增 commit

**如果脱敏做成第二个 commit，git 历史里会保留脱敏前的原文**，
任何人 `git log -p` 就能翻出来，脱敏等于白做。

```bash
git add -A
git commit --amend -q -m "初始提交信息"
git reflog expire --expire=now --all
git gc --prune=now --aggressive
```

验证（三项都必须过）：
```bash
git rev-list --count HEAD          # 应为 1
git fsck --unreachable             # 应为 0 个不可达对象
# 遍历所有可达对象扫描敏感词
git rev-list --objects --all | while read sha _; do git cat-file -p "$sha"; done
```

## 四、凭据：不要让用户经过「我发码 → 等用户」

**踩过两次**：gh 的设备码只有约 15 分钟有效期，
我发码后要等用户回消息，而异步沟通中用户可能几十分钟后才看到，码早过期
（报 `context deadline exceeded`）。

**正确做法：让用户自己在终端跑交互式命令，节奏由他定。**

```bat
"C:\Users\qazws\.workbuddy\binaries\gh\bin\gh.exe" auth login
```
提示依次选：GitHub.com → HTTPS → Yes（配置 git 凭据）→ Login with a web browser。
它会自动开浏览器，用户粘个码即可，**没有时间压力**。

之后建库+推送只需几秒，不再涉及有时效的中间态。

> 通用原则：凡「有时效的一次性凭据」，不要让用户经过
> 「我 → 用户 → 我」的往返，应让用户在本地一次性完成。

### 免管理员安装 gh

不用 winget（要提权），直接下 release 的 zip 解压即用：
```
https://api.github.com/repos/cli/cli/releases/latest  → 找 *_windows_amd64.zip
```
解压到 `~/.workbuddy/binaries/gh/`，可执行文件在 `bin/gh.exe`。

## 五、推送

```bash
gh repo create <name> --public --source=. --remote=origin --push \
   --description "..."
```
仓库名不要含单位、地点、供应商信息。

## 六、推送后核验远端（不能只看推送成功）

```bash
git ls-remote --heads origin          # 远端 SHA
git rev-parse HEAD                    # 本地 SHA，两者必须一致
git show origin/main:<path>           # 逐个把远端内容拉回来扫敏感词
```

### ⚠️ `git rev-list --all` 含远端跟踪引用，推送前扫必然误报

扫描"全历史"时报仍有命中，但工作区明明已经清干净 —— 十有八九是这个原因：
**`--all` 包含 `origin/main`**，而它还指着 amend / 重建之前的提交。

**正确顺序：先提交 → 再推送 → 然后扫 `--all`。**
只想确认"本次要推的内容干净"，扫 `HEAD` 就够：

```bash
git rev-list --objects HEAD      # 本次要推的
git rev-list --objects --all     # 含远端跟踪引用，推送后才准
```

### 另外两个 git 的坑

| 现象 | 原因 | 处理 |
|---|---|---|
| 扫描脚本漏掉中文名文件、静默跳过 | `git ls-files` 对非 ASCII 路径做**八进制转义**，`os.path.join` 找不到文件 | 用 `git -c core.quotePath=false ls-files -z` |
| ignore 规则明明写了却不生效 | `git rm -r --cached .` 可能报 exit=0 但没真清空 | 改用 `git reset` |

> **扫描脚本每跳过一个文件，都必须在报告里显式列出来。**
> 静默的 `except: continue` 会让"通过"变成假象 —— 本次就因此漏掉了主犯文件，
> 白跑了两轮才定位到。

## 六点五、⚠️ 同步文件进仓库：必须用白名单，不能用通配

**踩过**：写了个「把技能同步进仓库」的脚本，用了
`for name in os.listdir(SKILLS)` —— 结果把 `~/.workbuddy/skills/` 下
**全部 8 个技能**推进了公开仓库，其中 6 个是第三方技能包
（含各自的 LICENSE、几十个 XSD schema），131 个不相关文件。

问题不只是多余：**第三方技能包有自己的许可，把别人的包重新发布到你的公开仓库是许可问题**。

**正确做法**：

```python
KEEP = {'skill-a', 'skill-b', 'skill-c'}      # 白名单，明确列出
for name in sorted(os.listdir(SKILLS)):
    if name not in KEEP:
        continue
```

并且**每个提交前断言文件数**，数量对不上就中止：

```python
assert len(staged) <= EXPECTED_MAX, f'待提交 {len(staged)} 个，超出预期，中止'
```

### 误推之后怎么救

1. **普通 `git revert` 或补一个删除提交，历史里仍然留着那些文件**——
   `git log -p` / `git rev-list --objects --all` 都能翻出来。
2. 要彻底清除必须**重写历史**：
   ```bash
   git reset --mixed <误提交之前的那个提交>
   git add -A && git commit -m "..."
   git push --force-with-lease origin main
   git reflog expire --expire=now --all && git gc --prune=now -q
   ```
3. 验证：遍历全部可达对象，确认被误推的路径/文件名**零命中**，
   且 `git fsck --unreachable` 为 0。
4. `--force-with-lease` 比 `--force` 安全：远端被别人动过时会失败而不是覆盖。

> 判断标准：**误推的是「不该发布的内容」还是「敏感内容」**。
> 前者重写历史即可；后者要按泄露处理，因为可能已被抓取，
> 还需考虑吊销密钥、通知相关方。

## 七、本机注意事项

- **Bash 工具的 PATH 常损坏**（`dirname`/`cat`/`head` 全 not found），
  一律用 Python subprocess 执行命令。
- **PowerShell 工具常吞 stdout**（只回 exit code 0）→ 让脚本写 UTF-8 文件再读。
- **调用 gh 要用参数列表，不要走 shell 字符串**：
  ```
  subprocess.run([GH, 'api', 'repos/o/r/git/trees/main?recursive=1'])
  ```
  走 shell 时 `?` 与内层引号会被 cmd.exe 吃掉，导致输出为空、误判为「无文件」。
- Python 里拼 f-string 调外部命令容易写出语法错误（`.0f` 误用在非数字上），
  写完先 `py_compile` 过一遍再跑。
