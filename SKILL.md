---
name: lab-report-pipeline
description: >-
  课程实验（机器感知、机器学习、图像处理、操作系统、软件工程等）从读要求到交 Word 报告：
  收模板·指导书·示例代码，按步骤留中间产物，图表与代码截图由脚本生成，
  数字只从日志/metrics 取，最后套模板渲染 docx 并自动核对交付。
  触发：做实验、写实验报告、实验指导书、实验报告模板、复现实验、跑通实验、
  实验截图、实验一/实验二/实验三、中间产物留档、绘制实验图、补实验文档。
license: GPL-3.0
---

# 实验交付流水线

一次做完：要求 → 代码 → 图 → 日志 → Word 报告。四件事都由命令保证，不靠记性：
**数字只从 metrics 取、图只由脚本重画、对位先在 `--check` 通过、交付前跑完全部门禁。**

## 硬规则（每次都遵守）

1. **不提 AI**：报告正文与格式里不出现"AI / 大模型 / 助手 / 由…生成"。除非实验要求明确鼓励使用 AI 工具，才可提及并说明用于何处。
2. **一次完工**：开工先填完「冻结清单」（章节映射、数字、图、表、截图），写报告前把数字定稿；交付前跑第 6 节全部门禁，不靠返工补。
3. **简短精炼**：正文用短句，删形容词和套话；一个结论一两句说清，不重复题目。
4. **附录只放该放的**：不写源代码、不写中间产物清单，除非实验要求明确要求。附录默认只放必要的补充图表或思考题。
5. **只交 Word**：最终产物是 `.docx`；不交 PDF（自检时可以临时转 PDF 看，看完删掉，不进交付目录）。
6. **截图白底纯代码**：报告里的代码/运行结果图用脚本重绘——白底、无窗口边框、无标题栏、无提示符、无页脚，除代码本身不出现其他文字。
7. **文档先过 pandoc**：遇到 pdf / ppt / docx 等材料，先 `doc2md.py`（内部首选 pandoc）转 md；pandoc 转不了（PDF 就是转不了）再抽文本层，仍不行才 OCR/看图。

## 平台与入口

脚本本体都是 Python（同一套代码跨平台）；入口两套，任选其一：

| 平台 | 调用 |
| --- | --- |
| Windows | `py scripts\doc2md.py …` |
| macOS / Linux | `./scripts/doc2md.sh …`（等价的 bash 入口，会自动挑 `python3`） |

正文命令以 PowerShell 写，把 `py scripts\x.py` 换成 `./scripts/x.sh` 就是 bash 版。
截图字体**跟报告模板走**：`--template 模板.docx`，别自己猜字号字体。

## 0. 开工三条命令

```powershell
py scripts/doctor.py                                   # 解释器/缺包/缺工具 + 修复命令
py scripts/doc2md.py 实验02-实验指导书.pdf --out docs    # 材料转 md（PDF 自动降级到 pymupdf）
py scripts/md2report.py --inspect 实验报告模板.docx       # 模板有哪几部分、哪几小节、版心多宽、什么字体
```

执行入口三条约定（踩过坑，照做省十几轮）：

- **用 doctor 报出来的 `$PY`**。Windows 上 `py -3` 常是另一个空解释器，`python` / `python3` 各是一套，包装错地方就会"装了却导不进来"。
- **pwsh 里不写 heredoc**：`python - <<'EOF'` 会 `ParserError`。脚本先写成 `.py` 再跑。
- **中文输出落文件再用 read 看**：Windows 控制台 GBK 会乱码。脚本只信 `VERDICT` / `DETAIL` 这类 ASCII 状态行。

## 0.1 收料（缺料就问）

| 材料 | 用途 |
| --- | --- |
| 报告模板 `.docx` | 报告套模板页面与章节；**必须有** |
| 实验要求（指导书/教材扫描件） | 定"做什么、交什么、几张图几张表" |
| 示例代码/数据（若有） | 定"从哪跑起" |

`py scripts/pdf_extract.py 课件.pdf --report` 只看体检：哪些页没有文本层。纯图页必须看图，关键参数常只在图里。

## 1. 冻结清单（写 md 之前填完）

| 项 | 怎么定 | 落在哪 |
| --- | --- | --- |
| 章节映射 | `md2report.py --check` 全部 OK | 终端输出 |
| 数字 | 全部写进 `metrics.json`，md 里用 `{{disk.block_size}}` 这类占位 | metrics.json |
| 图 | 编号 + 来源 + 宽度（由 `shots.py` 反推，别手填 15cm） | `outputs/figure_widths.json` |
| 表 | 编号 + 表题（表题独占一段、在表上方） | md |
| 截图 | 每张代码图一个符号名，每张运行图一对起止标记 | `shots.json` |

清单没填完就写正文，后面必然重排重画。

## 2. 工作流

| 步 | 做什么 | 产物 |
| --- | --- | --- |
| 1 骨架 | 按实验号建目录，先写 README 占位 | 目录树 |
| 2 跑通 | 环境自检 + 第一步出数字 | `日志.txt` |
| 3 主线+案例 | 每个子任务：脚本 → 图 → `metrics.json` → 日志 | 四件套 |
| 4 图与截图 | 图写清参数图注；代码/输出用 `shots.py`（底层 `terminal_shot.py`） | `*.png` |
| 5 报告 | 先写 md（数字用占位）→ `--check` → `md2report.py` | `.docx` |
| 6 门禁 | `report_check.py` + `pagelayout.py` 双 PASS | 通过记录 |
| 7 收尾 | README 定稿（目录树/复现命令/关键数字/偏差/限制），清缓存 | 可交付目录 |

四件套缺一个，报告那一节就没有可信来源。

```
实验<N>\
├─ README.md            中间产物说明（templates/实验README.md）
├─ 实验报告.md / .docx  源与交付物
├─ 01_<子任务>\
│  ├─ 脚本.py / 日志.txt / metrics.json / *.png
├─ 终端截图\            由 shots.py 从日志重绘
└─ pipeline.sh         一条命令重跑（templates/pipeline.sh）
```

**所有步骤串成一条命令**（模板 `templates/pipeline.sh`）：编译 → 跑场景 → 会话记录 → 指标 → 图 → 截图 → docx → 门禁。
改任何源码后跑这一条，别挑着重跑（只重跑一半会得到"图是新的、日志是旧的"）。

程序输出格式也在这一步定死，否则会话记录与截图排版要返工：

- 每条命令的输出自成一段：提示符单独一行、以 `> ` 结尾，打印后 `fflush`（否则提示符和首行输出粘在一起，切不开）；
- 表格化输出用定宽格式（`%3d` / `%-16s`），每行不超过 100 列；
- 列表类输出按固定列数换行，不依赖终端宽度；
- 行尾不留悬空分隔符（`a -> b` 不要写成 `a -> b -> `）；
- 超长数据（几十 KB 的写入内容）在记录里标"（此处截断，共 N 字节）"。

## 3. 出图

- 图的图注写参数（窗长、秩、学习率、采样率、版本、块大小）。
- matplotlib 中文字体按平台设：Windows `Microsoft YaHei`、macOS `PingFang SC`、Linux `Noto Sans CJK SC`，并设 `axes.unicode_minus=False`。
- 字号按成品反推：`成品 pt ≈ 图内字号(px) × 插图宽(cm) ÷ 画布宽(px) × 28.35`，目标 6–9 pt。`shots.py` 会替你算好并写回 md。
- 图宽：版心宽度取 `--inspect` 打印的值（各校 14.66 / 16.5 都有，别写死）；常规 12–13 cm，宽扁图 14–14.2 cm，竖长图 10–11.6 cm。

## 3.1 Mermaid / draw.io 图（样式文件已备好，别手调）

| 文件 | 用途 |
| --- | --- |
| `templates/diagrams/mermaid-config.json` | `mmdc -c` 读的主题：白底、黑线、宋体、直线折线 |
| `templates/diagrams/mermaid-theme.mmd` | 没有 CLI 时，粘在 `.mmd` 第一行的 `%%{init}%%` 指令 |
| `templates/diagrams/sample-flow.mmd` | 带样式的流程图示例，照抄结构 |
| `templates/diagrams/lab-style.drawio` | draw.io 图元样板：开始/处理/判断/数据/文档/注释/UML 类，打开复制即用 |
| `templates/diagrams/drawio-style.json` | drawio skill 的样式预设（拷到 `~/.drawio-skill/styles/` 后 `--theme lab-report`） |

```powershell
npx -y @mermaid-js/mermaid-cli@11 -i 图\流程.mmd -o 图\流程.png `
  -c templates\diagrams\mermaid-config.json -w 1000 -b white
```

- **先 mermaid 后 draw.io**：mermaid 三行写完结构；对齐、分组、复杂 UML 再进 draw.io 精修。别一上来就拖框。
- 只改文字，不改样式：字号 12px、线宽 1px、白底黑线、无阴影，配色顶多灰度。
- 竖向流程图别按 12 cm 插，字会大得离谱。按宽高比定宽度：竖排通常 5–8 cm，横排（`flowchart LR`）12–13 cm，时序图 12–14 cm。
- 反算：`宽cm = 目标pt × 画布宽px ÷ (图内字号px × 28.35)`。实测 376×1348 的竖图、字号 24px，要 9 pt → 宽 5.0 cm（高 17.9 cm，正好一页）。
- 字体跟正文一致（模板宋体就 `SimSun`）；macOS 换成 `Songti SC`，别混用。

## 4. 代码/运行结果图（白底、只有代码）

```powershell
# 批量：按符号取代码区间、按起止标记取运行片段，并把宽度写回 md
py scripts/shots.py --spec shots.json --apply-md 实验报告.md

# 单张（字体跟模板走）：./scripts/terminal_shot.sh 是等价的 bash 入口
py scripts/terminal_shot.py --out 终端截图\code.png --code main.py --first-line 1 --last-line 40 --template 模板.docx
py scripts/terminal_shot.py --out 终端截图\run.png --log 日志.txt --split 30 --template 模板.docx
```

`shots.json` 长这样（代码图给符号名，别给行号——改一行就漂）：

```json
{"src": "src/filesys.cpp", "outdir": "figures",
 "code": [{"out": "code_write.png", "symbol": "int write("}],
 "run":  [{"out": "run_fat.png", "text": "outputs/transcripts/04.txt",
           "from": "> write big.txt", "to": "> read big.txt", "max": 20}]}
```

默认白底、无边框、无标题栏、无行号；需要时才加 `--lineno` / `--chrome "main.py"` / `--theme dark`。
长日志用 `--split N` 切成多张；含个人路径时加 `--mask "C:\\Users\\[^\\]+"`。
没模板可读时按平台自动找字体；要硬指定就 `--cjk-font 宋体` / `--font consolas`（名字或字体文件路径都行）。

## 5. 报告

```powershell
py scripts/md2report.py 实验报告.md 模板.docx --check --metrics outputs/metrics.json
py scripts/md2report.py 实验报告.md 模板.docx 实验报告.docx --overwrite-cover --metrics outputs/metrics.json
py scripts/report_check.py 实验报告.md --docx 实验报告.docx --metrics outputs/metrics.json --body-width 16.5
py scripts/pagelayout.py 实验报告.docx --body-width 16.5 --out .cache/layout.txt
# 要覆盖模板字体时：--body-font 黑体 --mono-font "Courier New"
```

- `--check` 不过就不要往下走：对位失败会直接中止，并删掉上一次的旧 docx（避免旧文件被当成新交付）。
- 字体默认跟随模板（读 Normal 样式与前几十段的 rFonts），读不到才退回 宋体/Consolas。

md 约定：顶部 `---` 填封面（课程名称/指导教师/实验日期/实验项目名称）；`# 实验报告`（三部分模板用 `# 实验预习报告`/`# 实验报告`/`# 附页`）；`## 一、xxx` 按中文序号对位红色小节标题；`表 n` 独占一段后接表格；`![图 n 说明](图.png){width=12cm}` **只写一个** `{width=}`；三反引号代码块。

## 6. 交付前自检（每条都有命令，不靠翻页猜）

| 检查 | 命令 |
| --- | --- |
| 章节对位齐全 | `md2report.py … --check` → PASS |
| 图/表编号连续、正文引用存在、图都进了 docx | `report_check.py … --docx` → PASS |
| 无"待填"占位；无 AI 字样（除要求） | 同上 |
| 数字与 `metrics.json`/日志一致；随机种子写死 | md 用 `{{…}}` 占位，渲染时取不到值即失败 |
| 每页填充率、插图跨页、空白页、末页只剩一行 | `pagelayout.py` → PASS |
| 图有图注且写明参数；图与图注同页；不超版心 | `report_check.py`（宽度）+ `pagelayout.py`（跨页） |
| 图样式统一（`templates/diagrams/` 那套） | 目视：字号、线宽、字体、配色全文一致 |
| 截图白底、只有代码；字体与模板一致 | `shots.py` 生成（默认不带装饰） |
| 附录没有源代码/中间产物（除非要求） | 目视 |
| 交付目录只有 `.docx` 与必要图片 | 删 `.cache`、`__pycache__`、`*.tmp.docx`、`~$*.docx`、临时 html |

## 7. 常见坑

1. `{width=12cm}{width=12cm}` 重复 → 整行退化成普通段落，图**静默丢失**（只能靠数图发现）。只写一个，或让 `shots.py --apply-md` 统一重写。
2. 模板首页标题含"实验报告"（如"XX大学实验报告 成绩"），附页标题是"…实验报告（附页）"：子串匹配会把两者认混，所有小节"对不上"。新版按部分互斥识别，先 `--inspect` 确认。
3. 模板没有 `Table Grid` 样式 → 赋样式 `KeyError`。`md2report.py` 会自动补一个带框线的。
4. docx 插入锚点要取小节标题**后**的空段，取标题本身会排到标题前面。
5. LibreOffice 会留 `~$xxx.docx` 锁，覆盖报 `PermissionError`：关掉 Word/LibreOffice；`md2report.py` 会先清锁再重试。
6. 代码图按行号截，改一行就漂 → 用 `shots.py` 按符号名取区间（大括号/缩进匹配）。
7. 封面字段一长就折行：填完按视觉宽度补空格，必要时轻微负字距压回一行。
8. 字号缩水：先按第 3 节公式算，别画完再发现 4 pt。
9. 中文 .ps1 用 `pwsh` 7 跑，Windows PowerShell 5.1 按 ANSI 读会报语法错。
10. 表格防拆页：表头重复 + `cantSplit` + 除末行 `keep_with_next`。
11. 图注跟图：图段 `keep_with_next`。
12. `pdf_extract.py --report` 点出的纯图页必须看图，别只读文本层。
13. macOS 上截图中文变方块：`brew install --cask font-noto-sans-cjk`，或 `--cjk-font 宋体`（会落到 Songti.ttc）。
14. `.sh` 带 CRLF 会报 `bad interpreter: /usr/bin/env bash^M`——仓库已用 `.gitattributes` 钉成 LF，别手工改回来。
15. mermaid 竖向图按 12 cm 插进 Word，字会大得比标题还显眼：按第 3.1 节公式算宽度，或改成 `flowchart LR`。

## 8. 脚本

| 脚本 | 作用 |
| --- | --- |
| `doctor.py` | 环境自检：解释器/库/工具 + 修复命令（中文详情写 UTF-8 文件，控制台只出 ASCII 状态） |
| `doc2md.py` | pdf/pptx/docx/html → md（pandoc 优先，PDF 降级抽文本+导页图） |
| `pdf_extract.py` | PDF → 逐页文本 + 逐页 PNG，点名纯图页 |
| `make_transcript.py` | `.in` + `.log` → 可读会话记录（配对由日志内容决定，失败的 write 不消耗数据行） |
| `shots.py` | 按符号生成白底代码图、按标记生成运行图，算宽度并写回 md |
| `terminal_shot.py` | 单个代码/日志 → 白底图片（字体跟模板） |
| `md2report.py` | md → 模板 docx（`--inspect` 看模板、`--check` 干跑、`--metrics` 填占位、对位失败即中止） |
| `report_check.py` | 图/表编号、引用、宽度、占位、AI 字样、docx 是否丢图的自动核对 |
| `pagelayout.py` | docx → PDF → 每页填充率 / 插图跨页 / 空白页 / 末页空白 |
| `_fonts.py` | 上面几个脚本共用的字体解析（名字→文件、从模板读字体），不直接调用 |

每个脚本都有等价的 `.sh` 入口（`./scripts/x.sh`），依赖：`pip install pillow python-docx pymupdf`；转换 docx/pptx 需要 `pandoc`（macOS: `brew install pandoc`）；`pagelayout.py` 需要 LibreOffice。

模板：`templates/实验README.md`（交付说明）、`templates/report_skeleton.md`（md 骨架）、`templates/pipeline.sh`（一条命令重跑）；图样式：`templates/diagrams/`；示例：`examples/demo_code.png`、`examples/demo_output_*.png`。

## 附：这些门禁是为什么加的

| 曾经怎么返工的 | 现在拦在哪 |
| --- | --- |
| 试了四个解释器才发现包装错地方 | `doctor.py` 一次列清 |
| 模板对位失败仍生成了 docx，白写一套自定义构建脚本 | `--check` + 对位失败中止 + 清旧产物 |
| 提示符和首行输出粘连、表格列数反复调 | 第 2 节的输出格式契约 |
| 改了源码 → 行号漂移 → 截图全错 | `shots.py --symbol` |
| 一行两个 `{width=}`，图静默消失，最后靠数图才发现 | `report_check.py` + 解析期报错 |
| 排版改了八轮才发现末页只剩一行 | `pagelayout.py` 一次给全 |
| 正文写"9 个场景"，实际 11 个 | `{{metrics…}}` 占位，取不到即失败 |
| 会话记录与日志错位，改了四遍配对逻辑 | `make_transcript.py` 按日志内容配对 |
