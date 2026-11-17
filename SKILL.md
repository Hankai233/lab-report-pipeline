---
name: lab-report-pipeline
description: >-
  课程实验（机器感知、机器学习、图像处理、软件工程等）从读要求到交 Word 报告的流水线：
  收模板·指导书·示例代码，按步骤留中间产物，图表与代码截图由脚本生成，
  数字只从日志/metrics 取，最后套模板渲染 docx 并自检。
  触发：做实验、写实验报告、实验指导书、实验报告模板、复现实验、跑通实验、
  实验截图、实验一/实验二/实验三、中间产物留档、绘制实验图、补实验文档。
license: GPL-3.0
---

# 实验交付流水线

一次做完：要求 → 代码 → 图 → 日志 → Word 报告。数字可追溯，图能重生成。

## 硬规则（每次都遵守）

1. **不提 AI**：报告正文与格式里不出现"AI / 大模型 / 助手 / 由…生成"。除非实验要求明确鼓励使用 AI 工具，才可提及并说明用于何处。
2. **一次完工**：开工先一次性对齐（材料、章节、编号、图宽、截图规格），写报告前把数字定稿；交付前用第 5 节清单核对，不靠返工补。
3. **简短精炼**：正文用短句，删形容词和套话；一个结论一两句说清，不重复题目。
4. **附录只放该放的**：不写源代码、不写中间产物清单，除非实验要求明确要求。附录默认只放必要的补充图表或思考题。
5. **只交 Word**：最终产物是 `.docx`；不交 PDF（自检时可以临时转 PDF 看，看完删掉，不进交付目录）。
6. **截图白底纯代码**：报告里的代码/运行结果图用脚本重绘——白底、无窗口边框、无标题栏、无提示符、无页脚，除代码本身不出现其他文字。
7. **文档先过 pandoc**：遇到 pdf / ppt / docx 等材料，先 `doc2md.py`（内部首选 pandoc）转 md；pandoc 转不了（PDF 就是转不了）再抽文本层，仍不行才 OCR/看图。

## 0. 收料（缺料就问）

| 材料 | 用途 |
| --- | --- |
| 报告模板 `.docx` | 报告套模板页面与章节；**必须有** |
| 实验要求（指导书/教材扫描件） | 定"做什么、交什么、几张图几张表" |
| 示例代码/数据（若有） | 定"从哪跑起" |

```powershell
py scripts/doc2md.py 实验02-实验指导书.pdf --out docs   # 先 pandoc，失败自动降级
py scripts/pdf_extract.py 课件.pdf --report            # 只看体检：哪些页是纯图
```

纯图页必须看图，关键参数常只在图里。

## 1. 工作流

| 步 | 做什么 | 产物 |
| --- | --- | --- |
| 1 骨架 | 按实验号建目录，先写 README 占位 | 目录树 |
| 2 跑通 | 环境自检 + 第一步出数字 | `日志.txt` |
| 3 主线+案例 | 每个子任务：脚本 → 图 → `metrics.json` → 日志 | 四件套 |
| 4 图与截图 | 图写清参数图注；代码/输出用 `terminal_shot.py` | `*.png` |
| 5 报告 | 先写 md（数字抄自 metrics/日志）→ `md2report.py` | `.docx` |
| 6 收尾 | README 定稿（目录树/复现命令/关键数字/偏差/限制），清缓存 | 可交付目录 |

四件套缺一个，报告那一节就没有可信来源。

```
实验<N>\
├─ README.md            中间产物说明（templates/实验README.md）
├─ 实验报告.md / .docx  源与交付物
├─ 01_<子任务>\{脚本.py, 日志.txt, metrics.json, *.png}
└─ 终端截图\
```

## 2. 出图

- 图的图注写参数（窗长、秩、学习率、采样率、版本）。
- 字号按成品反推：`成品 pt ≈ 图内字号(px) × 插图宽(cm) ÷ 画布宽(px) × 28.35`，目标 6–9 pt。
- 图宽：版心 14.66 cm；常规 12–13 cm，宽扁图 14–14.2 cm，竖长图 10–11.6 cm。

## 3. 代码/运行结果图（白底、只有代码）

```powershell
py scripts/terminal_shot.py --out 终端截图\code.png --code main.py --first-line 1 --last-line 40
py scripts/terminal_shot.py --out 终端截图\run.png --log 日志.txt --split 30
```

默认白底、无边框、无标题栏、无行号；需要时才加 `--lineno` / `--chrome "main.py"` / `--theme dark`。
长日志用 `--split N` 切成多张；含个人路径时加 `--mask "C:\\Users\\[^\\]+"`。

## 4. 报告

```powershell
py scripts/md2report.py --inspect "模板.docx"              # 看模板能对位哪些标题
py scripts/md2report.py 实验报告.md 模板.docx 实验报告.docx
py scripts/md2report.py 实验报告.md 模板.docx out.docx --map 章节映射.json
```

md 约定：顶部 `---` 填封面（课程名称/指导教师/实验日期/实验项目名称）；`# 实验报告`（三部分模板用 `# 实验预习报告`/`# 实验报告`/`# 附页`）；`## 一、xxx` 自动对位红色小节标题；`表 n` 独占一段后接表格；`![图注](图.png){width=12cm}`；三反引号代码块。

## 5. 交付前自检

- [ ] 要求里每个编号实验/问题都有对应小节；图、表编号连续且正文引用一致
- [ ] 无"待填"占位；封面字段单行不折行；无 AI 字样
- [ ] 数字与 `metrics.json`/日志逐项一致；随机种子写死
- [ ] 图有图注且写明参数；图与图注同页；不超版心
- [ ] 截图白底、只有代码
- [ ] 附录没有源代码/中间产物（除非要求）
- [ ] 交付目录只有 `.docx` 与必要图片，无 PDF、无 `__pycache__`、无临时 html
- [ ] 打开 docx 翻一遍：无孤行、无空白页、无错位

## 6. 常见坑

1. 字号缩水：先按第 2 节公式算，别画完再发现 4 pt。
2. docx 插入锚点要取小节标题**后**的空段，取标题本身会排到标题前面。
3. 封面字段一长就折行：填完按视觉宽度补空格，必要时轻微负字距压回一行。
4. LibreOffice 转换留 `~$xxx.docx` 锁文件，覆盖前先删。
5. 中文 .ps1 用 `pwsh` 7 跑，Windows PowerShell 5.1 按 ANSI 读会报语法错。
6. matplotlib 中文要显式设 `Microsoft YaHei` 且 `axes.unicode_minus=False`。
7. 表格防拆页：表头重复 + `cantSplit` + 除末行 `keep_with_next`。
8. 图注跟图：图段 `keep_with_next`。
9. `pdf_extract.py --report` 点出的纯图页必须看图，别只读文本层。

## 7. 脚本

| 脚本 | 作用 |
| --- | --- |
| `scripts/doc2md.py` | pdf/pptx/docx/html → md（pandoc 优先，PDF 降级抽文本+导页图） |
| `scripts/pdf_extract.py` | PDF → 逐页文本 + 逐页 PNG，点名纯图页 |
| `scripts/terminal_shot.py` | 代码/日志 → 白底图片（只含代码） |
| `scripts/md2report.py` | md → 模板 docx（`--inspect` 查标题、`--map` 指定映射） |

模板：`templates/实验README.md`、`templates/report_skeleton.md`；示例：`examples/demo_code.png`、`examples/demo_output_*.png`。
