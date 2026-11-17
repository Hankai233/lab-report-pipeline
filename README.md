# lab-report-pipeline

> 一个把「做实验」翻译成「交报告」的技能。

**一句话**：它从老师那个扫得有点歪的 PDF 里抠出要求，帮你把脚本、日志、图、数字串成一条**能追溯**的线，最后塞进学校的 Word 模板里。它不替你跑实验，但它会阻止你在截止前一晚才想起来「图没有图注」。

给 QLU 计算机专业的同学用。其他人也能用，只是你们的模板可能没这么难伺候。

---

## 这玩意儿到底在治什么病

凌晨两点，你面前有三个哲学问题：

1. **这张图插进 Word 是几号字？** —— 按成品反推一下就知道：`成品 pt ≈ 图内字号(px) × 插图宽(cm) ÷ 画布宽(px) × 28.35`。目标 6–9 pt。你现在画的那个大概是 4 pt。
2. **报告里那个 0.923 是哪来的？** —— 上周二的终端。已经关了。这就是为什么本技能要求每个子任务都留一份 `metrics.json`。
3. **为什么我插的图跑到标题上面去了？** —— 锚点取成了小节标题本身。要取标题**后面**那个空段。

这三个问题在 [`SKILL.md`](SKILL.md) 里各有一节。剩下的 20 几个坑也都写着，都是在真材料上撞出来的。

---

## 三十秒上手

把目录放进 DSH 的 `skills/` 下（`SKILL.md` 带 frontmatter，放进去就认），然后跟它说人话：

```text
做实验三，材料和模板都在 ./实验三/
```

它会认这些词：**做实验 / 写实验报告 / 实验指导书 / 实验报告模板 / 复现实验 / 跑通实验 / 实验截图 / 中间产物留档**。

交付物只有一个：**`.docx`**。没有 PDF，没有「我先发你一版看看」，没有 `最终版_最终版_真的最终版.docx`。

---

## 四个脚本，各管一摊

脚本本体是 Python，一套代码三平台通吃；入口按你手边的终端挑一个：

```powershell
py scripts/doc2md.py 指导书.pdf --out docs          # Windows PowerShell
```
```bash
./scripts/doc2md.sh 指导书.pdf --out docs           # macOS / Linux（等价，自动挑 python3）
```

| 脚本 | 干什么 | 大白话 |
| --- | --- | --- |
| `scripts/doc2md.py` / `.sh` | PDF / PPTX / DOCX / HTML → Markdown | 先上 pandoc。顺便说：**pandoc 读不了 PDF**，报错原文是「PDF is not a supported input format」，翻译过来是「别指望我」。所以它自己会降级到 pymupdf 抽文本，抽不出来就逐页导图。 |
| `scripts/pdf_extract.py` / `.sh` | PDF → 逐页文本 + 逐页 PNG，并点名「哪几页是纯图」 | 纯图页 = 关键要求都在图里、而你只 grep 了文本层、最后漏做一半的那几页。 |
| `scripts/terminal_shot.py` / `.sh` | 代码 / 日志 → 白底图片 | 默认没有标题栏、没有提示符、没有页脚、没有行号。因为老师模板里已经有一个「截图」的位置了，不需要再来一个 PowerShell 窗口。 |
| `scripts/md2report.py` / `.sh` | md → 学校的 docx | **以你的模板为底稿**，只往里插正文/表格/图/代码块。所以页面设置永远不会画错——这是它唯一敢打包票的事。 |

```powershell
py scripts/pdf_extract.py 实验02-实验指导书.pdf --report           # 先体检：哪几页是纯图
py scripts/doc2md.py 实验02-实验指导书.pdf --out docs             # 再转 md
py scripts/terminal_shot.py --out 终端截图\code.png --code main.py --template 模板.docx
py scripts/md2report.py --inspect "模板.docx"                     # 模板能对位哪些标题、用了什么字体
py scripts/md2report.py 实验报告.md 模板.docx 实验报告.docx
```

依赖：`pip install pillow python-docx pymupdf`；转 docx/pptx 还要 `pandoc`（macOS `brew install pandoc`）。

### 字体这块，认模板不认我

`terminal_shot.py` 出图时可以直接读你的报告模板：

```bash
./scripts/terminal_shot.py --out 终端截图/code.png --code main.py --template 模板.docx
```

它从模板的 Normal 样式和前几十段里取字体，中文用模板正文那套、西文用模板的代码字体，于是截图和正文是同一种字，不是我随手挑的 Consolas。`md2report.py` 同理——插进 docx 的正文/代码也跟随模板，而不是硬写宋体。

没模板可读时按平台自动找（Windows 微软雅黑/Consolas、macOS PingFang/Menlo、Linux Noto CJK/DejaVu）；要硬指定就 `--cjk-font 宋体 --font couriernew`，名字和字体文件路径都收。

### 图：Mermaid 写结构，draw.io 收尾，样式文件先备好

`templates/diagrams/` 里那几份就是"提前写好的样式"，别每次重画一遍：

| 文件 | 什么时候用 |
| --- | --- |
| `mermaid-config.json` | `mmdc -c` 传给 mermaid CLI，出白底黑线宋体的图 |
| `mermaid-theme.mmd` | 没有 CLI，只能贴进 Typora / 在线编辑器——把这一行粘到 `.mmd` 最前面 |
| `sample-flow.mmd` | 一个带样式的完整流程图，照抄结构就行 |
| `lab-style.drawio` | draw.io 图元样板：开始/处理/判断/数据/文档/注释/UML 类，打开复制即用 |
| `drawio-style.json` | 给 drawio skill 用的预设，拷进 `~/.drawio-skill/styles/` 后 `--theme lab-report` |

```bash
npx -y @mermaid-js/mermaid-cli@11 -i 图/流程.mmd -o 图/流程.png \
  -c templates/diagrams/mermaid-config.json -w 1000 -b white
```

一条纪律：**先用 mermaid 把结构写对，画不动了再上 draw.io 精修。** 顺序反过来，你会在第四个框上开始调对齐，然后一晚上没了。

还有个坑：竖向流程图别按 12 cm 插进 Word，24px 的字会被放大成标题。按 `宽cm = 目标pt × 画布宽px ÷ (图内字号px × 28.35)` 算——实测一张 376×1348 的竖图要 9 pt 就只该占 5 cm 宽（高度 17.9 cm，刚好一页）。

---

## 七条铁律（附理由，省得你问为什么）

1. **报告里不提 AI** —— 报告是给老师看的，不是给面试官看的。
2. **一次完工，少返工** —— 返工的成本永远比你多花十分钟先把材料对齐高。
3. **简短精炼** —— 「本文对……进行了深入研究」占掉一整行，信息量等于零。
4. **附录不放源代码** —— 你贴 300 行，老师会数你注释里的 TODO。
5. **只交 docx** —— 「我发你 PDF 了」在教务处的语境下约等于「我没交」。
6. **截图白底、只有代码** —— 深色主题是你喜欢，不是评分标准。
7. **文档先 pandoc，失败再抽文本，再不行才看图** —— 顺序别乱，看图的成本最高。

---

## 目录

```text
lab-report-pipeline/
├─ SKILL.md              硬规则 + 六步工作流 + 交付自检 + 25 个坑 + 脚本说明
├─ scripts/              四个 Python 脚本 + 各自的 .sh 入口 + 共用字体解析 _fonts.py
├─ templates/            报告 md 骨架、实验目录 README 骨架、diagrams/ 里的 mermaid 与 draw.io 样式
├─ examples/             真实输出的示例图（白底纯代码，可直接插 Word）
├─ .gitattributes        换行策略：.sh 一律 LF（不然 macOS 报 bad interpreter）
└─ LICENSE               GPL-3.0
```

---

## 三个好笑的坑（全文在第 6 节，这里挑最冤的）

- **中文 `.ps1` 必须用 `pwsh` 7 跑。** Windows PowerShell 5.1 会按 ANSI 读你的 UTF-8，然后当场报语法错，让你怀疑自己写的是不是 PowerShell。
- **`~$` 开头的文件是 LibreOffice 的锁。** 它长得像你的报告，但它不是你的报告。覆盖前先删。
- **matplotlib 里中文默认全是方框。** 不是字体没装，是你没设字体（Win `Microsoft YaHei` / macOS `PingFang SC` / Linux `Noto Sans CJK SC`）和 `axes.unicode_minus=False`。
- **`.sh` 用 CRLF 存盘会报 `bad interpreter: /usr/bin/env bash^M`。** 那个 `^M` 是回车，藏在你 shebang 的末尾，肉眼看不见。仓库用 `.gitattributes` 钉死了 LF。

---

## FAQ

**Q：它能帮我写实验报告吗？**
A：能帮你写报告里的「格式」和「数字」那部分。结论写什么——那得你真的做了实验。

**Q：能绕过查重吗？**
A：不能。它连 AI 都不让提，更不会帮你绕过查重。

**Q：学校模板年年换，还能用吗？**
A：能。它不改你的模板，只往里面塞内容。模板从「预习 / 报告 / 附页」三段变成单段，它也认（会自动降级）。

**Q：那为什么不直接截图呢？**
A：因为你的截图里有用户名、有壁纸、有回收站图标，还有 31 条未读通知。

**Q：我在 macOS 上，中文出图是方块。**
A：那是系统没给等宽字体配中文。要么 `brew install --cask font-noto-sans-cjk`，要么直接 `--cjk-font 宋体`——它会在本机落到 `Songti.ttc`，字和模板正文对得上。最省事的还是 `--template 模板.docx`，让模板自己说用哪套字。

**Q：中间产物会不会很占地方？**
A：会。所以每个子任务只留四件套：`脚本.py / 日志.txt / metrics.json / *.png`。四件套缺一个，报告那一节就没有可信来源——这句话它是认真的。

---

## 免责声明

本项目不提供实验灵感、不保证分数、也不负责解释为什么你在 `--last-line 40` 那里断掉的那半行，恰好是全班只有你抽到的那组参数。

## License

[GPL-3.0](LICENSE)。

MIT 的座右铭是「随便用，别提我就行」，GPL 的座右铭是「随便用，改好了也得让别人随便用」。

所以：你自己改着玩、藏着不发，没人管你；可一旦你把这套流程改进了又发给同学，那份改过的代码也得跟着开源。库是大家的，这个道理和实验数据一样——不能只有你一个人见过原始日志。
