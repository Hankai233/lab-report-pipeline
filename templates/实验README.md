# <课程名> 实验<N> · 中间产物说明

> 这个文件放在每次实验的交付目录根上。批阅的人先看它，就知道"东西在哪、怎么复现、哪些地方和指导书不一样"。

- 运行日期：**YYYY-MM-DD**
- 运行环境：<OS> / Python <ver> / <关键库版本> / <CUDA 或纯 CPU>
- 运行方式：`<一条命令>`
- 所有数字均来自本次运行的 `<日志.txt>`／`<metrics.json>`，所有图均由本次运行的脚本生成

---

## 1. 目录树

```text
实验<N>\
├─ README.md                    本文件
├─ 实验报告.docx                交付物（由 md 渲染）
├─ 实验报告.md                  报告源（md2report 的输入）
├─ 01_<子任务>\
│  ├─ <脚本>.py                 本步骤的实现
│  ├─ 日志.txt                  本次运行控制台输出（UTF-8）
│  ├─ metrics.json              本步骤全部测量值（报告数值的唯一来源）
│  └─ <图>.png                  本步骤产出的图
├─ 02_<子任务>\ ...
└─ 终端截图\                    由 terminal_shot.py 从日志生成
```

## 2. 文件与产物说明

| 文件 | 说明 |
| --- | --- |
| `<脚本>.py` | 做什么、输入输出是什么、关键参数 |
| `日志.txt` | 本次运行的控制台输出，报告里的"运行结果"截图由它重绘 |
| `metrics.json` | 报告所有数字的出处；改脚本后重跑即可刷新报告 |
| `<图>.png` | 该图对应的参数与含义（一句话） |

## 3. 完整复现命令

一条命令跑完（见 `pipeline.sh`），下面拆开写便于理解：

```bash
# 0) 环境自检：认解释器/缺包/缺工具（中文详情在 doctor_report.txt）
$PY scripts/doctor.py

# 1) 主线实验：编译 + 跑各场景，日志落 outputs/logs/
./run_lab.sh

# 2) 会话记录 + 指标（报告数字的唯一来源）
$PY scripts/make_transcript.py --in outputs/scenarios --log outputs/logs --out outputs/transcripts
$PY make_metrics.py

# 3) 图与白底截图（宽度按目标字号反推并写回 md）
$PY make_figures.py
$PY scripts/shots.py --spec shots.json --apply-md 实验报告.md

# 4) 报告：先干跑对位，再生成，最后过门禁
$PY scripts/md2report.py 实验报告.md 实验报告模板.docx --check --metrics outputs/metrics.json
$PY scripts/md2report.py 实验报告.md 实验报告模板.docx 实验报告.docx --overwrite-cover --metrics outputs/metrics.json
$PY scripts/report_check.py 实验报告.md --docx 实验报告.docx --metrics outputs/metrics.json --body-width 14.66
$PY scripts/pagelayout.py 实验报告.docx --body-width 14.66 --out .cache/layout.txt
```

`report_check.py` 与 `pagelayout.py` 都必须 PASS 再交付；不 PASS 就改，别靠肉眼翻页。

## 4. 关键结果一览（本次运行）

| 项目 | 结果 |
| --- | --- |
| <指标> | <数值> |

## 5. 与指导书/参考代码的偏差（重要）

1. **<偏差点>**：为什么改、改了之后得到什么、指导书原样会得到什么。
2. ……

## 6. 未做的事 / 已知限制

- 例：本机无 GPU，未测 CUDA 版本；某图只跑了一个随机种子。
