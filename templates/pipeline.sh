#!/usr/bin/env bash
# 一条命令重跑整条流水线（实验交付用；复制到实验目录后按注释改成你的命令）
#
# 为什么必须是"一条命令"：分步手动重跑很容易只重跑一半（图是新的、日志是旧的），
# 报告里的数字和截图就会互相矛盾，只能整体返工。
#
# 用法：./pipeline.sh            # 全跑
#       SKIP_LAB=1 ./pipeline.sh # 只重生成报告，不重跑实验（改了文字不想重跑程序）
set -euo pipefail
cd "$(dirname "$0")"

PY=${PY:-python3}
SKILL=${SKILL:-"$HOME/.dsh/skills/lab-report-pipeline/scripts"}
MD=${MD:-实验报告.md}
TPL=${TPL:-实验报告模板.docx}
OUT=${OUT:-实验报告.docx}
BODY_W=${BODY_W:-14.66}        # 取 md2report --inspect 打印的版心宽

step() { printf '\n== %s\n' "$*"; }

step "0 环境自检"
$PY "$SKILL/doctor.py" || { echo "环境不满足：看 doctor_report.txt，别硬跑"; exit 1; }

step "1 主体实验（改成你的编译与运行命令）"
if [ "${SKIP_LAB:-0}" != "1" ]; then
    ./run_lab.sh          # ← 编译 + 跑各场景，落日志到 outputs/logs/、磁盘映像到 outputs/images/
fi

step "2 会话记录（命令行回填进日志）"
$PY "$SKILL/make_transcript.py" --in outputs/scenarios --log outputs/logs --out outputs/transcripts

step "3 指标（报告数字的唯一来源）"
$PY make_metrics.py                 # → outputs/metrics.json

step "4 插图 + 白底截图（按符号取代码区间、按标记取运行片段）"
$PY make_figures.py                 # matplotlib 图
$PY "$SKILL/shots.py" --spec shots.json --apply-md "$MD"

step "5 报告：先干跑对位，再生成"
$PY "$SKILL/md2report.py" "$MD" "$TPL" --check --metrics outputs/metrics.json
$PY "$SKILL/md2report.py" "$MD" "$TPL" "$OUT" --overwrite-cover --metrics outputs/metrics.json

step "6 门禁：内容核对 + 摆版"
$PY "$SKILL/report_check.py" "$MD" --docx "$OUT" --metrics outputs/metrics.json --body-width "$BODY_W"
$PY "$SKILL/pagelayout.py" "$OUT" --body-width "$BODY_W" --out .cache/layout.txt

step "done"
echo "交付前再确认目录里没有 .cache/ __pycache__/ ~\$*.docx"
