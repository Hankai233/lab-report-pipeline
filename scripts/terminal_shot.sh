#!/usr/bin/env bash
# terminal_shot.sh —— terminal_shot.py 的 POSIX 入口（Windows 上直接跑 python scripts/terminal_shot.py）
# 字体按平台自动找（Windows 雅黑/Consolas、macOS PingFang/Menlo、Linux Noto CJK/DejaVu），
# 找不到时才需要 --cjk-font / --font 指定。
# 用法：./scripts/terminal_shot.sh --out 终端截图/code.png --code main.py --first-line 1 --last-line 40
set -euo pipefail

here="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"

py=""
for cand in python3 python; do
  if command -v "$cand" >/dev/null 2>&1; then py="$cand"; break; fi
done
if [ -z "$py" ]; then
  printf '找不到 python3，请先装 Python 3.9+（macOS: brew install python）\n' >&2
  exit 127
fi

export PYTHONIOENCODING=utf-8
exec "$py" "$here/terminal_shot.py" "$@"
