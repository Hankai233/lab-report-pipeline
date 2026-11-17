#!/usr/bin/env bash
# doc2md.sh —— doc2md.py 的 POSIX 入口（Windows 上直接跑 python scripts/doc2md.py）
# 用法：./scripts/doc2md.sh 实验02-实验指导书.pdf --out docs
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

export PYTHONIOENCODING=utf-8   # 中文报错别炸在 locale 上
exec "$py" "$here/doc2md.py" "$@"
