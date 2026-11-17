#!/usr/bin/env bash
# pdf_extract.sh —— pdf_extract.py 的 POSIX 入口（Windows 上直接跑 python scripts/pdf_extract.py）
# 用法：./scripts/pdf_extract.sh 实验02-实验指导书.pdf --report
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
exec "$py" "$here/pdf_extract.py" "$@"
