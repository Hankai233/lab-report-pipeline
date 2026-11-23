#!/usr/bin/env bash
# shots.sh —— shots.py 的 POSIX 入口（Windows 上直接跑 python scripts/shots.py）
# 用法：./scripts/shots.sh <参数...>    # ./scripts/shots.sh --help 看完整参数
set -euo pipefail

here="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"

py=""
for cand in python3 python; do
  if command -v "$cand" >/dev/null 2>&1; then py="$cand"; break; fi
done
if [ -z "$py" ]; then
  printf "找不到 python3，请先装 Python 3.9+" >&2
  exit 127
fi

export PYTHONIOENCODING=utf-8
exec "$py" "$here/shots.py" "$@"
