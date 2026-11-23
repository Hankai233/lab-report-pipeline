# -*- coding: utf-8 -*-
"""一次生成报告里的全部“白底纯代码 / 纯运行输出”图，并把插图宽度写回 md。

为什么不用行号：源码一改，行号就漂，图和代码对不上，只能重来一遍。这里按**符号**取区间
（大括号匹配 / Python 缩进匹配），改代码后重跑即对。

为什么宽度要算：`成品 pt ≈ 图内字号(px) × 插图宽(cm) ÷ 画布宽(px) × 28.35`。
写死 15 cm 会让字小的图只有 4–6 pt，只能重排。这里按目标字号反推宽度，并保证不超版心。

idempotent：同一个 md 反复执行，结果不变（先剥掉旧 `{width=}` 再写一个）。

用法
----
    python shots.py --spec shots.json            # 出图 + 算宽度 + 写 figure_widths.json
    python shots.py --spec shots.json --apply-md 实验报告.md
    python shots.py --spec shots.json --only code_write

spec 示例（shots.json）
-----------------------
    {
      "src": "src/filesys.cpp",
      "outdir": "figures",
      "code": [{"out": "code_write.png", "symbol": "int write("},
               {"out": "code_read.png",  "symbol": "int  read(",
                "first_line": 0, "last_line": 0}],
      "run":  [{"out": "run1_dir.png", "text": "outputs/transcripts/01.txt",
                "from": "> mkdir docs", "to": "> stat docs", "max": 20}]
    }

退出码：0 成功；1 有图没生成；2 参数/spec 错误。
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

MAX_W = 14.6        # 版心宽（cm），学校模板常见 14.66；超过会被 Word 缩回
MAX_H = 25.5        # 版心高（cm）
TARGET_PT = 9.0     # 目标成品字号上限
WIDTH_RX = re.compile(r"\{width=[\d.]+\s*cm\}")


def wide(s: str) -> int:
    return sum(2 if ord(c) > 0x2E7F else 1 for c in s.expandtabs(4))


def find_symbol_range(lines: list[str], symbol: str) -> tuple[int, int]:
    """按符号名定位函数定义，返回 1-based 闭区间。"""
    start = None
    for i, line in enumerate(lines):
        t = line.lstrip()
        if t.startswith(symbol) and not t.rstrip().endswith(";"):
            start = i
            break
    if start is None:
        raise SystemExit(f"[spec] 找不到符号定义：{symbol!r}（确认含 'def '/'函数名(' 前缀）")

    depth, seen, j = 0, False, start
    while j < len(lines):
        depth += lines[j].count("{") - lines[j].count("}")
        if "{" in lines[j]:
            seen = True
        if seen and depth == 0:
            return start + 1, j + 1
        j += 1
    if seen:
        raise SystemExit(f"[spec] 大括号不配对：{symbol!r}")

    indent = len(lines[start]) - len(lines[start].lstrip())
    for k in range(start + 1, len(lines)):
        s = lines[k]
        if not s.strip():
            continue
        if len(s) - len(s.lstrip()) <= indent:
            return start + 1, k
    return start + 1, len(lines)


def shot(out: Path, cols: int, code: Path | None = None, log: Path | None = None,
         span: tuple[int, int] | None = None, fs: int = 16, scale: int = 2,
         tool: Path | None = None) -> bool:
    cmd = [sys.executable, str(tool), "--out", str(out), "--cols", str(cols),
           "--fs", str(fs), "--scale", str(scale)]
    if code is not None:
        a, b = span or (1, 0)
        cmd += ["--code", str(code), "--first-line", str(a), "--last-line", str(b),
                "--highlight", "(?!)"]           # 代码图不做任何着色
    else:
        cmd += ["--log", str(log)]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       errors="replace")
    if r.returncode != 0:
        print(f"FAIL {out.name}: {(r.stderr or r.stdout).strip()[:200]}")
        return False
    return True


def recommend(p: Path, fs: int, scale: int) -> tuple[float, float]:
    from PIL import Image
    w, h = Image.open(p).size
    fontpx = fs * scale
    width = min(MAX_W, MAX_H * w / h, TARGET_PT * w / (fontpx * 28.35))
    width = max(6.0, round(width * 2) / 2)
    return width, fontpx * width / w * 28.35


def apply_widths(md: Path, widths: dict[str, float]) -> int:
    text = md.read_text(encoding="utf-8")
    changed = 0

    def fix(m: re.Match) -> str:
        nonlocal changed
        head, path = m.group(1), m.group(2)
        name = path.split("/")[-1]
        if name not in widths:
            return m.group(0)
        new = f"{head}{{width={widths[name]:g}cm}}"
        if new != m.group(0):
            changed += 1
        return new

    pat = re.compile(r"(!\[[^\]]*\]\(([^)\s]+)\))(?:\{width=[\d.]+\s*cm\})*")
    new = pat.sub(fix, text)
    if new != text:
        md.write_text(new, encoding="utf-8")
    return changed


def run_excerpt(text: Path, frm: str, to: str | None, cap: int | None) -> list[str]:
    lines = text.read_text(encoding="utf-8").splitlines()
    i = next((k for k, l in enumerate(lines) if l.startswith(frm)), None)
    if i is None:
        raise SystemExit(f"[spec] 找不到起始标记：{frm!r} in {text.name}")
    j = len(lines)
    if to:
        j = next((k for k, l in enumerate(lines) if k > i and l.startswith(to)), len(lines))
    sel = lines[i:j]
    if cap:
        sel = sel[:cap]
    while sel and (not sel[-1].strip() or sel[-1].startswith("> ")):
        sel.pop()
    return sel


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="按符号生成代码图/运行图并算插图宽度")
    ap.add_argument("--spec", required=True)
    ap.add_argument("--apply-md", help="把算出的宽度写回该 md")
    ap.add_argument("--only", help="只重做名字包含该片段的图")
    ap.add_argument("--fs", type=int, default=16)
    ap.add_argument("--scale", type=int, default=2)
    o = ap.parse_args(argv or sys.argv[1:])

    spec_path = Path(o.spec)
    if not spec_path.exists():
        print(f"找不到 spec：{spec_path}", file=sys.stderr)
        return 2
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    root = spec_path.parent
    outdir = root / spec.get("outdir", "figures")
    outdir.mkdir(parents=True, exist_ok=True)
    excdir = outdir / "excerpts"
    tool = Path(__file__).with_name("terminal_shot.py")

    widths: dict[str, float] = {}
    rows, fails = [], 0

    src = spec.get("src")
    src_path = (root / src) if src else None
    code_lines = src_path.read_text(encoding="utf-8").splitlines() if src_path else []

    for item in spec.get("code", []):
        name = item["out"]
        if o.only and o.only not in name:
            continue
        out = outdir / name
        if item.get("symbol"):
            span = find_symbol_range(code_lines, item["symbol"])
        else:
            span = (item.get("first_line") or 1, item.get("last_line") or len(code_lines))
        seg = code_lines[span[0] - 1:span[1]]
        cols = max(34, min(100, max(wide(l) for l in seg) + 1))
        if shot(out, cols, code=src_path, span=span, fs=o.fs, scale=o.scale, tool=tool):
            w, pt = recommend(out, o.fs, o.scale)
            widths[name] = w
            rows.append((name, f"{span[0]}-{span[1]}", len(seg), w, pt))
        else:
            fails += 1

    for item in spec.get("run", []):
        name = item["out"]
        if o.only and o.only not in name:
            continue
        text = (root / item["text"])
        sel = run_excerpt(text, item["from"], item.get("to"), item.get("max"))
        excdir.mkdir(parents=True, exist_ok=True)
        exc = excdir / (Path(name).stem + ".txt")
        exc.write_text("\n".join(sel) + "\n", encoding="utf-8")
        cols = max(40, min(104, max(wide(l) for l in sel) + 1))
        out = outdir / name
        if shot(out, cols, log=exc, fs=o.fs, scale=o.scale, tool=tool):
            w, pt = recommend(out, o.fs, o.scale)
            widths[name] = w
            rows.append((name, f"{len(sel)} 行", len(sel), w, pt))
        else:
            fails += 1

    (root / "outputs").mkdir(exist_ok=True)
    fw = root / "outputs" / "figure_widths.json"
    fw.write_text(json.dumps(widths, ensure_ascii=False, indent=1), encoding="utf-8")

    print(f"{'图片':<26}{'区间/行数':>10}{'行':>5}{'宽cm':>7}{'字号pt':>8}")
    for name, rng, n, w, pt in rows:
        print(f"{name:<26}{rng:>10}{n:>5}{w:>7.1f}{pt:>8.1f}")

    if o.apply_md:
        md = root / o.apply_md
        changed = apply_widths(md, widths)
        print(f"\nWROTE_WIDTH {changed} 处 → {md.name}")
    print(f"WIDTHS {fw}")
    print("VERDICT " + ("PASS" if not fails else f"FAIL: {fails} 张未生成"))
    return 0 if not fails else 1


if __name__ == "__main__":
    raise SystemExit(main())
