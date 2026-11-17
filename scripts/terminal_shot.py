# -*- coding: utf-8 -*-
"""把一段代码 / 终端输出渲染成图片（默认：纯白底、无窗口装饰、只显示代码本身）。

报告里需要"代码/运行结果"的图时用它重绘，而不是截屏：分辨率可控、可脱敏、换机器能重出。
默认遵守"除代码外不出现任何其他文字"——没有标题栏、没有提示符、没有页脚、没有行号。

用法
----
    # 代码文件（只显示代码）
    python terminal_shot.py --out 图/code.png --code main.py --first-line 1 --last-line 40

    # 终端输出（只显示输出文本）
    python terminal_shot.py --out 图/run.png --log 日志.txt --split 40

    # 直接给文本
    python terminal_shot.py --out demo.png --text "np.arange(5)" "array([0, 1, 2, 3, 4])"

    # 需要行号 / 深色 / 脱敏 / 窗口装饰时才显式加参数
    python terminal_shot.py --code main.py --out c.png --lineno
    python terminal_shot.py --log run.log --out r.png --theme dark
    python terminal_shot.py --log run.log --out r.png --mask "C:\\\\Users\\\\[^\\\\]+"
    python terminal_shot.py --text "..." --out t.png --chrome "main.py"
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

FONTS = {
    "latin": [r"C:\Windows\Fonts\consola.ttf", r"C:\Windows\Fonts\lucon.ttf",
              "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"],
    "cjk": [r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simsun.ttc",
            r"C:\Windows\Fonts\simhei.ttf",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"],
}

THEMES = {  # light 默认：白底黑字，适合 Word 打印
    "light": {"bg": (255, 255, 255), "bar": (245, 245, 245), "bar_text": (60, 60, 60),
              "text": (24, 24, 24), "dim": (130, 130, 130), "err": (190, 40, 40),
              "warn": (170, 120, 10), "ok": (20, 120, 60), "border": (215, 215, 215)},
    "dark": {"bg": (24, 26, 30), "bar": (39, 42, 48), "bar_text": (200, 204, 210),
             "text": (226, 230, 236), "dim": (140, 148, 158), "err": (240, 118, 118),
             "warn": (238, 200, 110), "ok": (126, 200, 128), "border": (60, 64, 70)},
}


def load_font(kind: str, size: int):
    for path in FONTS[kind]:
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    return ImageFont.load_default()


def is_wide(ch: str) -> bool:
    return ord(ch) > 0x2E7F or ch in "“”‘’—…"


def wrap_line(line: str, cols: int) -> list[str]:
    if cols <= 8:
        return [line]
    out, cur, w = [], "", 0
    for ch in line.expandtabs(4):
        cw = 2 if is_wide(ch) else 1
        if w + cw > cols:
            out.append(cur)
            cur, w = ch, cw
        else:
            cur += ch
            w += cw
    out.append(cur)
    return out


def read_lines(a) -> list[str]:
    if a.code:
        src = Path(a.code).read_text(encoding="utf-8", errors="replace").splitlines()
        lo = max(1, a.first_line or 1)
        hi = a.last_line or len(src)
        return [src[i - 1] for i in range(lo, min(hi, len(src)) + 1)]
    if a.text:
        return list(a.text)
    if a.log:
        return Path(a.log).read_text(encoding="utf-8", errors="replace").splitlines()
    if not sys.stdin.isatty():
        return sys.stdin.read().splitlines()
    return []


def draw(lines, o) -> Image.Image:
    th = THEMES[o.theme]
    s = max(1, o.scale)
    fs = o.fs * s
    cell = int(fs * 0.62)
    line_h = int(fs * 1.5)
    pad = int(fs * 0.9)
    bar_h = int(fs * 1.9) if o.chrome else 0

    f_lat, f_cjk = load_font("latin", fs), load_font("cjk", fs)
    gutter = int(cell * (len(str(len(lines))) + 2)) if o.lineno else 0
    W = o.cols * cell + pad * 2 + gutter
    H = bar_h + max(1, len(lines)) * line_h + pad * 2

    img = Image.new("RGB", (W, H), th["bg"])
    d = ImageDraw.Draw(img)
    if o.chrome:
        d.rectangle([0, 0, W, bar_h], fill=th["bar"])
        f_bar = load_font("cjk", int(fs * 0.85))
        d.text((pad, bar_h / 2), o.chrome, font=f_bar, fill=th["bar_text"],
               anchor="lm")

    err_rx = re.compile(o.highlight) if o.highlight else None
    y = bar_h + pad
    for i, line in enumerate(lines):
        x = pad
        if o.lineno:
            d.text((x, y), f"{i + 1}", font=f_lat, fill=th["dim"], anchor="lt")
            x += gutter
        color = th["text"]
        if err_rx and err_rx.search(line):
            color = th["err"]
        elif re.search(r"(WARNING|警告|WARN)", line):
            color = th["warn"]
        elif re.search(r"(PASS|通过|完成|saved|Saved|保存)", line):
            color = th["ok"]
        for ch in line:
            if is_wide(ch):
                d.text((x, y), ch, font=f_cjk, fill=color)
                x += cell * 2
            else:
                d.text((x, y), ch, font=f_lat, fill=color)
                x += cell
        y += line_h

    if o.include_frame:
        d.rectangle([0, 0, W - 1, H - 1], outline=th["border"])
    return img


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="代码/终端输出 → 图片（默认白底、只含代码）")
    ap.add_argument("--out", required=True, help="输出 PNG")
    ap.add_argument("--code", help="代码文件（只渲染代码行）")
    ap.add_argument("--log", help="终端输出日志")
    ap.add_argument("--text", nargs="*", help="直接给出的若干行")
    ap.add_argument("--cols", type=int, default=96, help="每行字符格数（默认 96）")
    ap.add_argument("--fs", type=int, default=16)
    ap.add_argument("--scale", type=int, default=2, help="放大倍数（默认 2）")
    ap.add_argument("--theme", choices=sorted(THEMES), default="light")
    ap.add_argument("--lineno", action="store_true", help="显示行号（默认不显示）")
    ap.add_argument("--chrome", metavar="TEXT", help="加一条标题栏并写下这段文字（默认不加）")
    ap.add_argument("--frame", dest="include_frame", action="store_true",
                    help="加一圈细边框（默认没有）")
    ap.add_argument("--split", type=int, default=0, help="每 N 行切一张")
    ap.add_argument("--no-wrap", dest="wrap", action="store_false")
    ap.add_argument("--first-line", type=int, default=1)
    ap.add_argument("--last-line", type=int, default=0)
    ap.add_argument("--highlight", default=r"(错误|失败|Error|ERROR|Traceback|Exception|Failed)")
    ap.add_argument("--mask", action="append", default=[], help="脱敏正则，可重复")
    ap.set_defaults(wrap=True)
    o = ap.parse_args(argv or sys.argv[1:])

    lines = read_lines(o)
    if not lines:
        print("没有内容：请给 --code / --log / --text 之一", file=sys.stderr)
        return 2
    for pat in o.mask:
        rx = re.compile(pat)
        lines = [rx.sub("***", ln) for ln in lines]
    if o.wrap:
        lines = [w for ln in lines for w in wrap_line(ln, o.cols)]

    chunks = [lines] if not o.split else \
        [lines[i:i + o.split] for i in range(0, len(lines), o.split)]
    out = Path(o.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    for i, chunk in enumerate(chunks, 1):
        img = draw(chunk, o)
        path = out if len(chunks) == 1 else out.with_name(f"{out.stem}_{i}{out.suffix}")
        img.save(path)
        print(f"[ok] {path}  {img.width}x{img.height}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
