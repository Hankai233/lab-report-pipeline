# -*- coding: utf-8 -*-
"""实验要求 PDF/扫描件 → 逐页文本 +（可选）逐页 PNG，并给出"哪几页需要看图"的体检结果。

为什么要这一步：实验要求常常是老师发的扫描件 PDF。先抽文本，能搜就能定位"要做什么"；
抽不出文本的页面自动标出来，提醒这些页必须用视觉读图（或转 PNG 再看），别漏掉关键要求。

用法
----
    # 抽文本，输出 <目录>/要求文本.txt
    python pdf_extract.py 实验要求.pdf --out docs

    # 抽文本 + 每页 PNG（dpi 120），便于逐页看
    python pdf_extract.py 3.2-3.6.pdf --out docs --png --dpi 120

    # 只看体检报告，不落盘
    python pdf_extract.py 3.2-3.6.pdf --report

输出
----
    docs/<stem>_文本.txt     每页前有 ===== PAGE n ===== 分隔
    docs/<stem>_pXX.png      逐页图（--png 时）
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pymupdf


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="PDF → 逐页文本 + 逐页 PNG")
    ap.add_argument("pdf", help="输入 PDF")
    ap.add_argument("--out", default=".", help="输出目录（默认当前目录）")
    ap.add_argument("--png", action="store_true", help="同时导出逐页 PNG")
    ap.add_argument("--dpi", type=int, default=120, help="PNG 分辨率，默认 120")
    ap.add_argument("--report", action="store_true", help="只打印体检结果，不写文件")
    ap.add_argument("--min-chars", type=int, default=40,
                    help="一页文字少于该值即视为“疑似纯图页”")
    opts = ap.parse_args(argv or sys.argv[1:])

    src = Path(opts.pdf)
    if not src.exists():
        print(f"找不到文件：{src}", file=sys.stderr)
        return 2
    doc = pymupdf.open(str(src))
    out_dir = Path(opts.out)
    stem = src.stem

    texts, thin = [], []
    for i, page in enumerate(doc, 1):
        t = page.get_text().strip()
        texts.append(f"===== PAGE {i} =====\n{t}")
        if len(t) < opts.min_chars:
            thin.append(i)

    total = sum(len(t) for t in texts)
    print(f"页数 {len(doc)}，抽出文字 {total} 字")
    if thin:
        print(f"疑似纯图页（文字 < {opts.min_chars} 字）：{thin}"
              f" → 这些页必须看图/OCR，建议加 --png 后逐页读图")
    else:
        print("所有页面都有可用文本层")

    if opts.report:
        return 0

    out_dir.mkdir(parents=True, exist_ok=True)
    txt = out_dir / f"{stem}_文本.txt"
    txt.write_text("\n\n".join(texts), encoding="utf-8")
    print(f"[ok] {txt}")

    if opts.png:
        for i, page in enumerate(doc, 1):
            png = out_dir / f"{stem}_p{i:02d}.png"
            page.get_pixmap(dpi=opts.dpi).save(str(png))
        print(f"[ok] {len(doc)} 张页面 PNG → {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
