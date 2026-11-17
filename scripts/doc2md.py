# -*- coding: utf-8 -*-
"""实验要求类文档（pdf / pptx / docx / html / epub）→ Markdown。

按顺序尝试，前一步能出可用文本就停：
  1. **pandoc** —— 首选。对 pptx / docx / html / epub 效果好，能保留标题层级与表格。
     （pandoc 读不了 PDF：`pandoc x.pdf` 会直接报 "PDF is not a supported input format"，
       所以 PDF 实际上会立即落到第 2 步，这符合"先试 pandoc、失败换手段"的约定。）
  2. **pymupdf 抽文本层**（仅 PDF）—— 扫描件没有文本层时会抽不到东西。
  3. **逐页 PNG + 提示** —— 抽不出时把每页渲染成图片，提示改用视觉/OCR 读图，并点名哪几页。

用法
----
    python doc2md.py 实验02-实验指导书.pdf --out docs
    python doc2md.py 课件.pptx --out docs --dpi 130
    python doc2md.py 指导书.docx --out docs --keep-html     # 保留 pandoc 的中间产物

输出（--out 目录下）
    <stem>.md          转换结果
    <stem>_pNN.png     需要看图时逐页导出的图片
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

PANDOC_FORMATS = {".pptx": "pptx", ".docx": "docx", ".html": "html", ".htm": "html",
                  ".epub": "epub", ".odt": "odt", ".rtf": "rtf"}


def run_pandoc(src: Path, out_md: Path, keep_html: bool) -> str | None:
    exe = shutil.which("pandoc")
    if not exe:
        print("[1/3] 未安装 pandoc，跳过", file=sys.stderr)
        return None
    fmt = PANDOC_FORMATS.get(src.suffix.lower())
    if fmt is None:
        print(f"[1/3] pandoc 不支持直接从 {src.suffix} 读取，跳过", file=sys.stderr)
        return None
    cmd = [exe, str(src), "-f", fmt, "-t", "gfm", "--wrap=none",
           "--extract-media", str(out_md.parent / (out_md.stem + "_media")),
           "-o", str(out_md)]
    print("[1/3] pandoc:", " ".join(cmd[:5]), "…")
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       errors="replace")
    if r.returncode != 0:
        print(f"      pandoc 失败：{r.stderr.strip()[:200]}", file=sys.stderr)
        return None
    text = out_md.read_text(encoding="utf-8", errors="replace")
    if len(text.strip()) < 80:
        print("      pandoc 输出过短，视为失败", file=sys.stderr)
        return None
    print(f"      成功：{len(text)} 字")
    return text


def run_pymupdf(src: Path) -> tuple[str, list[int]]:
    import pymupdf
    doc = pymupdf.open(str(src))
    pages, thin = [], []
    for i, page in enumerate(doc, 1):
        t = page.get_text().strip()
        pages.append(f"## 第 {i} 页\n\n{t}" if t else f"## 第 {i} 页\n\n（本页无文本层）")
        if len(t) < 40:
            thin.append(i)
    return "\n\n".join(pages), thin


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="实验要求文档 → Markdown（pandoc 优先）")
    ap.add_argument("src", help="pdf / pptx / docx / html / epub")
    ap.add_argument("--out", default=".", help="输出目录")
    ap.add_argument("--dpi", type=int, default=120, help="抽不出文本时的页面图 dpi")
    ap.add_argument("--keep-html", action="store_true", help="保留 pandoc 的中间文件")
    o = ap.parse_args(argv or sys.argv[1:])

    src = Path(o.src)
    if not src.exists():
        print(f"找不到 {src}", file=sys.stderr)
        return 2
    out_dir = Path(o.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    md_path = out_dir / f"{src.stem}.md"

    text = run_pandoc(src, md_path, o.keep_html) if src.suffix.lower() in PANDOC_FORMATS else None

    if text is None and src.suffix.lower() == ".pdf":
        print("[2/3] pymupdf 抽文本层")
        text, thin = run_pymupdf(src)
        md_path.write_text(text, encoding="utf-8")
        print(f"      {len(text)} 字；无文本层页：{thin or '无'}")
        if thin:
            print("[3/3] 这些页需要看图：导出逐页 PNG")
            import pymupdf
            doc = pymupdf.open(str(src))
            for i in thin:
                doc[i - 1].get_pixmap(dpi=o.dpi).save(
                    str(out_dir / f"{src.stem}_p{i:02d}.png"))
            print(f"      已导出 {len(thin)} 张 → {out_dir}")
            print("      下一步：用视觉/OCR 读这些页，把要点补进 md 对应位置")

    if text is None:
        print(f"[!] 未能转换 {src.name}：请改用 OCR/多模态识读", file=sys.stderr)
        return 1

    if not md_path.exists():
        md_path.write_text(text, encoding="utf-8")
    print(f"[ok] {md_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
