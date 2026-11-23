# -*- coding: utf-8 -*-
"""docx 摆版自检：一次量出页数、每页填充率、插图跨页、图注跑页、空白页、末页空白。

为什么要有：交付前“打开翻一遍”靠肉眼，往往翻到第 6 遍才发现第 15 页只剩一行。
这里把肉眼判断变成数字，一次改完再跑一次就收敛。

用法
----
    python pagelayout.py 实验报告.docx
    python pagelayout.py 实验报告.docx --out .cache/layout.txt --json .cache/layout.json
    python pagelayout.py 实验报告.docx --soffice "C:\\Program Files\\LibreOffice\\program\\soffice.exe"

需要 LibreOffice（或已安装 soffice 在 PATH）。没有也能用，只是量不了摆版。

退出码：0 通过（或仅有 warn）；1 有硬问题；2 转换失败。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

WIN = [r"C:\Program Files\LibreOffice\program\soffice.exe",
       r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]
BODY_W = 14.66      # 版心宽 cm（与 SKILL.md 一致）


def find_soffice(explicit: str | None) -> str | None:
    for cand in (explicit, os.environ.get("LAB_SOFFICE")):
        if cand and Path(cand).exists():
            return cand
    for name in ("soffice", "libreoffice", "soffice.exe"):
        exe = shutil.which(name)
        if exe:
            return exe
    for p in WIN:
        for cand in (p, "/mnt/" + p[0].lower() + p[2:].replace("\\", "/")):
            if Path(cand).exists():
                return cand
    return None


def native(p: Path) -> str:
    """WSL 里调 Windows 的 soffice.exe 时，路径必须翻成 Windows 形式。"""
    if os.name == "nt":
        return str(p)
    if shutil.which("wslpath"):
        r = subprocess.run(["wslpath", "-w", str(p)], capture_output=True, text=True)
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout.strip()
    return str(p)


def to_pdf(docx: Path, soffice: str, workdir: Path) -> Path:
    profile = workdir / "lo_profile"
    profile.mkdir(parents=True, exist_ok=True)
    win_exe = soffice.endswith(".exe") and os.name != "nt"
    cmd = [soffice, "--headless", "--norestore", "--nolockcheck",
           f"-env:UserInstallation={profile.as_uri() if not win_exe else 'file:///' + native(profile).replace(chr(92), '/')}",
           "--convert-to", "pdf",
           "--outdir", native(workdir) if win_exe else str(workdir),
           native(docx.resolve()) if win_exe else str(docx.resolve())]
    subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    pdf = workdir / (docx.stem + ".pdf")
    for _ in range(120):
        if pdf.exists() and pdf.stat().st_size > 0:
            return pdf
        time.sleep(0.5)
    raise SystemExit(f"[pdf] 转换失败：{pdf}（检查 soffice 能否启动、docx 是否被 Word 占用）")


def analyse(pdf: Path, body_w_cm: float) -> dict:
    import pymupdf
    doc = pymupdf.open(str(pdf))
    pages = []
    for i, page in enumerate(doc, 1):
        r = page.rect
        h = r.height
        blocks = [b for b in page.get_text("blocks") if b[4].strip()]
        imgs = [imu["bbox"] for imu in page.get_image_info()]
        texts = [b for b in blocks]
        bottom = max([b[3] for b in texts] + [im[3] for im in imgs] + [0])
        fill = bottom / h if h else 0
        issues = []
        if not blocks and not imgs:
            issues.append("blank")
        for (x0, y0, x1, y1) in imgs:
            wcm = (x1 - x0) * 2.54 / 72
            hcm = (y1 - y0) * 2.54 / 72
            if wcm > body_w_cm + 0.15:
                issues.append(f"img_wide:{wcm:.1f}cm")
            if hcm > 20.0:
                issues.append(f"img_tall:{hcm:.1f}cm")
            if y0 < 1.5 or y1 > h - 1.5:
                issues.append("img_straddle")
            below = [b for b in texts if b[1] > y1 - 1]
            if y1 > h * 0.82 and not below:
                issues.append("caption_next_page?")
        if texts and not imgs and fill < 0.12 and i != 1:
            issues.append("nearly_empty")
        pages.append({"page": i, "fill": round(fill, 3), "images": len(imgs),
                      "text_blocks": len(texts), "issues": sorted(set(issues))})
    # 末页
    if pages and pages[-1]["fill"] < 0.20 and len(pages) > 1:
        pages[-1]["issues"].append("last_page_short")
    hard = {"img_straddle", "img_wide", "blank", "blank_page"}
    fails = [p for p in pages if any(str(i).split(":")[0] in hard for i in p["issues"])]
    return {"pages": len(pages), "body_width_cm": body_w_cm,
            "detail": pages, "fail_pages": [p["page"] for p in fails]}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="docx → PDF → 摆版体检")
    ap.add_argument("docx")
    ap.add_argument("--soffice", default=None)
    ap.add_argument("--out", default=".cache/layout_report.txt", help="UTF-8 详情")
    ap.add_argument("--json", default=None)
    ap.add_argument("--body-width", type=float, default=BODY_W)
    ap.add_argument("--keep-pdf", action="store_true")
    o = ap.parse_args(argv or sys.argv[1:])

    docx = Path(o.docx)
    if not docx.exists():
        print(f"找不到 {docx}", file=sys.stderr)
        return 2
    soffice = find_soffice(o.soffice)
    if not soffice:
        print("MISS soffice：装 LibreOffice 或传 --soffice 路径；本项跳过", file=sys.stderr)
        return 0

    workdir = Path(o.keep_pdf and docx.parent or tempfile.mkdtemp(prefix="layout_"))
    workdir.mkdir(parents=True, exist_ok=True)
    pdf = to_pdf(docx.resolve(), soffice, workdir)
    res = analyse(pdf, o.body_width)

    lines = [f"# 摆版自检 {docx.name}", f"页数 {res['pages']}，版心宽 {res['body_width_cm']} cm", ""]
    lines.append("| 页 | 填充率 | 图 | 文本块 | 问题 |")
    lines.append("| --- | --- | --- | --- | --- |")
    for p in res["detail"]:
        lines.append(f"| {p['page']} | {p['fill']*100:.0f}% | {p['images']} | "
                     f"{p['text_blocks']} | {' / '.join(p['issues']) or '-'} |")
    lines.append("")
    if res["fail_pages"]:
        lines.append(f"硬问题页：{res['fail_pages']}")
        lines.append("img_straddle=插图跨页；img_wide=超版心；blank=空白页；"
                     "caption_next_page?=图注可能跑到下一页；last_page_short=末页近乎空白")
    else:
        lines.append("无硬问题。")
    Path(o.out).parent.mkdir(parents=True, exist_ok=True)
    Path(o.out).write_text("\n".join(lines) + "\n", encoding="utf-8")
    if o.json:
        Path(o.json).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")

    for p in res["detail"]:
        flag = ",".join(p["issues"]) or "-"
        print(f"PAGE {p['page']:>2} fill={p['fill']*100:5.1f}% img={p['images']} {flag}")
    print(f"\nPAGES {res['pages']}")
    print(f"DETAIL {Path(o.out).resolve()}")
    print("VERDICT " + ("PASS" if not res["fail_pages"] else
                        "FAIL pages=" + ",".join(map(str, res["fail_pages"]))))
    return 0 if not res["fail_pages"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
