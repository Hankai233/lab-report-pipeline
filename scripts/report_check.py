# -*- coding: utf-8 -*-
"""报告交付前的自动核对：图/表编号、插图文件、宽度、占位、AI 字样、docx 渲染是否丢东西。

为什么要有：md 里一行 `![图](x.png){width=12cm}{width=12cm}` 会让渲染器把整行当普通段落，
图**静默消失**（Word 里只剩一行字）；编号跳号、正文引用不存在的图、残留"待填"，肉眼很难一次
看全。这里逐条查，一次报全。

用法
----
    python report_check.py 实验报告.md --docx 实验报告.docx --metrics outputs/metrics.json
    python report_check.py 实验报告.md            # 只查 md

退出码：0 全部通过；1 有失败项。
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import re
import sys
from collections import Counter
from pathlib import Path

BODY_W = 14.66
AI_WORDS = ("AI", "大模型", "人工智能", "ChatGPT", "GPT-", "Copilot", "助手生成")
PLACEHOLDER = re.compile(r"待填|待补充|待完善|TODO|TBD|XXX+|此处省略|（略）|\(略\)|待定")
FIG_CAP = re.compile(r"!\[图\s*(\d+)[^\]]*\]")
TAB_CAP = re.compile(r"^表\s*(\d+)\s", re.M)
REF_FIG = re.compile(r"图\s*(\d+)")
REF_TAB = re.compile(r"表\s*(\d+)")
IMG = re.compile(r"!\[([^\]]*)\]\(([^)\s]+)\)(\{[^}]*\})?")
TOKEN = re.compile(r"\{\{\s*([\w.]+)\s*\}\}")


def load_md2report(md: Path):
    for base in (md.parent / "scripts", Path(__file__).parent):
        p = base / "md2report.py"
        if p.exists():
            spec = importlib.util.spec_from_file_location("m2r", p)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    raise SystemExit("找不到 md2report.py（应在同目录 scripts/ 下）")


def get(d: dict, dotted: str):
    cur = d
    for k in dotted.split("."):
        if not isinstance(cur, dict) or k not in cur:
            return None
        cur = cur[k]
    return cur


def check(md: Path, docx: Path | None, metrics: Path | None, allow_ai: bool,
          body_w: float = BODY_W) -> list[tuple[str, str, str]]:
    """返回 [(level, code, message)]，level ∈ OK / WARN / FAIL。"""
    out: list[tuple[str, str, str]] = []
    text = md.read_text(encoding="utf-8")
    m2r = load_md2report(md)

    # ---- 1. 解析（与渲染器同一套逻辑，能提前暴露静默丢图）
    try:
        meta, parts = m2r.parse_md(md)
    except SystemExit as e:                       # md2report 对格式错误直接退出
        out.append(("FAIL", "parse", str(e)))
        return out
    except Exception as e:                        # noqa: BLE001
        out.append(("FAIL", "parse", f"{type(e).__name__}: {e}"))
        return out

    figs, tabs, blocks_total = [], [], 0
    for part, secs in parts.items():
        for sec in secs:
            blocks = m2r.parse_blocks(sec["lines"])
            blocks_total += len(blocks)
            for b in blocks:
                if b[0] == "figure":
                    figs.append((part, b[2], b[1]))
                elif b[0] == "table":
                    tabs.append((part, b[1], len(b[3])))
    out.append(("OK", "blocks", f"{blocks_total} 个内容块，图 {len(figs)}，表 {len(tabs)}"))

    # ---- 2. 插图文件存在
    lost = [p for _part, p, _cap in figs if not (md.parent / p).exists()]
    if lost:
        out.append(("FAIL", "figure_missing", "找不到图片：" + ", ".join(lost)))
    else:
        out.append(("OK", "figure_exists", f"{len(figs)} 张图全部存在"))

    # ---- 3. md 里的图片行是否都能被识别（防 `}` 重复/多属性导致静默丢图）
    raw_imgs = IMG.findall(text)
    if len(raw_imgs) != len(figs):
        names = {p.split("/")[-1] for _p, p, _c in figs}
        unparsed = [m.group(0)[:70] for m in IMG.finditer(text)
                    if m.group(2).split("/")[-1] not in names]
        out.append(("FAIL", "figure_parse",
                    f"md 有 {len(raw_imgs)} 个图片行，只解析出 {len(figs)} 张；"
                    f"疑似格式错误行：{unparsed[:3]}"))
    else:
        out.append(("OK", "figure_parse", "图片行全部被识别"))

    # ---- 4. 图/表编号连续 & 题注存在
    caps = [int(n) for n in FIG_CAP.findall(text)]
    if caps:
        want = list(range(1, len(caps) + 1))
        if caps != want:
            out.append(("FAIL", "figure_number", f"图编号应为 1..{len(caps)}，实际 {sorted(caps)}"))
        else:
            out.append(("OK", "figure_number", f"图 1..{len(caps)} 连续"))
    no_cap = [p for _part, p, cap in figs if not str(cap).startswith("图")]
    if no_cap:
        out.append(("WARN", "figure_caption", "这些图没有以「图 n」开头的题注：" + ", ".join(no_cap)))

    tno = [int(n) for n in TAB_CAP.findall(text)]
    if tno and tno != list(range(1, len(tno) + 1)):
        out.append(("FAIL", "table_number", f"表编号应为 1..{len(tno)}，实际 {tno}"))
    elif tno:
        out.append(("OK", "table_number", f"表 1..{len(tno)} 连续"))

    # ---- 5. 正文引用 vs 实际编号
    body = re.sub(r"!\[[^\]]*\]\([^)]+\)(\{[^}]*\})?", "", text)
    body = TAB_CAP.sub("", body)
    ref_fig = {int(n) for n in REF_FIG.findall(body)}
    bad = sorted(n for n in ref_fig if n not in set(caps))
    if bad:
        out.append(("FAIL", "figure_ref", f"正文引用了不存在的图：{bad}（实有 {sorted(set(caps))}）"))
    elif ref_fig:
        out.append(("OK", "figure_ref", f"正文引用图 {sorted(ref_fig)} 都存在"))

    ref_tab = {int(n) for n in REF_TAB.findall(body)}
    badt = sorted(n for n in ref_tab if n not in set(tno))
    if badt:
        out.append(("FAIL", "table_ref", f"正文引用了不存在的表：{badt}（实有 {tno}）"))

    # ---- 6. 宽度：重复属性 / 超版心
    width_n = Counter(m.group(2).split("/")[-1] for m in IMG.finditer(text) if m.group(3))
    dup = [k for k, v in width_n.items() if v > 1]
    if dup:
        out.append(("FAIL", "width_dup", f"同一图出现多次 width：{dup}"))
    wide = []
    for m in IMG.finditer(text):
        if not m.group(3):
            continue
        wm = re.search(r"width=([\d.]+)cm", m.group(3))
        if wm and float(wm.group(1)) > body_w + 0.01:
            wide.append((m.group(2).split("/")[-1], float(wm.group(1))))
    if wide:
        out.append(("FAIL", "width_overflow", f"超版心 {body_w} cm：{wide}"))

    # ---- 7. 占位 / AI 字样
    hits = sorted({m.group(0) for m in PLACEHOLDER.finditer(text)})
    if hits:
        out.append(("FAIL", "placeholder", f"残留占位字样：{hits}"))
    if not allow_ai:
        ai = [w for w in AI_WORDS if w in text]
        if ai:
            out.append(("FAIL", "ai_word", f"出现 AI 相关字样：{ai}（实验要求明确鼓励才能写）"))

    # ---- 8. 封面字段
    need = ("课程名称", "指导教师", "实验日期", "实验项目名称")
    miss = [k for k in need if k not in meta]
    if miss:
        out.append(("WARN", "cover", f"frontmatter 缺：{miss}"))
    else:
        out.append(("OK", "cover", "封面四项齐全"))

    # ---- 9. metrics 占位符
    toks = set(TOKEN.findall(text))
    if toks and not metrics:
        out.append(("FAIL", "metrics", f"md 有 {len(toks)} 个 {{{{...}}}} 占位，但没传 --metrics"))
    elif toks and metrics:
        data = json.loads(metrics.read_text(encoding="utf-8"))
        bad = [t for t in sorted(toks) if get(data, t) is None]
        if bad:
            out.append(("FAIL", "metrics", f"metrics 里没有这些键：{bad}"))
        else:
            out.append(("OK", "metrics", f"{len(toks)} 个占位符都能取到值"))

    # ---- 10. docx 落地核对
    if docx:
        if not docx.exists():
            out.append(("FAIL", "docx", f"找不到 {docx}"))
        else:
            try:
                from docx import Document
                d = Document(str(docx))
            except Exception as e:                # noqa: BLE001
                out.append(("FAIL", "docx", f"打不开：{type(e).__name__}: {e}"))
                return out
            n_img = len(d.inline_shapes)
            if n_img != len(figs):
                out.append(("FAIL", "docx_figure",
                            f"docx 里 {n_img} 张图，md 里 {len(figs)} 张（有图没插进去）"))
            else:
                out.append(("OK", "docx_figure", f"docx 插图 {n_img} 张与 md 一致"))
            if len(d.tables) != len(tabs):
                out.append(("WARN", "docx_table",
                            f"docx 表格 {len(d.tables)} 个，md 里 {len(tabs)} 个"))
            junk = [p.text.strip()[:60] for p in d.paragraphs
                    if p.text.strip().startswith("![") or "{{" in p.text]
            if junk:
                out.append(("FAIL", "docx_junk", f"docx 里有未渲染的 md 残留：{junk[:3]}"))
            else:
                out.append(("OK", "docx_junk", "没有未渲染的 md 残留"))
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="报告交付前自动核对")
    ap.add_argument("md")
    ap.add_argument("--docx", default=None)
    ap.add_argument("--metrics", default=None)
    ap.add_argument("--out", default=".cache/report_check.txt", help="UTF-8 详情")
    ap.add_argument("--allow-ai", action="store_true", help="实验要求鼓励用 AI 时才加")
    ap.add_argument("--body-width", type=float, default=BODY_W,
                    help="版心宽 cm，取 md2report --inspect 打印的值")
    o = ap.parse_args(argv or sys.argv[1:])

    md = Path(o.md)
    if not md.exists():
        print(f"找不到 {md}", file=sys.stderr)
        return 2
    rows = check(md, Path(o.docx) if o.docx else None,
                 Path(o.metrics) if o.metrics else None, o.allow_ai, o.body_width)

    lines = [f"# 报告核对 {md.name}", ""]
    for level, code, msg in rows:
        lines.append(f"- [{level}] {code}: {msg}")
    bad = [r for r in rows if r[0] == "FAIL"]
    lines.append("")
    lines.append(f"结论：{'通过' if not bad else f'{len(bad)} 项失败'}")
    Path(o.out).parent.mkdir(parents=True, exist_ok=True)
    Path(o.out).write_text("\n".join(lines) + "\n", encoding="utf-8")

    for level, code, msg in rows:
        print(f"{level:<4} {code:<16} {msg[:110]}")
    print(f"\nDETAIL {Path(o.out).resolve()}")
    print("VERDICT " + ("PASS" if not bad else f"FAIL {len(bad)}"))
    return 0 if not bad else 1


if __name__ == "__main__":
    raise SystemExit(main())
