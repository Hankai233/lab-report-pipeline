# -*- coding: utf-8 -*-
"""把结构化 Markdown 渲染进"学校实验报告模板 docx"里，输出可交付的 Word。

设计：**以模板 docx 为底稿**，保留它的 A4 页面、页边距、封面字段、一~六节标题与附页；
只往各节标题下插入正文、表格、图片和代码块。这样不需要用代码去复刻一个学校模板，
页面设置永远不会画错。

输入 md 约定
-----------
1. 顶部元信息（可选，填封面）::

       ---
       课程名称: 机器感知技术
       指导教师: 董安明
       实验日期: 2026年11月5日
       实验项目名称: 实验二 数学与信号处理基础
       ---

2. 一级标题切分"部分"：`# 实验预习报告` / `# 实验报告` / `# 附页`。
   模板里没有预习部分时，只写 `# 实验报告` 即可（单部分模式）。

3. 二级标题对应模板里的"一、xxx / 二、xxx"：文字含关键字即可自动对位
   （如 md 写 `## 二、实验原理`，模板写 `二、 实验原理` 也能对上，
   也可用 --map map.json 精确指定）。

4. 支持的块：`###/####` 小标题、段落、`- ` 项目符号、`1. ` 编号、
   `| 表格 |`、`![图注](路径){width=12cm}`、```代码块```、`> 引用`。
   紧跟在 `表 x ...` 段落后的 Markdown 表格渲染为"表题在上 + 带框表格"。

用法
----
    python md2report.py 报告.md 模板.docx 输出.docx
    python md2report.py 报告.md 模板.docx 输出.docx --map 章节映射.json
    python md2report.py --inspect 模板.docx        # 先看看模板里能对位哪些标题
    python md2report.py 报告.md 模板.docx out.docx --overwrite-cover  # 连模板里预填的值也覆盖

自检（建议每次都做）::

    soffice --headless --convert-to pdf --outdir .cache 报告.docx
    python -c "import pymupdf;d=pymupdf.open('.cache/报告.pdf');[p.get_pixmap(dpi=100).save(f'.cache/p{i}.png') for i,p in enumerate(d,1)]"
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BODY_FONT = "宋体"
MONO_FONT = "Consolas"
BODY_SIZE = Pt(12)
SMALL_SIZE = Pt(10.5)
CODE_SIZE = Pt(9.5)

PART_HINTS = {                    # 模板里三部分标题的识别关键字
    "pre": ("实验预习报告", "预习报告"),
    "main": ("实验报告",),
    "appendix": ("附页", "附录"),
}
COVER_LABELS = ("课程名称", "指导教师", "实验日期", "学部（学院）", "院（系）", "学院",
                "专业班级", "实验地点", "学生姓名", "姓名", "学号", "同组人",
                "实验项目名称", "实验题目")
CN_NUM = "一二三四五六七八九十"


# --------------------------------------------------------------------- md 解析
def split_front_matter(text: str):
    m = re.match(r"^---\s*\n(.*?)\n---\s*\n", text, re.S)
    if not m:
        return {}, text
    meta = {}
    for line in m.group(1).splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            meta[k.strip()] = v.strip()
    return meta, text[m.end():]


def parse_blocks(lines):
    blocks, i, n = [], 0, len(lines)
    while i < n:
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        if s.startswith("```"):
            lang = s[3:].strip()
            body, i = [], i + 1
            while i < n and not lines[i].strip().startswith("```"):
                body.append(lines[i].rstrip("\n"))
                i += 1
            i += 1
            blocks.append(("code", lang, body))
            continue
        if s.startswith("####"):
            blocks.append(("h4", s.lstrip("#").strip()))
            i += 1
            continue
        if s.startswith("###"):
            blocks.append(("h3", s.lstrip("#").strip()))
            i += 1
            continue
        if s.startswith(">"):
            blocks.append(("quote", s.lstrip("> ").strip()))
            i += 1
            continue
        # 表格
        if s.startswith("|") and i + 1 < n and re.match(r"^\s*\|[\s:|-]+\|\s*$", lines[i + 1]):
            header = [c.strip() for c in s.strip("|").split("|")]
            i += 2
            rows = []
            while i < n and lines[i].strip().startswith("|"):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            caption = ""
            if blocks and blocks[-1][0] == "p" and re.match(
                    rf"^表\s*[\d{CN_NUM}]", blocks[-1][1]):
                caption = blocks.pop()[1]
            blocks.append(("table", caption, header, rows))
            continue
        # 图片
        m = re.match(r"^!\[(.*?)\]\((.+?)\)(?:\{width=([\d.]+)cm\})?\s*$", s)
        if m:
            blocks.append(("figure", m.group(1), m.group(2),
                           float(m.group(3)) if m.group(3) else None))
            i += 1
            continue
        # 列表
        if re.match(r"^[-*]\s+", s):
            items = []
            while i < n and re.match(r"^\s*[-*]\s+", lines[i]):
                items.append(re.sub(r"^\s*[-*]\s+", "", lines[i]).strip())
                i += 1
            blocks.append(("bullets", items))
            continue
        if re.match(r"^\d+[.)]\s+", s):
            items = []
            while i < n:
                mm = re.match(r"^\s*\d+[.)]\s+(.*)$", lines[i])
                if not mm:
                    break
                items.append(mm.group(1).strip())
                i += 1
            blocks.append(("numbers", items))
            continue
        buf = [s]
        i += 1
        while i < n:
            t = lines[i].strip()
            if not t or t.startswith(("#", "|", "![", ">", "```")) or \
               re.match(r"^([-*]\s|\d+[.)]\s)", t):
                break
            buf.append(t)
            i += 1
        blocks.append(("p", "".join(buf)))
    return blocks


def parse_md(path: Path):
    meta, text = split_front_matter(path.read_text(encoding="utf-8"))
    parts, current, section = {}, None, None
    for raw in text.splitlines():
        line = raw.rstrip()
        if line.startswith("# ") and not line.startswith("## "):
            title = line[2:].strip()
            if "预习" in title:
                current = "pre"
            elif "附页" in title or "附录" in title:
                current = "appendix"
            else:
                current = "main"
            parts.setdefault(current, [])
            section = None
            continue
        if line.startswith("## "):
            section = {"title": line[3:].strip(), "lines": []}
            parts.setdefault(current or "main", []).append(section)
            continue
        if section is not None:
            section["lines"].append(line)
        elif current and line.strip():
            section = {"title": "__intro__", "lines": [line]}
            parts.setdefault(current, []).append(section)
    if "main" not in parts and parts:
        parts["main"] = parts.pop(next(iter(parts)))
    return meta, parts


# --------------------------------------------------------------------- 渲染
class Renderer:
    def __init__(self, template: Path, section_map: dict | None = None,
                 overwrite_cover: bool = False):
        self.doc = Document(str(template))
        self.map = section_map or {}
        self.overwrite_cover = overwrite_cover

    # ---- 基础样式
    def style(self, run, size=BODY_SIZE, bold=False, mono=False, italic=False):
        run.font.size = size
        run.font.bold = bold
        run.font.italic = italic
        name = MONO_FONT if mono else BODY_FONT
        run.font.name = name
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = rPr.makeelement(qn("w:rFonts"), {})
            rPr.append(rFonts)
        for attr in ("w:eastAsia", "w:ascii", "w:hAnsi"):
            rFonts.set(qn(attr), name)
        run.font.color.rgb = RGBColor(0, 0, 0)

    def runs(self, p, text, size=BODY_SIZE, bold=False):
        """行内 **加粗** 与 `代码`。"""
        for tok in re.split(r"(\*\*.+?\*\*|`[^`]+`)", text):
            if not tok:
                continue
            if tok.startswith("**") and tok.endswith("**"):
                self.style(p.add_run(tok[2:-2]), size=size, bold=True)
            elif tok.startswith("`") and tok.endswith("`"):
                self.style(p.add_run(tok[1:-1]), size=size, mono=True, bold=bold)
            else:
                self.style(p.add_run(tok), size=size, bold=bold)

    def para_after(self, anchor, indent=True, align=None, before=0, after=0,
                   size=BODY_SIZE, spacing=WD_LINE_SPACING.ONE_POINT_FIVE):
        p = self.doc.add_paragraph()
        anchor._p.addnext(p._p)
        pf = p.paragraph_format
        pf.line_spacing_rule = spacing
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.first_line_indent = Pt(size.pt * 2) if indent else Pt(0)
        if align is not None:
            pf.alignment = align
        return p

    # ---- 模板结构识别
    def find_part_title(self, part):
        hints = PART_HINTS[part]
        cands = [p for p in self.doc.paragraphs if p.text.strip()]
        # 先找完全含 hints 且最短的段落（避免命中正文里的同名句）
        for hint in hints:
            hits = [p for p in cands if hint in p.text]
            if part == "main":
                hits = [p for p in hits if not any(h in p.text for h in PART_HINTS["pre"])]
            if hits:
                return min(hits, key=lambda p: len(p.text.strip()))
        return None

    def part_headings(self, part) -> list:
        title = self.find_part_title(part)
        paras = self.doc.paragraphs
        start = 0
        if title is not None:
            start = next(i for i, q in enumerate(paras) if q._p is title._p) + 1
        heads = []
        for p in paras[start:]:
            txt = p.text.strip()
            if not txt:
                continue
            if title is not None and len(txt) < 40 and \
               any(h in txt for k, hs in PART_HINTS.items() if k != part for h in hs):
                if re.match(rf"^[{CN_NUM}]+[、.]", txt) is None:
                    break
            if re.match(rf"^[{CN_NUM}]+[、.]\s*\S", txt) and len(txt) < 40:
                heads.append(p)
        return heads

    def score(self, md_title: str, head_text: str) -> int:
        """md 小节标题与模板小节标题的匹配分：公共子串越长越可信。"""
        a = re.sub(r"[\s：:（）()、.，,]", "", md_title)
        b = re.sub(r"[\s：:（）()、.，,]", "", head_text)
        best = 0
        for i in range(len(a)):
            for j in range(i + best + 1, len(a) + 1):
                if a[i:j] in b:
                    best = j - i
                else:
                    break
        return best

    def match_heading(self, part, md_title, heads):
        if self.map:
            for k, v in self.map.items():
                if k in md_title:
                    return next((h for h in heads if v in h.text), None)
        best, best_p = 0, None
        for h in heads:
            sc = self.score(md_title, h.text)
            if sc > best:
                best, best_p = sc, h
        return best_p if best >= 2 else None

    # ---- 块渲染
    def render(self, anchor, blocks, base: Path):
        cur = anchor
        for b in blocks:
            kind = b[0]
            if kind == "h3":
                p = self.para_after(cur, indent=False, before=6, after=3)
                self.runs(p, b[1], bold=True)
                self.outline(p, 2)
                cur = p
            elif kind == "h4":
                p = self.para_after(cur, indent=False, before=4, after=2)
                self.runs(p, b[1], bold=True)
                cur = p
            elif kind == "p":
                p = self.para_after(cur)
                self.runs(p, b[1])
                cur = p
            elif kind == "quote":
                p = self.para_after(cur, indent=False, before=2, after=2)
                p.paragraph_format.left_indent = Pt(24)
                self.runs(p, b[1], size=SMALL_SIZE)
                cur = p
            elif kind == "bullets":
                for it in b[1]:
                    p = self.para_after(cur, indent=False)
                    p.paragraph_format.left_indent = Pt(24)
                    p.paragraph_format.first_line_indent = Pt(-12)
                    self.style(p.add_run("• "))
                    self.runs(p, it)
                    cur = p
            elif kind == "numbers":
                for i, it in enumerate(b[1], 1):
                    p = self.para_after(cur, indent=False)
                    p.paragraph_format.left_indent = Pt(24)
                    p.paragraph_format.first_line_indent = Pt(-24)
                    head = f"（{i}）" if len(b[1]) <= 6 else f"{i}. "
                    self.style(p.add_run(head))
                    self.runs(p, it)
                    cur = p
            elif kind == "code":
                cur = self.render_code(cur, b[1], b[2])
            elif kind == "table":
                cur = self.render_table(cur, b[1], b[2], b[3])
            elif kind == "figure":
                cur = self.render_figure(cur, b[1], b[2], b[3], base)
        return cur

    def render_code(self, anchor, lang, body):
        for line in body or [""]:
            p = self.para_after(anchor, indent=False, before=0, after=0,
                                size=CODE_SIZE, spacing=WD_LINE_SPACING.SINGLE)
            p.paragraph_format.left_indent = Pt(18)
            pPr = p._p.get_or_add_pPr()
            shd = pPr.makeelement(qn("w:shd"),
                                  {qn("w:val"): "clear", qn("w:fill"): "F5F5F5"})
            pPr.append(shd)
            self.style(p.add_run(line if line else " "), size=CODE_SIZE, mono=True)
            anchor = p
        p = self.para_after(anchor, indent=False, after=4, spacing=WD_LINE_SPACING.SINGLE)
        return p

    def render_table(self, anchor, caption, header, rows):
        if caption:
            p = self.para_after(anchor, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, before=6)
            self.runs(p, caption, size=SMALL_SIZE, bold=True)
            p.paragraph_format.keep_with_next = True
            anchor = p
        table = self.doc.add_table(rows=1, cols=len(header))
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = table._tbl.tblPr
        tblW = tblPr.find(qn("w:tblW"))
        if tblW is None:
            tblW = tblPr.makeelement(qn("w:tblW"), {})
            tblPr.append(tblW)
        tblW.set(qn("w:type"), "pct")
        tblW.set(qn("w:w"), "5000")

        def put(cell, text, bold=False):
            cell.text = ""
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = Pt(1)
            para.paragraph_format.space_after = Pt(1)
            para.paragraph_format.line_spacing = 1.0
            self.runs(para, text, size=SMALL_SIZE, bold=bold)

        for j, h in enumerate(header):
            put(table.rows[0].cells[j], h, bold=True)
        for row in rows:
            cells = table.add_row().cells
            for j, val in enumerate(row[:len(header)]):
                put(cells[j], val)
        trPr = table.rows[0]._tr.get_or_add_trPr()
        trPr.append(trPr.makeelement(qn("w:tblHeader"), {}))
        for row in table.rows:
            rp = row._tr.get_or_add_trPr()
            rp.append(rp.makeelement(qn("w:cantSplit"), {}))
        for row in table.rows[:-1]:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.paragraph_format.keep_with_next = True
        anchor._p.addnext(table._tbl)
        spacer = self.para_after(self.doc.paragraphs[0], indent=False,
                                 spacing=WD_LINE_SPACING.SINGLE, after=6)
        table._tbl.addnext(spacer._p)
        return spacer

    def render_figure(self, anchor, caption, rel, width_cm, base: Path):
        img = (base / rel).resolve()
        if not img.exists():
            p = self.para_after(anchor, indent=False)
            self.runs(p, f"[缺图: {rel}]")
            return p
        p = self.para_after(anchor, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, before=6,
                            spacing=WD_LINE_SPACING.SINGLE)
        p.paragraph_format.keep_with_next = True
        p.add_run().add_picture(str(img), width=Cm(width_cm or 12.0))
        cp = self.para_after(p, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, after=6,
                             spacing=WD_LINE_SPACING.SINGLE)
        self.runs(cp, caption, size=SMALL_SIZE)
        return cp

    def outline(self, p, level):
        pPr = p._p.get_or_add_pPr()
        old = pPr.find(qn("w:outlineLvl"))
        if old is not None:
            pPr.remove(old)
        pPr.append(pPr.makeelement(qn("w:outlineLvl"), {qn("w:val"): str(level)}))

    # ---- 封面字段
    def fill_cover(self, part, values, overwrite=False):
        """把 frontmatter 的值填进封面下划线里。

        模板的两种形态都要兼容：
        - 空下划线占位（多个 run 都是空格 + underline）→ 填进第一个 run，其余清空；
        - 已经有默认值（如"软件工程""吕国华"写在下划线 run 里）→ 默认不动，
          只有 overwrite=True 才替换。
        """
        title = self.find_part_title(part)
        if title is None or not values:
            return
        idx = next(i for i, q in enumerate(self.doc.paragraphs) if q._p is title._p)
        # 封面字段可能在部分标题之前（如“一、实验目的”在上方）或之后（如“实验预习报告”在下方），两边都扫
        win = self.doc.paragraphs[max(0, idx - 4): idx + 16]
        for p in win:
            runs = p.runs
            plans, acc, i = [], "", 0
            while i < len(runs):
                acc += runs[i].text
                key = acc.strip(" :：　")
                if key in COVER_LABELS:
                    if values.get(key):
                        j, group = i + 1, []
                        while j < len(runs) and runs[j].font.underline:
                            group.append(runs[j])
                            j += 1
                        if group:
                            current = "".join(r.text for r in group).strip()
                            if not current or overwrite:
                                plans.append((key, group))
                    acc = ""
                elif not any(lbl.startswith(key) or key.startswith(lbl[:2])
                             for lbl in COVER_LABELS):
                    acc = ""
                i += 1
            if not plans:
                continue

            def vw(s):
                return sum(2 if ord(c) > 127 else 1 for c in s)

            total = vw(p.text)
            for key, group in plans:
                group[0].text = values[key] + " "
                for r in group[1:]:
                    r.text = ""
            left = total - sum(vw(r.text) for r in runs)
            if left > 3:
                tail = p.add_run(" " * (left - 3))
                tail.font.size = Pt(12)
            else:
                deficit = 3 - left
                for key, group in reversed(plans):
                    if deficit <= 0:
                        break
                    extra = len(group[0].text) - len(values[key])
                    take = min(deficit, max(0, extra))
                    if take > 0:
                        group[0].text = group[0].text[:len(group[0].text) - take]
                        deficit -= take
                if deficit > 0:
                    n = sum(max(1, len(r.text)) for r in p.runs)
                    twips = -max(2, int(deficit * 120 / n))
                    for r in p.runs:
                        rPr = r._element.get_or_add_rPr()
                        sp = rPr.makeelement(qn("w:spacing"), {qn("w:val"): str(twips)})
                        rPr.append(sp)

    # ---- 组装
    def build(self, meta, parts, base: Path):
        three = self.find_part_title("pre") is not None
        targets = ("pre", "main", "appendix") if three else ("main",)
        for part in targets:
            if self.find_part_title(part) is None:
                continue
            self.fill_cover(part, meta, overwrite=self.overwrite_cover)
            if part == "appendix":
                self.render_appendix(parts.get(part, []), base)
                continue
            heads = self.part_headings(part)
            for h in heads:
                self.outline(h, 0)
            for sec in parts.get(part, []):
                if sec["title"] in ("__intro__", ""):
                    continue
                head = self.match_heading(part, sec["title"], heads)
                blocks = parse_blocks(sec["lines"])
                if not blocks:
                    continue
                if head is None:
                    print(f"[跳过] 模板中无对应小节：{part} / {sec['title']}",
                          file=sys.stderr)
                    continue
                self.render(head, blocks, base)
        # 附页另起一页
        if three:
            p = self.find_part_title("appendix")
            if p is not None:
                p.paragraph_format.page_break_before = True
        return self.doc

    def render_appendix(self, sections, base: Path):
        """附页是自由结构：模板里没有一、二、三标题，按 md 小节顺序往下写。"""
        anchor = self.find_part_title("appendix")
        if anchor is None:
            return
        # 清掉附页标题后的占位空段
        nxt = anchor._p.getnext()
        while nxt is not None and nxt.tag == qn("w:p"):
            text = "".join(t.text or "" for t in nxt.iter(qn("w:t")))
            if text.strip():
                break
            nxt.getparent().remove(nxt)
            nxt = anchor._p.getnext()
        self.outline(anchor, 0)
        for sec in sections:
            if sec["title"] not in ("__intro__", ""):
                p = self.para_after(anchor, indent=False, before=6, after=3)
                self.runs(p, sec["title"], bold=True)
                self.outline(p, 1)
                anchor = p
            anchor = self.render(anchor, parse_blocks(sec["lines"]), base)


def inspect(template: Path):
    doc = Document(str(template))
    print(f"# {template}")
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        if any(h in t for hs in PART_HINTS.values() for h in hs) or \
           re.match(rf"^[{CN_NUM}]+[、.]", t):
            print(f"  [{len(t):>2}] {t}")
    sec = doc.sections[0]
    print(f"  页面 {sec.page_width.cm:.2f}×{sec.page_height.cm:.2f} cm，"
          f"左右边距 {sec.left_margin.cm:.2f}/{sec.right_margin.cm:.2f} cm，"
          f"版心 {sec.page_width.cm - sec.left_margin.cm - sec.right_margin.cm:.2f} cm")


def main(argv):
    if len(argv) >= 2 and argv[1] == "--inspect":
        inspect(Path(argv[2]))
        return 0
    if len(argv) < 4:
        print(__doc__)
        return 2
    md, tpl, out = Path(argv[1]), Path(argv[2]), Path(argv[3])
    section_map = None
    overwrite = "--overwrite-cover" in argv
    if "--map" in argv:
        section_map = json.loads(Path(argv[argv.index("--map") + 1]).read_text("utf-8"))
    meta, parts = parse_md(md)
    doc = Renderer(tpl, section_map, overwrite_cover=overwrite).build(meta, parts, md.parent)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    print(f"[ok] {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
