# -*- coding: utf-8 -*-
"""字体解析：字体名 → 字体文件；以及从 docx 模板里读出它实际用的字体。

两个脚本共用：
  - `md2report.py` 用 `docx_fonts()` 让插进 docx 的正文/代码沿用模板字体；
  - `terminal_shot.py` 用 `find_file()` 把模板字体落到具体文件，截图与正文看起来是同一套字。

只用标准库；只有 `docx_fonts()` 需要 python-docx（没装时返回空 dict，调用方走默认值）。
"""
from __future__ import annotations

import sys
from pathlib import Path

# 字体名（小写、去空格）→ 各平台候选文件，第一个存在的胜出。
# 顺序按"平台内最像的等价字体"排：Windows 用原名，macOS/Linux 用系统自带的等价字。
FONT_FILES: dict[str, list[str]] = {
    # ---- 中文
    "宋体": [r"C:\Windows\Fonts\simsun.ttc",
             "/System/Library/Fonts/Supplemental/Songti.ttc",
             "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
             "/usr/share/fonts/truetype/arphic/uming.ttc"],
    "新宋体": [r"C:\Windows\Fonts\simsun.ttc"],
    "黑体": [r"C:\Windows\Fonts\simhei.ttf",
             "/System/Library/Fonts/STHeiti Medium.ttc",
             "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
             "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"],
    "微软雅黑": [r"C:\Windows\Fonts\msyh.ttc",
                 "/System/Library/Fonts/PingFang.ttc",
                 "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
                 "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"],
    "楷体": [r"C:\Windows\Fonts\simkai.ttf",
             "/System/Library/Fonts/Supplemental/Kaiti.ttc",
             "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
             "/usr/share/fonts/truetype/arphic/ukai.ttc"],
    "仿宋": [r"C:\Windows\Fonts\simfang.ttf",
             "/System/Library/Fonts/Supplemental/STFangsong.ttf",
             "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc"],
    "等线": [r"C:\Windows\Fonts\Deng.ttf",
             "/System/Library/Fonts/PingFang.ttc",
             "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"],
    # ---- 西文等宽（代码用）
    "consolas": [r"C:\Windows\Fonts\consola.ttf",
                 "/System/Library/Fonts/Menlo.ttc",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"],
    "couriernew": [r"C:\Windows\Fonts\cour.ttf",
                   "/System/Library/Fonts/Supplemental/Courier New.ttf",
                   "/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf",
                   "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"],
    "menlo": ["/System/Library/Fonts/Menlo.ttc",
              r"C:\Windows\Fonts\consola.ttf",
              "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"],
    "dejavusansmono": ["/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
                       "/usr/share/fonts/TTF/DejaVuSansMono.ttf",
                       "/System/Library/Fonts/Menlo.ttc",
                       r"C:\Windows\Fonts\consola.ttf"],
    # ---- 西文比例（正文用）
    "timesnewroman": [r"C:\Windows\Fonts\times.ttf",
                      "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
                      "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf"],
}

ALIASES = {
    "simsun": "宋体", "宋体（正文）": "宋体", "songti": "宋体", "songtisc": "宋体",
    "simhei": "黑体", "heitisc": "黑体", "heiti": "黑体",
    "msyh": "微软雅黑", "microsoftyahei": "微软雅黑", "pingfang": "微软雅黑",
    "pingfangsc": "微软雅黑",
    "simkai": "楷体", "kaiti": "楷体", "kaitisc": "楷体",
    "simfang": "仿宋", "fangsong": "仿宋",
    "dengxian": "等线", "deng": "等线",
    "courier": "couriernew", "couriernew": "couriernew",
    "consola": "consolas",
    "times": "timesnewroman", "timesroman": "timesnewroman",
    "dejavusansmono": "dejavusansmono",
}

# 模板里出现这些名字时，说明"代码"本来就用等宽西文字体
MONO_HINTS = {"consolas", "couriernew", "courier", "consola", "menlo",
              "dejavusansmono", "lucidaconsole", "monaco"}

DEFAULT_BODY = "宋体"
DEFAULT_MONO = "consolas"


def norm(name: str) -> str:
    """字体名归一：去空格、去引号、转小写，再套别名。返回的一律是小写或中文名。"""
    key = (name or "").strip().strip("'\"").replace(" ", "").lower()
    return ALIASES.get(key, key)


def find_file(name_or_path: str) -> str | None:
    """字体名或字体文件路径 → 实际存在的文件路径；找不到返回 None。"""
    if not name_or_path:
        return None
    p = Path(name_or_path)
    if p.is_file():                      # 直接给了路径
        return str(p)
    key = norm(name_or_path)
    if key in FONT_FILES:                # 归一后命中表（中文名归一后仍是中文名）
        for cand in FONT_FILES[key]:
            if Path(cand).is_file():
                return cand
        return None
    # 名字没在表里：仍然按路径试一次
    return str(p) if p.is_file() else None


def docx_fonts(path: str | Path) -> dict[str, str]:
    """从 docx 模板读字体名：{'body': 正文字体, 'mono': 西文字体, 'size_pt': 正文字号}。

    读 Normal 样式的 rFonts（eastAsia / ascii / hAnsi），读不到再扫前几十个段落的 run。
    没有 python-docx 或模板打不开时返回空 dict，调用方用默认值。
    """
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except Exception:
        return {}

    try:
        doc = Document(str(path))
    except Exception as exc:
        print(f"[warn] 读不了模板 {path}：{exc}", file=sys.stderr)
        return {}

    out: dict[str, str] = {}

    def take(el, key_prefix=""):
        """从 rPr/rFonts 取字体名。"""
        if el is None:
            return
        rpr = el.find(qn("w:rPr")) if el.tag != qn("w:rPr") else el
        if rpr is None:
            return
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            return
        east = rfonts.get(qn("w:eastAsia"))
        ascii_ = rfonts.get(qn("w:ascii")) or rfonts.get(qn("w:hAnsi"))
        if east and "body" not in out:
            out["body"] = east
        if ascii_ and "mono" not in out and norm(ascii_) in MONO_HINTS:
            out["mono"] = ascii_
        if ascii_ and "ascii" not in out:
            out["ascii"] = ascii_

    # 1) Normal 样式
    try:
        take(doc.styles["Normal"].element)
    except Exception:
        pass
    # 2) 正文前若干段落（模板常把字体直接设在 run 上，而不是样式）
    for p in list(doc.paragraphs)[:80]:
        if out.get("body") and out.get("mono"):
            break
        take(p._p.find(qn("w:pPr")))
        for r in p.runs[:3]:
            take(r._r)
    return out


def describe(name: str) -> str:
    """给日志用：说明这个名字最终落到哪个文件。"""
    f = find_file(name)
    return f"{name} → {f}" if f else f"{name} → （本机没有等价字体）"
