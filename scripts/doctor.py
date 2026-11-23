# -*- coding: utf-8 -*-
"""一条命令看清这台机器能用哪个解释器、缺哪个库、缺哪个外部工具，并给出修复命令。

为什么需要它：实验报告流水线会同时用到 python-docx / pymupdf / Pillow / matplotlib 与
pandoc / LibreOffice。Windows 上常有多个解释器（`py -3` 指向的解释器往往一个包都没有），
逐个试错要花十几轮。先跑这个脚本，一次拿到结论。

约定
----
- 控制台只输出 ASCII 状态（避免 Windows 控制台 GBK 乱码）；
- 中文详情写入 UTF-8 报告文件（默认 ./doctor_report.txt），用 read 工具查看。

用法
----
    python doctor.py                      # 自检 + 写 doctor_report.txt
    python doctor.py --out "$env:TEMP/doctor.txt"
    python doctor.py --no-wsl             # 不探测 WSL 解释器
    python doctor.py --json               # 机器可读结果（stdout）

退出码：0 全部硬需求满足；1 有缺项。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

HARD = [("docx", "python-docx", "渲染/读取 Word 模板"),
        ("pymupdf", "pymupdf", "读 PDF 文本层、docx→PDF 自检"),
        ("PIL", "Pillow", "白底代码/运行图（terminal_shot.py）"),
        ("matplotlib", "matplotlib", "报告插图")]
SOFT = [("lxml", "lxml", "docx 加表格样式时更稳")]

TOOLS = [("pandoc", "docx/pptx/网页 → md（PDF 不需要它）"),
         ("soffice", "docx → PDF，摆版自检 pagelayout.py 用"),
         ("pdftotext", "PDF 抽文本的备用手段")]

WIN_SOFFICE = [r"C:\Program Files\LibreOffice\program\soffice.exe",
               r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]


def probe_module(name: str) -> tuple[bool, str]:
    try:
        mod = __import__(name)
        ver = getattr(mod, "__version__", "")
        return True, str(ver)
    except Exception as e:                     # noqa: BLE001 - 任何导入异常都算缺
        return False, f"{type(e).__name__}: {e}"[:120]


def probe_tool(name: str) -> tuple[bool, str]:
    for cand in ([name, name + ".exe", "libreoffice"] if name == "soffice" else [name]):
        exe = shutil.which(cand)
        if exe:
            return True, exe
    if name == "soffice":
        for p in WIN_SOFFICE:
            # WSL 里路径要换成 /mnt/c/... 才看得见 Windows 的 soffice
            for cand in (p, "/mnt/" + p[0].lower() + p[2:].replace("\\", "/")):
                if Path(cand).exists():
                    return True, cand
    return False, ""


def probe_interp(cmd: list[str]) -> dict:
    code = "import sys,json;" \
           "d={};\n" \
           "for m in ('docx','pymupdf','PIL','matplotlib'):\n" \
           "    try:\n" \
           "        __import__(m); d[m]=1\n" \
           "    except Exception: d[m]=0\n" \
           "print(json.dumps({'v':sys.version.split()[0],'exe':sys.executable,'m':d}))"
    try:
        r = subprocess.run(cmd + ["-c", code], capture_output=True, text=True,
                           timeout=40, encoding="utf-8", errors="replace")
        if r.returncode != 0:
            return {"ok": False, "err": (r.stderr or "").strip().splitlines()[-1][:120]
                    if (r.stderr or "").strip() else "run failed"}
        return {"ok": True, **json.loads(r.stdout.strip().splitlines()[-1])}
    except Exception as e:                     # noqa: BLE001
        return {"ok": False, "err": f"{type(e).__name__}: {e}"[:120]}


def find_interpreters(no_wsl: bool) -> list[tuple[str, dict]]:
    out = [("current", probe_interp([sys.executable]))]
    seen = {os.path.normcase(sys.executable)}
    for cand in ("python", "python3"):
        exe = shutil.which(cand)
        if exe and os.path.normcase(exe) not in seen:
            seen.add(os.path.normcase(exe))
            out.append((cand, probe_interp([exe])))
    if os.name == "nt" and shutil.which("py"):
        r = probe_interp(["py", "-3"])
        out.append(("py -3", r))
    if not no_wsl and shutil.which("wsl"):
        code = ("import sys,json;d={}\n"
                "for m in ('docx','pymupdf','PIL','matplotlib'):\n"
                "    try:\n"
                "        __import__(m);d[m]=1\n"
                "    except Exception:d[m]=0\n"
                "print(json.dumps({'v':sys.version.split()[0],'exe':sys.executable,'m':d}))")
        try:
            r = subprocess.run(["wsl", "-e", "python3", "-c", code], capture_output=True,
                               text=True, timeout=40, encoding="utf-8", errors="replace")
            if r.returncode == 0 and r.stdout.strip():
                out.append(("wsl python3",
                            {"ok": True, **json.loads(r.stdout.strip().splitlines()[-1])}))
            else:
                out.append(("wsl python3", {"ok": False, "err": "wsl 不可用或无 python3"}))
        except Exception as e:                 # noqa: BLE001
            out.append(("wsl python3", {"ok": False, "err": f"{type(e).__name__}"}))
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="实验报告流水线环境自检")
    ap.add_argument("--out", default="doctor_report.txt", help="UTF-8 详情报告")
    ap.add_argument("--no-wsl", action="store_true")
    ap.add_argument("--json", action="store_true", help="只打印 JSON")
    o = ap.parse_args(argv or sys.argv[1:])

    mods = {m: probe_module(m) for m, _, _ in HARD + SOFT}
    tools = {t: probe_tool(t) for t, _ in TOOLS}
    interps = find_interpreters(o.no_wsl)

    missing_hard = [m for m, _, _ in HARD if not mods[m][0]]
    missing_soft = [m for m, _, _ in SOFT if not mods[m][0]]
    py = sys.executable

    result = {"interpreter": py, "python": sys.version.split()[0],
              "modules": {m: {"ok": mods[m][0], "detail": mods[m][1]} for m in mods},
              "tools": {t: {"ok": tools[t][0], "path": tools[t][1]} for t in tools},
              "interpreters": {k: v for k, v in interps},
              "missing_hard": missing_hard, "missing_soft": missing_soft,
              "note": "详情写入 UTF-8 报告文件，见 --out"}

    if o.json:
        print(json.dumps(result, ensure_ascii=False, indent=1))
        return 1 if missing_hard else 0

    L: list[str] = ["# 环境自检", "", f"当前解释器：{py}（Python {sys.version.split()[0]}）",
                    f"工作目录：{Path.cwd()}", ""]

    L.append("## 解释器与库")
    L.append("| 解释器 | 版本 | docx | pymupdf | Pillow | matplotlib |")
    L.append("| --- | --- | --- | --- | --- | --- |")
    for label, r in interps:
        if not r.get("ok"):
            L.append(f"| {label} | — | 不可用：{r.get('err', '')} | | | |")
            continue
        mm = r.get("m", {})
        cells = ["OK" if mm.get(k) else "MISS" for k in ("docx", "pymupdf", "PIL", "matplotlib")]
        L.append(f"| {label} | {r.get('v', '')} | " + " | ".join(cells) + " |")
    L.append("")

    L.append("## 硬需求（缺一不可）")
    for m, pkg, why in HARD:
        ok, detail = mods[m]
        mark = "OK  " if ok else "MISS"
        L.append(f"- [{mark}] {pkg}（{why}）{detail if ok else '→ ' + detail}")
    L.append("")

    L.append("## 软需求")
    for m, pkg, why in SOFT:
        ok, detail = mods[m]
        L.append(f"- [{'OK  ' if ok else 'MISS'}] {pkg}（{why}）")
    L.append("")

    L.append("## 外部工具")
    for t, why in TOOLS:
        ok, path = tools[t]
        L.append(f"- [{'OK  ' if ok else 'MISS'}] {t}（{why}）{path if ok else ''}")
    L.append("")

    if missing_hard:
        L.append("## 修复命令（装到当前解释器）")
        L.append(f'    "{py}" -m pip install ' +
                 " ".join(n for m, n, _ in HARD if m in missing_hard) +
                 " " + " ".join(n for m, n, _ in SOFT if m in missing_soft))
        L.append("")
        L.append("Windows 上若 `py -3` 与 `python` 不是同一个解释器，务必把上面这条命令里的")
        L.append("解释器路径换成上面表格里 MISS 最少的那一行，否则会出现“装了却导不进来”。")
        L.append("")
        L.append("WSL 里装（Debian/Ubuntu，PEP 668 需要 --break-system-packages）：")
        L.append("    python3 -m pip install --break-system-packages pymupdf python-docx "
                 "matplotlib pillow lxml")
        L.append("")
    else:
        L.append("硬需求全部满足。")
        best = None
        for label, r in interps:
            if r.get("ok") and all(r.get("m", {}).get(m) for m in ("docx", "pymupdf", "PIL")):
                best = (label, r["exe"])
                break
        if best:
            L.append(f"建议执行入口：{best[0]} → {best[1]}")
            L.append(f'    $PY = "{best[1]}"     # 后面所有脚本都用这个解释器')
        L.append("")

    if not tools["soffice"][0]:
        L.append("## 关于 soffice")
        L.append("pagelayout.py（摆版自检）需要 LibreOffice 或 Word。没有也能出报告，只是")
        L.append("不能自动量“每页填充率/插图跨页/末页空白”，需要人工翻 docx。")
        L.append("")

    L.append("## 输出编码")
    L.append("Windows 控制台默认 GBK，Python 打印中文会乱码。跑脚本时加：")
    L.append("    $env:PYTHONIOENCODING=\"utf-8\"   # 或把结果重定向到文件后用 read 查看")

    Path(o.out).write_text("\n".join(L) + "\n", encoding="utf-8")

    for m, _, _ in HARD:
        ok = mods[m][0]
        print(f"{'OK  ' if ok else 'MISS'} {m:<12} {mods[m][1] if ok else ''}")
    for t, _ in TOOLS:
        ok = tools[t][0]
        print(f"{'OK  ' if ok else 'MISS'} {t:<12} {tools[t][1] if ok else ''}")
    print(f"\nDETAIL_FILE {Path(o.out).resolve()}")
    print("VERDICT " + ("PASS" if not missing_hard else "FAIL: " + ",".join(missing_hard)))
    return 0 if not missing_hard else 1


if __name__ == "__main__":
    raise SystemExit(main())
