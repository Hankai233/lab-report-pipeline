# -*- coding: utf-8 -*-
"""把脚本化输入（.in）与实际输出（.log）合成为可读的会话记录（transcript）。

为什么需要它：报告里的“运行结果”图不该是屏幕截图（分辨率低、有窗口装饰、换机器重不出来）。
做法是让程序跑一遍脚本化输入落日志，再按日志里的提示符把命令行回填进去——与真实交互一致，
且能一键重生成。

关键点（返工都出在这里）
----------------------
1. **配对必须由日志内容决定**：被拒绝的 write 不会读取数据行，所以“write 命令后面一定有
   一行数据”是错的。本脚本只在块内出现 data 提示时才消耗下一行。
2. **提示符必须独占一行**：程序打印提示符后要换行再 flush，否则提示符与首行输出粘在一起，
   这里就切不开。约定见 SKILL.md「产物格式契约」。
3. 超长数据行（几十 KB 的写入内容）标 `（此处截断，共 N 字节）`，不整段贴进报告。

用法
----
    python make_transcript.py --in outputs/scenarios --log outputs/logs --out outputs/transcripts
    python make_transcript.py --in a.in --log a.log --out a.txt
    python make_transcript.py ... --prompt-regex '^/?\\S*> ?' --data-marker '请输入|input:'

退出码：0 成功；2 参数错误。
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

ANSI = re.compile(r"\x1b\[[0-9;?]*[A-Za-z]")
DEFAULT_PROMPT = r"^[^\s>]{0,120}> ?"          # 形如 "path> " 的提示符
DEFAULT_DATA = r"请输入|请输入写入内容|input[:：]|Input[:：]"
MAXDATA = 72                                    # 超过即标注截断
WIDTH = 104


def wrap(line: str, width: int = WIDTH) -> list[str]:
    if len(line) <= width:
        return [line]
    return [line[i:i + width] for i in range(0, len(line), width)]


def split_blocks(lines: list[str], prompt: re.Pattern) -> tuple[list[str], list[list[str]]]:
    """返回（首个提示符之前的内容, 每个命令块）。块首行是提示符之后的那段输出。"""
    pre, blocks, cur = [], [], None
    for line in lines:
        m = prompt.match(line)
        if m:
            cur = [line[m.end():]]
            blocks.append(cur)
        elif cur is None:
            pre.append(line)
        else:
            cur.append(line)
    return pre, blocks


def build(inp: Path, log: Path, prompt_rx: str, data_rx: str, width: int,
          head: int = 3) -> list[str]:
    cmds = [l.rstrip("\r") for l in inp.read_text(encoding="utf-8").splitlines()]
    raw = ANSI.sub("", log.read_text(encoding="utf-8", errors="replace"))
    pre, blocks = split_blocks(raw.splitlines(), re.compile(prompt_rx))
    data_re = re.compile(data_rx)

    out = [w for line in pre[:head] for w in wrap(line, width)]
    if out:
        out.append("")
    idx = 0
    for blk in blocks:
        cmd = cmds[idx] if idx < len(cmds) else "<未匹配到命令>"
        idx += 1
        out.append("> " + cmd)
        payload = None
        if any(data_re.search(x) for x in blk):
            payload = cmds[idx] if idx < len(cmds) else ""
            idx += 1
        for line in blk:
            out.extend(wrap(line, width))
            if payload is not None and data_re.search(line):
                shown = payload if len(payload) <= MAXDATA else payload[:48] + " ……"
                note = "" if len(payload) <= MAXDATA else f"（此处截断，共 {len(payload.encode())} 字节）"
                out.append("  " + shown + note)
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=".in + .log → 可读会话记录")
    ap.add_argument("--in", dest="inp", required=True, help="输入脚本 .in 或目录")
    ap.add_argument("--log", required=True, help="日志 .log 或目录")
    ap.add_argument("--out", required=True, help="输出 .txt 或目录")
    ap.add_argument("--prompt-regex", default=DEFAULT_PROMPT)
    ap.add_argument("--data-marker", default=DEFAULT_DATA)
    ap.add_argument("--width", type=int, default=WIDTH)
    ap.add_argument("--head", type=int, default=3, help="开头保留几行初始化输出")
    o = ap.parse_args(argv or sys.argv[1:])

    inp, log, out = Path(o.inp), Path(o.log), Path(o.out)
    if inp.is_dir():
        pairs = []
        for f in sorted(inp.glob("*.in")):
            lg = (log / (f.stem + ".log")) if log.is_dir() else log
            if lg.exists():
                pairs.append((f, lg, out / (f.stem + ".txt") if out.suffix == "" else out))
        if not pairs:
            print(f"没有找到配对：{inp}/*.in ↔ {log}/*.log", file=sys.stderr)
            return 2
        out.mkdir(parents=True, exist_ok=True)
        for f, lg, dst in pairs:
            dst = dst if dst.suffix else out / (f.stem + ".txt")
            lines = build(f, lg, o.prompt_regex, o.data_marker, o.width, o.head)
            dst.write_text("\n".join(lines) + "\n", encoding="utf-8")
            print(f"OK {dst.name} {len(lines)} 行")
        return 0

    if not log.exists():
        print(f"找不到日志：{log}", file=sys.stderr)
        return 2
    out.parent.mkdir(parents=True, exist_ok=True)
    lines = build(inp, log, o.prompt_regex, o.data_marker, o.width, o.head)
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"OK {out} {len(lines)} 行")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
