# 示例产物

`demo_code.png`、`demo_output_1..4.png` 是脚本的真实输出，都是**白底、只有代码/输出文字**，
无标题栏、无边框、无提示符、无页脚——直接可以插进 Word。

生成命令：

```powershell
# 代码（只截代码行）
py scripts/terminal_shot.py --out examples/demo_code.png `
   --code "E:\projects\machine\ch01\case01_minimal_pipeline\main.py" --first-line 1 --last-line 26 --cols 84

# 运行输出（长日志按 18 行切 4 张）
py scripts/terminal_shot.py --out examples/demo_output.png `
   --log "E:\projects\machine\docs\实验二\01_信号时频分析\日志.txt" --split 18 --cols 84
```

需要时才算的开关：`--lineno`（行号）、`--chrome "文件名"`（标题栏）、`--theme dark`（深色）、
`--frame`（细边框）、`--mask`（脱敏）。
