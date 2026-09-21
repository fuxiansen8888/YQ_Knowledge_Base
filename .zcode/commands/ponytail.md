---
description: 切换 ponytail 懒人开发模式强度（lite/full/ultra）或关闭（off）
---

请执行以下步骤：

1. 运行 `node .zcode/hooks/persona-config.js set ponytail "$1"` 设置强度。
   - 若未带参数，先运行 `node .zcode/hooks/persona-config.js get ponytail` 显示当前强度，再告知用户可选值：`lite` / `full` / `ultra` / `off`。
   - 若命令因 cwd 不对而失败，先 `cd` 到项目根目录再执行。
2. 向用户报告：新强度已写入 `.ponytail/config.json`，将在**下次新会话**自动生效；本次会话可用口头指令 "ponytail" 临时开启、说 "normal mode" 关闭。
