---
description: 关闭 caveman 与 ponytail 模式，恢复正常表达
---

请执行以下步骤：

1. 运行 `node .zcode/hooks/persona-config.js set caveman off`。
2. 运行 `node .zcode/hooks/persona-config.js set ponytail off`。
3. 若命令因 cwd 不对而失败，先 `cd` 到项目根目录再执行。
4. 向用户报告：两个模式已关闭，**下次新会话**起恢复正常风格；本次会话说 "normal mode" 即可立即恢复。
