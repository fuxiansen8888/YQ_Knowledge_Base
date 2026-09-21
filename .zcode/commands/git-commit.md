---
description: 提交当前改动并推送远程（不打 tag）
---

请按以下流程执行：

1. 执行提交前检查：
   - `git status` — 查看所有变更文件
   - `git diff --stat` — 确认改动范围
   - `git log --oneline -3` — 查看最近 commit 风格
2. 按需 add 文件：
   - 禁止 `git add .` 或 `git add -A`
   - 明确指定要提交的文件路径
3. commit message 格式：`[模块名] 简要描述`
4. 推送当前分支到远程
5. 最后 `git status` 确认状态
