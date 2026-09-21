---
description: 提交小版本（修订号+1）并打 tag 推送远程
---

请按以下流程执行：

1. 执行 commit/tag 前必做检查（三条命令全部跑完再继续）：
   - `git log --oneline -5`
   - `git tag -l | sort -V | tail -5`
   - `git diff --stat`
2. 根据检查结果确认：
   - 改动范围是否符合预期（只包含目标文件）
   - 版本号在最新 tag 基础上修订号 +1（如 v1.0.7 → v1.0.8）
3. 提交改动：
   - 禁止 `git add .` 或 `git add -A`，明确指定文件
   - commit message 格式：`[模块名] 简要描述`
4. 打 tag：`v<x>.<y>.<z>`
5. 推送分支和 tag 到远程
