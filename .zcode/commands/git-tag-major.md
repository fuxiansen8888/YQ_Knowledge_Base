---
description: 提交大版本（主版本+1，次版本和修订号归零）并打 tag 推送远程
---

请按以下流程执行：

1. 执行 commit/tag 前必做检查（三条命令全部跑完再继续）：
   - `git log --oneline -5`
   - `git tag -l | sort -V | tail -5`
   - `git diff --stat`
2. 根据检查结果确认：
   - 改动范围是否符合预期（只包含目标文件）
   - 版本号在最新 tag 基础上主版本 +1，次版本和修订号归零（如 v1.8.0 → v2.0.0）
3. 提交改动：
   - 禁止 `git add .` 或 `git add -A`，明确指定文件
   - commit message 格式：`[模块名] 简要描述`
4. 打 tag：`v<x>.0.0`
5. 推送分支和 tag 到远程
