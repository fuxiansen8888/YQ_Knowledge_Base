---
name: git-workflow-guide
description: |
  Git 提交、切分支、打 tag、解决冲突等场景的完整指南。
  when_to_use: 执行 git 提交/切分支/打 tag、commit message 规范、版本号规范、push 被拒等异常处理时。
---

# Git 工作流

## 核心原则

- 先同步 main/master 最新代码，再创建并使用自己的工作分支
- **禁止在 main/master 上直接 commit**：main 仅用于 pull 和 merge，所有改动必须落在功能分支
- 禁止使用 `git add .` 或 `git add -A`，明确指定要提交的文件
- 禁止 `git push --force`

## 日常四步走

```bash
# 1. 切到 main 并拉最新
git checkout main
git pull origin main

# 2. 创建自己的功能分支（首次）
git checkout -b <分支名>

# 3. 正常开发 + 提交
git add <具体文件路径>
git commit -m "[模块名] 简要描述"

# 4. 推送
git push origin <分支名>
```

## Commit Message 规范

格式：`[模块名] 简要描述`

- 模块名用中文或英文，保持项目内一致
- 描述简短清晰，说明做了什么
- 示例：`[股票分析] 添加K线均线叠加功能`

## 分支命名

- 功能分支：`feat/<描述>` 或 `<姓名>/<描述>`
- 修复分支：`fix/<描述>`

## Tag 版本管理

- 格式：`v<主版本>.<次版本>.<修订号>`（如 v1.2.3）
- 重大变更 → 主版本 +1
- 新功能 → 次版本 +1
- Bug 修复 → 修订号 +1

## 异常处理

- 误在 main 上 commit：`git stash` → `git switch -c <新分支>` → `git stash pop`
- Push 被拒：先 `git pull --rebase` 再 push
- 冲突解决：逐文件查看，理解双方改动后再合并，不盲目覆盖
