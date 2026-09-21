---
description: 同步 main/master 最新代码到当前分支并推送远程
---

请按以下流程执行：

1. 记录当前分支名：`git branch --show-current`
2. 暂存当前工作区（如有未提交改动）：`git stash`
3. 切到 main 拉最新：
   ```
   git checkout main
   git pull origin main
   ```
4. 切回自己的分支：`git checkout <之前记录的分支名>`
5. 合并 main：`git merge main`
   - 如果有冲突，逐个文件查看冲突内容，优先保留自己的改动
   - 解决完所有冲突后 `git add` 冲突文件并 `git commit`
