#!/bin/bash
# PreToolUse(Bash) 守门：
# 1. 禁止 git add . / ./ / -A / --all，强制明确指定文件。
# 2. 禁止在 main/master 上直接 git commit。
# 3. 禁止 git push --force / -f。
# 退出码约定（Claude Code）：exit 2 = 拦截；exit 0 = 放行；exit 1 = 非阻塞错误（工具仍执行）。

input=$(cat)
command=$(printf '%s' "$input" | jq -r '.tool_input.command // ""')

# 命中 git add . / git add ./ / git add -A / git add --all → 拦截
if printf '%s' "$command" | grep -qE '(^|[;&|[:space:]])git +add +(-A|--all|\.($|[ /])|\. *$)'; then
  echo "🚫 禁止 git add . / ./ / -A / --all —— 请明确指定文件路径。" >&2
  exit 2
fi

# 禁止在 main/master 上直接 commit
if printf '%s' "$command" | grep -qE '(^|[;&|[:space:]])git +commit([[:space:]]|$)'; then
  branch=$(git -C "$CLAUDE_PROJECT_DIR" branch --show-current 2>/dev/null)
  if [ "$branch" = "master" ] || [ "$branch" = "main" ]; then
    echo "🚫 禁止在 $branch 分支直接 commit。请先切到功能分支后再提交。" >&2
    exit 2
  fi
fi

# 禁止 force push
if printf '%s' "$command" | grep -qE '(^|[;&|[:space:]])git +push([^;&|]*[[:space:]])(-f|--force|--force-with-lease)([[:space:]]|$)'; then
  echo "🚫 禁止 git push --force / -f / --force-with-lease。请使用正常 push。" >&2
  exit 2
fi

exit 0
