#!/bin/bash
# PostToolUse(Bash) 通知 hook：
# 长任务（python3 脚本、git push、pip install 等）完成后发送 macOS 桌面通知。
# 退出码：exit 0 = 放行（PostToolUse 不阻塞，仅做通知副作用）。

input=$(cat)
command=$(printf '%s' "$input" | jq -r '.tool_input.command // ""')

is_long_task=false

# python3 执行脚本（排除 --version 等瞬时命令）
if printf '%s' "$command" | grep -qE '(^|[;&|[:space:]])python3[[:space:]]+((-m|[^;&|]+\.py)[^;&|]*)'; then
  is_long_task=true
fi

# git push
if printf '%s' "$command" | grep -qE '(^|[;&|[:space:]])git +push([[:space:]]|$)'; then
  is_long_task=true
fi

# pip install
if printf '%s' "$command" | grep -qE '(^|[;&|[:space:]])pip3? +install([[:space:]]|$)'; then
  is_long_task=true
fi

# curl/wget 下载
if printf '%s' "$command" | grep -qE '(^|[;&|[:space:]])(curl|wget)[[:space:]]'; then
  is_long_task=true
fi

if [ "$is_long_task" = true ]; then
  short_cmd=$(printf '%s' "$command" | cut -c1-60)
  osascript -e "display notification \"$short_cmd\" with title \"Claude Code\" subtitle \"任务完成 ✓\"" 2>/dev/null &
fi

exit 0
