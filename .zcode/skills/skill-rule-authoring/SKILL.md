---
name: skill-rule-authoring
description: |
  创建、编写、注册新的 Skill 或 Rule，把现有工作流沉淀成可复用的 Skill。
  when_to_use: 创建/迁移/调试 Skill 或 Rule、排查 "规则没生效"、不确定该写 Rule 还是 Skill 时。
---

# Skill / Rule 搭建标准流程

## 第一步：选载体

```
Q1: 是"流程"还是"约束"？
    ├─ 流程（多步骤、有分支）→ 【Skill】
    └─ 约束（几行到几十行规则）→ Q2

Q2: 能否用文件路径 glob 描述"什么时候需要"？
    ├─ 能（编辑某类文件/某目录时）→ 【.claude/rules/】
    └─ 不能（全局 / 错误触发型）→ 【AGENTS.md】

补充：模型会主动绕过的命令拦截 → Hook（.claude/settings.json 注册）
```

## 第二步：写 Skill

### Skill 文件结构

```
.claude/skills/<skill-name>/
├── SKILL.md          # 必需：技能定义
└── references/        # 可选：参考文档
```

### SKILL.md 模板

```markdown
---
name: my-skill
description: |
  一句话描述这个技能做什么。
  when_to_use: 触发条件 / 关键词。
---

# 技能标题

## 触发场景
- 用户说 "XXX"
- 用户问 "YYY"

## 执行流程

### Phase 1: ...
### Phase 2: ...
### Phase 3: ...
```

## 第三步：写 Rule

### Rule 文件结构

```markdown
---
paths:
  - "**/*.py"
  - "**/*.html"
---

# 规则标题

规则内容...
```

`paths` 指定 glob 匹配的文件路径，只有编辑匹配文件时规则才生效。

## 第四步：验证

- Skill：在对话中说出触发关键词，确认 Skill 被正确加载
- Rule：编辑匹配路径的文件，确认规则被引用

## 选择指南

| 场景 | 载体 | 原因 |
|------|------|------|
| 多步骤操作流程 | Skill | 需要分支和步骤 |
| 代码编写规范 | Rule（带 paths） | 编辑特定文件时生效 |
| 全局行为准则 | AGENTS.md | 始终加载 |
| 命令拦截 | Hook + settings.json | 在工具执行前/后触发 |
