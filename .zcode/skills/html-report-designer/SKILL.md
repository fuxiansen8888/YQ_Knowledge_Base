---
name: html-report-designer
description: |
  把分析结果做成自包含 HTML 报告（单文件，内联 CSS/JS，可直接浏览器打开）。
  适用场景：数据可视化报告、分析报告、趋势分析、数据对比等。
  when_to_use: 用户要求 "生成报告""做个报告""出个 HTML""数据报告""分析报告" 时。
---

# HTML 数据分析报告设计与生成

输出**单文件 self-contained HTML**，内联所有 CSS/JS，可直接浏览器打开或分享。

## 触发模式

### 模式 A：承接数据分析流程
数据查询/分析完成后，用户同意生成报告时触发。此时已有数据上下文。

### 模式 B：独立触发
用户直接要求生成报告，需先确认数据来源和展示目标。

## 设计原则

1. **自包含**：所有 CSS/JS/数据内联，不依赖 CDN 或外部文件
2. **响应式**：适配不同屏幕宽度
3. **可分享**：一个 HTML 文件即完整报告

## 代码规范

- 数据用 `<script type="application/json">` 块内嵌，JS 通过 JSON.parse 读取
- 每个图表独立 try-catch，单图报错不影响其他
- 离线可用：禁止 CDN 加载 ECharts，内联或 base64
- 详见 `.claude/rules/html-echart-safety.md`

## 常用图表库

- ECharts（推荐，内联使用）
- Chart.js（轻量替代）

## 报告结构模板

```
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><title>报告标题</title></head>
<body>
  <!-- 头部摘要 -->
  <!-- 图表区域 -->
  <!-- 数据表格 -->
  <!-- 结论 -->
  <script type="application/json" id="data">...</script>
  <script>/* 内联 echarts.min.js 或图表库 */</script>
  <script>/* 渲染逻辑 */</script>
</body>
</html>
```
