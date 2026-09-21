---
paths:
  - "**/*report*.py"
  - "**/*html*.py"
  - "**/*.html"
---

# HTML / ECharts 报告安全（三条铁律）

Python 生成含 ECharts 的 HTML 报告时，违反任一条 = 图表静默空白：

1. **数据与 JS 分离**：数据 `json.dumps` 进 `<script type="application/json">` 块，JS 用 `JSON.parse` 读取，f-string 只注入标量。禁止 f-string 内嵌 JS 对象字面量。
2. **每个图表 try-catch 隔离**：单图报错不影响其他图表渲染。
3. **离线可用**：禁止 CDN 加载 ECharts（离线环境不可用），必须内联 `echarts.min.js` 或 base64；图表数据禁止外部 fetch，必须内嵌。
