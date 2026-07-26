# MCP Agent 流程图设计

## 目标

为 MCPTox 文献汇报制作一张可在 Microsoft Visio 中继续编辑的 MCP Agent 流程图，并将其作为矢量图插入当前 PPT 第 3 页上半部分，替换现有 Attack Surface Comparison 占位内容。

## Visio 源文件

- 文件名：`MCP_Agent_流程_攻击面.vsdx`。
- 页面采用横向画布，白底，使用可编辑的圆角矩形、连接线、文本和标注对象。
- 所有节点、箭头、文字和攻击面标注均为独立 Visio 形状，不使用栅格截图。

## 图形结构

- 主链：`用户` → `LLM Agent`。
- `LLM Agent` 向下分支为 `Reasoning` 与 `MCP Tool`。
- 工具支路为：`MCP Tool` → `Tool Description / Metadata` → `Tool Response`。
- `Tool Description / Metadata` 采用淡红填充、深红描边；右侧以红色虚线标注 `攻击面`。
- 图下方放置一句中文说明：`MCP通过标准化工具调用接口提升Agent扩展能力，但工具描述本身也进入LLM上下文，成为新的攻击入口。`

## 视觉与嵌入规则

- 常规节点使用 Persona 主题的深蓝描边、白底或淡蓝底；箭头为蓝灰色，攻击链路仅以红色强调。
- 术语节点使用英文，用户输入和攻击提示使用中文，字号在嵌入后的第 3 页仍可清晰阅读。
- 使用 Visio 导出的矢量 EMF 或 SVG 作为 PPT 图像，保持缩放清晰度。
- 仅替换第 3 页上半部分的现有占位区域；下半部分的“安全挑战 / 研究意义”两栏不修改。

## 验收标准

- `.vsdx` 可由 Microsoft Visio 正常打开、编辑并保存。
- 流程图在 PPT 第 3 页上半部分完整显示，不遮挡页眉、下方两栏或页码。
- 攻击面位于 Tool Description / Metadata，且红色强调清晰可辨。
- 更新后的 PPT 另存为新文件，不覆盖现有版本。
