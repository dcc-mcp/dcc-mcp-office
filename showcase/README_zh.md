# dcc-mcp-office 全家桶 Showcase

这里展示的是可执行证据，不是路线图拼贴。每个 Demo 都同时提供结构化输入、
可编辑 Office 产物、真实渲染预览、Host/MCP 结果、验证说明和 SHA-256。

| Demo | 证明的能力 | 入口 |
|---|---|---|
| 模板优先的 Deck Pipeline | Presentation IR → 外置 `brand://` 模板 → 可编辑 PPTX → PowerPoint 原生预览 → 溢出报告 | [预览与产物](./deck-pipeline/) · [元数据](./deck-pipeline/metadata.json) |
| 品牌模板对比 | 同一份 Presentation IR → 三个带 Aptos Display/Aptos 字体角色的版本化 `brand://` 包 → 三份可编辑 PPTX → 12 张 PowerPoint 渲染 → 确定性质量门 | [模板与预览](./template-gallery/) · [质量报告](./template-gallery/quality-report.json) |
| 富图片语义布局 | 六张原创编辑式素材 → 图片封面 + `image_left_text_right` + 非对称 `image_grid` → 可编辑 PPTX 媒体 → PowerPoint 原生渲染 | [PPTX 与预览](./image-rich-deck/) · [素材输入](./image-rich-deck/assets/) · [生成清单](./image-rich-deck/asset-manifest.json) |
| 生产能力 Dashboard | Workbook IR → 图片化 KPI 区 + 原生能力图表 + 可编辑流程轨道 → XLSX → Excel 原生 PDF 预览 | [预览与产物](./production-dashboard/) · [元数据](./production-dashboard/metadata.json) |
| Word 管理层简报 | 结构化内容 + 原创编辑式横幅 → Aptos 字体层级 → 可编辑 DOCX → Word 检查 → 原生 PDF → 全页视觉验收 | [预览与产物](./word-executive-brief/) · [元数据](./word-executive-brief/metadata.json) |
| 安全全局文本替换 | PowerPoint + Word + Excel dry-run → 确认门 → 字节级检查点 → 提交后复验 | [前后文件](./global-text-replace/) · [元数据](./global-text-replace/metadata.json) |
| 混合 Office 批量转 PDF | PowerPoint、Word、Excel 独立 COM Sidecar → 同一批次的已验证 PDF | [输入与 PDF](./batch-to-pdf/) · [元数据](./batch-to-pdf/metadata.json) |

完整复现需要 Windows 交互式用户会话，并安装 PowerPoint、Word、Excel：

```powershell
vx run build
$env:PATH = "<poppler-bin>;$env:PATH"
python scripts/capture_showcases.py --with-office --force
python scripts/validate_showcases.py
```

没有桌面 Office 时，`vx run self-test` 仍可验证确定性的 Presentation IR →
Open XML → inspect 往返。Hosted CI 只校验仓库中已提交的 Gallery，不会把
这一步冒充成真实 Office COM 验收。
