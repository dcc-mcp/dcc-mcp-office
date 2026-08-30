# dcc-mcp-office

DCC-MCP 生态的 Office 通用底座：为 PowerPoint / Word / Excel / Outlook /
Visio / Project / Access 提供任务级 MCP 能力（批量转 PDF、批量替换、从零
生成文档、操作当前文档等）。

本仓库实现[架构方案](./docs/proposals/office-automation-platform-v1.0.md)中的
**共享核心**：`office-rpc/1` 线协议、各应用文档 IR、C# STA COM Sidecar
运行时（`office-host`）、Open XML 批处理 Worker、Microsoft Graph 连接器和
Office 通用 Skill Pack。应用专属适配层放在薄的兄弟仓库中（如同
`dcc-mcp-maya` 依赖 `dcc-mcp-core` 一样依赖本仓库）：

| 仓库 | 范围 | 状态 |
|---|---|---|
| **dcc-mcp-office**（本仓库） | 协议 / IR / C# 运行时 / Open XML / Graph / 通用技能 | 仓库 M1 完成；M2 共享核心已具备 |
| dcc-mcp-PowerPoint | Deck 生成、Slide 编排、评审 Deck | 方案 Phase 1；外部适配器 |
| dcc-mcp-word | 重排、域、目录 | 方案 Phase 2；外部适配器 |
| dcc-mcp-excel | 计算、表格、图表、Graph Workbook | 方案 Phase 2；外部适配器 |
| dcc-mcp-outlook | 草稿、日历 | 规划中的方案 Phase 3；外部适配器 |

## 可执行 Showcase 展廊

![dcc-mcp-office 将受治理的结构化输入转换为可编辑的 PowerPoint、Word、Excel 与经过验证的 PDF 证据](./docs/images/office-suite-showcase.webp)

仓库内置七个真实案例，覆盖当前 PowerPoint、Word、Excel 全链路。每个
案例都提供可编辑源文件、Office 原生渲染、脱敏后的 `office-rpc/1` 记录与
SHA-256 清单；展示结果来自本仓库的 Open XML 与桌面 COM 路径，不使用替代
渲染器冒充 Office 结果。

| 演示系统 | 文档与数据 |
|---|---|
| [![同一个故事通过三套项目自有品牌模板生成](./showcase/template-gallery/preview.png)](./showcase/template-gallery/) | [![精美 Word 简报经原生 Word 导出](./showcase/word-executive-brief/preview.png)](./showcase/word-executive-brief/) |
| [![原创现代图片素材进入语义化幻灯片布局](./showcase/image-rich-deck/preview.png)](./showcase/image-rich-deck/) | [![可编辑 Excel 能力仪表盘经原生 Excel 渲染](./showcase/production-dashboard/preview.png)](./showcase/production-dashboard/) |

进入[完整 Showcase 展廊](./showcase/README_zh.md)，还可以查看模板优先的
Deck 流水线、跨 Office 批量转 PDF、带确认门禁与检查点的全局替换，以及
对应的可复现验证证据。Graph、Office.js 和后续应用适配器仍不在当前展廊的
实证范围内。

## 核心原则

1. Rust 网关（既有 `dcc-mcp-core`）负责控制面，C# 负责 Office COM 数据面；
2. 每个 Office 应用独立 Sidecar 进程，共享同一套 C# Runtime；
3. 批量静态处理优先 Open XML，最终高保真渲染优先 Office COM；
4. Agent 只见任务级 Capability，不见原始 COM 成员；
5. 所有复杂写入具备检查点、验证、预览和可追踪结果。

## 快速开始

C# 开发使用 [vx](https://github.com/loonghao/vx)（通用版本执行器）：`vx.toml`
锁定 .NET 8 LTS SDK，`vx.lock` 保证可复现安装。CI 通过官方
[`loonghao/vx`](https://github.com/loonghao/vx) GitHub Action 使用同一套工具链。

```bash
cargo test
vx setup            # 安装 vx.toml 中锁定的 .NET 8 LTS SDK
vx run build        # 使用 vx 管理的 SDK 构建 office-host
vx run self-test    # 宿主自检（编译 + 检查往返，无需安装 Office）
```

## CI / CD

- `ci.yml` — Rust 质量门禁 + 技能 lint + C# 宿主构建自检（Windows，走 `loonghao/vx`）。
- `publish-host.yml` — 每次 dotnet 变更发布单文件 `dcc-office-host.exe` 制品。
- `release.yml` — release-please 自动发版：合并发版 PR 后自动打 tag，用 vx
  工具链构建宿主并挂载 `dcc-office-host.exe` 到 GitHub Release。

详见 [README.md](./README.md) 与 [AGENTS.md](./AGENTS.md)。

## 许可

MIT
