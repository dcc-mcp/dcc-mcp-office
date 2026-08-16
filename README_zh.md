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
| **dcc-mcp-office**（本仓库） | 协议 / IR / C# 运行时 / Open XML / Graph / 通用技能 | M0 骨架 |
| dcc-mcp-PowerPoint | Deck 生成、Slide 编排、评审 Deck | M0 骨架 |
| dcc-mcp-word | 重排、域、目录 | Phase 2 占位 |
| dcc-mcp-excel | 计算、表格、图表、Graph Workbook | Phase 2 占位 |
| dcc-mcp-outlook | 草稿、日历 | Phase 3 占位 |

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
