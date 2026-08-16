# AGENTS.md — dcc-mcp-office Agent Navigation Map

> Progressive disclosure: this file is a **map**, not an encyclopedia.
> Follow the links for depth. Stay here for breadth.

## 30-Second Summary

`dcc-mcp-office` is the **shared Office core** for the DCC-MCP ecosystem. It
implements the Office Automation Platform proposal: a versioned JSON-RPC wire
protocol (`office-rpc/1`) spoken between the existing Rust gateway
(`dcc-mcp-core`) and a new per-application C# sidecar (`office-host.exe`)
that drives desktop Office over COM; Open XML handles bulk structural work;
Microsoft Graph handles cloud files (Phase 3); UIA/Computer Use fallback is
reused from `dcc-mcp-computer-use`, not rebuilt here.

**Current status:** M0 scaffold. Protocol/IR schemas and the C# host skeleton
compile and are contract-tested; no COM attachment yet (M1).

## Repo Map

| Path | What it is | Read when |
|---|---|---|
| `crates/office-protocol` | Wire schema: handshake, capability manifest, error codes, progress/events, pipe naming | any RPC work |
| `crates/office-ir` | Common envelope + application IRs (presentation now, word/workbook next) | generation pipelines |
| `crates/office-client` | Rust client for `office-host.exe`; will implement `HostRpcClient` from `dcc-mcp-host-rpc` | sidecar wiring (M1) |
| `crates/office-tools` | Task-level MCP tool registry (names, app, phase) | adding a capability |
| `crates/office-jobs` | Job phases layered on `dcc-mcp-job` (approval/checkpoint) | batch operations |
| `crates/office-graph` | Graph connector stub (Phase 3) | cloud scenarios |
| `crates/office-security` | Default-deny policy: ExecuteMso whitelist, AutomationSecurity, risk levels | any write path |
| `crates/office-testkit` | Contract-test helpers | test work |
| `dotnet/Office.Automation.Runtime` | STA dispatcher, COM lifecycle, named-pipe server | sidecar internals |
| `dotnet/Office.Automation.OpenXml` | Batch structural worker — compiler, never a renderer | bulk edits |
| `dotnet/Office.Automation.Host` | `office-host.exe` entry (`--app=powerpoint|word|...`) | host startup |
| `skills/` | Office-wide skill packs (SKILL.md, dcc-mcp-skills format) | workflows |
| `manifests/` | Capability manifest examples | provider registration |
| `templates/` | Brand template registry (`brand://` URIs) | deck/report generation |
| `docs/adr/` | Decision records (001-006) | "why is it built this way" |
| `docs/proposals/office-automation-platform-v1.0.md` | Full platform proposal | architecture questions |

## Dependency Map

- `dcc-mcp-core`: gateway, job engine, artefact store, skills runtime,
  process lifecycle, telemetry — all **reused**, not re-implemented.
- `dcc-mcp-host-rpc`: the `HostRpcClient` trait this repo's
  `office-client` will implement (wiring lands in M1; M0 keeps crates
  dependency-free except serde).
- `dcc-mcp-computer-use`: UIA/visual fallback — Office contributes only
  semantic profiles + fallback policy (Phase 3+).
- Application adapters (`dcc-mcp-PowerPoint`, `dcc-mcp-word`, ...) consume
  this repo's published crates + the `office-host` binaries.

## Build / Test

```bash
cargo test                          # protocol/IR/security unit + round-trip tests
vx setup                            # install the pinned .NET 8 LTS SDK (vx.toml/vx.lock)
vx run build                        # C# host build via the vx-managed SDK
vx run self-test                    # office-host self-test (no Office required)
```

The C# toolchain is managed by [vx](https://github.com/loonghao/vx): `vx.toml`
pins the SDK, `vx.lock` freezes it. CI (`ci.yml`, `publish-host.yml`,
`release.yml`) installs vx through the `loonghao/vx` action so local and CI
SDKs stay identical; releases are automated by release-please.

CI matrix (proposal §22.3): M365 current channel, Office LTSC, 32/64-bit,
Windows 11, zh-CN + en-US — golden-file and visual-snapshot tests live under
`tests/` and run against the published `office-host` binaries.

## Conventions

- Engineering agreement (first principles / contract-first / SOLID / Clean
  Architecture / no code smells): [CONTRIBUTING.md](./CONTRIBUTING.md).
- Crate names are `dcc-mcp-office-*` (dirs are `office-*`).
- Protocol changes bump the schema version and get a contract test first.
- Everything touching COM enforces the security policy from
  `crates/office-security` (deny by default: VBA, macros, external links,
  arbitrary ExecuteMso).
- Skills define all eight items required by the proposal §11.3 (input
  contract, plan, provider choice, confirmation, validation, fallback,
  naming, agent-visible summary).
