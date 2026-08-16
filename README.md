# dcc-mcp-office

Office 通用底座 for the DCC-MCP ecosystem — task-level MCP capabilities over
PowerPoint / Word / Excel / Outlook / Visio / Project / Access.

This repository implements the **shared core** of the
[DCC-MCP Office Automation Platform proposal](./docs/proposals/office-automation-platform-v1.0.md):
the `office-rpc/1` wire protocol, application document IRs, the C# STA COM
sidecar runtime (`office-host`), the Open XML batch worker, the Microsoft
Graph connector and the Office-wide skill packs. Application-specific
adapters live in thin sibling repos and depend on this one the way
`dcc-mcp-maya` depends on `dcc-mcp-core`.

| Repo | Scope | Status |
|---|---|---|
| **dcc-mcp-office** (this repo) | shared protocol / IR / C# runtime / Open XML / Graph / generic skills | M1 COM sidecar MVP |
| dcc-mcp-PowerPoint | deck generate, slide compose, review decks | M1 consumes |
| dcc-mcp-word | reflow, fields, TOC | M1 consumes (COM legs) |
| dcc-mcp-excel | calculate, tables, charts, Graph workbook | M1 consumes (COM legs) |
| dcc-mcp-outlook | drafts, calendar | Phase 3 (placeholder) |

## Layout

```text
crates/
  office-protocol   wire schema: office-rpc/1, handshake, capability manifest,
                    error codes, job progress, events (schema-only, no I/O)
  office-ir         document IRs: common envelope + presentation / word / workbook
  office-client     Rust-side client for office-host.exe (namedpipe://)
  office-tools      task-level MCP tool registry (office.batch.convert, ...)
  office-jobs       approval/checkpoint job phases + per-item aggregation
  office-graph      Microsoft Graph connector (OneDrive/SharePoint/Workbook)
  office-security   default-deny policy: ExecuteMso whitelist, AutomationSecurity
  office-testkit    JSON-RPC fixtures + FakeSidecar with fault injection
dotnet/
  Office.Automation.Runtime   STA dispatcher (message pump, IOleMessageFilter
                              busy retry, modal-dialog detection, soft timeouts),
                              COM object lifecycle rules
  Office.Automation.Com       per-app COM backends: PowerPoint/Word/Excel
                              (attach + security defaults, batch PDF export,
                              replace-text dry-run/commit, inspect, slide
                              previews with overflow detection)
  Office.Automation.OpenXml   batch structural worker (compiler, not renderer)
  Office.Automation.Host      office-host.exe entry point (--app=powerpoint|...)
                              + office-rpc/1 named-pipe JSON-RPC server
skills/            office-wide skill packs (SKILL.md)
manifests/         capability manifest examples + input JSON Schemas
templates/         brand template registry (brand:// URIs, registry.json)
tests/             golden-files / visual-snapshots / compatibility / security / stress
tests/fixtures/    small docx/xlsx fixtures for the Rust ↔ C# contract tests
docs/adr/          architecture decision records
docs/proposals/    the platform proposal this repo implements
```

## Quickstart

C# development uses [vx](https://github.com/loonghao/vx) (universal version
executor): `vx.toml` pins the .NET 8 LTS SDK and `vx.lock` makes installs
reproducible. CI installs the same toolchain through the official
[`loonghao/vx`](https://github.com/loonghao/vx) GitHub Action.

```bash
cargo test            # protocol/IR round-trips, policy defaults, no Office required
vx setup              # install the pinned .NET 8 LTS SDK into the vx store
vx run build          # build office-host with the vx-managed SDK
vx run self-test      # host self-test (compile + inspect round-trip, no Office needed)
```

With Microsoft Office installed, the real COM legs run too — the Rust client
drives a live host over the named pipe and validates PDF export, replace-text
and slide previews end to end:

```powershell
dotnet build dotnet/Office.Automation.Host          # or: vx run build
dotnet/Office.Automation.Host/bin/Debug/net8.0-windows/dcc-office-host.exe --app=powerpoint --self-test-com
$env:DCC_OFFICE_HOST_EXE = "F:githubdcc-mcp-officedotnetOffice.Automation.HostinDebug
et8.0-windowsdcc-office-host.exe"
cargo test -p dcc-mcp-office-client --test pipe_contract   # skips gracefully without Office
```

Prefer a system SDK? The same commands work without the `vx` prefix:

```bash
dotnet build dotnet/Office.Automation.Host
```

## CI / CD

- `.github/workflows/ci.yml` — Rust fmt/clippy/test, skill lint + dashboard
  e2e, and the C# host build/self-test (Windows, via `loonghao/vx`).
- `.github/workflows/publish-host.yml` — publishes the single-file
  `dcc-office-host.exe` as a CI artifact on every dotnet change.
- `.github/workflows/release.yml` — release-please automated versioning and
  releases: merging the release PR cuts the tag, builds the host with the vx
  toolchain and attaches `dcc-office-host.exe` to the GitHub release.

## Roadmap

| Milestone | Content | Acceptance (proposal §26/§27) |
|---|---|---|
| M0 | office-rpc/1 schema + C# host skeleton (handshake/STA queue, no COM) | contract tests green without Office |
| M1 | Open XML worker + COM sidecar MVP: named-pipe server, STA busy/modal ladder, batch convert / replace_text dry-run+commit / inspect / slide previews+overflow (PowerPoint, Word, Excel) | the 12 MVP criteria in §27 |
| M2 | deck generate pipeline (template registry → IR → OpenXML → COM finalize → previews → validation loop) — core pieces live here (IRs, brand:// registry, compile + finalize + render); the pipeline ships in dcc-mcp-PowerPoint | shipped in dcc-mcp-PowerPoint |
| M3 | Graph connector + Office.js add-in | cloud file scenarios |
| M4 | Visio/Project/Access + ecosystem (marketplace, SUA pack) | §26 Phase 4/5 |

## License

MIT
