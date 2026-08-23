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
  office-protocol   office-rpc/1 catalog + typed wire DTOs for handshake,
                    commands, jobs, progress, events, and error codes
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
                              + office-rpc/1 named-pipe JSON-RPC server,
                              bounded in-memory job tracker and event producer
skills/            office-wide skill packs (SKILL.md)
manifests/         catalog-referenced input JSON Schemas
templates/         external brand template packages and release catalog
tests/             golden-files / visual-snapshots / compatibility / security / stress
tests/fixtures/    small docx/xlsx fixtures for the Rust ↔ C# contract tests
docs/adr/          architecture decision records
docs/proposals/    the platform proposal this repo implements
```

The Host applies the catalog-owned default-deny policy again at the desktop
boundary. File access is confined to the process-bound `--workspace-root`
(the Host working directory when omitted); a request may echo that root but
cannot replace it. In-place replace commits and existing-output overwrites
require a byte-exact checkpoint plus structured `confirmation` proof. Desktop
COM is refused in Session 0, and audit records contain values read back from
the live Office application rather than unverified intent. Capabilities with
no explicit overwrite mode refuse existing destination artifacts.

`batch.convert` and `batch.replace_text` are asynchronous at the sidecar
boundary: `office.command.execute` returns a `job_id` immediately, callers poll
`office.job.get`, and `office.job.cancel` is observed between files. The pipe
delivers `office.job.progress` notifications plus job/application/document,
security, and modal events. `office.host.ping` reports Host state without
attaching to or starting an Office application. Job state is intentionally
process-local for M1; durable recovery remains the later `dcc-mcp-job` adapter.

Presentation templates are external, versioned folder packages. The Host
discovers packages from the release bundle, `%LOCALAPPDATA%\dcc-mcp\office-templates`,
or repeatable `--template-dir=<path>` arguments and advertises only validated,
materialized packages in the handshake. See
[`templates/README.md`](./templates/README.md) for the package contract and the
resolvable `brand://dcc-mcp/studio-light` example.

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
- `.github/workflows/publish-host.yml` — builds the complete Host bundle as a
  short-lived verification artifact on every distribution-related change.
- `.github/workflows/release.yml` — release-please automated versioning and
  releases: merging the release PR creates the immutable Rust source tag and
  attaches the Host binaries, install bundle, manifest, checksums, SPDX SBOM,
  and GitHub attestations. See
  [`docs/distribution.md`](./docs/distribution.md).

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
