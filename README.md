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
| **dcc-mcp-office** (this repo) | shared protocol / IR / C# runtime / Open XML / Graph / generic skills | M0 scaffold |
| dcc-mcp-PowerPoint | deck generate, slide compose, review decks | M0 scaffold |
| dcc-mcp-word | reflow, fields, TOC | Phase 2 (placeholder) |
| dcc-mcp-excel | calculate, tables, charts, Graph workbook | Phase 2 (placeholder) |
| dcc-mcp-outlook | drafts, calendar | Phase 3 (placeholder) |

## Layout

```text
crates/
  office-protocol   wire schema: office-rpc/1, handshake, capability manifest,
                    error codes, job progress, events (schema-only, no I/O)
  office-ir         document IRs: common envelope + presentation/word/workbook
  office-client     Rust-side client for office-host.exe (namedpipe://)
  office-tools      task-level MCP tool registry (office.batch.convert, ...)
  office-jobs       approval/checkpoint job phases layered on dcc-mcp-job
  office-graph      Microsoft Graph connector (OneDrive/SharePoint/Workbook)
  office-security   default-deny policy: ExecuteMso whitelist, AutomationSecurity
  office-testkit    contract-test helpers for the office-rpc surface
dotnet/
  Office.Automation.Runtime   STA dispatcher, COM lifecycle, named-pipe server
  Office.Automation.OpenXml   batch structural worker (compiler, not renderer)
  Office.Automation.Host      office-host.exe entry point (--app=powerpoint|...)
skills/            office-wide skill packs (SKILL.md)
manifests/         capability manifest examples (powerpoint-desktop, ...)
templates/         brand template registry (brand:// URIs)
tests/             golden-files / visual-snapshots / compatibility / security / stress
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
| M1 | Open XML worker + COM sidecar MVP: batch convert / replace_text / inspect | the 12 MVP criteria in §27 |
| M2 | deck generate pipeline (template registry → IR → OpenXML → COM finalize → previews → validation loop) | shipped in dcc-mcp-PowerPoint |
| M3 | Graph connector + Office.js add-in | cloud file scenarios |
| M4 | Visio/Project/Access + ecosystem (marketplace, SUA pack) | §26 Phase 4/5 |

## License

MIT
