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

```bash
cargo test            # protocol/IR round-trips, policy defaults, no Office required
dotnet build dotnet/Office.Automation.Host
```

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
