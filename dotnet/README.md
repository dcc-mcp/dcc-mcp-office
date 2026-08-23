# dotnet — C# Office Automation Runtime

C# side of the DCC-MCP Office stack (ADR-001): desktop COM data plane behind
the Rust gateway. Rust keeps the control plane; C# owns STA dispatch, COM
interop, Open XML batching and the named-pipe server.

| Project | Role | Phase |
|---|---|---|
| `Office.Automation.Runtime` | STA dispatcher, COM lifecycle, pipe server (M1) | M0 skeleton |
| `Office.Automation.OpenXml` | batch structural worker — compiler, not renderer | M1 |
| `Office.Automation.Host` | `dcc-office-host` executable (`--app=...`) | M0 skeleton |

## Build

```powershell
dotnet build dotnet/Office.Automation.Host
```

## Publish (self-contained, proposal §9.1/§22.1)

```powershell
dotnet publish dotnet/Office.Automation.Host -c Release -r win-x64 --self-contained
```

The release packager creates the proposal §22.1 aliases and ships them with the
catalog, schemas, templates, checksums, SBOM, and provenance in the repository's
versioned GitHub Release. The canonical install and discovery contract is
documented in [`docs/distribution.md`](../docs/distribution.md).

## Process model (proposal §8)

- Never drive Office COM from Session 0; run in the interactive user session.
- One sidecar process per application (isolation: a Word modal dialog must
  not block Excel).
- OfficeInstanceResolver (M1): enumerate Running Object Table → map HWND to
  pid → enumerate open documents → attach or create per policy.
