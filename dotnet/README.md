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

Publish aliases (proposal §22.1): copy `dcc-office-host.exe` to
`dcc-office-powerpoint-host.exe`, `dcc-office-word-host.exe`,
`dcc-office-excel-host.exe` — the binary switches behaviour on `--app`.

Distribution goes through `dcc-mcp-release-artifacts`; application adapters
(`dcc-mcp-PowerPoint`, ...) download/verify/launch the host, following the
Unity sidecar-launcher precedent.

## Process model (proposal §8)

- Never drive Office COM from Session 0; run in the interactive user session.
- One sidecar process per application (isolation: a Word modal dialog must
  not block Excel).
- OfficeInstanceResolver (M1): enumerate Running Object Table → map HWND to
  pid → enumerate open documents → attach or create per policy.
