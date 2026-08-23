# dotnet — C# Office Automation Runtime

C# side of the DCC-MCP Office stack (ADR-001): desktop COM data plane behind
the Rust gateway. Rust keeps the control plane; C# owns STA dispatch, COM
interop, Open XML batching and the named-pipe server.

| Project | Role | Repository milestone |
|---|---|---|
| `Office.Automation.Runtime` | STA dispatcher, message pump, busy retry, modal detection, and COM lifecycle rules | M1 complete |
| `Office.Automation.Com` | PowerPoint, Word, and Excel COM backends for inspect, convert, replace, render, and security defaults | M1 complete |
| `Office.Automation.OpenXml` | batch structural worker — compiler, not renderer | M1 complete |
| `Office.Automation.Host` | `dcc-office-host` executable, named-pipe server, command router, jobs, events, policy, and audit | M1 complete |

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
- Handshake probes the application's registered ProgID without launching it.
  The first real COM command lazily creates an isolated Automation instance,
  which the sidecar owns and closes during shutdown or recovery.
