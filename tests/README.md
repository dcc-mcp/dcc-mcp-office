# tests — Test Matrix (proposal §24)

## Current automated coverage

Hosted CI runs the C# runtime, COM-classification, Open XML, Host stdio, and
Rust ↔ C# named-pipe contracts without requiring Office. The Host contracts
use `--openxml-only`, so they remain deterministic even on a developer machine
with desktop Office installed. The real PowerPoint/Word/Excel COM contracts
are marked `#[ignore]` instead of silently returning success; run them on a
provisioned Windows desktop with:

```powershell
$env:DCC_OFFICE_HOST_EXE = "<path-to-dcc-office-host.exe>"
vx run test-office-com
```

The hosted runner emits a warning and job summary for this coverage boundary.
A scheduled M365/LTSC self-hosted lane is still to be provisioned. The matrix
directories below describe that planned lane; empty `.gitkeep` directories are
not evidence that a matrix cell ran.

| Dir | Content |
|---|---|
| `golden-files/` | minimal / large / image-chart-table / macro / corrupt / protected / external-link / multilingual / legacy-format / special-font files per app; asserted: structure before/after, openability, page/slide counts, text deltas, formula results, export hashes |
| `visual-snapshots/` | per-slide PNG, per-page PDF/PNG, sheet/range previews; baseline diffs, font substitution, overflow/out-of-bounds detection |
| `compatibility/` | Office version matrix (proposal §22.3): M365 current, LTSC, 32/64-bit, Windows 11, zh-CN + en-US, classic vs new Outlook |
| `security/` | macro/XLM blocking, AutomationSecurity restore, ExecuteMso whitelist, path traversal, workspace confinement |
| `stress/` | fault injection: Busy, modal, file lock, Protected View, COM call rejected, sidecar crash, gateway restart, Graph 429, network drop, concurrent user edit, disk full |

Agent evaluation (§24.5) runs on the fixed task set with `dcc-mcp-tester`:
tool-selection accuracy, unnecessary visual-degradation rate, batch success
rate, human-acceptable deck rate, edit-scope accuracy, high-risk confirmation
rate, failure-explanation completeness.
