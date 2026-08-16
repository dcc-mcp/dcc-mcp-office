# tests — Test Matrix (proposal §24)

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
