# docs

- `adr/` — architecture decision records (001-007)
- `proposals/` — the DCC-MCP Office Automation Platform proposal v1.0 that
  this repository implements

## Delivery vocabulary

The repository uses three independent planning dimensions. They are not
aliases: a tool can be high priority (`P0`) while its application workflow is
scheduled for a later delivery phase, and one repository milestone can support
more than one proposal phase.

| Vocabulary | Meaning | Values |
|---|---|---|
| Proposal delivery phase | Cross-repository delivery sequence defined in proposal §26 | Phase 0 through Phase 5 |
| Repository milestone | Implementation checkpoint for this shared-core repository | M0 through M4 |
| Tool priority | Support/investment priority from proposal §7; stored in the historical `RegistryEntry.phase` field | P0 through P2 |

| Proposal delivery phase | Shared-repository milestone contribution | Tool priority relationship |
|---|---|---|
| Phase 0 — unified foundation | M0 defined the wire/host skeleton; M1 completed the shared runtime | Common discovery, document, batch, and job tools are P0 |
| Phase 1 — PowerPoint | M1 supplies COM operations; M2 supplies the shared deck pipeline core | PowerPoint tools are P0 except slideshow control, which is P1 |
| Phase 2 — Word and Excel | M1 supplies COM operations; M2 supplies Word and Workbook IRs | Word and Excel tools are P0 |
| Phase 3 — cloud and Office in-app entry | M3 is the matching repository checkpoint | Priority depends on the application: P0 for existing core apps, P1 for Outlook/OneNote |
| Phase 4 — Visio, Project, and Access | M4 is the matching repository checkpoint | Visio is P1; Project and Access are P2 |
| Phase 5 — ecosystem | M4 includes the initial ecosystem tail; later work is not separately numbered here | No one-to-one tool priority |

Use `Phase` for the proposal sequence, `M` for this repository's checkpoints,
and `P` only for tool support priority.
