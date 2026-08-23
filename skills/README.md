# skills — Office-wide Skill Packs

Stable workflows are fixed as skills (proposal §11.3): agents invoke a skill
instead of re-assembling tool sequences every time. Format follows
`dcc-mcp-skills` (SKILL.md, frontmatter: name/description/dcc/version/...).

Each skill defines all eight items required by proposal §11.3:

1. input contract
2. planning steps
3. provider choice
4. safety confirmation
5. validation rules
6. failure compensation
7. artifact naming
8. agent-visible summary

## Packs

| Skill | Status |
|---|---|
| `office-batch-to-pdf` | **implemented** — COM sidecar batch PDF export with validation (repository M1) |
| `office-global-text-replace` | **implemented** — dry-run/commit replace across pptx/docx/xlsx (body/headers/footers/notes) |
| `office-generate-production-dashboard` | **implemented** — Workbook IR → styled XLSX with summary formulas + bar chart (see `examples/capability-dashboard.json`) |

PowerPoint-specific packs (`powerpoint-deck`, `powerpoint-review`,
`dcc-review-deck-from-renders`) live in `dcc-mcp-PowerPoint`.
