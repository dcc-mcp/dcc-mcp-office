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

## Planned packs

| Skill | Status |
|---|---|
| `office-batch-to-pdf` | drafted (M1 tooling pending) |
| `office-global-text-replace` | planned (M1) |
| `office-brand-template-migration` | planned (M2) |
| `office-document-redaction` | planned (M2) |
| `office-generate-executive-deck` | planned (M2) |
| `office-generate-technical-report` | planned (Phase 2) |
| `office-generate-production-dashboard` | planned (Phase 2) |

PowerPoint-specific packs (`powerpoint-deck`, `powerpoint-review`,
`dcc-review-deck-from-renders`) live in `dcc-mcp-PowerPoint`.
