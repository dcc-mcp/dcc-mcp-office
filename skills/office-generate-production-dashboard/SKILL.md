---
name: office-generate-production-dashboard
description: >-
  Generate a styled production dashboard workbook (XLSX) from a Workbook IR
  JSON: worksheets with tables, header styling, freeze panes, summary
  formulas and a bar chart. Use whenever the agent must produce a
  data-facing Excel deliverable from structured rows.
license: MIT
allowed-tools: Bash Read
metadata:
  dcc-mcp:
    dcc: office
    layer: domain
    stage: authoring
    version: 0.1.0
    tags:
      - office
      - excel
      - dashboard
      - xlsx
    search-hint: >-
      generate dashboard, make excel, xlsx report, capability ledger,
      production dashboard
    tools: tools.yaml
---

# office-generate-production-dashboard (Authoring stage)

Excel generation through the designed pipeline (proposal §15.6): Data
Contract → Workbook IR → base workbook (openpyxl as the Open XML builder) →
formulas → XLSX artifact. Charts and pivots stay declarative in the IR;
native recalc/refresh belongs to the Excel COM sidecar (Phase 2).

## Contract boundary

This script is an offline authoring adapter for a **new draft artifact**, not
an `office-rpc/1` provider and not an alternative capability namespace. It
consumes the shared Workbook IR envelope but does not report a Host audit,
mutate an open workbook, run native Excel validation, or claim that an
`office.command.execute` call succeeded. Those operations must go through the
catalog-mapped MCP/sidecar path when the Excel generation capability lands.

## Input contract

- `input` — path to a Workbook IR JSON
  (`schema_version: office-ir/1.0`, `kind: workbook`) with
  `worksheets[{name, tables[{name, headers, rows}]}]` and optional
  `charts[{sheet, type, title, categories, values}]`

## Scripts

- `generate_dashboard` — Workbook IR → styled XLSX with summary formulas

## Validation rules

- contract enforced at load: headers/rows shape, unknown chart types
- artifact must exist and be non-empty

## Agent-visible summary

Artifact path, sheets written, row counts, and whether human review is
recommended.
