# templates — Brand Template Registry

Template-first generation, never free-coordinate drawing (proposal §15.4).
Templates are addressed by `brand://` URIs with pinned versions:

```text
brand://studio/review-v3
brand://dcc-mcp/product-launch-v2
brand://internal/technical-report-v1
```

Importing a template extracts (proposal §15.4): theme colors, theme fonts,
masters, layouts, placeholders, logo safe areas, title hierarchy, chart
styles, margins and aspect ratio.

Layouts gain semantic tags the agent selects instead of guessing coordinates:

```text
title_cover section_cover two_columns image_left_text_right comparison
timeline kpi_dashboard full_bleed_image technical_architecture
```

| Dir | Content |
|---|---|
| `registry.json` | machine-readable registry (schema `brand-registry/1.0`): URI → package source, version, kind, semantic layout list |
| `presentations/` | PPTX brand templates (POTX/PPTX) |
| `documents/` | DOCX templates with Content Controls as anchors |
| `workbooks/` | XLSX templates with named ranges / tables |
| `diagrams/` | Visio stencils (Phase 4) |

## Current packages

| URI | Kind | Source | Status |
|---|---|---|---|
| `brand://dcc-mcp/default` | presentation | embedded Open XML skeletons (master, 11 layouts: title_cover, section_cover, two_columns, comparison, timeline, kpi_dashboard, technical_architecture, image_left_text_right, image_grid, closing, bullets) + brand logo | shipped — `deck.compile` default |

`deck.compile` refuses URIs outside the registry (OFFICE_CAPABILITY_UNSUPPORTED)
and warns when a slide's `semantic_layout` is not part of the resolved package.
