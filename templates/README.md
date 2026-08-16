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
| `presentations/` | PPTX brand templates (POTX/PPTX) |
| `documents/` | DOCX templates with Content Controls as anchors |
| `workbooks/` | XLSX templates with named ranges / tables |
| `diagrams/` | Visio stencils (Phase 4) |
