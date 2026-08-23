# Brand template packages

`deck.compile` resolves `brand://` URIs to validated, materialized packages.
The Host advertises exactly those packages in
`capability_manifest.template_packages`; an advertised URI is therefore
executable, not just a catalog entry.

## Install and select

Folder packages are discovered recursively from:

1. each repeatable `--template-dir=<path>` argument;
2. the release bundle's `templates/` directory;
3. `%LOCALAPPDATA%\dcc-mcp\office-templates`.

The process working directory is never searched. To use the repository's
example during development:

```powershell
dcc-office-host.exe --app=powerpoint --openxml-only --template-dir=templates
```

Then pass `"template":"brand://dcc-mcp/studio-light"` to `deck.compile`, or
set `template.uri` in the Deck IR. An unknown or invalid package is refused.

## Package contract

Each package directory contains `package.json` conforming to
[`package.schema.json`](./package.schema.json). It pins a semantic version,
maps public semantic layout names to built-in renderers, and may override the
inherited master, layout, theme, slide, notes, style, and logo files. Paths
must remain inside the package; absolute paths, traversal, and escaping
symlinks are refused.

The complete renderer set is:

```text
title_cover section_cover two_columns comparison timeline kpi_dashboard
technical_architecture image_left_text_right image_grid closing bullets
```

`presentations/studio-light` is a working external package. It inherits the
default skeleton and media while replacing the palette, fonts, semantic
layout map, and brand name. Copy it outside the release bundle to create a
package that can be updated without rebuilding the Host.

## Catalog

`registry.json` is the release catalog used by documentation and packaging.
Runtime discovery is package-driven and the handshake is the authoritative
materialized capability surface, preventing a stale catalog from promising an
unavailable template.

The `documents/`, `workbooks/`, and `diagrams/` directories are reserved for
future package kinds; they are not current runtime capabilities.
