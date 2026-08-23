# manifests — Office RPC contract source

[`crates/office-protocol/office-rpc.catalog.json`](../crates/office-protocol/office-rpc.catalog.json)
is the machine-readable source of truth for the
sidecar contract: wire capability names and semvers, task-level MCP mappings,
handler IDs, app/execution-mode availability, input schemas, and error subsets.
Rust embeds it through `dcc-mcp-office-protocol`; the C# Host embeds the same
file and derives both dispatch and handshake manifests from it.

The namespace boundary is intentional:

- Agents call task-level MCP tools such as `office.batch.convert` and
  `powerpoint.slide.render`.
- The gateway maps them through the catalog's `mcp_tool` field to compact
  sidecar wire names such as `batch.convert` and `slide.render`.
- `office-host` accepts only catalog wire names. There is no implicit prefix
  stripping or second handwritten mapping.

The exact Office-free handshake shape is replayed from
`tests/office-free-rpc/handshake.expected.json`; installed Office identity is
filled at runtime and is therefore not duplicated in a static example file.

## Input schemas

`schemas/` carries the draft-07 input contracts embedded and enforced by the
Host before a handler runs. The catalog is the only capability-to-schema map:

| Wire capability | MCP tool | Schema |
|---|---|---|
| `deck.compile` | `powerpoint.deck.generate` | `schemas/deck-compile.schema.json` |
| `document.inspect` | `office.document.inspect` | `schemas/document-inspect.schema.json` |
| `batch.convert` | `office.batch.convert` | `schemas/batch-convert.schema.json` |
| `batch.replace_text` | `office.batch.replace_text` | `schemas/batch-replace-text.schema.json` |
| `slide.render` | `powerpoint.slide.render` | `schemas/slide-render.schema.json` |

Every capability also references `schemas/command-result.schema.json`; the
Host validates the handler result against it before audit enrichment.
