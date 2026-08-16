# manifests — Capability Manifests

Example manifests the sidecars report at handshake and re-report on reconnect
(proposal §10.3, §12.2). Shape matches
`dcc-mcp-office-protocol::CapabilityManifest` (Rust) 1:1.

Manifests are **reported by the runtime**, never hand-written in production:
the host fills `application` (version/bitness/language) from the installed
Office and `capabilities` from per-app probes. The example files document
the expected shape for tests and offline environments.

## Input schemas

`schemas/` carries JSON Schema (draft-07) input contracts the gateway
validates before dispatch — the same contracts the Rust protocol crate types
and the host's parsers implement:

| Schema | Capability |
|---|---|
| `schemas/batch-convert.schema.json` | `office.batch.convert` |
| `schemas/batch-replace-text.schema.json` | `office.batch.replace_text` (dry_run defaults to true) |
| `schemas/slide-render.schema.json` | `slide.render` (PowerPoint) |
