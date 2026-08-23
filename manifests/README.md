# manifests — Office RPC contract source

[`crates/office-protocol/office-rpc.catalog.json`](../crates/office-protocol/office-rpc.catalog.json)
is the machine-readable source of truth for the
sidecar contract: wire capability names and semvers, task-level MCP mappings,
handler IDs, app/execution-mode availability, input schemas, error subsets, and
the default-deny security policy.
Rust embeds it through `dcc-mcp-office-protocol`; the C# Host embeds the same
file and derives both dispatch and handshake manifests from it.

The namespace boundary is intentional:

- Agents call task-level MCP tools such as `office.batch.convert` and
  `powerpoint.slide.render`.
- The reference stdio MCP server (and, later, the gateway adapter) maps them
  through the catalog's `mcp_tool` field to compact sidecar wire names such as
  `batch.convert` and `slide.render`.
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

The catalog also owns every request/response RPC schema:

| RPC method | Params | Result |
|---|---|---|
| `office.host.handshake` | `schemas/handshake-params.schema.json` | `schemas/handshake-result.schema.json` |
| `office.host.ping` | `schemas/empty-object.schema.json` | `schemas/sidecar-status.schema.json` |
| `office.host.shutdown` | `schemas/empty-object.schema.json` | `schemas/shutdown-result.schema.json` |
| `office.job.get` | `schemas/job-id.schema.json` | `schemas/job-status.schema.json` |
| `office.job.cancel` | `schemas/job-id.schema.json` | `schemas/job-cancel-result.schema.json` |
| `office.command.execute` | `schemas/command-params.schema.json` | `schemas/command-result.schema.json` |

Batch command submissions validate against the normal command-result schema
with `backend: "job"`, `job_id`, and `phase`. Their terminal command result is
returned from `office.job.get.result`; progress and completion are also sent as
JSON-RPC notifications on the same pipe.

`schemas/command-params.schema.json` owns the outer
`office.command.execute` envelope, including `document`, `policy`, and the
structured confirmation proof. The Host validates it before applying the
catalog policy or dispatching a capability.

## Write-safety envelope

- `--workspace-root` binds every input/output path when the Host starts;
  omission uses the Host working directory, never an unrestricted filesystem
  root. A request's `policy.workspace_root` may only echo that bound value.
- `document.expected_revision` is currently rejected with
  `OFFICE_CAPABILITY_UNSUPPORTED` because revision tracking is not yet real.
- `batch.replace_text` with `dry_run: false`, and `batch.convert` with
  `overwrite: "overwrite"`, require
  `confirmation: {action: "overwrite_original", confirmed: true,
  confirmed_by, confirmed_at}` and create byte-exact checkpoint artifacts for
  every pre-existing file before the first destructive write.
- `deck.compile` and `slide.render` have no overwrite mode, so the Host
  refuses existing destination artifacts instead of silently replacing them.
- A soft timeout on work that may have written returns or aggregates
  `indeterminate: true`; callers must re-inspect instead of retrying blindly.
