# ADR 007 — Reference stdio MCP server

- **Status**: Accepted
- **Date**: 2026-08-24
- **Related**: Issue #26, ADR 001, ADR 004

## Context

The repository shipped the typed Rust pipe client and the C# Office Host, but
clients had no executable MCP endpoint until the gateway lifecycle adapter in
`dcc-mcp-core` could consume a published `dcc-mcp-host-rpc` crate. Shipping
only the sidecar made the implemented task-level catalog inaccessible to a
normal MCP client.

## Decision

Ship `dcc-mcp-office-mcp-server.exe` as a thin, interim stdio adapter:

- it owns one `office-host.exe` child for one explicitly selected Office app;
- a live `office.host.handshake` filters the catalog-owned MCP tool set;
- the canonical input/output schemas and security policy are enforced before
  `office-client` sends `office-rpc/1` commands;
- asynchronous batch capabilities add `office.job.get` and
  `office.job.cancel`;
- sidecar `OFFICE_*` failures remain structured MCP tool errors; and
- the executable ships beside the Host in the checksummed, attested release
  bundle.

The server contains no Office behavior, no duplicate capability registry, and
no COM dependency. Its bridge is injected for Office-free protocol tests.

## Consequences

The Office platform has a directly consumable MCP surface while core gateway
wiring remains release-train blocked. Once the gateway owns the same lifecycle
contract, this executable can remain a reference/debug adapter or be retired
without changing the catalog, schemas, sidecar, or application adapters.
