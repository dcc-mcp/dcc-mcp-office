# Architecture health review v0.2.1 — closeout

- **Status:** closed
- **Tracking issue:** [#28](https://github.com/dcc-mcp/dcc-mcp-office/issues/28)
- **Review baseline:** `c5d6694` (`v0.2.1`)
- **Implementation evidence reviewed through:** `8600ff4`

The v0.2.1 architecture review identified thirteen concrete findings across
reliability, iteration sustainability, and deployment flexibility. Each
finding was handled by a dedicated pull request with its own validation. This
document is the durable closeout ledger; the linked issues and pull requests
retain the detailed acceptance evidence and discussion.

## Resolved findings

### A. Design flaws and reliability

| Finding | Merged change | Result |
|---|---|---|
| [#15](https://github.com/dcc-mcp/dcc-mcp-office/issues/15) STA soft-timeout race | [#29](https://github.com/dcc-mcp/dcc-mcp-office/pull/29) | Timeout completion and dispatcher recovery were hardened so late work cannot signal disposed state or poison subsequent requests. |
| [#16](https://github.com/dcc-mcp/dcc-mcp-office/issues/16) localized COM classification | [#31](https://github.com/dcc-mcp/dcc-mcp-office/pull/31) | COM failures are classified by HRESULT, and modal ownership detection is covered without English-message assumptions. |
| [#17](https://github.com/dcc-mcp/dcc-mcp-office/issues/17) unbounded client operations | [#32](https://github.com/dcc-mcp/dcc-mcp-office/pull/32) | Connect, handshake, read, and write operations are bounded; notifications are buffered and reconnect behavior is production code. |
| [#18](https://github.com/dcc-mcp/dcc-mcp-office/issues/18) unreachable write-safety guarantees | [#36](https://github.com/dcc-mcp/dcc-mcp-office/pull/36) | Revision guards, checkpoints, confirmations, indeterminate outcomes, readback-backed audit, and the second policy gate are enforced. |

### B. Iteration sustainability

| Finding | Merged change | Result |
|---|---|---|
| [#19](https://github.com/dcc-mcp/dcc-mcp-office/issues/19) drifted wire contracts | [#35](https://github.com/dcc-mcp/dcc-mcp-office/pull/35) | The Rust protocol catalog and schemas are canonical, with generated C# contracts and parity tests. |
| [#20](https://github.com/dcc-mcp/dcc-mcp-office/issues/20) schema-only jobs and events | [#38](https://github.com/dcc-mcp/dcc-mcp-office/pull/38) | Batch commands return jobs, support get/cancel, publish progress/events, and keep ping side-effect free. |
| [#21](https://github.com/dcc-mcp/dcc-mcp-office/issues/21) misleading test coverage | [#34](https://github.com/dcc-mcp/dcc-mcp-office/pull/34) | C# unit and Office-free integration layers run in CI; real-Office tests are explicit ignored contracts with a reported boundary. |
| [#22](https://github.com/dcc-mcp/dcc-mcp-office/issues/22) divergent release versions | [#33](https://github.com/dcc-mcp/dcc-mcp-office/pull/33) | Rust, wire, assembly, file, and CLI versions share one release value and are checked in CI. |
| [#26](https://github.com/dcc-mcp/dcc-mcp-office/issues/26) absent MCP surface | [#42](https://github.com/dcc-mcp/dcc-mcp-office/pull/42) | A reference stdio MCP server owns sidecar lifecycle, derives live tools from the catalog, and ships in the release bundle. |
| [#27](https://github.com/dcc-mcp/dcc-mcp-office/issues/27) documentation drift | [#43](https://github.com/dcc-mcp/dcc-mcp-office/pull/43) | Runtime, CI, roadmap, generated-output, and terminology claims are aligned and guarded by documentation contract tests. |

### C. Deployment flexibility

| Finding | Merged change | Result |
|---|---|---|
| [#23](https://github.com/dcc-mcp/dcc-mcp-office/issues/23) no consumable distribution | [#39](https://github.com/dcc-mcp/dcc-mcp-office/pull/39) | Releases contain the Host and MCP server, source contract, locator rules, manifests, checksums, SBOM, and attestations. |
| [#24](https://github.com/dcc-mcp/dcc-mcp-office/issues/24) decorative template registry | [#40](https://github.com/dcc-mcp/dcc-mcp-office/pull/40) | Versioned on-disk template packages support bundled, user, and repeatable sideload directories without recompiling the Host. |
| [#25](https://github.com/dcc-mcp/dcc-mcp-office/issues/25) fixed configuration and weak operability | [#41](https://github.com/dcc-mcp/dcc-mcp-office/pull/41) | Layered settings, structured correlated logs, bounded globbing, cancellation, parent watch, graceful shutdown, and launch-free handshake probing are implemented. |

## Remaining boundaries

These are explicit delivery boundaries rather than unresolved findings from
the v0.2.1 review:

- Hosted CI is Office-free. The real M365/LTSC, 32/64-bit, and locale matrix
  remains a separately provisioned desktop lane; ignored COM contracts and CI
  summaries make that boundary visible.
- The reference stdio MCP server is the current consumable surface. Direct
  `dcc-mcp-host-rpc` gateway lifecycle integration still follows the external
  `dcc-mcp-core` release train; no crates.io package was discoverable when this
  closeout was recorded on 2026-08-24.
- Graph, Office.js, Outlook/OneNote, Visio, Project, Access, and the broader
  ecosystem remain later proposal phases. Their planned status is not evidence
  of an implementation gap in the completed M1 review scope.

Any regression in a resolved contract or newly discovered gap should be filed
as a new focused issue rather than reopening this historical umbrella.
