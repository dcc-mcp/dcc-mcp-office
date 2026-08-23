# CONTRIBUTING — dcc-mcp-office

Engineering agreement for this repository. Applies to all code and review.

## 1. 第一性原理（First principles）

Start from domain invariants, not from tools:

- The **wire contract** (`office-rpc/1` in `crates/office-protocol`) and the
  **document IRs** (`crates/office-ir`) are the spine of this project.
- Everything else — COM, Open XML, Graph, Office.js, UIA — is an
  *implementation behind a contract*. An implementation may be replaced
  without changing the contract.
- A capability is a **contract**, not a call: input schema + output schema +
  error set + revision semantics + validation rules (proposal §3.2).

## 2. 契约精神（Contract-first）

- Every cross-boundary message is versioned and round-trip tested; a schema
  change is a contract change and gets a contract test first.
- The error-code set is **closed** (`OfficeErrorCode`): new failures extend
  the enum deliberately, never ad-hoc strings.
- Never degrade silently: unsupported input returns
  `OFFICE_CAPABILITY_UNSUPPORTED` with a reason; an uncertain write result
  returns `indeterminate: true` — never a guessed success/failure.
- Artifact/Job/results follow the proposal shapes (§16/§17) so agents always
  read back *what changed*, not just "succeeded".
- `crates/office-protocol/office-rpc.catalog.json` is the only capability,
  Office error-code, and default security-policy registry. Add or version a
  capability or policy there, reference an embedded schema, and let the Host
  manifest/dispatch/policy gate plus Rust mappings derive from it; never add a
  second handwritten registry.
- In-place writes require a byte-exact checkpoint and structured confirmation
  proof. `expected_revision` requests are refused until revision tracking is
  real; no adapter may accept and ignore an optimistic-concurrency guard.
- File paths are confined to the process-bound workspace root at the Host
  boundary; request policy may never widen that root.
  Session 0 cannot start desktop COM automation. Audit security fields report
  read-back state, not desired settings.

### Versioning policy

- The workspace release version is the provider and host build version.
  release-please stamps both `Cargo.toml` and
  `dotnet/Directory.Build.props`; `Cargo.lock`, assembly metadata,
  `--version`, handshake `provider_version`, and audit `host_version` must
  remain equal to it.
- Capability versions are independent semantic versions. Bump **major** for
  incompatible input/output schema or behavior, **minor** for backward-
  compatible optional fields or execution modes, and **patch** for fixes that
  preserve the capability contract. A capability version changes only with
  its contract, not with every provider release.

## 3. SOLID

- **S** — one crate, one responsibility: protocol ≠ IR ≠ security ≠ tools
  ≠ client. In C#: `StaDispatcher` schedules, `ComObjectLifecycle` owns
  RCW rules, the host owns process/pipe — never one god class.
- **O** — open for extension via the capability registry
  (`crates/office-tools`): adding a capability never edits dispatch logic.
- **L** — sidecar clients implement `dcc_mcp_host_rpc::HostRpcClient`
  (M1); any backend must be substitutable behind the same trait.
- **I** — thin interfaces: clients see only handshake/command/progress, not
  COM plumbing.
- **D** — depend on abstractions: Rust crates depend on the protocol schema,
  never on COM types; the C# runtime implements the protocol as its
  outermost adapter.

## 4. Clean Architecture

Dependency rule: source code dependencies point **inward** only.

```text
  dcc-mcp-core gateway (use cases)
        |
        v
  office-mcp-server / office-client / office-tools / office-jobs
                                                (application + outer adapter)
        |
        v
  office-protocol + office-ir                  (domain core: pure schemas)
        ^
        |
  C# Runtime / OpenXml / Graph / Office.js     (outer adapters, implement the
                                                protocol — never imported by
                                                the domain)
```

- Domain crates have zero I/O dependencies (M0 rule: only serde).
- The reference MCP server is a replaceable outer adapter: it derives tool
  contracts from `office-protocol`, applies `office-security`, and reaches the
  Host only through `office-client`.
- COM interop, pipe I/O and HTTP live in adapters only.
- Tests assert the dependency direction stays clean.

## 5. No Code Smells

Hard gates (run before every commit):

```bash
cargo fmt --check
cargo clippy --all-targets -- -D warnings
cargo test
vx run build
vx run format-check
vx run test-dotnet
```

House rules:

- `#![forbid(unsafe_code)]` in every crate.
- No empty catch blocks, no `unwrap()` on untrusted data paths, no magic
  strings for error/state values — use the closed enums.
- No COM references in async queues; copy to DTOs at event boundaries
  (proposal §9.3).
- Doc comments on public API; decisions go to `docs/adr/`, not comments.
