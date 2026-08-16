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
  office-client / office-tools / office-jobs   (application layer)
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
- COM interop, pipe I/O and HTTP live in adapters only.
- Tests assert the dependency direction stays clean.

## 5. No Code Smells

Hard gates (run before every commit):

```bash
cargo fmt --check
cargo clippy --all-targets -- -D warnings
cargo test
dotnet format --verify-no-changes
dotnet build
```

House rules:

- `#![forbid(unsafe_code)]` in every crate.
- No empty catch blocks, no `unwrap()` on untrusted data paths, no magic
  strings for error/state values — use the closed enums.
- No COM references in async queues; copy to DTOs at event boundaries
  (proposal §9.3).
- Doc comments on public API; decisions go to `docs/adr/`, not comments.
