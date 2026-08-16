# ADR 001 — Rust Gateway + C# COM Sidecar

- **Status**: Accepted
- **Date**: 2026-08-16
- **Related**: proposal §1, §4, §9, §28 (ADR-001 in proposal)

## Context

The platform proposal requires deep desktop Office automation (COM object
models) behind the DCC-MCP MCP surface. The Rust gateway already exists in
`dcc-mcp-core` and owns routing, capability discovery, jobs, artifacts,
policy and lifecycle.

## Decision

- Gateway stays Rust (existing `dcc-mcp-core`, no fork).
- Office desktop providers run as C# sidecar processes (`office-host.exe`),
  one per application, speaking `office-rpc/1` over named pipes.
- C# owns STA dispatch, COM interop and Open XML batching; Rust owns the
  control plane.
- No self-built Rust Office automation runtime.

## Consequences

- Office crashes/busy states cannot block the gateway (process isolation).
- C# interop maturity (RCW handling, events, STA) is leveraged instead of
  re-implemented in Rust.
- Distribution must ship a self-contained .NET runtime artifact via
  `dcc-mcp-release-artifacts`.
