# ADR 003 — Open XML + COM 混合后端

- **Status**: Accepted
- **Date**: 2026-08-16
- **Related**: proposal §6.1/§6.2, ADR-003 in proposal

## Context

Bulk structural edits want throughput and Office-free machines; final layout,
rendering and PDF export want native Office fidelity. Neither backend alone
satisfies both.

## Decision

- Open XML (C#) = structural compiler and batch engine for closed files.
- Desktop COM = native completion, rendering, export and live-document ops.
- Microsoft Graph = cloud files and conversion (Phase 3).
- Routing is explicit and reported in every result (`backend` field).

## Consequences

- Batch throughput without launching Office; high-fidelity output where it
  matters.
- Two implementations to maintain for overlapping operations — mitigated by
  task-level capabilities hiding the split from agents.
