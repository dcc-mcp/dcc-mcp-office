# ADR 002 — 每个 Office 应用独立 Sidecar 进程

- **Status**: Accepted
- **Date**: 2026-08-16
- **Related**: proposal §8.2, ADR-002 in proposal

## Context

Word, Excel, PowerPoint, Outlook, Visio, Project and Access share one runtime
design but have independent UI threads, dialogs and workload profiles.

## Decision

- One `office-host` process per application (`--app=...`); same physical
  binary, process isolation enforced.
- Each process owns its STA queue, message pump and Office instance.

## Consequences

- A Word modal dialog cannot block Excel; a huge Excel recalculation cannot
  block PowerPoint.
- Per-app crash recovery, version probing, logs and health status.
- Slightly higher process footprint — acceptable for isolation guarantees.
