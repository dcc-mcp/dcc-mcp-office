# ADR 004 — 任务级 MCP Capability

- **Status**: Accepted
- **Date**: 2026-08-16
- **Related**: proposal §3.2, §11, ADR-004 in proposal

## Context

Exposing raw COM members (PowerPoint.Shape.TextFrame.TextRange.Text, ...) to
agents would leak Office version differences, explode the MCP schema and make
permissioning, validation and retry policy intractable.

## Decision

- Agents see task-level capabilities only: `office.batch.convert`,
  `office.document.patch`, `powerpoint.deck.generate`, ... (registry in
  `crates/office-tools`).
- Backends (Open XML / COM / Graph / Office.js / UIA) stay swappable behind
  one capability.

## Consequences

- Smaller, stable MCP schema; policy/audit unified at the gateway.
- Fewer low-level COM round-trips per agent action.
- Capability coverage is deliberate work per app — fine-grained COM power is
  intentionally not exposed.
