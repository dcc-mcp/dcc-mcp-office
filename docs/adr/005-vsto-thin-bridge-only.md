# ADR 005 — VSTO 仅作可选薄桥

- **Status**: Accepted
- **Date**: 2026-08-16
- **Related**: proposal §9.1, ADR-005 in proposal

## Context

VSTO depends on .NET Framework and Microsoft is not migrating the classic
COM add-in platform to .NET 5+; it is also Windows-only.

## Decision

- Core automation logic never lives in VSTO.
- In-Office UI prefers Office.js (cross-platform); a thin VSTO bridge is
  allowed only for Windows-specific gaps, never as the runtime host.

## Consequences

- The runtime is self-contained modern .NET, publishable and testable
  without Office add-in installation.
- Office.js remains the cross-platform entry (Phase 3).
