# ADR 006 — 底座 / 应用适配 / 技能包的仓库拆分

- **Status**: Accepted
- **Date**: 2026-08-16
- **Related**: proposal §5, §23, ADR-006 in proposal

## Context

The proposal sketches one monolithic `dcc-mcp-office` repository. The
DCC-MCP ecosystem convention is: one repo per DCC target adapter
(`dcc-mcp-maya`), optional skill-pack repos, `dcc-mcp-core` as shared
base, and `dcc-mcp-catalog.yml` as the index. One monolith would break
release cadence and the "thin adapter" pattern.

## Decision

- `dcc-mcp-office` (this repo): shared Office core — protocol, IR, C#
  runtime, Open XML, Graph, generic skills, manifests, templates, ADRs.
- Per-application thin adapters: `dcc-mcp-PowerPoint` (now),
  `dcc-mcp-word` / `dcc-mcp-excel` (Phase 2), `dcc-mcp-outlook`
  (Phase 3) — they consume Rust crates from the matching immutable release
  tag plus the version-compatible `office-host` binaries.
- Skill packs live in the owning repo; a separate
  `dcc-mcp-office-skills` repo is created only if pack volume demands it.
- SUA sharing (proposal §5): extract a neutral core later, keeping exactly
  one COM implementation — two COM implementations are forbidden.

## Consequences

- Ecosystem catalog/release tooling works unchanged; adapters stay thin.
- Cross-repo version pinning needed (mirrors `dcc-mcp-core` dependency
  conventions).
- Visio/Project/Access add app semantics to the core + a thin repo each,
  without touching PowerPoint's cadence.
