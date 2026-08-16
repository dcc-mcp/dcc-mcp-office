# Changelog

## [0.1.1](https://github.com/dcc-mcp/dcc-mcp-office/compare/v0.1.0...v0.1.1) (2026-08-16)


### Features

* full layout parity in the C# host (all 10 layouts + pictures + brand logo) ([#4](https://github.com/dcc-mcp/dcc-mcp-office/issues/4)) ([26b0968](https://github.com/dcc-mcp/dcc-mcp-office/commit/26b0968282c08d9cae4ca84f26b7ad62f1bfdc72))
* production dashboard skill + CI ([#1](https://github.com/dcc-mcp/dcc-mcp-office/issues/1)) ([8f2b6b5](https://github.com/dcc-mcp/dcc-mcp-office/commit/8f2b6b5044a36a8ba888c7a61d77063505d8690c))
* self-implemented Open XML host (M1, zero NuGet dependencies) ([#3](https://github.com/dcc-mcp/dcc-mcp-office/issues/3)) ([db9fd47](https://github.com/dcc-mcp/dcc-mcp-office/commit/db9fd47dcd5cbbf9b1a14d81bc59599eb3b579d4))
* vx-managed .NET toolchain, CI via loonghao/vx, and release-please auto-release ([#6](https://github.com/dcc-mcp/dcc-mcp-office/issues/6)) ([6909a31](https://github.com/dcc-mcp/dcc-mcp-office/commit/6909a31f5ff9534b1cfb0ff1df287582cc02ab50))
* vx-managed .NET toolchain, CI via loonghao/vx, and release-please auto-release ([#6](https://github.com/dcc-mcp/dcc-mcp-office/issues/6)) ([6909a31](https://github.com/dcc-mcp/dcc-mcp-office/commit/6909a31f5ff9534b1cfb0ff1df287582cc02ab50))


### Bug Fixes

* release-please simple release type with Cargo.toml extra-file ([#7](https://github.com/dcc-mcp/dcc-mcp-office/issues/7)) ([08710db](https://github.com/dcc-mcp/dcc-mcp-office/commit/08710dba3567979f56dc156a1979af97789e043e))
* release-please simple release type with Cargo.toml extra-file ([#7](https://github.com/dcc-mcp/dcc-mcp-office/issues/7)) ([08710db](https://github.com/dcc-mcp/dcc-mcp-office/commit/08710dba3567979f56dc156a1979af97789e043e))
* skill spec compliance + official lint gate ([#2](https://github.com/dcc-mcp/dcc-mcp-office/issues/2)) ([58b666c](https://github.com/dcc-mcp/dcc-mcp-office/commit/58b666c695541fbdcaf52a5d433e5ce41b123a83))

## [0.1.0] - unreleased

### Added

- M0 scaffold: `office-protocol` / `office-ir` / `office-tools` /
  `office-jobs` / `office-security` crates with schema drafts and
  round-trip tests.
- C# skeleton: `Office.Automation.Runtime` (STA dispatcher),
  `Office.Automation.OpenXml`, `Office.Automation.Host` (`office-host`).
- `office-batch-to-pdf` skill pack, capability manifest example, template
  registry layout, ADR-001..006, platform proposal record.
