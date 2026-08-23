# Changelog

## [0.2.3](https://github.com/dcc-mcp/dcc-mcp-office/compare/v0.2.2...v0.2.3) (2026-08-23)


### Features

* add external office templates ([#40](https://github.com/dcc-mcp/dcc-mcp-office/issues/40)) ([0efe573](https://github.com/dcc-mcp/dcc-mcp-office/commit/0efe5730b1c5e757e6a0e4ae3b8e5f04a28ecfa8))
* add office job runtime ([#38](https://github.com/dcc-mcp/dcc-mcp-office/issues/38)) ([e71797d](https://github.com/dcc-mcp/dcc-mcp-office/commit/e71797d82b065750f280541798dd7ffd2458c1b5))
* add reference Office MCP server ([08290bd](https://github.com/dcc-mcp/dcc-mcp-office/commit/08290bdd3294f6907bae287c58939e3e57efa4c6))
* add verified office distribution ([#39](https://github.com/dcc-mcp/dcc-mcp-office/issues/39)) ([6582b06](https://github.com/dcc-mcp/dcc-mcp-office/commit/6582b063a60ee23f3b9029070a567736be46da48))
* improve Office host operability ([db29df5](https://github.com/dcc-mcp/dcc-mcp-office/commit/db29df541b10535295d31b5927ec1fcb3cb4fdd4)), closes [#25](https://github.com/dcc-mcp/dcc-mcp-office/issues/25)


### Bug Fixes

* enforce Office write safety contracts ([#36](https://github.com/dcc-mcp/dcc-mcp-office/issues/36)) ([75b5f24](https://github.com/dcc-mcp/dcc-mcp-office/commit/75b5f2432c21c0473e87eec9bae761920798a6b4))


### Documentation

* align Office documentation contracts ([8600ff4](https://github.com/dcc-mcp/dcc-mcp-office/commit/8600ff4c4a413160295b9b739892a49e4c1a06a1))
* close architecture health review ([4f37926](https://github.com/dcc-mcp/dcc-mcp-office/commit/4f3792616eb4c6d6cff31baec658a6ad958449ca))

## [0.2.2](https://github.com/dcc-mcp/dcc-mcp-office/compare/v0.2.1...v0.2.2) (2026-08-23)


### Bug Fixes

* bound Office client pipe operations ([#32](https://github.com/dcc-mcp/dcc-mcp-office/issues/32)) ([d8c8b42](https://github.com/dcc-mcp/dcc-mcp-office/commit/d8c8b42a2061644b99b76cec38430531f7200e8c))
* classify COM failures by HRESULT ([#31](https://github.com/dcc-mcp/dcc-mcp-office/issues/31)) ([33d35a6](https://github.com/dcc-mcp/dcc-mcp-office/commit/33d35a6649e7ca1c9a562719d706a6bbfd78759e))
* harden STA timeout recovery ([#29](https://github.com/dcc-mcp/dcc-mcp-office/issues/29)) ([6b8cf42](https://github.com/dcc-mcp/dcc-mcp-office/commit/6b8cf4278c4181c9b394f8d6c4923660cec9bfcd))
* synchronize Office host versions ([#33](https://github.com/dcc-mcp/dcc-mcp-office/issues/33)) ([6c9db25](https://github.com/dcc-mcp/dcc-mcp-office/commit/6c9db25782b354c09fed501d92b2fedda1848a86))


### Code Refactoring

* centralize Office RPC contracts ([#35](https://github.com/dcc-mcp/dcc-mcp-office/issues/35)) ([9f0597f](https://github.com/dcc-mcp/dcc-mcp-office/commit/9f0597fe3e5e2d73a5192e03d729baa2c3683756))

## [0.2.1](https://github.com/dcc-mcp/dcc-mcp-office/compare/v0.2.0...v0.2.1) (2026-08-16)


### Bug Fixes

* resolve COM paths to absolute before render/batch/inspect ([#13](https://github.com/dcc-mcp/dcc-mcp-office/issues/13)) ([01f6345](https://github.com/dcc-mcp/dcc-mcp-office/commit/01f63459f45d46c6f8f33871b8bad9245fc2ae00))

## [0.2.0](https://github.com/dcc-mcp/dcc-mcp-office/compare/v0.1.1...v0.2.0) (2026-08-16)


### Features

* M1 COM sidecar MVP — named-pipe server, per-app COM backends, batch + render capabilities ([#11](https://github.com/dcc-mcp/dcc-mcp-office/issues/11)) ([b4dda5d](https://github.com/dcc-mcp/dcc-mcp-office/commit/b4dda5d5c0ea5ee5a4af9cebd72927885bce5fbb))


### Bug Fixes

* align intra-workspace version pins with the 0.2.0 release ([39f8759](https://github.com/dcc-mcp/dcc-mcp-office/commit/39f8759a9fb0071dfe88f2e5116a8703b04f0270))
* drop version pins on intra-workspace path deps so release-please bumps stay green ([c698098](https://github.com/dcc-mcp/dcc-mcp-office/commit/c69809893a10bcd1c9c020fc627b4eac395f5654))

## [0.1.1](https://github.com/dcc-mcp/dcc-mcp-office/compare/v0.1.0...v0.1.1) (2026-08-16)


### Features

* full layout parity in the C# host (all 11 layouts + pictures + brand logo) ([#4](https://github.com/dcc-mcp/dcc-mcp-office/issues/4)) ([26b0968](https://github.com/dcc-mcp/dcc-mcp-office/commit/26b0968282c08d9cae4ca84f26b7ad62f1bfdc72))
* production dashboard skill + CI ([#1](https://github.com/dcc-mcp/dcc-mcp-office/issues/1)) ([8f2b6b5](https://github.com/dcc-mcp/dcc-mcp-office/commit/8f2b6b5044a36a8ba888c7a61d77063505d8690c))
* self-implemented Open XML host (M1, zero NuGet dependencies) ([#3](https://github.com/dcc-mcp/dcc-mcp-office/issues/3)) ([db9fd47](https://github.com/dcc-mcp/dcc-mcp-office/commit/db9fd47dcd5cbbf9b1a14d81bc59599eb3b579d4))
* vx-managed .NET toolchain, CI via loonghao/vx, and release-please auto-release ([#6](https://github.com/dcc-mcp/dcc-mcp-office/issues/6)) ([6909a31](https://github.com/dcc-mcp/dcc-mcp-office/commit/6909a31f5ff9534b1cfb0ff1df287582cc02ab50))


### Bug Fixes

* release-please PAT token + pin loonghao/vx action version ([#9](https://github.com/dcc-mcp/dcc-mcp-office/issues/9)) ([efdebb3](https://github.com/dcc-mcp/dcc-mcp-office/commit/efdebb3b598ed378b6879a4383b64107d7e0cbb3))
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
