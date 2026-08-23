# Office Host distribution and discovery

Versioned GitHub Releases are the durable binary distribution channel. Each
release publishes:

- `dcc-office-host.exe`, plus PowerPoint, Word, and Excel alias executables;
- `dcc-mcp-office-mcp-server.exe`, the reference stdio MCP entry point;
- `dcc-mcp-office-host-<version>-win-x64.zip`, containing the binaries,
  protocol catalog, input schemas, materialized template packages, install contract, release
  manifest, license, and SPDX 2.3 SBOM;
- `release-manifest.json`, the SPDX SBOM, and `SHA256SUMS`;
- GitHub build-provenance and SBOM attestations.

The binaries are not Authenticode-signed. Verify both the SHA-256 digest and
GitHub attestation before installing them:

```powershell
gh release download v0.2.2 --repo dcc-mcp/dcc-mcp-office --pattern "dcc-mcp-office-host-0.2.2-win-x64.zip" --pattern SHA256SUMS
gh attestation verify dcc-mcp-office-host-0.2.2-win-x64.zip --repo dcc-mcp/dcc-mcp-office
$expected = (Select-String -LiteralPath SHA256SUMS -Pattern "  dcc-mcp-office-host-0.2.2-win-x64.zip$").Line.Split("  ")[0]
$actual = (Get-FileHash -Algorithm SHA256 -LiteralPath dcc-mcp-office-host-0.2.2-win-x64.zip).Hash.ToLowerInvariant()
if ($actual -ne $expected) { throw "SHA-256 verification failed" }
```

Install the canonical executables at:

```text
%LOCALAPPDATA%\dcc-mcp\office-host\<version>\dcc-office-host.exe
%LOCALAPPDATA%\dcc-mcp\office-host\<version>\dcc-mcp-office-mcp-server.exe
```

Copy the bundle's `templates` directory beside the executable to install its
versioned packages. Independently managed packages may instead be placed in
`%LOCALAPPDATA%\dcc-mcp\office-templates` or supplied with a repeatable
`--template-dir=<path>` argument. The Host never searches its working
directory and advertises only packages that it successfully validates and
materializes.

`dcc-mcp-office-client` locates the executable in this strict order:

1. `DCC_OFFICE_HOST_EXE` (an explicit missing path fails closed);
2. `dcc-office-host.exe` beside the running gateway executable;
3. the versioned per-user installation above.

Discovery never downloads or launches a process. The gateway lifecycle owner
starts the located binary with the application, pipe, and workspace boundary,
for example:

```powershell
dcc-office-host.exe --app=powerpoint --pipe-name=<pipe> --workspace-root=<root>
```

The subsequent `office.host.handshake` enforces provider compatibility: before
1.0 the client and Host must share major and minor versions; from 1.0 onward
they must share the major version. A missing or incompatible Host is an
installation error and must not silently fall back to another executable.

## Reference MCP server

The release bundle is directly consumable by any client that can launch a
stdio MCP server. Bind one server process to one Office application and one
existing workspace directory. For example:

```json
{
  "mcpServers": {
    "office-powerpoint": {
      "command": "<install-root>\\dcc-mcp-office-mcp-server.exe",
      "args": [
        "--app=powerpoint",
        "--workspace-root=<existing-workspace>"
      ]
    }
  }
}
```

`--app` is required and currently accepts `powerpoint`, `word`, or `excel`.
`--workspace-root` defaults to the server working directory. The server finds
the matching Host through the strict locator order above; `--host=<path>` is
available for an explicit installation. MCP JSON-RPC uses stdout exclusively,
while lifecycle diagnostics use stderr.

At startup the server handshakes with the Host, exposes only capabilities
advertised for the selected application, and derives every tool's input and
output schema from the canonical catalog. Handshake does not attach COM or
launch an empty Office window. Invalid schemas and policy relaxations fail
before dispatch; sidecar failures preserve their `OFFICE_*` code, data, and
retryability in the MCP tool result.

Rust crates use the release's immutable Git tag as their source distribution
channel. Pin the tag explicitly instead of tracking a branch:

```toml
dcc-mcp-office-client = { git = "https://github.com/dcc-mcp/dcc-mcp-office", tag = "v0.2.2" }
```

The other package names follow the same `dcc-mcp-office-*` names listed in the
workspace manifest. Release tags are treated as immutable, and the crate
version, Host version, release tag, protocol dependency version, and .NET
metadata are stamped in lockstep by release-please. Consumers must use the
matching Host release.
