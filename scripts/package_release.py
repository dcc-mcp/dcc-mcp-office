#!/usr/bin/env python3
"""Build the canonical, self-describing Office Host release bundle."""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import tempfile
import tomllib
import zipfile
from datetime import datetime, timezone
from pathlib import Path


CANONICAL_HOST = "dcc-office-host.exe"
MCP_SERVER = "dcc-mcp-office-mcp-server.exe"
HOST_ALIASES = [
    "dcc-office-powerpoint-host.exe",
    "dcc-office-word-host.exe",
    "dcc-office-excel-host.exe",
]
ALIAS_APPS = dict(zip(HOST_ALIASES, ("powerpoint", "word", "excel"), strict=True))


def workspace_version(repository: Path) -> str:
    with (repository / "Cargo.toml").open("rb") as stream:
        return str(tomllib.load(stream)["workspace"]["package"]["version"])


def file_digest(path: Path, algorithm: str) -> str:
    digest = hashlib.new(algorithm)
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sha256(path: Path) -> str:
    return file_digest(path, "sha256")


def write_json(path: Path, value: object) -> None:
    path.write_text(
        json.dumps(value, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def inventory(root: Path) -> list[dict[str, object]]:
    return [
        {
            "path": path.relative_to(root).as_posix(),
            "sha1": file_digest(path, "sha1"),
            "sha256": sha256(path),
            "size": path.stat().st_size,
        }
        for path in sorted(candidate for candidate in root.rglob("*") if candidate.is_file())
    ]


def make_sbom(version: str, files: list[dict[str, object]]) -> dict[str, object]:
    created = (
        datetime.now(timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    )
    spdx_files = []
    relationships = [
        {
            "spdxElementId": "SPDXRef-DOCUMENT",
            "relationshipType": "DESCRIBES",
            "relatedSpdxElement": "SPDXRef-Package",
        }
    ]
    for index, entry in enumerate(files, start=1):
        spdx_id = f"SPDXRef-File-{index}"
        spdx_files.append(
            {
                "SPDXID": spdx_id,
                "fileName": f"./{entry['path']}",
                "checksums": [
                    {"algorithm": "SHA1", "checksumValue": entry["sha1"]},
                    {"algorithm": "SHA256", "checksumValue": entry["sha256"]},
                ],
            }
        )
        relationships.append(
            {
                "spdxElementId": "SPDXRef-Package",
                "relationshipType": "CONTAINS",
                "relatedSpdxElement": spdx_id,
            }
        )
    return {
        "SPDXID": "SPDXRef-DOCUMENT",
        "spdxVersion": "SPDX-2.3",
        "dataLicense": "CC0-1.0",
        "name": f"dcc-mcp-office-host-{version}",
        "documentNamespace": (
            f"https://github.com/dcc-mcp/dcc-mcp-office/releases/tag/v{version}/sbom"
        ),
        "creationInfo": {"created": created, "creators": ["Tool: package_release.py"]},
        "packages": [
            {
                "SPDXID": "SPDXRef-Package",
                "name": "dcc-mcp-office-host",
                "versionInfo": version,
                "downloadLocation": "NOASSERTION",
                "filesAnalyzed": True,
                "licenseConcluded": "MIT",
                "licenseDeclared": "MIT",
                "copyrightText": "NOASSERTION",
                "packageVerificationCode": {
                    "packageVerificationCodeValue": hashlib.sha1(
                        "".join(sorted(str(entry["sha1"]) for entry in files)).encode(
                            "ascii"
                        )
                    ).hexdigest()
                },
            }
        ],
        "files": spdx_files,
        "relationships": relationships,
    }


def write_archive(source: Path, destination: Path, prefix: str) -> None:
    with zipfile.ZipFile(
        destination,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=9,
    ) as package:
        files = sorted(candidate for candidate in source.rglob("*") if candidate.is_file())
        for path in files:
            name = f"{prefix}/{path.relative_to(source).as_posix()}"
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            with path.open("rb") as source_stream, package.open(info, "w") as archive_stream:
                shutil.copyfileobj(source_stream, archive_stream, length=1024 * 1024)


def package_release(
    host_dir: Path,
    mcp_server: Path,
    output_dir: Path,
    repository: Path,
    version: str,
) -> list[Path]:
    source_host = host_dir / CANONICAL_HOST
    if not source_host.is_file():
        raise FileNotFoundError(f"published host is missing: {source_host}")
    if not mcp_server.is_file():
        raise FileNotFoundError(f"published MCP server is missing: {mcp_server}")
    output_dir.mkdir(parents=True, exist_ok=True)

    host_assets = [
        output_dir / CANONICAL_HOST,
        *[output_dir / name for name in HOST_ALIASES],
    ]
    for destination in host_assets:
        shutil.copyfile(source_host, destination)
    mcp_server_asset = output_dir / MCP_SERVER
    shutil.copyfile(mcp_server, mcp_server_asset)
    executable_assets = [*host_assets, mcp_server_asset]

    package_name = f"dcc-mcp-office-host-{version}-win-x64"
    sbom_name = f"dcc-mcp-office-host-{version}.spdx.json"
    with tempfile.TemporaryDirectory(prefix="dcc-office-package-") as temporary:
        staging = Path(temporary)
        binary_directory = staging / "bin"
        binary_directory.mkdir()
        for asset in executable_assets:
            shutil.copyfile(asset, binary_directory / asset.name)

        shutil.copytree(repository / "manifests", staging / "manifests")
        catalog = repository / "crates" / "office-protocol" / "office-rpc.catalog.json"
        if catalog.is_file():
            shutil.copyfile(catalog, staging / "manifests" / catalog.name)
        shutil.copytree(repository / "templates", staging / "templates")
        shutil.copyfile(
            repository / "docs" / "distribution.md",
            staging / "INSTALL.md",
        )
        shutil.copyfile(
            repository / "docs" / "operations.md",
            staging / "OPERATIONS.md",
        )
        shutil.copyfile(repository / "LICENSE", staging / "LICENSE")

        manifest = {
            "schema": "dcc-mcp-office-release/1",
            "version": version,
            "platform": "win-x64",
            "canonical_host": CANONICAL_HOST,
            "mcp_server": MCP_SERVER,
            "aliases": ALIAS_APPS,
            "install_root": r"%LOCALAPPDATA%\dcc-mcp\office-host\<version>",
            "locator_order": [
                "DCC_OFFICE_HOST_EXE",
                "gateway_sibling",
                "versioned_install",
            ],
            "protocol_version": "office-rpc/1",
            "provider_compatibility": "same minor before 1.0; same major from 1.0",
            "files": inventory(staging),
        }
        write_json(staging / "release-manifest.json", manifest)
        write_json(staging / sbom_name, make_sbom(version, inventory(staging)))

        manifest_asset = output_dir / "release-manifest.json"
        sbom_asset = output_dir / sbom_name
        shutil.copyfile(staging / "release-manifest.json", manifest_asset)
        shutil.copyfile(staging / sbom_name, sbom_asset)
        archive_asset = output_dir / f"{package_name}.zip"
        write_archive(staging, archive_asset, package_name)

    assets = [*executable_assets, archive_asset, manifest_asset, sbom_asset]
    checksum_asset = output_dir / "SHA256SUMS"
    checksum_asset.write_text(
        "".join(f"{sha256(asset)}  {asset.name}\n" for asset in assets),
        encoding="utf-8",
        newline="\n",
    )
    return [*assets, checksum_asset]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--host-dir", type=Path, required=True)
    parser.add_argument("--mcp-server", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--version")
    parser.add_argument(
        "--repository",
        type=Path,
        default=Path(__file__).resolve().parents[1],
    )
    arguments = parser.parse_args()
    repository = arguments.repository.resolve()
    assets = package_release(
        arguments.host_dir.resolve(),
        arguments.mcp_server.resolve(),
        arguments.output_dir.resolve(),
        repository,
        arguments.version or workspace_version(repository),
    )
    print("\n".join(str(asset) for asset in assets))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
