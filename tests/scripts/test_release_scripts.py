import hashlib
import importlib.util
import json
import re
import tempfile
import tomllib
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]


def load_script(name: str):
    path = ROOT / "scripts" / name
    spec = importlib.util.spec_from_file_location(path.stem, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class SourceDistributionTests(unittest.TestCase):
    def test_distribution_documents_immutable_tag_dependencies(self):
        documentation = (ROOT / "docs" / "distribution.md").read_text(encoding="utf-8")
        dependency = (
            'dcc-mcp-office-client = { git = '
            '"https://github.com/dcc-mcp/dcc-mcp-office", tag = "v0.2.2" }'
        )
        self.assertIn(
            dependency,
            documentation,
        )
        self.assertIn("Release tags are treated as immutable", documentation)


class ReleasePleaseConfigTests(unittest.TestCase):
    def test_release_please_updates_every_workspace_lock_entry(self):
        workspace = tomllib.loads((ROOT / "Cargo.toml").read_text(encoding="utf-8"))
        workspace_version = workspace["workspace"]["package"]["version"]
        workspace_packages = {
            tomllib.loads((ROOT / member / "Cargo.toml").read_text(encoding="utf-8"))[
                "package"
            ]["name"]
            for member in workspace["workspace"]["members"]
        }

        lock = tomllib.loads((ROOT / "Cargo.lock").read_text(encoding="utf-8"))
        locked_versions = {
            package["name"]: package["version"]
            for package in lock["package"]
            if package["name"] in workspace_packages
        }
        self.assertEqual(set(locked_versions), workspace_packages)
        self.assertEqual(set(locked_versions.values()), {workspace_version})

        config = json.loads(
            (ROOT / "release-please-config.json").read_text(encoding="utf-8")
        )
        lock_updaters = [
            entry
            for entry in config["packages"]["."]["extra-files"]
            if entry["path"] == "Cargo.lock"
        ]
        self.assertEqual(len(lock_updaters), 1)
        self.assertEqual(lock_updaters[0]["type"], "toml")
        configured_packages = set(
            re.findall(
                r"@\.name\.value === '([^']+)'",
                lock_updaters[0]["jsonpath"],
            )
        )
        self.assertEqual(configured_packages, workspace_packages)


class TemplateCatalogTests(unittest.TestCase):
    def test_every_file_catalog_entry_resolves_to_a_matching_package(self):
        template_root = ROOT / "templates"
        catalog = json.loads(
            (template_root / "registry.json").read_text(encoding="utf-8")
        )

        for entry in catalog["templates"]:
            if not entry["source"].startswith("file://"):
                continue
            package_path = template_root / entry["source"].removeprefix("file://")
            package = json.loads(
                (package_path / "package.json").read_text(encoding="utf-8")
            )
            self.assertEqual(package["uri"], entry["uri"])
            self.assertEqual(package["version"], entry["version"])
            self.assertEqual(set(package["layouts"]), set(entry["layouts"]))


class ReleasePackageTests(unittest.TestCase):
    def test_workspace_version_requires_no_rust_toolchain(self):
        module = load_script("package_release.py")
        with tempfile.TemporaryDirectory() as temporary:
            repository = Path(temporary)
            (repository / "Cargo.toml").write_text(
                '[workspace.package]\nversion = "7.8.9"\n',
                encoding="utf-8",
            )

            self.assertEqual(module.workspace_version(repository), "7.8.9")

    def test_package_contains_aliases_contracts_checksums_and_sbom(self):
        module = load_script("package_release.py")
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            host_dir = root / "host"
            output_dir = root / "release"
            repository = root / "repository"
            host_dir.mkdir()
            (host_dir / module.CANONICAL_HOST).write_bytes(b"office-host-fixture")
            mcp_server = root / module.MCP_SERVER
            mcp_server.write_bytes(b"office-mcp-server-fixture")
            (repository / "manifests").mkdir(parents=True)
            (repository / "manifests" / "batch.json").write_text("{}", encoding="utf-8")
            (repository / "templates").mkdir()
            (repository / "templates" / "registry.json").write_text("{}", encoding="utf-8")
            package = repository / "templates" / "presentations" / "studio-light"
            package.mkdir(parents=True)
            (package / "package.json").write_text("{}", encoding="utf-8")
            (package / "theme.xml").write_text("<theme/>", encoding="utf-8")
            (repository / "docs").mkdir()
            (repository / "docs" / "distribution.md").write_text("install", encoding="utf-8")
            (repository / "docs" / "operations.md").write_text("operate", encoding="utf-8")
            (repository / "LICENSE").write_text("MIT", encoding="utf-8")

            assets = module.package_release(
                host_dir,
                mcp_server,
                output_dir,
                repository,
                "0.2.2",
            )

            expected_executables = [module.CANONICAL_HOST, *module.HOST_ALIASES]
            for name in expected_executables:
                self.assertEqual((output_dir / name).read_bytes(), b"office-host-fixture")
            self.assertEqual(
                (output_dir / module.MCP_SERVER).read_bytes(),
                b"office-mcp-server-fixture",
            )

            archive = output_dir / "dcc-mcp-office-host-0.2.2-win-x64.zip"
            with zipfile.ZipFile(archive) as package:
                names = set(package.namelist())
            prefix = "dcc-mcp-office-host-0.2.2-win-x64/"
            self.assertTrue(all(f"{prefix}bin/{name}" in names for name in expected_executables))
            self.assertIn(f"{prefix}bin/{module.MCP_SERVER}", names)
            self.assertIn(f"{prefix}manifests/batch.json", names)
            self.assertIn(f"{prefix}templates/registry.json", names)
            self.assertIn(
                f"{prefix}templates/presentations/studio-light/package.json",
                names,
            )
            self.assertIn(
                f"{prefix}templates/presentations/studio-light/theme.xml",
                names,
            )
            self.assertIn(f"{prefix}INSTALL.md", names)
            self.assertIn(f"{prefix}OPERATIONS.md", names)
            self.assertIn(f"{prefix}release-manifest.json", names)
            self.assertIn(f"{prefix}dcc-mcp-office-host-0.2.2.spdx.json", names)

            manifest = json.loads(
                (output_dir / "release-manifest.json").read_text(encoding="utf-8")
            )
            self.assertEqual(manifest["schema"], "dcc-mcp-office-release/1")
            self.assertEqual(manifest["version"], "0.2.2")
            self.assertEqual(set(manifest["aliases"]), set(module.HOST_ALIASES))
            self.assertEqual(manifest["mcp_server"], module.MCP_SERVER)

            sbom_name = "dcc-mcp-office-host-0.2.2.spdx.json"
            sbom = json.loads((output_dir / sbom_name).read_text(encoding="utf-8"))
            self.assertEqual(sbom["spdxVersion"], "SPDX-2.3")
            self.assertRegex(
                sbom["packages"][0]["packageVerificationCode"][
                    "packageVerificationCodeValue"
                ],
                r"^[0-9a-f]{40}$",
            )
            self.assertIn(
                {
                    "spdxElementId": "SPDXRef-DOCUMENT",
                    "relationshipType": "DESCRIBES",
                    "relatedSpdxElement": "SPDXRef-Package",
                },
                sbom["relationships"],
            )

            checksums = {}
            for line in (output_dir / "SHA256SUMS").read_text(encoding="utf-8").splitlines():
                digest, name = line.split("  ", 1)
                checksums[name] = digest
            checked_assets = {
                asset.name for asset in assets if asset.name != "SHA256SUMS"
            }
            self.assertEqual(set(checksums), checked_assets)
            for name, digest in checksums.items():
                actual = hashlib.sha256((output_dir / name).read_bytes()).hexdigest()
                self.assertEqual(actual, digest)


if __name__ == "__main__":
    unittest.main()
