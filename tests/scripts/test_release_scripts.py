import hashlib
import importlib.util
import json
import tempfile
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
            (repository / "manifests").mkdir(parents=True)
            (repository / "manifests" / "batch.json").write_text("{}", encoding="utf-8")
            (repository / "templates").mkdir()
            (repository / "templates" / "registry.json").write_text("{}", encoding="utf-8")
            (repository / "docs").mkdir()
            (repository / "docs" / "distribution.md").write_text("install", encoding="utf-8")
            (repository / "LICENSE").write_text("MIT", encoding="utf-8")

            assets = module.package_release(host_dir, output_dir, repository, "0.2.2")

            expected_executables = [module.CANONICAL_HOST, *module.HOST_ALIASES]
            for name in expected_executables:
                self.assertEqual((output_dir / name).read_bytes(), b"office-host-fixture")

            archive = output_dir / "dcc-mcp-office-host-0.2.2-win-x64.zip"
            with zipfile.ZipFile(archive) as package:
                names = set(package.namelist())
            prefix = "dcc-mcp-office-host-0.2.2-win-x64/"
            self.assertTrue(all(f"{prefix}bin/{name}" in names for name in expected_executables))
            self.assertIn(f"{prefix}manifests/batch.json", names)
            self.assertIn(f"{prefix}templates/registry.json", names)
            self.assertIn(f"{prefix}INSTALL.md", names)
            self.assertIn(f"{prefix}release-manifest.json", names)
            self.assertIn(f"{prefix}dcc-mcp-office-host-0.2.2.spdx.json", names)

            manifest = json.loads(
                (output_dir / "release-manifest.json").read_text(encoding="utf-8")
            )
            self.assertEqual(manifest["schema"], "dcc-mcp-office-release/1")
            self.assertEqual(manifest["version"], "0.2.2")
            self.assertEqual(set(manifest["aliases"]), set(module.HOST_ALIASES))

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
