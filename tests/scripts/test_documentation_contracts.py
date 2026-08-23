import re
import subprocess
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]


def read(relative_path: str) -> str:
    return (ROOT / relative_path).read_text(encoding="utf-8")


class DocumentationContractTests(unittest.TestCase):
    def test_host_contract_example_is_repository_relative(self):
        documentation = read("README.md")

        self.assertIn(
            '$env:DCC_OFFICE_HOST_EXE = (Resolve-Path "dotnet/Office.Automation.Host/'
            'bin/Debug/net8.0-windows/dcc-office-host.exe").Path',
            documentation,
        )
        self.assertNotRegex(
            documentation,
            r"\$env:DCC_OFFICE_HOST_EXE\s*=\s*['\"]?[A-Za-z]:",
        )

    def test_changelog_entries_are_unique_within_each_release(self):
        releases = re.split(r"(?=^## \[)", read("CHANGELOG.md"), flags=re.MULTILINE)

        for release in releases:
            entries = re.findall(r"^\* .+$", release, flags=re.MULTILINE)
            self.assertEqual(len(entries), len(set(entries)))

    def test_dotnet_project_and_process_model_match_the_implementation(self):
        documentation = read("dotnet/README.md")

        for project in (
            "Office.Automation.Runtime",
            "Office.Automation.Com",
            "Office.Automation.OpenXml",
            "Office.Automation.Host",
        ):
            self.assertIn(f"`{project}`", documentation)
        self.assertNotIn("M0 skeleton", documentation)
        self.assertNotIn("OfficeInstanceResolver", documentation)
        self.assertIn("named-pipe server", documentation)
        self.assertIn("lazily creates", documentation)

    def test_phase_vocabulary_has_one_canonical_crosswalk(self):
        index = read("docs/README.md")

        for label in (
            "Proposal delivery phase",
            "Repository milestone",
            "Tool priority",
            "Phase 0",
            "Phase 5",
            "M0",
            "M4",
            "P0",
            "P2",
        ):
            self.assertIn(label, index)
        self.assertIn("docs/README.md#delivery-vocabulary", read("README.md"))
        self.assertIn("docs/README.md#delivery-vocabulary", read("AGENTS.md"))
        self.assertIn("support priority", read("crates/office-tools/src/lib.rs"))

    def test_planned_surfaces_are_not_described_as_present(self):
        agents = read("AGENTS.md")
        tests = read("tests/README.md")
        proposal = read("docs/proposals/office-automation-platform-v1.0.md")

        self.assertIn("not provisioned", agents)
        self.assertIn("empty `.gitkeep`", agents)
        self.assertIn("No evaluation runner is selected yet", tests)
        self.assertNotIn("dcc-mcp-tester", tests)
        self.assertIn("conceptual target layout", proposal)
        self.assertIn("`addins/` is planned for Phase 3", proposal)

    def test_skill_index_lists_only_existing_packs(self):
        documentation = read("skills/README.md")

        for ghost_pack in (
            "office-brand-template-migration",
            "office-document-redaction",
            "office-generate-executive-deck",
            "office-generate-technical-report",
        ):
            self.assertNotIn(ghost_pack, documentation)

    def test_generated_dashboard_output_is_untracked_by_design(self):
        generated = ROOT / "examples" / "output" / "draft-dcc-mcp-capability-dashboard.xlsx"
        relative_path = generated.relative_to(ROOT).as_posix()
        tracked = subprocess.run(
            ["git", "ls-files", "--cached", "--", relative_path],
            cwd=ROOT,
            capture_output=True,
            check=True,
            text=True,
        )

        self.assertEqual("", tracked.stdout.strip())
        self.assertIn("/examples/output/", read(".gitignore").splitlines())

    def test_dashboard_ci_command_has_explicit_line_continuations(self):
        workflow = read(".github/workflows/ci.yml")

        continuation = chr(92)
        expected = "\n".join(
            (
                f"generate_dashboard.py {continuation}",
                f"            --input examples/capability-dashboard.json {continuation}",
                "            --out /tmp/dash",
            )
        )
        self.assertIn(expected, workflow)


if __name__ == "__main__":
    unittest.main()
