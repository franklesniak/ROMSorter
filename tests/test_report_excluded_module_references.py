"""Exercise excluded-module cleanup reporting."""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT_PATH = REPO_ROOT / ".template-sync" / "scripts" / "report_excluded_module_references.py"
MARKER_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-marker.schema.json"
MANIFEST_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-manifest.schema.json"
SOURCE_REPO = "https://github.com/franklesniak/copilot-repo-template.git"
FULL_SHA = "0123456789abcdef0123456789abcdef01234567"
INCLUDED_MODULES = [
    "baseline",
    "agent-instructions",
    "github-actions",
    "github-platform",
    "github-templates",
    "markdown",
    "template-sync-support",
    "yaml",
]
MODULE_DEFINITIONS = {
    "baseline": "Baseline files.",
    "agent-instructions": "Agent instruction files.",
    "github-platform": "GitHub platform files.",
    "github-actions": "GitHub Actions workflows.",
    "github-templates": "GitHub collaboration surfaces.",
    "template-onboarding": "Template onboarding files.",
    "template-sync-support": "Template sync support files.",
    "markdown": "Markdown files.",
    "powershell": "PowerShell files.",
    "json": "JSON files.",
    "yaml": "YAML files.",
    "schema": "Schema files.",
    "python": "Python project files.",
    "terraform": "Terraform files.",
}


def _write_text(repo_root: Path, relative_path: str, text: str = "placeholder\n") -> None:
    """Write a UTF-8 text file below a fixture repository root."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _write_yaml(repo_root: Path, relative_path: str, data: dict[str, Any]) -> None:
    """Write a YAML file below a fixture repository root."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data, sort_keys=False), encoding="utf-8")


def _copy_schemas(repo_root: Path) -> None:
    """Copy real template-sync schemas into a fixture repository."""
    schemas_dir = repo_root / "schemas"
    schemas_dir.mkdir(parents=True, exist_ok=True)
    for source_path in (MARKER_SCHEMA_PATH, MANIFEST_SCHEMA_PATH):
        shutil.copyfile(source_path, schemas_dir / source_path.name)


def _run_git(repo_root: Path, *args: str) -> None:
    """Run a git command in a fixture repository."""
    subprocess.run(
        ["git", *args],
        cwd=repo_root,
        check=True,
        capture_output=True,
        text=True,
    )


def _run_report(repo_root: Path, *extra_args: str) -> subprocess.CompletedProcess[str]:
    """Run the excluded-module cleanup reporter."""
    return subprocess.run(
        [
            sys.executable,
            str(SCRIPT_PATH),
            "--repo-root",
            str(repo_root),
            *extra_args,
        ],
        check=False,
        capture_output=True,
        text=True,
    )


def test_reporter_script_entrypoint_help_smoke() -> None:
    """The excluded-module reporter remains executable through its script entry point."""
    result = subprocess.run(
        [sys.executable, str(SCRIPT_PATH), "--help"],
        check=False,
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    assert "usage: report_excluded_module_references.py" in result.stdout


def _manifest() -> dict[str, Any]:
    """Build a schema-valid manifest fixture with cross-module cleanup surfaces."""
    return {
        "template_manifest": {
            "version": 2,
            "modules": [
                {"name": name, "description": description}
                for name, description in MODULE_DEFINITIONS.items()
            ],
            "path_mappings": [
                {"pattern": ".template-sync/marker.yml", "requires_all": ["template-sync-support"]},
                {
                    "pattern": ".template-sync/manifest.yml",
                    "requires_all": ["template-sync-support"],
                },
                {
                    "pattern": ".template-sync/scripts/**",
                    "requires_all": ["template-sync-support"],
                },
                {"pattern": "README.md", "requires_all": ["baseline"]},
                {
                    "pattern": "AGENTS.md",
                    "requires_all": ["agent-instructions"],
                    "notes": "Contains Markdown-safe *-reference-only inline blocks.",
                },
                {"pattern": ".pre-commit-config.yaml", "requires_all": ["baseline"]},
                {"pattern": ".github/dependabot.yml", "requires_all": ["github-platform"]},
                {"pattern": ".github/ISSUE_TEMPLATE/**", "requires_all": ["github-templates"]},
                {
                    "pattern": ".github/pull_request_template.md",
                    "requires_all": ["github-templates"],
                },
                {
                    "pattern": ".github/workflows/data-ci.yml",
                    "requires_all": ["github-actions"],
                    "requires_any": ["yaml", "schema", "template-sync-support"],
                },
                {"pattern": "package.json", "requires_all": ["markdown"]},
                {"pattern": "templates/json/**", "requires_all": ["json"]},
                {"pattern": "schemas/**", "requires_all": ["schema"]},
                {"pattern": "pyproject.toml", "requires_all": ["python"]},
                {"pattern": "src/copilot_repo_template/**", "requires_all": ["python"]},
                {"pattern": ".github/workflows/python-ci.yml", "requires_all": ["python"]},
                {"pattern": "modules/**", "requires_all": ["terraform"]},
            ],
            "filtering": {
                "default_semantics": "AND",
                "requires_any_semantics": "OR",
                "path_matching": "most_specific_match_wins",
                "same_specificity_action": "union_modules",
                "unmapped_action": "surface_for_owner",
            },
            "notes": {
                "downstream_retention": "Downstream repositories keep marker data for syncs.",
            },
        }
    }


def _marker(
    *,
    local_overrides: list[dict[str, str]] | None = None,
    deferred_candidates: list[dict[str, str]] | None = None,
) -> dict[str, Any]:
    """Build a schema-valid marker fixture."""
    template_sync: dict[str, Any] = {
        "source_repo": SOURCE_REPO,
        "last_reviewed_template_commit": FULL_SHA,
        "included_modules": INCLUDED_MODULES,
    }
    if local_overrides is not None:
        template_sync["local_overrides"] = local_overrides
    if deferred_candidates is not None:
        template_sync["deferred_protected_candidates"] = deferred_candidates
    return {"template_sync": template_sync}


def _write_common_repo(
    repo_root: Path,
    *,
    marker: dict[str, Any] | None = None,
    include_marker: bool = True,
    include_reference_content: bool = True,
) -> None:
    """Write a partial-adoption fixture repository."""
    _run_git(repo_root, "init")
    _copy_schemas(repo_root)
    _write_yaml(repo_root, ".template-sync/manifest.yml", _manifest())
    if include_marker:
        _write_yaml(repo_root, ".template-sync/marker.yml", marker or _marker())
    _write_text(repo_root, ".template-sync/scripts/report_excluded_module_references.py")
    _write_text(
        repo_root,
        ".pre-commit-config.yaml",
        (
            "repos:\n"
            "  - repo: https://github.com/pre-commit/pre-commit-hooks\n"
            "    rev: v6.0.0\n"
            "    hooks:\n"
            "      - id: check-json\n"
            "# template-sync: begin python-only\n"
            "  - repo: https://github.com/psf/black\n"
            "    rev: 26.5.1\n"
            "    hooks:\n"
            "      - id: black\n"
            "  - repo: https://github.com/astral-sh/ruff-pre-commit\n"
            "    rev: v0.15.15\n"
            "    hooks:\n"
            "      - id: ruff-check\n"
            "# template-sync: end python-only\n"
            "# template-sync: begin terraform-only\n"
            "  - repo: local\n"
            "    hooks:\n"
            "      - id: terraform-fmt\n"
            "        entry: python .github/scripts/terraform_hooks.py fmt\n"
            "# template-sync: end terraform-only\n"
        ),
    )
    _write_text(
        repo_root,
        ".github/workflows/data-ci.yml",
        (
            "name: Data file linting\n"
            "on: pull_request\n"
            "permissions: {}\n"
            "jobs:\n"
            "  data-file-linting:\n"
            "    runs-on: ubuntu-latest\n"
            "    steps:\n"
            "      - name: Run check-json\n"
            "        run: pre-commit run check-json --all-files\n"
            "      # template-sync: begin schema-only\n"
            "      - name: Run validate-example-config-schema\n"
            "        run: pre-commit run validate-example-config-schema --all-files\n"
            "      # template-sync: end schema-only\n"
        ),
    )
    _write_text(
        repo_root,
        ".github/dependabot.yml",
        (
            "version: 2\n"
            "updates:\n"
            '  - package-ecosystem: "pip"\n'
            '    directory: "/"\n'
            "    schedule:\n"
            '      interval: "weekly"\n'
            '  - package-ecosystem: "github-actions"\n'
            '    directory: "/"\n'
            "    schedule:\n"
            '      interval: "weekly"\n'
            '  - package-ecosystem: "pre-commit"\n'
            '    directory: "/"\n'
            "    schedule:\n"
            '      interval: "weekly"\n'
        ),
    )
    _write_text(
        repo_root,
        "AGENTS.md",
        (
            "# Agent Instructions\n\n"
            "<!-- template-sync: begin python-reference-only -->\n"
            "Python validation command: `pytest tests/ -v`.\n"
            "<!-- template-sync: end python-reference-only -->\n"
            "Terraform validation command: `terraform test`.\n"
            "Pyright validation command: `python -m pyright --project pyrightconfig.json`.\n"
            'Split pytest command: `pytest tests/ -m "not slow" -v --cov --cov-report=term-missing`.\n'
        ),
    )
    _write_text(
        repo_root,
        ".github/ISSUE_TEMPLATE/bug_report.yml",
        (
            "name: Bug Report\n"
            "description: Report a bug\n"
            "body:\n"
            "  - type: markdown\n"
            "    attributes:\n"
            "      value: |\n"
            "        See [schema](../../schemas/example-config.schema.json).\n"
        ),
    )
    _write_text(
        repo_root,
        ".github/ISSUE_TEMPLATE/config.yml",
        (
            "blank_issues_enabled: true\n"
            "contact_links:\n"
            "  - name: Schema Guide\n"
            "    url: https://github.com/OWNER/REPO/blob/HEAD/schemas/README.md\n"
            "    about: Read schema guidance\n"
        ),
    )
    _write_text(
        repo_root,
        ".github/pull_request_template.md",
        (
            "## Checklist\n\n"
            "### General\n\n"
            "- [ ] Baseline checks pass\n\n"
            "### Python-Specific (if applicable)\n\n"
            "- [ ] `pytest` passes locally\n\n"
            "### Schema-Specific (if applicable)\n\n"
            "- [ ] If a `check-jsonschema` hook was changed, schema docs were reviewed\n"
        ),
    )
    _write_text(repo_root, "templates/json/example.json", "{}\n")
    if include_reference_content:
        _write_text(
            repo_root,
            "README.md",
            (
                "# Downstream\n\n"
                "See [JSON starter](templates/json/example.json).\n\n"
                "See [upstream JSON]"
                "(https://github.com/franklesniak/copilot-repo-template/blob/HEAD/"
                "templates/json/example.json).\n\n"
                "<!-- template-sync: begin json-reference-only -->\n"
                "JSON is mentioned as reference-only documentation.\n"
                "<!-- template-sync: end json-reference-only -->\n"
            ),
        )
    else:
        _write_text(repo_root, "README.md", "# Downstream\n")


def _section_entries(output: str, heading: str) -> set[str]:
    """Return bullet entries rendered under a named output section."""
    entries: set[str] = set()
    in_section = False
    for line in output.splitlines():
        if line and not line.startswith(" ") and line.endswith(":"):
            in_section = line == f"{heading}:"
            continue
        if in_section and line.startswith("  - "):
            entries.add(line.removeprefix("  - ").strip())
    return entries


def _finding_lines(output: str) -> set[str]:
    """Return rendered finding lines without their bullet marker."""
    return _section_entries(output, "Findings")


def test_marker_excluding_optional_modules_reports_cleanup_scope(tmp_path: Path) -> None:
    """A marker-derived report explains retained references to excluded modules."""
    _write_common_repo(tmp_path)

    result = _run_report(tmp_path)

    assert result.returncode == 0, result.stderr
    assert "State source: marker (.template-sync/marker.yml)" in result.stdout

    retained_modules = _section_entries(result.stdout, "Retained modules")
    excluded_modules = _section_entries(result.stdout, "Excluded modules")
    findings = _finding_lines(result.stdout)

    assert "template-sync-support" in retained_modules
    assert "python" in excluded_modules
    assert "terraform" in excluded_modules
    assert "json" in excluded_modules
    assert "schema" in excluded_modules
    assert "Excluded module scopes:" in result.stdout
    assert "python | manifest-owned paths:" in result.stdout
    assert "workflow files/jobs: .github/workflows/python-ci.yml" in result.stdout

    # Inline-block markers live on their own lines, so the scope scan must match
    # per line; a whole-file regex match would miss every retained cross-module
    # reference and inline-block family below.
    excluded_scopes = _section_entries(result.stdout, "Excluded module scopes")
    python_scope = next(
        (line for line in excluded_scopes if line.startswith("python | ")),
        "",
    )
    assert (
        "retained cross-module paths: .pre-commit-config.yaml; AGENTS.md" in python_scope
    ), python_scope
    assert "inline block families: python-only; python-reference-only" in python_scope, python_scope

    assert any(
        line.startswith(
            "manifest-owned-path | required_cleanup | json | templates/json/example.json |"
        )
        for line in findings
    )
    assert any(
        line.startswith(
            "inline-block.stale | required_cleanup | python | .pre-commit-config.yaml:6 |"
        )
        for line in findings
    )
    assert any(
        "pre-commit-hook-group.stale | required_cleanup | python | "
        ".pre-commit-config.yaml:6 | Hooks: black, ruff-check." in line
        for line in findings
    )
    assert any(
        "workflow-validation-command.stale | required_cleanup | schema | "
        ".github/workflows/data-ci.yml:10 | "
        "Command: pre-commit run validate-example-config-schema --all-files." in line
        for line in findings
    )
    assert any(
        line.startswith(
            "dependabot-ecosystem.stale | required_cleanup | python | "
            ".github/dependabot.yml | pip scans"
        )
        for line in findings
    )
    assert any(
        line.startswith("markdown-link.excluded-target | required_cleanup | json | README.md:3 |")
        for line in findings
    )
    assert any(
        line.startswith(
            "markdown-link.excluded-target | required_cleanup | schema | "
            ".github/ISSUE_TEMPLATE/bug_report.yml:7 |"
        )
        for line in findings
    )
    assert any(
        line.startswith(
            "contact-link.excluded-target | required_cleanup | schema | "
            ".github/ISSUE_TEMPLATE/config.yml:4 |"
        )
        for line in findings
    )
    assert any(
        line.startswith(
            "collaboration-template.prose-reference | required_cleanup | "
            "python | .github/pull_request_template.md:"
        )
        and "`pytest` documents excluded python tooling." in line
        for line in findings
    )
    assert any(
        line.startswith(
            "collaboration-template.prose-reference | required_cleanup | "
            "schema | .github/pull_request_template.md:"
        )
        and "check-jsonschema documents excluded schema tooling." in line
        for line in findings
    )
    assert any(
        line.startswith(
            "markdown-link.upstream-reference | likely_false_positive_documented_reference | "
            "json | README.md:5 |"
        )
        for line in findings
    )
    assert any(
        line.startswith(
            "inline-block.reference-only | optional_reference_only_cleanup | "
            "json | README.md:7 |"
        )
        for line in findings
    )
    assert any(
        line.startswith(
            "inline-block.reference-only | protected_file_authorization_needed | "
            "python | AGENTS.md:3 |"
        )
        for line in findings
    )
    assert any(
        line.startswith(
            "manifest.reference-only-note | documented_reference_only_retention | "
            "python | AGENTS.md:3 |"
        )
        for line in findings
    )
    assert any(
        line.startswith(
            "protected-document.prose-reference | protected_file_authorization_needed | "
            "terraform | AGENTS.md:6 | terraform test documents excluded terraform tooling."
        )
        for line in findings
    )
    assert any(
        line.startswith(
            "protected-document.prose-reference | protected_file_authorization_needed | "
            "python | AGENTS.md:7 | python -m pyright documents excluded python tooling."
        )
        for line in findings
    )
    assert any(
        line.startswith(
            "protected-document.prose-reference | protected_file_authorization_needed | "
            "python | AGENTS.md:8 | pytest tests/ -m documents excluded python tooling."
        )
        for line in findings
    )
    assert any(
        "validation-command.documented-reference | "
        "likely_false_positive_documented_reference | json | "
        ".pre-commit-config.yaml | check-json remains" in line
        for line in findings
    )


def test_explicit_included_modules_without_marker_matches_marker_findings(
    tmp_path: Path,
) -> None:
    """Pre-marker explicit module input produces the same deterministic findings."""
    marker_repo = tmp_path / "marker"
    explicit_repo = tmp_path / "explicit"
    marker_repo.mkdir()
    explicit_repo.mkdir()
    _write_common_repo(marker_repo)
    _write_common_repo(explicit_repo, include_marker=False)

    marker_result = _run_report(marker_repo)
    explicit_result = _run_report(
        explicit_repo,
        *(argument for module in INCLUDED_MODULES for argument in ("--included-module", module)),
    )

    assert marker_result.returncode == 0, marker_result.stderr
    assert explicit_result.returncode == 0, explicit_result.stderr
    assert "State source: explicit --included-module input" in explicit_result.stdout
    assert _finding_lines(explicit_result.stdout) == _finding_lines(marker_result.stdout)


def test_pre_adoption_catalog_is_protected_data_not_stale_prose(tmp_path: Path) -> None:
    """A retained catalog keeps excluded obligations while actual prose is reported."""
    root = tmp_path / "fixture"
    root.mkdir()
    _write_common_repo(root, include_marker=False)
    catalog_path = ".template-sync/instruction-contracts.yml"
    prose_paths = [".cursor/rules/policy.mdc", ".github/instructions/nested/policy"]
    manifest = _manifest()
    manifest["template_manifest"]["path_mappings"].extend(
        [
            {"pattern": catalog_path, "requires_all": ["template-sync-support"]},
            *({"pattern": path, "requires_all": ["baseline"]} for path in prose_paths),
        ]
    )
    _write_yaml(root, ".template-sync/manifest.yml", manifest)
    _write_yaml(
        root,
        catalog_path,
        {
            "instruction_contracts": [
                {
                    "path": "AGENTS.md",
                    "requires_modules": ["agent-instructions"],
                    "required_phrases": ["Preserve review authority."],
                }
            ]
        },
    )
    for path in prose_paths:
        _write_text(root, path, "Use AGENTS.md for agent instructions.\n")
    _run_git(root, "add", ".")
    assert (root / "AGENTS.md").is_file()
    assert not (root / ".template-sync/marker.yml").exists()
    args = ["--included-module", "baseline", "--included-module", "template-sync-support"]
    result = _run_report(root, *args)
    assert result.returncode == 0, result.stdout + result.stderr
    assert "State source: explicit --included-module input" in result.stdout
    findings = _finding_lines(result.stdout)
    prose_findings = [
        line for line in findings if line.startswith("protected-document.prose-reference |")
    ]
    assert not any(f"| {catalog_path}:" in line for line in prose_findings)
    for path in prose_paths:
        assert any(
            f"| {path}:1 | AGENTS.md references excluded path AGENTS.md." in line
            for line in prose_findings
        ), prose_findings
    assert any(
        line.startswith(
            "manifest-owned-path | protected_file_authorization_needed | "
            "agent-instructions | AGENTS.md |"
        )
        for line in findings
    ), findings

    mutant_root = tmp_path / "mutant"
    mutant_dir = mutant_root / ".template-sync/scripts"
    # Preserve the deployed dependency layout so the mutant executes the real
    # shared primitives before exercising the independent catalog/prose oracle.
    for relative in (".template-sync/scripts", ".github/scripts"):
        destination = mutant_root / relative
        destination.mkdir(parents=True)
        for source in (REPO_ROOT / relative).glob("*.py"):
            shutil.copyfile(source, destination / source.name)
    mutant = mutant_dir / SCRIPT_PATH.name
    source_text = mutant.read_text(encoding="utf-8")
    original = "if not is_protected_prose_path(relative_path):"
    assert source_text.count(original) == 1
    mutant.write_text(
        source_text.replace(original, "if not is_protected_instruction_path(relative_path):"),
        encoding="utf-8",
    )
    mutated = subprocess.run(
        [sys.executable, str(mutant), "--repo-root", str(root), *args],
        check=False,
        capture_output=True,
        text=True,
    )
    assert mutated.returncode == 0, mutated.stdout + mutated.stderr
    assert any(
        line.startswith("protected-document.prose-reference |") and f"| {catalog_path}:" in line
        for line in _finding_lines(mutated.stdout)
    ), "Restoring the broad prose predicate must fail the no-catalog-prose oracle."


def test_ambiguous_marker_and_explicit_included_modules_fail(tmp_path: Path) -> None:
    """Marker-derived and explicit module state cannot both be supplied."""
    _write_common_repo(tmp_path)

    result = _run_report(tmp_path, "--included-module", "baseline")

    assert result.returncode == 1
    assert "Ambiguous included-module state" in result.stderr


def test_local_overrides_and_deferred_decisions_are_report_categories(
    tmp_path: Path,
) -> None:
    """Marker waivers and deferred protected decisions are reported separately."""
    _write_common_repo(
        tmp_path,
        marker=_marker(
            local_overrides=[
                {
                    "path": "templates/json/",
                    "reason": "Downstream keeps JSON examples as local documentation.",
                    "default_decision": "SKIP",
                }
            ],
            deferred_candidates=[
                {
                    "path": "AGENTS.md",
                    "source_commit": FULL_SHA,
                    "reason": "Owner review pending for protected reference-only guidance.",
                }
            ],
        ),
        include_reference_content=False,
    )

    result = _run_report(tmp_path)

    assert result.returncode == 0, result.stderr
    findings = _finding_lines(result.stdout)
    assert any(
        line.startswith(
            "marker.local-override | local_override_recorded | json | "
            "templates/json/ | SKIP; reason:"
        )
        for line in findings
    )
    assert any(
        line.startswith(
            "marker.deferred-protected-decision | deferred_decision_recorded | "
            "unknown | AGENTS.md |"
        )
        for line in findings
    )
    assert not any(
        line.startswith(
            "manifest-owned-path | required_cleanup | json | templates/json/example.json |"
        )
        for line in findings
    )


def test_invalid_explicit_module_is_runtime_failure(tmp_path: Path) -> None:
    """Invalid input exits nonzero because no deterministic module source exists."""
    _write_common_repo(tmp_path, include_marker=False)

    result = _run_report(tmp_path, "--included-module", "not-a-module")

    assert result.returncode == 1
    assert "not defined by the manifest" in result.stderr


def test_reference_reporter_preserves_physical_lines_and_literal_markers() -> None:
    """Normal and malformed-marker fallback scans keep non-CR/LF characters literal."""
    program = r"""
import sys
sys.path.insert(0, sys.argv[1])
import report_excluded_module_references as reporter
for literal in ("\v", "\f", "\x1c", "\x1d", "\x1e", "\x85", "\u2028", "\u2029", "\u00a0", "\u2003", "\x1f"):
    begin = literal + "<!-- template-sync: begin python-reference-only -->"
    end = literal + "<!-- template-sync: end python-reference-only -->"
    text = begin + "\n[code](src/example.py)\n" + end + "\n"
    assert reporter.lines_outside_inline_blocks(text, relative_path="README.md") == (
        (1, begin), (2, "[code](src/example.py)"), (3, end))
    malformed = "<!-- template-sync: begin unknown-fixture-only -->"
    text = malformed + "\nleft" + literal + "right\n"
    assert reporter.lines_outside_inline_blocks(text, relative_path="README.md") == (
        (1, malformed), (2, "left" + literal + "right"))
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_inline_block_invalid_finding_reports_location_once(tmp_path: Path) -> None:
    """A structural inline-block error reports its location once in the finding."""
    _write_common_repo(tmp_path)
    _write_text(
        tmp_path,
        ".pre-commit-config.yaml",
        "repos: []\n# template-sync: begin python-only\n",
    )

    result = _run_report(tmp_path)

    findings = _finding_lines(result.stdout)
    invalid = [line for line in findings if line.startswith("inline-block.invalid")]
    assert invalid, result.stdout
    assert invalid[0].count(".pre-commit-config.yaml") == 1, invalid[0]
    assert "Unclosed template-sync inline marker: python-only" in invalid[0]


def test_dependabot_local_override_keeps_ecosystem_from_stale(tmp_path: Path) -> None:
    """A locally overridden scanned file keeps its Dependabot ecosystem non-stale."""
    _write_common_repo(
        tmp_path,
        marker=_marker(
            local_overrides=[
                {
                    "path": "pyproject.toml",
                    "reason": "Downstream keeps Python packaging metadata as a local file.",
                    "default_decision": "SKIP",
                }
            ],
        ),
    )
    _write_text(tmp_path, "pyproject.toml", '[project]\nname = "downstream"\n')

    result = _run_report(tmp_path)

    assert result.returncode == 0, result.stderr
    findings = _finding_lines(result.stdout)
    # The pip ecosystem maps to the excluded ``python`` module, but the retained
    # pyproject.toml is an intentional local override, so the ecosystem entry is
    # still justified and must not be reported as stale.
    assert not any(
        line.startswith("dependabot-ecosystem.stale | required_cleanup | python |")
        for line in findings
    )


def test_dependabot_github_actions_directory_surface_is_detected(tmp_path: Path) -> None:
    """A retained workflow keeps the github-actions ecosystem from being stale."""
    marker = {
        "template_sync": {
            "source_repo": SOURCE_REPO,
            "last_reviewed_template_commit": FULL_SHA,
            "included_modules": [
                "baseline",
                "agent-instructions",
                "github-platform",
                "github-templates",
                "markdown",
                "template-sync-support",
                "yaml",
            ],
            "local_overrides": [
                {
                    "path": ".github/workflows/data-ci.yml",
                    "reason": "Downstream keeps the data-file CI workflow.",
                    "default_decision": "SKIP",
                }
            ],
        }
    }
    _write_common_repo(tmp_path, marker=marker)

    result = _run_report(tmp_path)

    assert result.returncode == 0, result.stderr
    findings = _finding_lines(result.stdout)
    # github-actions is excluded, but a workflow under the ".github/workflows/"
    # directory surface is retained (here via override), so the ecosystem entry
    # is still justified. The scanned surface must be matched as a directory
    # prefix, not as a literal file path that is never present.
    assert not any(
        line.startswith("dependabot-ecosystem.stale | required_cleanup | github-actions |")
        for line in findings
    )


def test_reporter_multiline_markdown_links_have_original_lines(tmp_path: Path) -> None:
    """The native reporter recognizes multiline links and keeps fence gaps inert."""
    _write_common_repo(tmp_path, include_reference_content=False)
    _write_text(
        tmp_path,
        "README.md",
        '# Downstream\n\n[JSON](templates/json/example.json\n "JSON\nexample")\n'
        '\n[Schema]:\n schemas/example-config.schema.json\n "Schema"\n'
        '\n[Not live](templates/json/example.json\n```text\nexample\n```\n "title")\n',
    )
    result = _run_report(tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    links = [
        line
        for line in _finding_lines(result.stdout)
        if line.startswith("markdown-link.") and " | README.md:" in line
    ]
    assert len(links) == 2, links
    assert any("README.md:3 |" in line and "templates/json/example.json" in line for line in links)
    assert any(
        "README.md:7 |" in line and "schemas/example-config.schema.json" in line for line in links
    )


def test_reporter_code_span_labels_keep_live_links_and_literal_examples(tmp_path: Path) -> None:
    """The native reporter shares code context, raw destinations and opening lines."""
    _write_common_repo(tmp_path, include_reference_content=False)
    _write_text(
        tmp_path,
        "README.md",
        "# Downstream\n\n"
        "[JSON `]`](templates/json/example.json)\n"
        "[Schema\n `[`](schemas/example-config.schema.json)\n"
        "`[Literal](templates/json/example.json)`\n"
        "![Image `]`](templates/json/example.json)\n"
        "[First](literal`target.md) [JSON](templates/json/example.json) and `tail\n"
        '[First](other.md "`title") [Schema](schemas/example-config.schema.json) and `tail\n',
    )
    result = _run_report(tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    links = [
        line
        for line in _finding_lines(result.stdout)
        if line.startswith("markdown-link.") and " | README.md:" in line
    ]
    assert len(links) == 4, links
    for number, module, target in (
        (3, "json", "templates/json/example.json"),
        (4, "schema", "schemas/example-config.schema.json"),
        (8, "json", "templates/json/example.json"),
        (9, "schema", "schemas/example-config.schema.json"),
    ):
        assert any(
            f"markdown-link.excluded-target | required_cleanup | {module} | README.md:{number} |"
            in line
            and target in line
            for line in links
        ), links


@pytest.mark.parametrize("multiline", [False, True], ids=["single-line", "multiline"])
def test_yaml_embedded_fenced_links_are_skipped(tmp_path: Path, multiline: bool) -> None:
    """Links inside fenced code blocks in YAML-embedded Markdown are ignored."""
    _write_common_repo(tmp_path)
    schema_link = "[schema](../../schemas/example-config.schema.json)."
    json_link = "[json](../../templates/json/example.json)."
    if multiline:
        schema_link = (
            '[schema](../../schemas/example-config.schema.json\n          "Schema title").'
        )
        json_link = '[json](../../templates/json/example.json\n          "JSON title").'
    _write_text(
        tmp_path,
        ".github/ISSUE_TEMPLATE/fenced_example.yml",
        (
            "name: Fenced Example\n"
            "description: Example with fenced Markdown\n"
            "body:\n"
            "  - type: markdown\n"
            "    attributes:\n"
            "      value: |\n"
            f"        Outside fence: {schema_link}\n"
            "        ```\n"
            f"        Inside fence: {json_link}\n"
            "        ```\n"
        ),
    )

    result = _run_report(tmp_path)

    assert result.returncode == 0, result.stderr
    findings = _finding_lines(result.stdout)
    # The link outside the (indented) fenced block is still reported.
    assert any(
        line.startswith(
            "markdown-link.excluded-target | required_cleanup | schema | "
            ".github/ISSUE_TEMPLATE/fenced_example.yml:7 |"
        )
        for line in findings
    )
    # The link inside the fenced block must be ignored even though the file is
    # YAML and the fence is indented inside a ``value: |`` block.
    hidden_line = 10 if multiline else 9
    assert not any(
        f".github/ISSUE_TEMPLATE/fenced_example.yml:{hidden_line}" in line for line in findings
    )


def test_markdown_blockquote_fenced_links_use_standard_fence_context(tmp_path: Path) -> None:
    """Links inside blockquote-contained Markdown fences are ignored for .md files."""
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "report_excluded_module_references_blockquote_fence", SCRIPT_PATH
    )
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)

    content = "".join(
        [
            "> ```\n",
            "> [inside](templates/json/example.json)\n",
            "> ```\n",
            "[outside](schemas/example-config.schema.json)\n",
        ]
    )
    md_path = tmp_path / "doc.md"
    md_path.write_text(content, encoding="utf-8")

    targets = [target for _, target in module.link_targets_outside_fences(md_path)]

    # The blockquote-fenced link is recognized as fenced (the standard Markdown
    # context handles the ``>`` container), so only the unfenced link is live.
    assert "schemas/example-config.schema.json" in targets
    assert "templates/json/example.json" not in targets


def test_fenced_registered_marker_examples_are_not_live_blocks(tmp_path: Path) -> None:
    """Markdown fenced marker examples are not reported as retained live markers."""
    _write_common_repo(tmp_path, include_reference_content=False)
    _write_text(
        tmp_path,
        "README.md",
        (
            "# Downstream\n\n"
            "```markdown\n"
            "<!-- template-sync: begin python-reference-only -->\n"
            "Example marker only.\n"
            "<!-- template-sync: end python-reference-only -->\n"
            "```\n"
        ),
    )

    result = _run_report(tmp_path)

    assert result.returncode == 0, result.stderr
    findings = _finding_lines(result.stdout)
    assert not any("README.md" in line and "inline-block" in line for line in findings)

    excluded_scopes = _section_entries(result.stdout, "Excluded module scopes")
    python_scope = next(
        (line for line in excluded_scopes if line.startswith("python | ")),
        "",
    )
    assert "README.md" not in python_scope
    assert "python-reference-only" in python_scope


def test_active_contact_link_urls_preserves_fragment_and_quoted_urls() -> None:
    """Contact-link URL extraction keeps ``#`` fragments and quoted scalars."""
    import importlib.util

    # Register under a test-scoped module name rather than the importable
    # ``report_excluded_module_references``. The sys.modules entry is required so
    # the reporter's dataclasses resolve during exec_module, and a unique name
    # keeps it from shadowing a real import of the reporter elsewhere in the
    # session (the script puts its own directory on sys.path).
    spec = importlib.util.spec_from_file_location(
        "report_excluded_module_references_under_test", SCRIPT_PATH
    )
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)

    text = (
        "contact_links:\n"
        "  - name: Unquoted Fragment\n"
        "    url: https://github.com/OWNER/REPO#support\n"
        "  - name: Quoted Fragment\n"
        '    url: "https://github.com/OWNER/REPO/blob/HEAD/file.md#L1-L5"\n'
        "  - name: Trailing Comment\n"
        "    url: https://example.com/docs # see the docs\n"
    )

    extracted = [
        url
        for _, url in module.active_contact_link_urls(
            text,
            relative_path=".github/ISSUE_TEMPLATE/config.yml",
        )
    ]

    # A ``#`` fragment (quoted or unquoted) must be retained, while a genuine
    # trailing YAML comment (``#`` after whitespace) is still stripped.
    assert extracted == [
        "https://github.com/OWNER/REPO#support",
        "https://github.com/OWNER/REPO/blob/HEAD/file.md#L1-L5",
        "https://example.com/docs",
    ]


@pytest.mark.parametrize(
    ("suffix", "body", "line_number"),
    [
        ("md", "    [JSON](templates/json/example.json)\n", None),
        ("mdc", "    [JSON](templates/json/example.json)\n", None),
        ("md", "Paragraph\n    [JSON](templates/json/example.json)\n", 2),
        ("md", "- Item\n\n      [JSON](templates/json/example.json)\n", None),
        ("md", "- Item\n\n    [JSON](templates/json/example.json)\n", 3),
        ("md", ">     [JSON](templates/json/example.json)\n", None),
        ("yml", "value: |\n    [JSON](templates/json/example.json)\n", 2),
        ("yaml", "value: |\n    [JSON](templates/json/example.json)\n", 2),
    ],
)
def test_reporter_preserves_contextual_code_and_embedded_indentation(
    tmp_path: Path, suffix: str, body: str, line_number: int | None
) -> None:
    """Markdown code is inert while structural YAML indentation remains live."""
    _write_common_repo(tmp_path, include_reference_content=False)
    relative_path = f"example.{suffix}"
    manifest = _manifest()
    manifest["template_manifest"]["path_mappings"].append(
        {"pattern": relative_path, "requires_all": ["baseline"]}
    )
    _write_yaml(tmp_path, ".template-sync/manifest.yml", manifest)
    _write_text(tmp_path, relative_path, body)
    result = _run_report(tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    links = [
        line
        for line in _finding_lines(result.stdout)
        if line.startswith("markdown-link.") and f" | {relative_path}:" in line
    ]
    if line_number is None:
        assert links == []
    else:
        assert len(links) == 1, links
        assert f"{relative_path}:{line_number} |" in links[0]
        assert "templates/json/example.json" in links[0]


@pytest.mark.parametrize(
    ("suffix", "body", "line_number"),
    [
        ("md", "<!-- [JSON](templates/json/example.json) -->\n", None),
        ("mdc", "<pre>\n[JSON](templates/json/example.json)\n</pre>\n", None),
        ("md", "Text <!-- [JSON](templates/json/example.json) -->\n", None),
        ("md", 'Text <span title="[JSON](templates/json/example.json)"> text</span>\n', None),
        ("md", "<span>[JSON](templates/json/example.json)</span>\n", 1),
        ("md", "[A <!-- ] --> B](templates/json/example.json)\n", 1),
        ("md", "<!--\n```\n-->\n[JSON](templates/json/example.json)\n", 4),
        ("yml", "value: |\n    [JSON](templates/json/example.json)\n", 2),
        ("yaml", "value: |\n    <!-- [JSON](templates/json/example.json) -->\n", None),
        ("yaml", "value: |\n    <span>[JSON](templates/json/example.json)</span>\n", 2),
    ],
)
def test_reporter_html_context_preserves_fixed_targets_and_embedded_behavior(
    tmp_path: Path, suffix: str, body: str, line_number: int | None
) -> None:
    """Native reports preserve Markdown literals, inline text and YAML indentation."""
    _write_common_repo(tmp_path, include_reference_content=False)
    relative_path = f"example.{suffix}"
    manifest = _manifest()
    manifest["template_manifest"]["path_mappings"].append(
        {"pattern": relative_path, "requires_all": ["baseline"]}
    )
    _write_yaml(tmp_path, ".template-sync/manifest.yml", manifest)
    _write_text(tmp_path, relative_path, body)
    result = _run_report(tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    links = [
        line
        for line in _finding_lines(result.stdout)
        if line.startswith("markdown-link.") and f" | {relative_path}:" in line
    ]
    if line_number is None:
        assert links == []
    else:
        assert len(links) == 1, links
        assert f"{relative_path}:{line_number} |" in links[0]
        assert "templates/json/example.json" in links[0]


@pytest.mark.parametrize(
    ("suffix", "body", "line_number"),
    [
        ("md", "[outer [JSON](templates/json/example.json)](kept.md)", 1),
        ("md", "[outer [kept](kept.md)](templates/json/example.json)", None),
        ("mdc", "![outer [JSON](templates/json/example.json)](image.png)", None),
        ("md", "[outer ![image](image.png)](templates/json/example.json)", 1),
        ("md", "[outer\n [JSON](templates/json/example.json)](kept.md)", 2),
        ("md", "[outer [Guide][ref]](templates/json/example.json)\n\n[ref]: kept.md", None),
        ("md", "[outer [Guide][missing]](templates/json/example.json)", 1),
        ("yaml", "value: |\n    [outer [JSON](templates/json/example.json)](kept.md)", 2),
    ],
)
def test_reporter_nested_link_activity_has_fixed_targets_and_lines(
    tmp_path: Path, suffix: str, body: str, line_number: int | None
) -> None:
    """The cleanup report follows live nested links without reporting literal parents."""
    _write_common_repo(tmp_path, include_reference_content=False)
    relative_path = f"example.{suffix}"
    manifest = _manifest()
    manifest["template_manifest"]["path_mappings"].append(
        {"pattern": relative_path, "requires_all": ["baseline"]}
    )
    _write_yaml(tmp_path, ".template-sync/manifest.yml", manifest)
    _write_text(tmp_path, relative_path, body + "\n")
    result = _run_report(tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    links = [
        line
        for line in _finding_lines(result.stdout)
        if line.startswith("markdown-link.excluded-target") and f" | {relative_path}:" in line
    ]
    if line_number is None:
        assert links == []
    else:
        assert len(links) == 1, links
        assert f"{relative_path}:{line_number} |" in links[0]
        assert "templates/json/example.json" in links[0]


@pytest.mark.parametrize(
    ("suffix", "body", "line_number"),
    [
        ("md", "[Guide]\n[Guide]: templates/json/example.json", None),
        ("mdc", "> [Guide]\n> [Guide]: templates/json/example.json", None),
        ("md", "- [Guide]\n  [Guide]: templates/json/example.json", None),
        ("md", "[Guide]\n\n[Guide]: templates/json/example.json", 3),
        ("md", 'Paragraph\n[Guide]: kept.md "[live](templates/json/example.json)"', 2),
        ("md", "# Heading\n[Guide]: templates/json/example.json", 2),
        ("md", "[kept]: kept.md\n[Guide]: templates/json/example.json", 2),
        ("yaml", "value: |\n  [Guide]: templates/json/example.json", 2),
        ("yml", "value: |\n  [Guide]\n  [Guide]: templates/json/example.json", 3),
        ("yaml", "value: |\n    [Guide]: templates/json/example.json", None),
        ("yml", "value: |\n    [Guide]\n    [Guide]: templates/json/example.json", None),
    ],
)
def test_reporter_definition_context_preserves_native_findings_and_embedded_yaml(
    tmp_path: Path, suffix: str, body: str, line_number: int | None
) -> None:
    """Reports apply Markdown block context while retaining explicit YAML scanning."""
    _write_common_repo(tmp_path, include_reference_content=False)
    relative_path = f"example.{suffix}"
    manifest = _manifest()
    manifest["template_manifest"]["path_mappings"].append(
        {"pattern": relative_path, "requires_all": ["baseline"]}
    )
    _write_yaml(tmp_path, ".template-sync/manifest.yml", manifest)
    _write_text(tmp_path, relative_path, body + "\n")
    result = _run_report(tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    links = [
        line
        for line in _finding_lines(result.stdout)
        if line.startswith("markdown-link.excluded-target") and f" | {relative_path}:" in line
    ]
    if line_number is None:
        assert links == []
    else:
        assert len(links) == 1, links
        assert f"{relative_path}:{line_number} |" in links[0]
        assert "templates/json/example.json" in links[0]
