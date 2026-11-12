"""Exercise the marker-aware sync candidate table generator."""

from __future__ import annotations

import importlib
import json
import shutil
import subprocess
import sys
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, cast

import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT_PATH = REPO_ROOT / ".template-sync" / "scripts" / "generate_sync_candidates.py"
SCRIPT_DIR = SCRIPT_PATH.parent
MARKER_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-marker.schema.json"
MANIFEST_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-manifest.schema.json"
MANIFEST_PATH = REPO_ROOT / ".template-sync" / "manifest.yml"
SOURCE_REPO = "https://github.com/franklesniak/copilot-repo-template.git"
MODULE_DEFINITIONS = {
    "baseline": "Baseline files.",
    "agent-instructions": "Agent instruction files.",
    "github-actions": "GitHub Actions workflows.",
    "template-sync-support": "Template sync support files.",
    "markdown": "Markdown files.",
    "python": "Python files.",
    "json": "JSON files.",
    "yaml": "YAML files.",
    "schema": "Schema files.",
}

if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

sync_candidates = cast(Any, importlib.import_module("generate_sync_candidates"))


def test_validation_commands_cover_every_manifest_module() -> None:
    """Retained-module summaries must have validation guidance for each module."""
    manifest = yaml.safe_load(MANIFEST_PATH.read_text(encoding="utf-8"))
    modules = {module["name"] for module in manifest["template_manifest"]["modules"]}

    assert modules <= set(sync_candidates.VALIDATION_COMMANDS_BY_MODULE)


def test_validation_commands_keep_dependabot_github_scoped_and_azure_manual() -> None:
    """Azure-only platform summaries must not require GitHub Dependabot validation."""
    github_commands = sync_candidates.VALIDATION_COMMANDS_BY_MODULE["github-platform"]
    azure_commands = sync_candidates.VALIDATION_COMMANDS_BY_MODULE["azure-devops-platform"]

    assert any("validate-dependabot-config" in command for command in github_commands)
    assert any("test_dependabot_schema.py" in command for command in github_commands)
    assert not any("dependabot" in command.lower() for command in azure_commands)
    assert any("dependency scanning" in command.lower() for command in azure_commands)
    assert any("dependency update policy" in command.lower() for command in azure_commands)


def test_validation_commands_keep_actionlint_github_actions_scoped() -> None:
    """Azure Pipelines guidance must not require the GitHub Actions linter."""
    github_actions_commands = sync_candidates.VALIDATION_COMMANDS_BY_MODULE["github-actions"]
    azure_pipeline_commands = sync_candidates.VALIDATION_COMMANDS_BY_MODULE["azure-pipelines"]

    assert any("actionlint" in command for command in github_actions_commands)
    assert "python .github/scripts/validate_workflow_security.py" in github_actions_commands
    assert not any("validate_workflow_security" in command for command in azure_pipeline_commands)
    assert not any("actionlint" in command for command in azure_pipeline_commands)
    assert any("Azure Pipelines" in command for command in azure_pipeline_commands)


def _run(command: list[str], cwd: Path) -> subprocess.CompletedProcess[str]:
    """Run a test command and capture text output."""
    return subprocess.run(
        command,
        cwd=cwd,
        check=False,
        capture_output=True,
        text=True,
    )


def _git(repo_root: Path, *args: str) -> str:
    """Run a Git command in a fixture repository and return stdout."""
    result = _run(["git", *args], repo_root)
    assert result.returncode == 0, result.stderr
    return result.stdout.strip()


def _write_text(repo_root: Path, relative_path: str, text: str = "placeholder\n") -> None:
    """Write a text file below a fixture repository root."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _write_yaml(repo_root: Path, relative_path: str, data: dict[str, Any]) -> None:
    """Write a YAML fixture below a fixture repository root."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data, sort_keys=False), encoding="utf-8")


def _symlink_or_skip(link_path: Path, target_path: Path) -> None:
    """Create a directory symlink, or skip when the platform refuses it."""
    try:
        link_path.symlink_to(target_path, target_is_directory=True)
    except OSError as error:
        pytest.skip(f"Symlink creation unavailable in this environment: {error}")


def _copy_schemas(repo_root: Path) -> None:
    """Copy the real marker and manifest schemas into a fixture repository."""
    schemas_dir = repo_root / "schemas"
    schemas_dir.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(MARKER_SCHEMA_PATH, schemas_dir / MARKER_SCHEMA_PATH.name)
    shutil.copyfile(MANIFEST_SCHEMA_PATH, schemas_dir / MANIFEST_SCHEMA_PATH.name)


def _manifest(
    *,
    version: int = 2,
    extra_mappings: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Build a small schema-valid manifest fixture."""
    path_mappings: list[dict[str, Any]] = [
        {"pattern": "README.md", "requires_all": ["baseline"]},
        {
            "pattern": ".github/copilot-instructions.md",
            "requires_all": ["agent-instructions"],
        },
        {
            "pattern": ".pre-commit-config.yaml",
            "requires_all": ["baseline"],
            "notes": "Contains a Python-only inline block for the python module.",
        },
        {"pattern": ".template-sync/scripts/**", "requires_all": ["template-sync-support"]},
        {"pattern": "templates/markdown/**", "requires_all": ["markdown"]},
        {"pattern": "templates/python/**", "requires_all": ["python"]},
    ]
    filtering: dict[str, str] = {
        "default_semantics": "AND",
        "path_matching": "most_specific_match_wins",
        "same_specificity_action": "union_modules",
        "unmapped_action": "surface_for_owner",
    }
    notes: dict[str, Any] = {
        "downstream_retention": "Downstream repositories keep marker data for syncs.",
    }
    if version == 1:
        path_mappings.append(
            {
                "pattern": ".github/workflows/data-ci.yml",
                "requires_all": ["github-actions", "yaml"],
            }
        )
        notes["known_limitations"] = [
            {
                "path": ".github/workflows/data-ci.yml",
                "description": "Version 1 cannot model alternative module relations.",
                "future_work": "Migrate to manifest version 2 relation semantics.",
            }
        ]
    elif version == 2:
        path_mappings.append(
            {
                "pattern": ".github/workflows/data-ci.yml",
                "requires_all": ["github-actions"],
                "requires_any": ["json", "yaml", "schema", "template-sync-support"],
            }
        )
        filtering["requires_any_semantics"] = "OR"
    else:
        raise ValueError(f"Unsupported manifest version: {version}")

    if extra_mappings is not None:
        path_mappings.extend(extra_mappings)

    return {
        "template_manifest": {
            "version": version,
            "modules": [
                {"name": name, "description": description}
                for name, description in MODULE_DEFINITIONS.items()
            ],
            "path_mappings": path_mappings,
            "filtering": filtering,
            "notes": notes,
        }
    }


def _marker(
    included_modules: list[str],
    *,
    last_reviewed_template_commit: str | None,
    local_overrides: list[dict[str, str]] | None = None,
    local_path_ownership: list[dict[str, str]] | None = None,
    deferred_candidates: list[dict[str, str]] | None = None,
    protected_decisions: list[dict[str, str]] | None = None,
) -> dict[str, Any]:
    """Build a small schema-valid marker fixture."""
    template_sync: dict[str, Any] = {
        "source_repo": SOURCE_REPO,
        "included_modules": included_modules,
    }
    if last_reviewed_template_commit is not None:
        template_sync["last_reviewed_template_commit"] = last_reviewed_template_commit
    if local_overrides is not None:
        template_sync["local_overrides"] = local_overrides
    if local_path_ownership is not None:
        template_sync["local_path_ownership"] = local_path_ownership
    if deferred_candidates is not None:
        template_sync["deferred_protected_candidates"] = deferred_candidates
    if protected_decisions is not None:
        template_sync["protected_file_decisions"] = protected_decisions
    return {"template_sync": template_sync}


def _init_repo(
    repo_root: Path,
    *,
    manifest_version: int = 2,
    extra_mappings: list[dict[str, Any]] | None = None,
) -> None:
    """Create a fixture Git repository with schemas and a manifest."""
    _git(repo_root, "init")
    _git(repo_root, "config", "user.name", "Template Sync Tests")
    _git(repo_root, "config", "user.email", "template-sync-tests@example.invalid")
    _copy_schemas(repo_root)
    _write_yaml(
        repo_root,
        ".template-sync/manifest.yml",
        _manifest(version=manifest_version, extra_mappings=extra_mappings),
    )


def _commit_all(repo_root: Path, message: str) -> str:
    """Commit all fixture repository changes and return the commit SHA."""
    _git(repo_root, "add", ".")
    _git(repo_root, "commit", "-m", message)
    return _git(repo_root, "rev-parse", "HEAD")


def _run_generator(repo_root: Path, *extra_args: str) -> subprocess.CompletedProcess[str]:
    """Run the candidate generator against a fixture repository."""
    return _run(
        [
            sys.executable,
            str(SCRIPT_PATH),
            "--repo-root",
            str(repo_root),
            *extra_args,
        ],
        repo_root,
    )


def _rest_result(
    status_code: int | None,
    body: Any = "",
    *,
    error: str | None = None,
    headers: tuple[tuple[str, str], ...] = (),
) -> Any:
    """Build a fake REST result for metadata tests."""
    response_body = body if isinstance(body, str) else json.dumps(body)
    return sync_candidates.RestResult(
        status_code=status_code,
        body=response_body,
        error=error,
        headers=headers,
    )


def _fixed_clock() -> datetime:
    """Return a deterministic timestamp for metadata ledger tests."""
    return datetime(2026, 6, 29, 12, 34, 56, tzinfo=UTC)


def _observation(metadata: Any, setting: str) -> Any:
    """Return one metadata observation from a test ledger."""
    observation = metadata.observation(setting)
    assert observation is not None
    return observation


def _metadata_rest_success_responses(
    repo_payload: dict[str, Any],
    labels_payload: list[dict[str, Any]] | dict[str, Any] | str | None = None,
    *,
    rulesets_payload: list[dict[str, Any]] | None = None,
    branch_protection_payload: dict[str, Any] | None = None,
    private_vulnerability_payload: dict[str, Any] | None = None,
) -> list[Any]:
    """Return the standard REST response sequence for full metadata discovery."""
    labels = [] if labels_payload is None else labels_payload
    rulesets = [] if rulesets_payload is None else rulesets_payload
    branch_protection = (
        {
            "required_status_checks": {"contexts": []},
            "enforce_admins": {"enabled": True},
        }
        if branch_protection_payload is None
        else branch_protection_payload
    )
    private_vulnerability = (
        {"enabled": False}
        if private_vulnerability_payload is None
        else private_vulnerability_payload
    )
    return [
        _rest_result(200, repo_payload),
        _rest_result(200, labels),
        _rest_result(200, rulesets),
        _rest_result(200, branch_protection),
        _rest_result(200, private_vulnerability),
    ]


class FakeRestClient:
    """Callable fake REST client that records deterministic request envelopes."""

    def __init__(self, responses: list[Any]) -> None:
        self.responses = responses
        self.requests: list[Any] = []

    def __call__(self, request: Any) -> Any:
        """Record one request and return the next queued response."""
        self.requests.append(request)
        if not self.responses:
            raise AssertionError(f"Unexpected REST request: {request}")
        return self.responses.pop(0)


def _summary_section_modules(output: str, heading: str) -> set[str]:
    """Return the backtick-quoted module names rendered under a `## <heading>` section."""
    modules: set[str] = set()
    in_section = False
    for line in output.splitlines():
        if line.startswith("## "):
            in_section = line == f"## {heading}"
            continue
        if in_section:
            stripped = line.strip()
            if stripped.startswith("- `") and stripped.endswith("`"):
                modules.add(stripped.removeprefix("- `").removesuffix("`"))
    return modules


def test_no_changes_reports_empty_range(tmp_path: Path) -> None:
    """A clean reviewed range produces a clear empty report."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=base_sha),
    )

    result = _run_generator(tmp_path, "--range-head", base_sha)

    assert result.returncode == 0, result.stderr
    assert "# Template Sync Candidate Table" in result.stdout
    assert "No changed paths found in the reviewed range." in result.stdout


def test_normal_mixed_change_set_reports_marker_and_manifest_decisions(tmp_path: Path) -> None:
    """Retained, excluded, cross-module, and inline-note rows appear together."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    _write_text(tmp_path, ".github/workflows/data-ci.yml", "name: data\n")
    _write_text(tmp_path, ".pre-commit-config.yaml", "repos: []\n")
    _write_text(tmp_path, "templates/python/app.py", "print('base')\n")
    base_sha = _commit_all(tmp_path, "base")

    _write_text(tmp_path, "README.md", "head\n")
    _write_text(tmp_path, ".github/workflows/data-ci.yml", "name: data updated\n")
    _write_text(tmp_path, ".pre-commit-config.yaml", "repos:\n  - repo: local\n")
    _write_text(tmp_path, "templates/python/app.py", "print('head')\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["baseline", "github-actions", "yaml", "template-sync-support"],
            last_reviewed_template_commit=base_sha,
        ),
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 0, result.stderr
    assert f"git diff --name-status -M {base_sha}..{head_sha} --" in result.stdout
    assert "| README.md | Modified | requires all: baseline | Retained |" in result.stdout
    assert (
        "requires all: github-actions; requires any: json, schema, template-sync-support, yaml"
        in result.stdout
    )
    assert "Cross-module manifest relation matched." in result.stdout
    assert (
        "| templates/python/app.py | Modified | requires all: python | Excluded |" in result.stdout
    )
    assert "Manifest note: Contains a Python-only inline block" in result.stdout


def test_manifest_v1_requires_all_rows_are_supported(tmp_path: Path) -> None:
    """Version 1 manifests retain requires_all-only relation semantics."""
    _init_repo(tmp_path, manifest_version=1)
    _write_text(tmp_path, ".github/workflows/data-ci.yml", "name: data\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, ".github/workflows/data-ci.yml", "name: data updated\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["github-actions", "yaml"], last_reviewed_template_commit=base_sha),
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 0, result.stderr
    assert (
        "| .github/workflows/data-ci.yml | Modified | "
        "requires all: github-actions, yaml | Retained |"
    ) in result.stdout


def test_explicit_range_base_allows_marker_without_last_reviewed(tmp_path: Path) -> None:
    """First-sync delta reviews can pass an explicit base when the marker lacks one."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, "README.md", "head\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=None),
    )

    result = _run_generator(tmp_path, "--range-base", base_sha, "--range-head", head_sha)

    assert result.returncode == 0, result.stderr
    assert f"- Range base: `{base_sha}` from `{base_sha}` (--range-base)" in result.stdout
    assert f"git diff --name-status -M {base_sha}..{head_sha} --" in result.stdout
    assert "| README.md | Modified | requires all: baseline | Retained |" in result.stdout


def test_matching_local_procedure_does_not_warn(tmp_path: Path) -> None:
    """The procedure warning is quiet when local text matches range-head upstream."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    _write_text(tmp_path, "TEMPLATE_UPDATE_PROCEDURE.md", "procedure v1\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, "README.md", "head\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=base_sha),
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 0, result.stderr
    assert "WARNING: Local `TEMPLATE_UPDATE_PROCEDURE.md` may be stale" not in result.stdout
    assert f"git show {head_sha}:TEMPLATE_UPDATE_PROCEDURE.md" not in result.stdout


def test_stale_local_procedure_warns_with_upstream_show_command(tmp_path: Path) -> None:
    """A stale local procedure produces a non-fatal warning with a copyable command."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    _write_text(tmp_path, "TEMPLATE_UPDATE_PROCEDURE.md", "procedure v1\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, "README.md", "head\n")
    _write_text(tmp_path, "TEMPLATE_UPDATE_PROCEDURE.md", "procedure v2\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_text(tmp_path, "TEMPLATE_UPDATE_PROCEDURE.md", "procedure v1\n")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=base_sha),
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 0, result.stderr
    assert "WARNING: Local `TEMPLATE_UPDATE_PROCEDURE.md` may be stale" in result.stdout
    assert f"git show {head_sha}:TEMPLATE_UPDATE_PROCEDURE.md" in result.stdout
    assert "| README.md | Modified | requires all: baseline | Retained |" in result.stdout


def test_missing_range_base_errors_when_marker_has_no_last_reviewed(tmp_path: Path) -> None:
    """The generator refuses to guess a baseline."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    _commit_all(tmp_path, "base")
    _write_text(tmp_path, "README.md", "head\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=None),
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 1
    assert "No range base was provided" in result.stderr
    assert "will not guess a baseline" in result.stderr


def test_default_range_head_uses_local_template_main_ref(tmp_path: Path) -> None:
    """When no head is supplied, the local template/main ref is used without fetching."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, "README.md", "head\n")
    head_sha = _commit_all(tmp_path, "head")
    _git(tmp_path, "update-ref", "refs/remotes/template/main", head_sha)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=base_sha),
    )

    result = _run_generator(tmp_path)

    assert result.returncode == 0, result.stderr
    assert f"- Range head: `{head_sha}` from `template/main`" in result.stdout


def test_write_candidates_writes_table_and_preserves_stdout(tmp_path: Path) -> None:
    """The optional output path receives the table while stdout keeps the full report."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, "README.md", "head\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=base_sha),
    )

    result = _run_generator(
        tmp_path,
        "--range-head",
        head_sha,
        "--write-candidates",
        ".cache/template-sync/candidates.md",
    )

    assert result.returncode == 0, result.stderr
    assert "# Template Sync Candidate Table" in result.stdout
    assert "- Saved candidate table: `.cache/template-sync/candidates.md`" in result.stdout
    assert "| README.md | Modified | requires all: baseline | Retained |" in result.stdout

    written_table = (tmp_path / ".cache/template-sync/candidates.md").read_text(encoding="utf-8")
    assert written_table.startswith("| Path | Change | Matched module relation |")
    assert "| README.md | Modified | requires all: baseline | Retained |" in written_table
    assert "# Template Sync Candidate Table" not in written_table


def test_write_candidates_rejects_paths_outside_repository(tmp_path: Path) -> None:
    """The optional output path uses the existing repo-root traversal guard."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, "README.md", "head\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=base_sha),
    )

    result = _run_generator(
        tmp_path,
        "--range-head",
        head_sha,
        "--write-candidates",
        "../candidates.md",
    )

    assert result.returncode == 1
    assert "Path escapes the repository root: ../candidates.md" in result.stderr


def test_ledger_only_reports_manifest_marker_and_todo_decisions(tmp_path: Path) -> None:
    """The adoption ledger can be generated without a git comparison range."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["baseline"],
            last_reviewed_template_commit=None,
            local_overrides=[
                {
                    "path": "README.md",
                    "reason": "Project-specific README.",
                    "default_decision": "SKIP",
                }
            ],
        ),
    )
    _write_text(
        tmp_path,
        "_TODO-repo-init.md",
        (
            "- [ ] Confirm private vulnerability reporting.\n"
            "- [x] Record adoption mode for protected files.\n"
        ),
    )

    result = _run_generator(
        tmp_path,
        "--ledger-only",
        "--write-ledger",
        ".cache/template-sync/adoption-ledger.md",
    )

    assert result.returncode == 0, result.stderr
    assert "Saved adoption ledger: `.cache/template-sync/adoption-ledger.md`" in result.stdout
    assert "# Template Adoption Ledger" in result.stdout
    assert "Generated snapshot; review artifact only." in result.stdout
    assert (
        "| README.md | all: baseline | local override | "
        "Marker local override defaults to `SKIP`: Project-specific README."
    ) in result.stdout
    assert "| .pre-commit-config.yaml | all: baseline | trim |" in result.stdout
    assert "module-scoped inline blocks" in result.stdout
    assert (
        "| .github/copilot-instructions.md | all: agent-instructions | "
        "needs maintainer decision | Protected path is not retained by included modules"
    ) in result.stdout
    assert (
        "| templates/python/** | all: python | skip | Required module(s) not included: python."
        in (result.stdout)
    )
    assert "| _TODO-repo-init.md | manual setup | manual TODO | open:" in result.stdout
    assert "[line 1](../../_TODO-repo-init.md#L1)" in result.stdout
    assert "pre-commit run --all-files" in result.stdout
    assert "manual first-adoption review" in result.stdout
    assert "No range base was provided" not in result.stderr

    written_ledger = (tmp_path / ".cache/template-sync/adoption-ledger.md").read_text(
        encoding="utf-8"
    )
    assert written_ledger.startswith("# Template Adoption Ledger")
    assert (
        ".template-sync/manifest.yml` and `.template-sync/marker.yml` remain the authoritative"
        in written_ledger
    )
    assert "| `_TODO-repo-init.md` link |" in written_ledger


def test_ledger_and_summary_report_local_path_ownership(tmp_path: Path) -> None:
    """Local path ownership records render separately from local overrides."""
    _init_repo(
        tmp_path,
        extra_mappings=[
            {"pattern": "schemas/**", "requires_all": ["schema"]},
        ],
    )
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["baseline"],
            last_reviewed_template_commit=None,
            local_path_ownership=[
                {
                    "path": "docs/",
                    "reason": "Project documentation is downstream-owned.",
                },
                {
                    "path": "schemas/local-project.schema.json",
                    "reason": "Project schema is downstream-owned.",
                    "overlap_exception_reason": "It lives near template schema examples.",
                },
            ],
        ),
    )

    ledger_result = _run_generator(tmp_path, "--ledger-only")
    summary_result = _run_generator(tmp_path, "--summary")

    assert ledger_result.returncode == 0, ledger_result.stderr
    assert (
        "| docs/ | local | local ownership | Marker local path ownership: "
        "Project documentation is downstream-owned."
    ) in ledger_result.stdout
    assert (
        "| schemas/local-project.schema.json | all: schema | local ownership | "
        "Marker local path ownership: Project schema is downstream-owned. "
        "Broad manifest overlap exception: It lives near template schema examples."
    ) in ledger_result.stdout
    assert "Marker local override defaults" not in ledger_result.stdout

    assert summary_result.returncode == 0, summary_result.stderr
    assert "## Local Path Ownership" in summary_result.stdout
    assert (
        "| docs/ | Project documentation is downstream-owned. | not recorded |"
        in summary_result.stdout
    )
    assert (
        "| schemas/local-project.schema.json | Project schema is downstream-owned. | "
        "It lives near template schema examples. |"
    ) in summary_result.stdout


def test_ledger_directory_ownership_reports_manifest_proximity(tmp_path: Path) -> None:
    """A directory ownership record over a broad manifest pattern reports its modules.

    ``selected_relation_for_path`` cannot match the directory path itself against a
    glob such as ``schemas/**``, so the ledger must derive the proximate modules
    from manifest mappings under the directory prefix instead of rendering ``local``.
    """
    _init_repo(
        tmp_path,
        extra_mappings=[
            {"pattern": "schemas/**", "requires_all": ["schema"]},
        ],
    )
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["baseline"],
            last_reviewed_template_commit=None,
            local_path_ownership=[
                {
                    "path": "schemas/",
                    "reason": "Project owns the schemas directory.",
                    "overlap_exception_reason": "Downstream owns this entire tree.",
                },
            ],
        ),
    )

    ledger_result = _run_generator(tmp_path, "--ledger-only")

    assert ledger_result.returncode == 0, ledger_result.stderr
    assert (
        "| schemas/ | all: schema | local ownership | "
        "Marker local path ownership: Project owns the schemas directory. "
        "Broad manifest overlap exception: Downstream owns this entire tree. "
        "Manifest proximity is documented by this ownership record."
    ) in ledger_result.stdout
    assert "| schemas/ | local |" not in ledger_result.stdout


def test_ledger_only_without_marker_uses_empty_marker_data(tmp_path: Path) -> None:
    """The upstream template can render a ledger before a marker exists."""
    _init_repo(tmp_path)

    result = _run_generator(tmp_path, "--ledger-only")

    assert result.returncode == 0, result.stderr
    assert "# Template Adoption Ledger" in result.stdout
    assert "- Marker: `.template-sync/marker.yml`" in result.stdout
    assert "- Included modules: none" in result.stdout
    assert "No range base was provided" not in result.stderr


def test_write_ledger_rejects_paths_outside_repository(tmp_path: Path) -> None:
    """The ledger snapshot path uses the existing repo-root traversal guard."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=None),
    )

    result = _run_generator(
        tmp_path,
        "--ledger-only",
        "--write-ledger",
        "../adoption-ledger.md",
    )

    assert result.returncode == 1
    assert "Path escapes the repository root: ../adoption-ledger.md" in result.stderr


def test_ledger_only_rejects_write_candidates_flag(tmp_path: Path) -> None:
    """`--ledger-only` cannot be combined with `--write-candidates`."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=None),
    )

    result = _run_generator(
        tmp_path,
        "--ledger-only",
        "--write-candidates",
        ".cache/template-sync/candidates.md",
    )

    assert result.returncode == 1
    assert "--write-candidates cannot be used with --ledger-only." in result.stderr


def test_ledger_flag_appends_ledger_after_candidate_table(tmp_path: Path) -> None:
    """`--ledger` (non-only) prints the candidate table followed by the adoption ledger."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, "README.md", "head\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=base_sha),
    )

    result = _run_generator(
        tmp_path,
        "--range-head",
        head_sha,
        "--ledger",
    )

    assert result.returncode == 0, result.stderr
    assert "# Template Sync Candidate Table" in result.stdout
    assert "| README.md | Modified | requires all: baseline | Retained |" in result.stdout
    assert "# Template Adoption Ledger" in result.stdout
    assert result.stdout.index("# Template Sync Candidate Table") < result.stdout.index(
        "# Template Adoption Ledger"
    )


def test_summary_reports_clean_adoption_without_unresolved_decisions(tmp_path: Path) -> None:
    """A clean marker produces a concise PR-description-ready summary."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions", "baseline"],
            last_reviewed_template_commit=None,
            protected_decisions=[
                {
                    "path": ".github/copilot-instructions.md",
                    "decision": "MERGE",
                    "adoption_mode": "minimal-preservation",
                    "authorization_basis": "Owner authorized protected edits on 2026-06-04.",
                    "authorized_scope": ".github/copilot-instructions.md only.",
                }
            ],
        ),
    )
    _write_text(
        tmp_path,
        "_TODO-repo-init.md",
        (
            "# Repository Initialization Checklist\n"
            "\n"
            "## Manual GitHub Settings\n"
            "\n"
            "- [x] Private vulnerability reporting decision recorded.\n"
            "\n"
            "## Maintainer Policy Decisions\n"
            "\n"
            "- [x] Adoption mode recorded.\n"
            "\n"
            "## Protected-File Adoption Decisions\n"
            "\n"
            "- [x] Protected-file edits authorized by maintainer.\n"
            "\n"
            "## Unresolved Settings\n"
            "\n"
            "- [x] No unresolved settings remain.\n"
        ),
    )

    result = _run_generator(tmp_path, "--summary")

    assert result.returncode == 0, result.stderr
    assert "# Template Adoption Summary" in result.stdout
    assert "# Template Adoption Ledger" not in result.stdout
    assert "## Included Modules" in result.stdout
    assert "## Excluded Modules" in result.stdout
    included_modules = _summary_section_modules(result.stdout, "Included Modules")
    excluded_modules = _summary_section_modules(result.stdout, "Excluded Modules")
    assert "agent-instructions" in included_modules
    assert "baseline" in included_modules
    assert "agent-instructions" not in excluded_modules
    assert "python" in excluded_modules
    assert "python" not in included_modules
    assert (
        "| .github/copilot-instructions.md | MERGE | minimal-preservation | authorized |"
        in result.stdout
    )
    assert (
        "- Unresolved maintainer decisions remain in available machine-readable state: No."
        in result.stdout
    )
    assert "No unchecked items in machine-interpretable checklist sections." in result.stdout
    assert "python .template-sync/scripts/run_first_adoption_checks.py" in result.stdout
    assert 'pytest tests/ -m "not slow" -v --cov --cov-report=term-missing' not in result.stdout


def test_summary_reports_deferred_overrides_and_documented_todo_items(
    tmp_path: Path,
) -> None:
    """Deferred marker records and documented unchecked TODOs remain visible."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["baseline"],
            last_reviewed_template_commit=None,
            local_overrides=[
                {
                    "path": "README.md",
                    "reason": "Owner has not chosen the downstream README treatment.",
                    "default_decision": "DEFER",
                }
            ],
            deferred_candidates=[
                {
                    "path": ".github/copilot-instructions.md",
                    "source_commit": "a" * 40,
                    "reason": "Awaiting protected-file authorization.",
                }
            ],
        ),
    )
    _write_text(
        tmp_path,
        "_TODO-repo-init.md",
        (
            "# Repository Initialization Checklist\n"
            "\n"
            "## Discoverable Repository State\n"
            "\n"
            "- [ ] Repository owner/name recorded from Git remote.\n"
            "\n"
            "## Maintainer Policy Decisions\n"
            "\n"
            "- [ ] Security vulnerability reporting channel selected.\n"
            "\n"
            "## Resolution Evidence\n"
            "\n"
            "- [ ] Evidence captured for resolved settings.\n"
        ),
    )

    result = _run_generator(tmp_path, "--summary")

    assert result.returncode == 0, result.stderr
    assert "| README.md | DEFER | Owner has not chosen the downstream README treatment. |" in (
        result.stdout
    )
    assert (
        "`README.md`: local override; Marker local override defaults to `DEFER`: "
        "Owner has not chosen the downstream README treatment."
    ) in result.stdout
    assert "Deferred protected candidate: aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa" in result.stdout
    assert (
        "- Unresolved maintainer decisions remain in available machine-readable state: Yes."
        in result.stdout
    )
    assert (
        "- Maintainer Policy Decisions: Security vulnerability reporting channel selected. "
        "(`_TODO-repo-init.md` line 9)"
    ) in result.stdout
    assert (
        "- Discoverable Repository State: Repository owner/name recorded from Git remote. "
        "(`_TODO-repo-init.md` line 5)"
    ) in result.stdout


def test_summary_surfaces_deferred_candidate_unmatched_by_manifest_rows(
    tmp_path: Path,
) -> None:
    """Deferred candidates sharing a templated reason are deduplicated by path, not text."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions", "baseline"],
            last_reviewed_template_commit=None,
            deferred_candidates=[
                {
                    "path": ".github/copilot-instructions.md",
                    "source_commit": "a" * 40,
                    "reason": "Awaiting protected-file authorization.",
                },
                {
                    "path": "AGENTS.md",
                    "source_commit": "b" * 40,
                    "reason": "Awaiting protected-file authorization.",
                },
            ],
        ),
    )

    result = _run_generator(tmp_path, "--summary")

    assert result.returncode == 0, result.stderr
    # The manifest-matched candidate is surfaced via its ledger row.
    assert "Deferred protected candidate: " + "a" * 40 in result.stdout
    # The candidate whose path matches no deferred-candidate ledger row must still be
    # surfaced even though it shares the templated reason text of the matched one.
    assert (
        "`AGENTS.md`: deferred protected candidate from `" + "b" * 40 + "`; "
        "Awaiting protected-file authorization."
    ) in result.stdout


def test_summary_reports_protected_file_removal_authorization(tmp_path: Path) -> None:
    """REMOVE-LOCAL protected decisions keep authorization fields in the summary."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            last_reviewed_template_commit=None,
            protected_decisions=[
                {
                    "path": "GEMINI.md",
                    "decision": "REMOVE-LOCAL",
                    "authorization_basis": "Owner explicitly authorized removing GEMINI.md.",
                    "authorized_scope": "GEMINI.md only.",
                    "reason": "Gemini agent not used by this repository.",
                }
            ],
        ),
    )

    result = _run_generator(tmp_path, "--summary")

    assert result.returncode == 0, result.stderr
    assert (
        "| GEMINI.md | REMOVE-LOCAL | not recorded | authorized | "
        "Owner explicitly authorized removing GEMINI.md. | GEMINI.md only. | "
        "not recorded | Gemini agent not used by this repository. |"
    ) in result.stdout


def test_summary_reports_non_interpretable_todo_without_inferring_tasks(
    tmp_path: Path,
) -> None:
    """A non-document-shaped TODO file is reported but not parsed as policy state."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=None),
    )
    _write_text(tmp_path, "_TODO-repo-init.md", "- [ ] Random local cleanup task\n")

    result = _run_generator(tmp_path, "--summary")

    assert result.returncode == 0, result.stderr
    assert "found; not machine-interpretable" in result.stdout
    assert (
        "`_TODO-repo-init.md` is present but not machine-interpretable; "
        "unchecked items were not parsed."
    ) in result.stdout
    assert "manual review is required" in result.stdout
    assert "Random local cleanup task" not in result.stdout


def test_summary_rejects_todo_when_title_is_not_first_content_line(
    tmp_path: Path,
) -> None:
    """A documented H1 that appears after preamble is not treated as interpretable."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=None),
    )
    _write_text(
        tmp_path,
        "_TODO-repo-init.md",
        (
            "Downstream preamble that is not the documented title.\n"
            "\n"
            "# Repository Initialization Checklist\n"
            "\n"
            "## Maintainer Policy Decisions\n"
            "\n"
            "- [ ] Security vulnerability reporting channel selected.\n"
        ),
    )

    result = _run_generator(tmp_path, "--summary")

    assert result.returncode == 0, result.stderr
    assert "found; not machine-interpretable" in result.stdout
    assert (
        "`_TODO-repo-init.md` is present but not machine-interpretable; "
        "unchecked items were not parsed."
    ) in result.stdout
    assert "manual review is required" in result.stdout
    assert "Security vulnerability reporting channel selected." not in result.stdout


@pytest.mark.parametrize("mode_flag", ["--ledger", "--ledger-only", "--preflight"])
def test_summary_is_mutually_exclusive_with_other_review_modes(
    tmp_path: Path,
    mode_flag: str,
) -> None:
    """`--summary` is part of the existing mutually-exclusive mode group."""
    _init_repo(tmp_path)

    result = _run_generator(tmp_path, "--summary", mode_flag)

    assert result.returncode == 2
    assert "not allowed with argument" in result.stderr


def test_preflight_without_marker_reports_questionnaire_and_reused_ledger(
    tmp_path: Path,
) -> None:
    """The read-only preflight mode works before a marker is created."""
    _init_repo(tmp_path)

    result = _run_generator(tmp_path, "--preflight")

    assert result.returncode == 0, result.stderr
    assert "# First-Adoption Preflight Report" in result.stdout
    assert "Read-only planning artifact." in result.stdout
    assert "Marker: `.template-sync/marker.yml` (not found;" in result.stdout
    assert "GitHub metadata: not requested" in result.stdout
    assert "Which vulnerability reporting channel should `SECURITY.md` publish?" in result.stdout
    assert "## Raw First-Adoption State" in result.stdout
    assert "### Missing State Files" in result.stdout
    assert "`_ADOPTION-DIFFICULTIES.md`" in result.stdout
    assert "## `_TODO-repo-init.md` Starter" in result.stdout
    assert "state: `[not yet asked, asked and deferred, or unavailable" in result.stdout
    assert "## Reused Adoption Ledger" in result.stdout
    assert "# Template Adoption Ledger" in result.stdout
    assert (
        "| .github/copilot-instructions.md | all: agent-instructions | "
        "needs maintainer decision | Protected path is not retained by included modules"
    ) in result.stdout
    assert "No range base was provided" not in result.stderr


def test_preflight_full_state_lists_entries_hidden_by_default(tmp_path: Path) -> None:
    """Default raw state output is bounded; full-state output lists every entry."""
    _init_repo(tmp_path)
    for index in range(12):
        _write_text(tmp_path, f"notes/file-{index:02}.md", "untracked\n")

    bounded = _run_generator(tmp_path, "--preflight")
    full = _run_generator(tmp_path, "--preflight", "--full-state")

    assert bounded.returncode == 0, bounded.stderr
    assert full.returncode == 0, full.stderr
    assert "Remainder:" in bounded.stdout
    assert "`notes/file-11.md`" not in bounded.stdout
    assert "`notes/file-11.md`" in full.stdout
    assert "Remainder:" not in full.stdout


def test_questionnaire_alias_with_marker_reports_protected_questions(
    tmp_path: Path,
) -> None:
    """`--questionnaire` aliases preflight and reuses protected ledger decisions."""
    _init_repo(tmp_path)
    _git(tmp_path, "remote", "add", "origin", "https://github.com/example/project.git")
    _write_text(tmp_path, ".github/copilot-instructions.md", "instructions\n")
    _write_text(tmp_path, "_TODO-repo-init.md", "- [ ] Existing adoption note\n")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["agent-instructions", "baseline"], last_reviewed_template_commit=None),
    )

    result = _run_generator(tmp_path, "--questionnaire")

    assert result.returncode == 0, result.stderr
    assert "Repository owner/name: `example/project`" in result.stdout
    assert (
        "Marker: `.template-sync/marker.yml` (found; included modules: "
        "`agent-instructions`, `baseline`;"
    ) in result.stdout
    assert "First-adoption checklist: `_TODO-repo-init.md` (found)" in result.stdout
    assert "Agent instruction files: `.github/copilot-instructions.md`" in result.stdout
    assert (
        "Does the maintainer authorize the selected action for "
        "`.github/copilot-instructions.md`?"
    ) in result.stdout
    assert (
        "| .github/copilot-instructions.md | all: agent-instructions | "
        "needs maintainer decision | Protected retained path requires explicit owner "
        "authorization."
    ) in result.stdout


def test_preflight_rejects_write_flags(tmp_path: Path) -> None:
    """Preflight remains read-only even though ledger mode can write snapshots."""
    _init_repo(tmp_path)

    result = _run_generator(tmp_path, "--preflight", "--write-ledger", "ledger.md")

    assert result.returncode == 1
    assert "--write-ledger cannot be used with --preflight." in result.stderr


def test_full_state_requires_preflight(tmp_path: Path) -> None:
    """The raw full-state switch belongs to the preflight report only."""
    _init_repo(tmp_path)

    result = _run_generator(tmp_path, "--ledger-only", "--full-state")

    assert result.returncode == 1
    assert "--full-state can only be used with --preflight." in result.stderr


def test_github_metadata_not_requested_skips_provider(tmp_path: Path) -> None:
    """GitHub-only metadata lookup is not attempted without explicit opt-in."""
    calls: list[list[str]] = []

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        calls.append(command)
        return subprocess.CompletedProcess(command, 0, stdout="{}", stderr="")

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=False,
        command_runner=runner,
    )

    assert calls == []
    assert metadata.requested is False
    assert metadata.available is False
    assert metadata.source == "not requested"


def test_github_metadata_provider_reports_visible_values(tmp_path: Path) -> None:
    """The optional GitHub metadata provider can be mocked without network access."""
    calls: list[list[str]] = []

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        calls.append(command)
        if command == ["gh", "api", "/repos/example/project"]:
            return subprocess.CompletedProcess(
                command,
                0,
                stdout=json.dumps(
                    {
                        "full_name": "example/project",
                        "visibility": "public",
                        "default_branch": "main",
                        "has_discussions": True,
                        "security_and_analysis": {
                            "secret_scanning": {"status": "enabled"},
                        },
                    }
                ),
                stderr="",
            )
        if command == [
            "gh",
            "api",
            "--paginate",
            "--slurp",
            "/repos/example/project/labels?per_page=100",
        ]:
            return subprocess.CompletedProcess(
                command,
                0,
                stdout=json.dumps([[{"name": "triage"}, {"name": "bug"}]]),
                stderr="",
            )
        if command == [
            "gh",
            "api",
            "--paginate",
            "--slurp",
            "/repos/example/project/rulesets?per_page=100",
        ]:
            return subprocess.CompletedProcess(
                command,
                0,
                stdout=json.dumps(
                    [
                        [
                            {
                                "id": 1,
                                "name": "default",
                                "target": "branch",
                                "enforcement": "active",
                                "bypass_actors": [],
                            }
                        ]
                    ]
                ),
                stderr="",
            )
        if command == ["gh", "api", "/repos/example/project/branches/main/protection"]:
            return subprocess.CompletedProcess(
                command,
                0,
                stdout=json.dumps(
                    {
                        "required_status_checks": {"contexts": []},
                        "enforce_admins": {"enabled": True},
                    }
                ),
                stderr="",
            )
        if command == ["gh", "api", "/repos/example/project/private-vulnerability-reporting"]:
            return subprocess.CompletedProcess(
                command,
                0,
                stdout=json.dumps({"enabled": False}),
                stderr="",
            )
        raise AssertionError(f"Unexpected command: {command}")

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        clock=_fixed_clock,
    )

    assert not any(command[:3] == ["gh", "label", "list"] for command in calls)
    assert metadata.requested is True
    assert metadata.available is True
    assert metadata.repository == "example/project"
    assert metadata.visibility == "public"
    assert metadata.default_branch == "main"
    assert metadata.discussions_enabled == "enabled"
    assert metadata.labels == ("bug", "triage")
    assert metadata.repository_source == "gh"
    assert metadata.visibility_source == "gh"
    assert metadata.default_branch_source == "gh"
    assert metadata.discussions_source == "gh"
    assert metadata.labels_source == "gh"
    assert _observation(metadata, "private_vulnerability_reporting").value is False
    assert _observation(metadata, "rulesets").pagination == "complete"
    assert _observation(metadata, "repository").observed_at == "2026-06-29T12:34:56Z"


def test_github_metadata_rest_fallback_reports_public_values_after_missing_gh(
    tmp_path: Path,
) -> None:
    """Missing gh is not a dead end for public-safe REST metadata."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        raise FileNotFoundError(2, "No such file or directory", command[0])

    rest_client = FakeRestClient(
        _metadata_rest_success_responses(
            {
                "full_name": "example/project",
                "visibility": "public",
                "default_branch": "main",
                "has_discussions": False,
                "security_and_analysis": {"secret_scanning": {"status": "disabled"}},
            },
            [{"name": "triage"}, {"name": "bug"}],
            rulesets_payload=[],
            private_vulnerability_payload={"enabled": False},
        )
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    assert metadata.requested is True
    assert metadata.available is True
    assert metadata.source == "rest"
    assert metadata.repository == "example/project"
    assert metadata.visibility == "public"
    assert metadata.default_branch == "main"
    assert metadata.discussions_enabled == "disabled"
    assert metadata.labels == ("bug", "triage")
    assert _observation(metadata, "discussions").value is False
    assert _observation(metadata, "private_vulnerability_reporting").value is False
    assert _observation(metadata, "rulesets").value == []
    assert _observation(metadata, "rulesets").pagination == "complete"
    assert metadata.repository_source == "rest"
    assert metadata.visibility_source == "rest"
    assert metadata.default_branch_source == "rest"
    assert metadata.discussions_source == "rest"
    assert metadata.labels_source == "rest"
    assert metadata.error is not None
    assert "gh unavailable: FileNotFoundError: No such file or directory" in metadata.error

    assert [request.method for request in rest_client.requests] == ["GET"] * 5
    assert rest_client.requests[0].url == "https://api.github.com/repos/example/project"
    assert (
        rest_client.requests[1].url
        == "https://api.github.com/repos/example/project/labels?per_page=100&page=1"
    )
    assert (
        rest_client.requests[2].url
        == "https://api.github.com/repos/example/project/rulesets?per_page=100&page=1"
    )
    assert (
        rest_client.requests[3].url
        == "https://api.github.com/repos/example/project/branches/main/protection"
    )
    assert (
        rest_client.requests[4].url
        == "https://api.github.com/repos/example/project/private-vulnerability-reporting"
    )
    headers = dict(rest_client.requests[0].headers)
    assert headers["Accept"] == "application/vnd.github+json"
    assert headers["X-GitHub-Api-Version"] == "2022-11-28"
    assert headers["User-Agent"] == "copilot-repo-template-preflight/1.0"
    assert "Authorization" not in headers
    assert "example" not in headers["User-Agent"]
    assert rest_client.requests[0].timeout == sync_candidates.GITHUB_REST_TIMEOUT_SECONDS

    rendered = sync_candidates.format_github_metadata(metadata)
    assert "Repository owner/name | `example/project` | observed complete (`rest`)" in rendered
    assert "Default branch | `main` | observed complete (`rest`)" in rendered
    assert "Labels | `bug`, `triage` | observed complete (`rest`)" in rendered
    assert "Repository rulesets | observed empty list | observed complete (`rest`)" in rendered
    assert "Private vulnerability reporting | `false` (disabled, observed)" in rendered


def test_github_metadata_rest_fallback_handles_unauthenticated_gh(
    tmp_path: Path,
) -> None:
    """Unauthenticated gh falls through to public REST reads."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh auth required\n")

    rest_client = FakeRestClient(
        _metadata_rest_success_responses(
            {
                "full_name": "example/project",
                "private": False,
                "default_branch": "trunk",
                "has_discussions": True,
            },
            [],
        )
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    assert metadata.available is True
    assert metadata.source == "rest"
    assert metadata.visibility == "public"
    assert metadata.default_branch == "trunk"
    assert metadata.discussions_enabled == "enabled"
    assert metadata.labels == ()
    assert metadata.labels_source == "rest"
    assert metadata.error is not None
    assert "gh auth required" in metadata.error
    # An auth failure must not be relabeled as gh being missing/broken: the
    # self-describing "gh ... unavailable" message is kept verbatim, not double
    # prefixed with "gh unavailable:".
    assert "gh unavailable: gh repo metadata" not in metadata.error


def test_github_metadata_rest_fallback_handles_malformed_gh_json(tmp_path: Path) -> None:
    """Malformed gh JSON is unusable, so public REST fallback is attempted."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 0, stdout="{", stderr="")

    rest_client = FakeRestClient(
        _metadata_rest_success_responses(
            {
                "full_name": "example/project",
                "visibility": "public",
                "default_branch": "main",
                "has_discussions": True,
            },
            [],
        )
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    assert metadata.available is True
    assert metadata.source == "rest"
    assert metadata.error is not None
    assert "gh repo metadata returned malformed JSON" in metadata.error


@pytest.mark.parametrize(
    ("status_code", "message", "expected_basis"),
    [
        (401, "Requires authentication", "auth-required"),
        (403, "Resource not accessible by integration", "permission-required"),
        (403, "API rate limit exceeded for 203.0.113.10", "rate-limited"),
    ],
)
def test_github_metadata_rest_auth_permission_and_rate_limit_are_not_version_retries(
    tmp_path: Path,
    status_code: int,
    message: str,
    expected_basis: str,
) -> None:
    """Auth, permission, and rate-limit responses remain distinguishable failures."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient([_rest_result(status_code, {"message": message})])

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
    )

    assert metadata.available is False
    # The compatibility ``source`` summary must stay in the legacy vocabulary
    # ("not requested" / "manual" / "gh" / "rest") and not leak a per-observation
    # basis token such as "auth-required" or "rate-limited".
    assert metadata.source == "manual"
    assert _observation(metadata, "repository").basis == expected_basis
    assert metadata.error is not None
    assert f"HTTP {status_code}: {message}" in metadata.error
    assert len(rest_client.requests) == 1
    assert "X-GitHub-Api-Version" in dict(rest_client.requests[0].headers)


def test_github_metadata_rest_timeout_degrades_to_manual_review(tmp_path: Path) -> None:
    """Transport failures are reported as unavailable metadata, not fatal errors."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient([_rest_result(None, error="TimeoutError: timed out")])

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
    )

    assert metadata.available is False
    assert metadata.error is not None
    assert "REST repo metadata unavailable: TimeoutError: timed out" in metadata.error


def test_github_metadata_rest_malformed_json_degrades_to_manual_review(
    tmp_path: Path,
) -> None:
    """Malformed REST JSON is reported without aborting the preflight."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient([_rest_result(200, "{")])

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
    )

    assert metadata.available is False
    assert metadata.error is not None
    assert "REST repo metadata returned malformed JSON" in metadata.error


def test_github_metadata_rest_unexpected_payload_types_are_reported(
    tmp_path: Path,
) -> None:
    """Unexpected REST payload shapes become diagnostics instead of crashes."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    repo_payload_client = FakeRestClient([_rest_result(200, [])])
    repo_payload_metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=repo_payload_client,
    )

    assert repo_payload_metadata.available is False
    assert repo_payload_metadata.error is not None
    assert "REST repo metadata returned a non-object JSON payload" in repo_payload_metadata.error

    labels_payload_client = FakeRestClient(
        _metadata_rest_success_responses(
            {
                "full_name": "example/project",
                "visibility": "public",
                "default_branch": "main",
                "has_discussions": True,
            },
            {"name": "not-a-list"},
        )
    )
    labels_payload_metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=labels_payload_client,
    )

    assert labels_payload_metadata.available is True
    assert labels_payload_metadata.labels_source == "manual"
    assert labels_payload_metadata.error is not None
    assert "REST labels metadata returned a non-list JSON payload" in labels_payload_metadata.error
    assert _observation(labels_payload_metadata, "labels").value is None
    assert _observation(labels_payload_metadata, "labels").basis == "malformed-response"


def test_github_metadata_rest_uses_ghes_api_base_override(tmp_path: Path) -> None:
    """An explicit API base allows GHES remotes without assuming github.com."""
    _init_repo(tmp_path)
    _git(tmp_path, "remote", "add", "origin", "https://github.company.com/octo/widget.git")

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient(
        _metadata_rest_success_responses(
            {
                "full_name": "octo/widget",
                "visibility": "internal",
                "default_branch": "main",
                "has_discussions": False,
            },
            [],
        )
    )

    discovery = sync_candidates.discover_repository_state(
        repo_root=tmp_path,
        marker_path=tmp_path / ".template-sync" / "marker.yml",
        todo_path=tmp_path / "_TODO-repo-init.md",
        include_github_metadata=True,
        github_api_base="https://github.company.com/api/v3",
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    assert discovery.owner_name == "octo/widget"
    assert discovery.github_metadata.available is True
    assert discovery.github_metadata.source == "rest"
    assert rest_client.requests[0].url == "https://github.company.com/api/v3/repos/octo/widget"


def test_github_metadata_gh_targets_ghes_host_via_hostname(tmp_path: Path) -> None:
    """A non-default API base routes authenticated gh reads to that host via --hostname."""
    commands: list[list[str]] = []

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        commands.append(command)
        if command[-1] == "/repos/octo/widget":
            return subprocess.CompletedProcess(
                command,
                0,
                stdout=json.dumps(
                    {
                        "full_name": "octo/widget",
                        "visibility": "internal",
                        "default_branch": "main",
                        "has_discussions": True,
                    }
                ),
                stderr="",
            )
        if "--paginate" in command:
            return subprocess.CompletedProcess(command, 0, stdout="[]", stderr="")
        return subprocess.CompletedProcess(
            command, 0, stdout=json.dumps({"enabled": False}), stderr=""
        )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "octo/widget",
        include_metadata=True,
        command_runner=runner,
        github_api_base="https://github.company.com/api/v3",
        clock=_fixed_clock,
    )

    assert metadata.available is True
    # gh was attempted first (not skipped) and observed the GHES repo.
    assert metadata.source == "gh"
    assert metadata.repository == "octo/widget"
    # Every gh api call targets the GHES host instead of assuming github.com.
    gh_calls = [command for command in commands if command[:2] == ["gh", "api"]]
    assert gh_calls
    assert all("--hostname" in command for command in gh_calls)
    assert all("github.company.com" in command for command in gh_calls)


def test_github_metadata_rest_retries_ghes_version_header_rejection(
    tmp_path: Path,
) -> None:
    """GHES version-header incompatibility gets one read-only retry without the header."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient(
        [
            _rest_result(400, {"message": "X-GitHub-Api-Version is not supported"}),
            _rest_result(
                200,
                {
                    "full_name": "example/project",
                    "visibility": "public",
                    "default_branch": "main",
                    "has_discussions": True,
                },
            ),
            _rest_result(200, []),
            _rest_result(200, []),
            _rest_result(
                200,
                {
                    "required_status_checks": {"contexts": []},
                    "enforce_admins": {"enabled": True},
                },
            ),
            _rest_result(200, {"enabled": False}),
        ]
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    assert metadata.available is True
    assert metadata.source == "rest"
    assert metadata.error is not None
    assert "rejected the GitHub API version header" in metadata.error
    # read_github_rest_json already prefixes the label, so the caller must not
    # add it again and produce "REST repo metadata: REST repo metadata ...".
    assert "REST repo metadata: REST repo metadata" not in metadata.error
    assert rest_client.requests[0].url == rest_client.requests[1].url
    assert "X-GitHub-Api-Version" in dict(rest_client.requests[0].headers)
    assert "X-GitHub-Api-Version" not in dict(rest_client.requests[1].headers)
    assert rest_client.requests[0].method == "GET"
    assert rest_client.requests[1].method == "GET"


def test_github_metadata_rest_repository_source_is_manual_without_full_name(
    tmp_path: Path,
) -> None:
    """When REST omits full_name, the repository is locally inferred and labeled manual."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient(
        _metadata_rest_success_responses(
            {
                "visibility": "public",
                "default_branch": "main",
                "has_discussions": True,
            },
            [],
        )
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    assert metadata.available is True
    assert metadata.source == "rest"
    # full_name was absent, so the GitHub metadata value stays unobserved/null
    # instead of treating the locally inferred owner/name as a remote observation.
    assert metadata.repository is None
    assert metadata.repository_source == "manual"
    assert _observation(metadata, "repository").value is None
    assert _observation(metadata, "repository").basis == "manual-review"
    # Fields that were present in the REST payload remain sourced from rest.
    assert metadata.visibility == "public"
    assert metadata.visibility_source == "rest"
    assert metadata.default_branch_source == "rest"


def test_github_metadata_rest_failed_version_header_retry_is_unavailable(
    tmp_path: Path,
) -> None:
    """If the GHES version-header retry also fails, metadata stays unavailable."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient(
        [
            _rest_result(400, {"message": "Unsupported API version header"}),
            _rest_result(404, {"message": "Not Found"}),
        ]
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
    )

    assert metadata.available is False
    assert metadata.error is not None
    assert "retried the same read-only endpoint without it" in metadata.error
    assert "Retry failed: HTTP 404: Not Found" in metadata.error
    assert len(rest_client.requests) == 2


def test_github_metadata_rest_paginates_collections_and_marks_rulesets_partial(
    tmp_path: Path,
) -> None:
    """REST collection pagination follows Link evidence and marks partial objects."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient(
        [
            _rest_result(
                200,
                {
                    "full_name": "example/project",
                    "visibility": "public",
                    "default_branch": "main",
                    "has_discussions": True,
                },
            ),
            _rest_result(
                200,
                [{"name": "bug"}],
                headers=(
                    (
                        "Link",
                        '<https://api.github.com/repositories/1/labels?page=2>; rel="next"',
                    ),
                ),
            ),
            _rest_result(200, [{"name": "triage"}]),
            _rest_result(
                200,
                [
                    {
                        "id": 1,
                        "name": "protect-main",
                        "target": "branch",
                        "enforcement": "active",
                    }
                ],
            ),
            _rest_result(200, {"required_status_checks": {"contexts": []}}),
            _rest_result(200, {"enabled": False}),
        ]
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    labels = _observation(metadata, "labels")
    rulesets = _observation(metadata, "rulesets")

    assert metadata.labels == ("bug", "triage")
    assert labels.pagination == "complete"
    assert rest_client.requests[1].url.endswith("/labels?per_page=100&page=1")
    assert rest_client.requests[2].url.endswith("/labels?per_page=100&page=2")
    assert rulesets.partial is True
    assert rulesets.partial_fields == ("rulesets[].bypass_actors",)


def test_github_metadata_rest_marks_collection_partial_after_late_page_failure(
    tmp_path: Path,
) -> None:
    """A later page failure keeps earlier collection items but marks the result partial."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient(
        [
            _rest_result(
                200,
                {
                    "full_name": "example/project",
                    "visibility": "public",
                    "default_branch": "main",
                    "has_discussions": True,
                },
            ),
            _rest_result(
                200,
                [{"name": "bug"}],
                headers=(
                    (
                        "Link",
                        '<https://api.github.com/repositories/1/labels?page=2>; rel="next"',
                    ),
                ),
            ),
            _rest_result(
                403,
                {"message": "Resource not accessible by integration"},
                headers=(
                    (
                        "X-Accepted-GitHub-Permissions",
                        "metadata=read; administration=read",
                    ),
                ),
            ),
            _rest_result(200, []),
            _rest_result(200, {"required_status_checks": {"contexts": []}}),
            _rest_result(200, {"enabled": False}),
        ]
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    labels = _observation(metadata, "labels")

    assert labels.observed is True
    assert labels.value == ["bug"]
    assert labels.pagination == "partial"
    assert any(key == "accepted_permissions" for key, _value in labels.diagnostics)
    assert "administration=read" in sync_candidates.github_observation_diagnostic_text(labels)


def test_github_metadata_rest_rate_limit_headers_are_classified(tmp_path: Path) -> None:
    """Rate-limit basis uses status plus headers, not status code alone."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient(
        [
            _rest_result(
                429,
                {"message": "You have exceeded a secondary rate limit."},
                headers=(
                    ("x-ratelimit-remaining", "0"),
                    ("x-ratelimit-reset", "1790000000"),
                    ("retry-after", "60"),
                ),
            )
        ]
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    repository = _observation(metadata, "repository")

    assert metadata.available is False
    assert repository.value is None
    assert repository.basis == "rate-limited"
    assert ("rate_limit_remaining", "0") in repository.diagnostics
    assert ("retry_after", "60") in repository.diagnostics


def test_github_metadata_rest_skips_branch_protection_without_default_branch(
    tmp_path: Path,
) -> None:
    """Branch protection is dependency-unobserved when the branch is unknown."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient(
        [
            _rest_result(
                200,
                {
                    "full_name": "example/project",
                    "visibility": "public",
                    "has_discussions": True,
                },
            ),
            _rest_result(200, []),
            _rest_result(200, []),
            _rest_result(200, {"enabled": False}),
        ]
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    branch_protection = _observation(metadata, "branch_protection")

    assert branch_protection.value is None
    assert branch_protection.basis == "dependency-unobserved"
    assert not any("/branches/" in request.url for request in rest_client.requests)


def test_github_metadata_rest_private_vulnerability_404_is_not_disabled(
    tmp_path: Path,
) -> None:
    """Endpoint-specific 404 leaves private reporting unknown, not disabled."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient(
        [
            *_metadata_rest_success_responses(
                {
                    "full_name": "example/project",
                    "visibility": "public",
                    "default_branch": "main",
                    "has_discussions": True,
                },
                [],
            )[:-1],
            _rest_result(404, {"message": "Not Found"}),
        ]
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    private_reporting = _observation(metadata, "private_vulnerability_reporting")
    rendered = sync_candidates.format_github_metadata(metadata)

    assert private_reporting.value is None
    assert private_reporting.basis == "not-found-or-not-visible"
    # An unavailable endpoint (payload is None) must keep the read's basis and must
    # not be described as a malformed object.
    assert all("was not an object" not in value for _key, value in private_reporting.diagnostics)
    assert "Private vulnerability reporting | `null` (not observed)" in rendered
    assert "Private vulnerability reporting | `false`" not in rendered


def test_github_metadata_rest_private_vulnerability_non_object_is_malformed(
    tmp_path: Path,
) -> None:
    """A present-but-non-object private-reporting payload is malformed, not unavailable."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 1, stdout="", stderr="gh unavailable\n")

    rest_client = FakeRestClient(
        [
            *_metadata_rest_success_responses(
                {
                    "full_name": "example/project",
                    "visibility": "public",
                    "default_branch": "main",
                    "has_discussions": True,
                },
                [],
            )[:-1],
            _rest_result(200, ["unexpected"]),
        ]
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    private_reporting = _observation(metadata, "private_vulnerability_reporting")
    assert private_reporting.observed is False
    assert private_reporting.basis == "malformed-response"
    assert any("was not an object" in value for _key, value in private_reporting.diagnostics)


def test_github_metadata_security_and_analysis_missing_is_manual_review(
    tmp_path: Path,
) -> None:
    """A successful payload without security_and_analysis is manual-review, not permission-required."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        raise FileNotFoundError(2, "No such file or directory", command[0])

    rest_client = FakeRestClient(
        _metadata_rest_success_responses(
            {
                "full_name": "example/project",
                "visibility": "public",
                "default_branch": "main",
                "has_discussions": True,
            },
            [],
        )
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    security = _observation(metadata, "security_and_analysis")
    assert security.observed is False
    # A missing field in an otherwise-successful payload is not proof of a permission
    # failure (it is also absent on older API versions, GHES, and for non-admin callers).
    assert security.basis == "manual-review"
    assert security.value is None
    message = dict(security.diagnostics).get("message", "")
    assert "security_and_analysis was not present" in message
    assert "admin" in message


def test_markdown_code_span_is_safe_for_backticks_pipes_and_newlines() -> None:
    """The code-span helper survives backticks, pipes, CR/LF, and edge backticks."""
    # No special characters: single-backtick span, no padding.
    assert sync_candidates.markdown_code_span("plain") == "`plain`"
    # Embedded backtick runs force a fence one backtick longer than the longest run.
    assert sync_candidates.markdown_code_span("a`b") == "``a`b``"
    assert sync_candidates.markdown_code_span("a``b") == "```a``b```"
    # Leading/trailing backticks require space padding (CommonMark).
    assert sync_candidates.markdown_code_span("`x`") == "`` `x` ``"
    # Pipes are escaped so the GFM table parser keeps them inside the cell.
    assert sync_candidates.markdown_code_span("a|b") == "`a\\|b`"
    # CR/LF are flattened so a value cannot break out of its table row.
    assert sync_candidates.markdown_code_span("a\r\nb") == "`a b`"
    assert sync_candidates.markdown_code_span("a\nb") == "`a b`"
    # Backslashes are preserved verbatim because they are literal inside a code span.
    assert sync_candidates.markdown_code_span("C:\\path") == "`C:\\path`"


def test_format_github_metadata_renders_backtick_unsafe_values_without_breaking_table() -> None:
    """Labels containing backticks or pipes stay inside one well-formed table cell."""
    observation = sync_candidates.GitHubObservation(
        setting="labels",
        value=["needs `review`", "area|core"],
        observed=True,
        basis="rest",
        source="REST repos/example/project/labels",
        observed_at="2026-06-29T12:34:56Z",
        pagination="complete",
    )
    metadata = sync_candidates.GitHubMetadata(requested=True, observations=(observation,))

    rendered = sync_candidates.format_github_metadata(metadata)
    labels_row = next(line for line in rendered.splitlines() if line.startswith("| Labels "))

    # Backticks/pipes in values must not add table columns: exactly four cells.
    assert labels_row.count(" | ") == 3
    # The backtick-containing label is fenced with a longer run and padded.
    assert "`` needs `review` ``" in labels_row
    # The pipe is escaped so it stays inside the cell instead of splitting the row.
    assert "`area\\|core`" in labels_row


def test_github_boolean_observation_classifies_missing_vs_malformed() -> None:
    """Missing boolean fields are manual-review; present non-bools are malformed."""
    observed = sync_candidates.github_boolean_observation(
        setting="discussions", value=True, source="REST", observed_at="t", basis="rest"
    )
    assert observed.observed is True
    assert observed.value is True
    assert observed.basis == "rest"

    missing = sync_candidates.github_boolean_observation(
        setting="discussions", value=None, source="REST", observed_at="t", basis="rest"
    )
    assert missing.observed is False
    assert missing.value is None
    assert missing.basis == "manual-review"

    malformed = sync_candidates.github_boolean_observation(
        setting="discussions", value="yes", source="REST", observed_at="t", basis="rest"
    )
    assert malformed.observed is False
    assert malformed.basis == "malformed-response"
    assert any("was not a boolean" in value for _key, value in malformed.diagnostics)


def test_github_metadata_missing_discussions_is_manual_review_not_malformed(
    tmp_path: Path,
) -> None:
    """A successful payload without a Discussions flag is unobserved manual-review, not malformed."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        raise FileNotFoundError(2, "No such file or directory", command[0])

    rest_client = FakeRestClient(
        _metadata_rest_success_responses(
            {
                "full_name": "example/project",
                "visibility": "public",
                "default_branch": "main",
                # has_discussions / hasDiscussionsEnabled intentionally omitted.
            },
            [],
        )
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    discussions = _observation(metadata, "discussions")
    assert discussions.observed is False
    assert discussions.value is None
    assert discussions.basis == "manual-review"


def test_link_header_has_next_is_case_insensitive() -> None:
    """rel="next" detection tolerates case variations in Link header parameters."""
    assert sync_candidates.link_header_has_next(
        (("Link", '<https://api.github.com/x?page=2>; rel="next"'),)
    )
    assert sync_candidates.link_header_has_next(
        (("Link", '<https://api.github.com/x?page=2>; REL="next"'),)
    )
    assert sync_candidates.link_header_has_next(
        (("link", '<https://api.github.com/x?page=2>; rel="Next"'),)
    )
    assert not sync_candidates.link_header_has_next(
        (("Link", '<https://api.github.com/x?page=1>; rel="prev"'),)
    )


def test_ruleset_summaries_skips_malformed_items_without_collapsing() -> None:
    """Malformed ruleset items are skipped with a diagnostic; the result stays a list."""
    summaries, _partial_fields, diagnostics = sync_candidates.ruleset_summaries(
        (
            {
                "id": 1,
                "name": "default",
                "target": "branch",
                "enforcement": "active",
                "bypass_actors": [],
            },
            "not-an-object",
        )
    )
    assert isinstance(summaries, list)
    assert len(summaries) == 1
    assert summaries[0]["name"] == "default"
    assert any("was not an object" in value for _key, value in diagnostics)


def test_github_api_base_validation_rejects_credential_bearing_urls() -> None:
    """The API base must not preserve or print credential-bearing URLs."""
    with pytest.raises(sync_candidates.CandidateGenerationError) as exc_info:
        sync_candidates.normalize_github_api_base("https://token@example.com/api/v3")

    assert "must not contain credentials" in str(exc_info.value)
    assert "token" not in str(exc_info.value)


def test_run_rest_request_rejects_non_get_methods() -> None:
    """The REST helper enforces its read-only contract by refusing non-GET methods."""
    request = sync_candidates.RestRequest(
        method="POST",
        url="https://api.github.com/repos/example/project",
        headers=sync_candidates.rest_headers(include_version_header=True),
        timeout=sync_candidates.GITHUB_REST_TIMEOUT_SECONDS,
    )

    with pytest.raises(ValueError) as exc_info:
        sync_candidates.run_rest_request(request)

    assert "GET" in str(exc_info.value)
    assert "POST" in str(exc_info.value)


def test_format_remotes_redacts_embedded_credentials() -> None:
    """Remote URLs with embedded user-info are redacted in the report."""
    remotes = (
        sync_candidates.RemoteInfo(
            name="origin",
            url="https://maintainer:ghp_secrettoken@github.com/example/project.git",
            purpose="fetch",
        ),
        sync_candidates.RemoteInfo(
            name="bare-token",
            url="https://ghp_anothersecret@github.com/example/project.git",
            purpose="push",
        ),
        sync_candidates.RemoteInfo(
            name="scheme-relative",
            url="//ghp_schemerelative@github.com/example/project.git",
            purpose="fetch",
        ),
        sync_candidates.RemoteInfo(
            name="ssh",
            url="git@github.com:example/project.git",
            purpose="push",
        ),
    )

    rendered = sync_candidates.format_remotes(remotes)

    assert "ghp_secrettoken" not in rendered
    assert "ghp_anothersecret" not in rendered
    assert "ghp_schemerelative" not in rendered
    assert "maintainer:" not in rendered
    assert "https://***@github.com/example/project.git" in rendered
    # Scheme-relative authorities must also be redacted, not just scheme:// URLs.
    assert "//***@github.com/example/project.git" in rendered
    # SCP-style SSH URLs cannot embed a password and are preserved verbatim.
    assert "git@github.com:example/project.git" in rendered


def test_redact_url_userinfo_handles_url_forms() -> None:
    """The canonical redactor covers scheme, scheme-relative, SCP, and malformed forms."""
    redact = sync_candidates.redact_url_userinfo

    assert (
        redact("https://maintainer:token@github.com/org/repo.git")
        == "https://***@github.com/org/repo.git"
    )
    assert (
        redact("https://ghp_token@github.com/org/repo.git") == "https://***@github.com/org/repo.git"
    )
    # Scheme-relative authority must also be redacted.
    assert redact("//user:token@host/path") == "//***@host/path"
    # SCP-style SSH remotes have no URL authority and are preserved.
    assert redact("git@github.com:org/repo.git") == "git@github.com:org/repo.git"
    # No user-info: unchanged.
    assert redact("https://github.com/org/repo.git") == "https://github.com/org/repo.git"
    # Unparseable authority (malformed IPv6): fail-safe redaction that drops the scheme.
    assert redact("https://user:secret@[::1/path") == "***@[::1/path"


def test_infer_repository_name_anchors_github_host() -> None:
    """Owner/name inference matches the URL authority, not any github substring."""

    def infer(url: str) -> str | None:
        return cast(
            str | None,
            sync_candidates.infer_repository_name(
                (sync_candidates.RemoteInfo(name="origin", url=url, purpose="fetch"),)
            ),
        )

    # Genuine GitHub remotes (scheme, SCP, user-info, and subdomain) resolve.
    assert infer("https://github.com/example/project.git") == "example/project"
    assert infer("git@github.com:example/project.git") == "example/project"
    assert infer("https://user:token@github.com/example/project") == "example/project"
    assert infer("https://team.github.com/example/project.git") == "example/project"
    # Look-alike or unrelated hosts MUST NOT be misclassified as GitHub.
    assert infer("https://notgithub.com/example/project.git") is None
    assert infer("https://github.com.evil.com/example/project.git") is None
    assert infer("git@gitlab.com:example/project.git") is None

    # The first genuine GitHub remote wins even if a look-alike precedes it.
    remotes = (
        sync_candidates.RemoteInfo(
            name="decoy", url="https://notgithub.com/decoy/repo.git", purpose="fetch"
        ),
        sync_candidates.RemoteInfo(
            name="origin", url="https://github.com/example/project.git", purpose="fetch"
        ),
    )
    assert sync_candidates.infer_repository_name(remotes) == "example/project"


def test_github_metadata_handles_non_object_json(tmp_path: Path) -> None:
    """Non-object JSON from gh repo view falls through to public REST fallback."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 0, stdout="null", stderr="")

    rest_client = FakeRestClient(
        _metadata_rest_success_responses(
            {
                "full_name": "example/project",
                "visibility": "public",
                "default_branch": "main",
                "has_discussions": True,
            },
            [],
        )
    )

    metadata = sync_candidates.discover_github_metadata(
        tmp_path,
        "example/project",
        include_metadata=True,
        command_runner=runner,
        rest_client=rest_client,
        clock=_fixed_clock,
    )

    assert metadata.requested is True
    assert metadata.available is True
    assert metadata.source == "rest"
    assert metadata.error is not None
    assert "non-object" in metadata.error


def test_list_directory_files_skips_symlinked_ancestor(tmp_path: Path, monkeypatch: Any) -> None:
    """A symlinked ancestor directory is not iterated outside the repo root."""
    outside = tmp_path / "outside" / "workflows"
    outside.mkdir(parents=True)
    (outside / "leak.yml").write_text("name: leak\n", encoding="utf-8")

    repo_root = tmp_path / "repo"
    repo_root.mkdir()
    # `.github` is a symlink whose target lives outside the repo root.
    _symlink_or_skip(repo_root / ".github", tmp_path / "outside")

    escaping_dir = (repo_root / ".github" / "workflows").resolve()
    original_iterdir = Path.iterdir

    def guarded_iterdir(self: Path) -> Any:
        if self.resolve() == escaping_dir:
            raise AssertionError(f"iterated outside repo root: {self}")
        return original_iterdir(self)

    monkeypatch.setattr(Path, "iterdir", guarded_iterdir)

    result = sync_candidates.list_directory_files(repo_root, ".github/workflows", (".yml", ".yaml"))

    assert result == ()


def test_path_has_symlink_component_detects_symlinked_ancestor(tmp_path: Path) -> None:
    """A symlinked ancestor (or leaf) is detected; fully real paths pass."""
    repo_root = tmp_path / "repo"
    (repo_root / "real" / "nested").mkdir(parents=True)
    (repo_root / "real" / "nested" / "file.txt").write_text("x\n", encoding="utf-8")
    outside = tmp_path / "outside"
    outside.mkdir()
    # `.github` is a symlink whose target lives outside the repo root.
    _symlink_or_skip(repo_root / ".github", outside)

    # A symlinked ancestor and a symlinked leaf are both rejected.
    assert sync_candidates.path_has_symlink_component(repo_root, ".github/workflows") is True
    assert sync_candidates.path_has_symlink_component(repo_root, ".github") is True
    # Fully real paths are accepted whether or not they exist.
    assert sync_candidates.path_has_symlink_component(repo_root, "real/nested") is False
    assert sync_candidates.path_has_symlink_component(repo_root, "real/nested/file.txt") is False
    assert sync_candidates.path_has_symlink_component(repo_root, "does/not/exist") is False


def test_repository_path_exists_rejects_symlinked_ancestor(tmp_path: Path) -> None:
    """A path reached through a symlinked ancestor is not treated as a repo path."""
    repo_root = tmp_path / "repo"
    repo_root.mkdir()
    outside = tmp_path / "outside"
    outside.mkdir()
    (outside / "config.yml").write_text("name: leak\n", encoding="utf-8")
    # `.github` is a symlink whose target lives outside the repo root.
    _symlink_or_skip(repo_root / ".github", outside)

    # The file exists *through* the symlink, but must be rejected as out-of-boundary.
    assert sync_candidates.repository_path_exists(repo_root, ".github/config.yml") is False

    # A genuine in-repo file is accepted.
    (repo_root / "real").mkdir()
    (repo_root / "real" / "in_repo.yml").write_text("ok\n", encoding="utf-8")
    assert sync_candidates.repository_path_exists(repo_root, "real/in_repo.yml") is True


def test_unmapped_paths_are_flagged(tmp_path: Path) -> None:
    """Changed paths outside the manifest surface as owner-review items."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, "docs/unmapped.md", "new\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=base_sha),
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 0, result.stderr
    assert "| docs/unmapped.md | Added | UNMAPPED | Unmapped |" in result.stdout
    assert "owner must assign or confirm a module" in result.stdout


def test_unknown_manifest_modules_are_flagged(tmp_path: Path) -> None:
    """Mappings that reference modules absent from the manifest are surfaced."""
    _init_repo(
        tmp_path,
        extra_mappings=[
            {"pattern": "future.md", "requires_all": ["future-module"]},
        ],
    )
    _write_text(tmp_path, "README.md", "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, "future.md", "new\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["baseline"], last_reviewed_template_commit=base_sha),
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 0, result.stderr
    assert "| future.md | Added | requires all: future-module | Excluded |" in result.stdout
    assert "Unknown module(s): future-module." in result.stdout


def test_local_overrides_are_reported(tmp_path: Path) -> None:
    """Marker local overrides appear as defaults, not automatic decisions."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, "README.md", "head\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["baseline"],
            last_reviewed_template_commit=base_sha,
            local_overrides=[
                {
                    "path": "README.md",
                    "reason": "Project-specific README.",
                    "default_decision": "SKIP",
                }
            ],
        ),
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 0, result.stderr
    assert "SKIP: Project-specific README." in result.stdout
    assert "Local override present; use it as a default" in result.stdout


def test_catalog_is_protected_without_becoming_an_agent_entry_point(tmp_path: Path) -> None:
    """Catalog evolution needs authority even when the agent module is not retained."""
    _init_repo(tmp_path)
    catalog_path = ".template-sync/instruction-contracts.yml"
    _write_text(tmp_path, catalog_path, "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, catalog_path, "head\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["template-sync-support"], last_reviewed_template_commit=base_sha),
    )
    result = _run_generator(tmp_path, "--range-head", head_sha)
    assert result.returncode == 0, result.stderr
    assert "| .template-sync/instruction-contracts.yml | Modified |" in result.stdout
    assert (
        "Protected instruction/governance file; explicit owner authorization is required."
        in result.stdout
    )
    assert catalog_path not in sync_candidates.discover_agent_instruction_files(tmp_path)
    _write_text(tmp_path, "AGENTS.md", "Agent guidance\n")
    assert "AGENTS.md" in sync_candidates.discover_agent_instruction_files(tmp_path)


def test_deferred_protected_candidates_and_protected_files_are_flagged(tmp_path: Path) -> None:
    """Deferred protected marker entries and protected paths are visible."""
    _init_repo(tmp_path)
    _write_text(tmp_path, ".github/copilot-instructions.md", "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, ".github/copilot-instructions.md", "head\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            last_reviewed_template_commit=base_sha,
            deferred_candidates=[
                {
                    "path": ".github/copilot-instructions.md",
                    "source_commit": head_sha,
                    "reason": "Awaiting owner authorization.",
                }
            ],
        ),
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 0, result.stderr
    assert "Awaiting owner authorization." in result.stdout
    assert "| .github/copilot-instructions.md | Modified |" in result.stdout
    assert (
        "Protected instruction/governance file; explicit owner authorization is required."
        in result.stdout
    )


def test_protected_file_decision_is_reported_in_candidate_table(tmp_path: Path) -> None:
    """Changed protected paths show matching protected decision records."""
    _init_repo(tmp_path)
    _write_text(tmp_path, ".github/copilot-instructions.md", "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, ".github/copilot-instructions.md", "head\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            last_reviewed_template_commit=base_sha,
            protected_decisions=[
                {
                    "path": ".github/copilot-instructions.md",
                    "decision": "MERGE",
                    "adoption_mode": "minimal-preservation",
                    "authorization_basis": "Owner authorized protected edits on 2026-05-27.",
                    "authorized_scope": ".github/copilot-instructions.md only.",
                }
            ],
        ),
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 0, result.stderr
    assert "Protected file decision" in result.stdout
    assert "decision=MERGE; adoption_mode=minimal-preservation" in result.stdout
    assert "Protected file decision record present in marker." in result.stdout


def test_ledger_reports_protected_decision_modes_and_remove_local_section(
    tmp_path: Path,
) -> None:
    """The adoption ledger distinguishes tailored decisions and removals."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            last_reviewed_template_commit=None,
            protected_decisions=[
                {
                    "path": ".github/copilot-instructions.md",
                    "decision": "TAKE",
                    "adoption_mode": "tailored",
                    "authorization_basis": "Owner authorized protected edits on 2026-05-27.",
                    "authorized_scope": ".github/copilot-instructions.md only.",
                    "tailored_authorization_basis": "Owner authorized a tailored rewrite.",
                },
                {
                    "path": "GEMINI.md",
                    "decision": "REMOVE-LOCAL",
                    "authorization_basis": "Owner explicitly authorized removing GEMINI.md.",
                    "authorized_scope": "GEMINI.md only.",
                    "reason": "Gemini agent not used by this repository.",
                },
            ],
        ),
    )

    result = _run_generator(tmp_path, "--ledger-only")

    assert result.returncode == 0, result.stderr
    assert (
        "| .github/copilot-instructions.md | all: agent-instructions | protected decision: TAKE |"
        in (result.stdout)
    )
    assert "tailored (authorized protected-file decision)" in result.stdout
    assert "tailored_authorization_basis: Owner authorized a tailored rewrite." in result.stdout
    assert "| GEMINI.md | unmapped | protected decision: REMOVE-LOCAL |" in result.stdout
    assert "## REMOVE-LOCAL Authorizations" in result.stdout
    assert "authorization_basis: Owner explicitly authorized removing GEMINI.md." in result.stdout
    assert "authorized_scope: GEMINI.md only." in result.stdout
    assert "reason: Gemini agent not used by this repository." in result.stdout


def test_ledger_reports_protected_decision_overlap_side_by_side(tmp_path: Path) -> None:
    """Ledger protected decision context shows compatible local override overlap."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            last_reviewed_template_commit=None,
            local_overrides=[
                {
                    "path": ".github/copilot-instructions.md",
                    "reason": "Use minimal-preservation with a local policy note.",
                    "default_decision": "MERGE",
                }
            ],
            protected_decisions=[
                {
                    "path": ".github/copilot-instructions.md",
                    "decision": "MERGE",
                    "adoption_mode": "minimal-preservation",
                    "authorization_basis": "Owner authorized protected edits on 2026-05-27.",
                    "authorized_scope": ".github/copilot-instructions.md only.",
                }
            ],
        ),
    )

    result = _run_generator(tmp_path, "--ledger-only")

    assert result.returncode == 0, result.stderr
    assert "## Protected File Decision Records" in result.stdout
    assert (
        "protected_file_decisions: decision=MERGE; adoption_mode=minimal-preservation"
        in result.stdout
    )
    assert (
        "local_overrides: path=.github/copilot-instructions.md; default_decision=MERGE; "
        "reason=Use minimal-preservation with a local policy note."
    ) in result.stdout


def test_authorized_protected_decision_does_not_require_maintainer_decision(
    tmp_path: Path,
) -> None:
    """Authorized TAKE/MERGE protected decisions do not re-flag maintainer review."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            last_reviewed_template_commit=None,
            protected_decisions=[
                {
                    "path": ".github/copilot-instructions.md",
                    "decision": "MERGE",
                    "adoption_mode": "minimal-preservation",
                    "authorization_basis": "Owner authorized protected edits on 2026-05-27.",
                    "authorized_scope": ".github/copilot-instructions.md only.",
                }
            ],
        ),
    )

    result = _run_generator(tmp_path, "--ledger-only")

    assert result.returncode == 0, result.stderr
    assert (
        "| .github/copilot-instructions.md | all: agent-instructions "
        "| protected decision: MERGE | Authorized minimal-preservation protected-file edit."
    ) in result.stdout
    assert (
        "authorization_basis: Owner authorized protected edits on 2026-05-27. "
        "authorized_scope: .github/copilot-instructions.md only. | Yes | No |"
    ) in result.stdout


def test_unmatched_protected_decision_for_non_protected_path_reports_no(
    tmp_path: Path,
) -> None:
    """Non-protected paths in protected_file_decisions report protected_file=No."""
    _init_repo(tmp_path)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            last_reviewed_template_commit=None,
            protected_decisions=[
                {
                    "path": "docs/non-protected-note.md",
                    "decision": "SKIP",
                    "reason": "Owner has decided not to adopt this file.",
                }
            ],
        ),
    )

    result = _run_generator(tmp_path, "--ledger-only")

    assert result.returncode == 0, result.stderr
    assert (
        "| docs/non-protected-note.md | unmapped | protected decision: SKIP "
        "| Protected decision `SKIP`: Owner has decided not to adopt this file. | No |"
    ) in result.stdout


def test_renamed_files_preserve_old_and_new_paths(tmp_path: Path) -> None:
    """Rename rows keep both paths and identify the rename status."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "templates/markdown/intro.md", "same content\n")
    base_sha = _commit_all(tmp_path, "base")
    _git(tmp_path, "mv", "templates/markdown/intro.md", "templates/markdown/getting-started.md")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["markdown"], last_reviewed_template_commit=base_sha),
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 0, result.stderr
    assert "templates/markdown/intro.md -> templates/markdown/getting-started.md" in result.stdout
    assert "Renamed (R100)" in result.stdout
    assert "Renamed from templates/markdown/intro.md." in result.stdout


def test_rename_across_module_boundaries_is_flagged(tmp_path: Path) -> None:
    """A rename whose old and new paths map to different modules is surfaced."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "templates/markdown/intro.md", "same content\n")
    base_sha = _commit_all(tmp_path, "base")
    (tmp_path / "templates/markdown/intro.md").unlink()
    _write_text(tmp_path, "templates/python/intro.md", "same content\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["markdown", "python"], last_reviewed_template_commit=base_sha),
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 0, result.stderr
    assert "templates/markdown/intro.md -> templates/python/intro.md" in result.stdout
    assert "Renamed from templates/markdown/intro.md." in result.stdout
    assert (
        "Rename crosses module mapping boundaries: old path resolves to "
        "requires all: markdown; new path resolves to requires all: python." in result.stdout
    )


def test_invalid_marker_schema_is_rejected(tmp_path: Path) -> None:
    """Marker and manifest files must pass the checked-in schemas before output."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "README.md", "base\n")
    base_sha = _commit_all(tmp_path, "base")
    _write_text(tmp_path, "README.md", "head\n")
    head_sha = _commit_all(tmp_path, "head")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        {
            "template_sync": {
                "source_repo": SOURCE_REPO,
                "last_reviewed_template_commit": base_sha,
                "included_modules": ["not-a-real-module"],
            }
        },
    )

    result = _run_generator(tmp_path, "--range-head", head_sha)

    assert result.returncode == 1
    assert "Schema validation failed for .template-sync/marker.yml" in result.stderr


@pytest.mark.parametrize(
    "agent_module",
    [
        "agent-copilot",
        "agent-codex",
        "agent-claude",
        "agent-cursor",
        "agent-gemini",
        "agent-hermes",
    ],
)
def test_agent_selection_guidance_does_not_require_optional_enforcement(agent_module: str) -> None:
    """Policy-only agents require review without silently selecting optional tooling."""
    modules = frozenset({"agent-instructions", agent_module})
    guidance = sync_candidates.validation_commands_for_modules(modules)
    assert "manual" in guidance
    for unrelated in ("python ", "pytest ", "npm ", "pre-commit ", "validate_instruction_profile"):
        assert unrelated not in guidance
    assert (
        sync_candidates.adoption_mode_for_modules(
            frozenset({agent_module}), (), False, "minimal-preservation"
        )
        == "minimal-preservation"
    )


def test_enforcement_guidance_uses_portable_profile_dispatcher() -> None:
    """Selected enforcement names its portable CLI independently of project/host modules."""
    guidance = sync_candidates.validation_commands_for_modules(
        frozenset({"instruction-enforcement"})
    )
    assert "python .github/scripts/validate_instruction_profile.py" in guidance
    assert "selected modules" in guidance
    assert "local exception declarations" in guidance
    for unrelated in (".template-sync/", "pytest ", "npm ", "pre-commit "):
        assert unrelated not in guidance
    assert (
        sync_candidates.adoption_mode_for_modules(
            frozenset({"instruction-enforcement"}), (), False, "tailored"
        )
        == "tailored"
    )
