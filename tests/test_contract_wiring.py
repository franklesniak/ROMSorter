"""Prove instruction and workflow contracts remain connected to blocking gates."""

from __future__ import annotations

import shlex
import subprocess
import sys
from pathlib import Path
from typing import Any, cast

import yaml  # type: ignore[import-untyped]

from tests import test_validate_instruction_contracts as fixtures
from tests import test_validate_marker as marker_fixtures
from tests._pytest_compat import pytest

pytestmark = pytest.mark.upstream_template_only
ROOT = Path(__file__).resolve().parents[1]
HOOKS = {
    "validate-instruction-contracts-upstream": [
        "python",
        ".template-sync/scripts/validate_instruction_contracts.py",
        "--mode",
        "upstream-template",
        "--skip-if-marker-present",
    ],
    "validate-instruction-contracts-downstream": [
        "python",
        ".template-sync/scripts/validate_instruction_contracts.py",
        "--mode",
        "downstream",
    ],
    # template-sync: begin github-actions-only
    "validate-workflow-security": ["python", ".github/scripts/validate_workflow_security.py"],
    # template-sync: end github-actions-only
}
REQUIRED_SECTIONS = {
    "### Immutable action pins and release comments",
    "## Agent Instruction Files",
    "## Agent Execution",
    "### Ownership and delegation",
    "### Continuity and recovery",
    "## Shared Review Governance",
    "### Finding inventory and decisions",
    "### Selected-action writing rule",
    "### Safe PR-head placement",
    "### Protected authority and deferral",
    "### Review inputs and attribution",
    "### Review recovery decisions",
    "### Polling and continuation",
    "### CI diagnosis and completion",
    "## Host-Specific PR Review Protocols",
    "### Azure DevOps Services with Azure Repos",
}


def read_yaml(relative: str) -> dict[str, Any]:
    """Read structured non-workflow fields used by independent wiring assertions."""
    return cast(dict[str, Any], yaml.safe_load((ROOT / relative).read_text(encoding="utf-8")))


def assert_hook(config: dict[str, Any], key: str) -> dict[str, Any]:
    """Check the actual always-running local entry, including bypass settings."""
    matches = [
        (repo, hook) for repo in config["repos"] for hook in repo["hooks"] if hook["id"] == key
    ]
    assert len(matches) == 1
    repo, hook = matches[0]
    assert repo["repo"] == "local"
    assert shlex.split(hook["entry"]) == HOOKS[key]
    assert hook["language"] == "python"
    assert hook.get("always_run") is True
    assert hook.get("pass_filenames") is False
    for field in ("files", "exclude", "stages", "args", "types", "types_or", "exclude_types"):
        assert field not in hook
    assert "default_stages" not in config
    return cast(dict[str, Any], hook)


def assert_ci(document: dict[str, Any], key: str, azure: bool = False) -> None:
    """Require an unmasked direct invocation from a non-advisory job and step."""
    if azure:
        steps = document["steps"]
    else:
        assert len(document["jobs"]) == 1
        job = next(iter(document["jobs"].values()))
        assert "if" not in job and "continue-on-error" not in job
        steps = job["steps"]
    command = f"pre-commit run {key} --all-files"
    matches = [
        step for step in steps if step.get("displayName" if azure else "name") == f"Run {key}"
    ]
    assert len(matches) == 1
    step = matches[0]
    assert step.get("bash" if azure else "run") == command
    for field in ("if", "condition", "continue-on-error", "continueOnError", "env"):
        assert field not in step


@pytest.mark.parametrize("key", list(HOOKS))
def test_authoritative_local_and_ci_routes(key: str) -> None:
    """Every applicable contract has a blocking pre-commit and Data CI route."""
    assert_hook(read_yaml(".pre-commit-config.yaml"), key)
    assert_ci(read_yaml(".github/workflows/data-ci.yml"), key)
    if key.startswith("validate-instruction-contracts-"):
        assert_ci(read_yaml(".azuredevops/pipelines/data-ci.yml"), key, azure=True)


@pytest.mark.parametrize("key", list(HOOKS))
@pytest.mark.parametrize(
    "mutation", ["delete", "rename", "mode", "always", "manual", "files", "args"]
)
def test_local_hook_bypass_mutations(key: str, mutation: str) -> None:
    """Removal, renaming, and dispatch bypasses are killed by the wiring oracle."""
    config = read_yaml(".pre-commit-config.yaml")
    hook = assert_hook(config, key)
    if mutation == "delete":
        for repo in config["repos"]:
            repo["hooks"] = [item for item in repo["hooks"] if item is not hook]
    elif mutation == "rename":
        hook["id"] += "-disabled"
    elif mutation == "mode":
        hook["entry"] = "python -c pass"
    else:
        changes: dict[str, dict[str, Any]] = {
            "always": {"always_run": False},
            "manual": {"stages": ["manual"]},
            "files": {"files": "^$"},
            "args": {"args": ["--help"]},
        }
        hook.update(changes[mutation])
    with pytest.raises(AssertionError):
        assert_hook(config, key)


@pytest.mark.parametrize("key", list(HOOKS))
@pytest.mark.parametrize("mutation", ["delete", "rename", "advisory", "skip", "mask", "echo"])
@pytest.mark.parametrize("azure", [False, True])
def test_ci_bypass_mutations(key: str, mutation: str, azure: bool) -> None:
    """CI cannot hide a hook's native failure through a conditional or shell suffix."""
    if azure and not key.startswith("validate-instruction-contracts-"):
        return
    doc = read_yaml(
        ".azuredevops/pipelines/data-ci.yml" if azure else ".github/workflows/data-ci.yml"
    )
    steps = doc["steps"] if azure else next(iter(doc["jobs"].values()))["steps"]
    name = "displayName" if azure else "name"
    command = "bash" if azure else "run"
    step = next(item for item in steps if item.get(name) == f"Run {key}")
    if mutation == "delete":
        steps.remove(step)
    elif mutation == "rename":
        step[command] = step[command].replace(key, key + "-disabled")
    elif mutation == "advisory":
        step["continueOnError" if azure else "continue-on-error"] = True
    elif mutation == "skip":
        step["condition" if azure else "if"] = False
    elif mutation == "mask":
        step[command] += " || true"
    else:
        step[command] = "echo passing"
    with pytest.raises(AssertionError):
        assert_ci(doc, key, azure)


@pytest.mark.parametrize("mode", ["upstream", "downstream"])
def test_configured_instruction_entry_propagates_failure(tmp_path: Path, mode: str) -> None:
    """Execute the configured command, with a real missing-clause defect and native exit."""
    fixtures._write_common_contract_repo(
        tmp_path, fixtures._contracts(required_headings=["## Required"])
    )
    fixtures._write_text(tmp_path, "CLAUDE.md", "# Agent\n")
    if mode == "downstream":
        fixtures._write_yaml(
            tmp_path, ".template-sync/marker.yml", fixtures._marker(["agent-instructions"])
        )
    entry = shlex.split(
        assert_hook(read_yaml(".pre-commit-config.yaml"), f"validate-instruction-contracts-{mode}")[
            "entry"
        ]
    )
    result = subprocess.run(
        [sys.executable, str(ROOT / entry[1]), *entry[2:], "--repo-root", str(tmp_path)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Required" in result.stdout


def assert_governance_inventory(catalog: dict[str, Any]) -> None:
    """An independent semantic inventory prevents deleting the catalog's obligations."""
    canonical = next(
        item
        for item in catalog["instruction_contracts"]
        if item["path"] == ".github/copilot-instructions.md"
    )
    assert REQUIRED_SECTIONS <= {section["heading"] for section in canonical["required_sections"]}
    for contract in catalog["instruction_contracts"]:
        assert "agent-instructions" in contract["requires_modules"]


@pytest.mark.parametrize("heading", sorted(REQUIRED_SECTIONS))
def test_catalog_obligation_removal_is_detected(heading: str) -> None:
    """Catalog edits cannot silently erase a mandatory P1 governance obligation."""
    catalog = read_yaml(".template-sync/instruction-contracts.yml")
    assert_governance_inventory(catalog)
    canonical = next(
        item
        for item in catalog["instruction_contracts"]
        if item["path"] == ".github/copilot-instructions.md"
    )
    canonical["required_sections"] = [
        section for section in canonical["required_sections"] if section["heading"] != heading
    ]
    with pytest.raises(AssertionError):
        assert_governance_inventory(catalog)


@pytest.mark.parametrize("field", ["authorization_basis", "anchor", "path"])
def test_waiver_misuse_fails_real_validator(tmp_path: Path, field: str) -> None:
    """Missing authorization and wrongly scoped waivers cannot hide a required clause."""
    fixtures._write_common_contract_repo(
        tmp_path, fixtures._contracts(required_headings=["## Required"])
    )
    fixtures._write_text(tmp_path, "CLAUDE.md", "# Agent\n")
    waiver = {
        "path": "CLAUDE.md",
        "anchor": "## Required",
        "reason": "Alternative protocol",
        "authorization_basis": "Owner explicitly authorized this fixture.",
    }
    waiver[field] = "" if field == "authorization_basis" else "unrelated"
    fixtures._write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        fixtures._marker(["agent-instructions"], waivers=[waiver]),
    )
    assert fixtures._run_validator(tmp_path, "--mode", "downstream").returncode == 1


def test_module_ownership_mutation_is_detected() -> None:
    """A valid but unrelated module cannot suppress the canonical governance catalog."""
    catalog = read_yaml(".template-sync/instruction-contracts.yml")
    assert_governance_inventory(catalog)
    catalog["instruction_contracts"][0]["requires_modules"] = ["terraform"]
    with pytest.raises(AssertionError):
        assert_governance_inventory(catalog)


@pytest.mark.parametrize("retained", [True, False])
def test_marker_module_mismatch_fails_real_cli(tmp_path: Path, retained: bool) -> None:
    """A retained missing control and an omitted leftover both fail marker validation."""
    marker_fixtures._run_git(tmp_path, "init", "-q")
    marker_fixtures._copy_schemas(tmp_path)
    manifest = marker_fixtures._manifest()
    manifest["template_manifest"]["path_mappings"].append(
        {
            "pattern": ".template-sync/instruction-contracts.yml",
            "requires_all": ["agent-instructions"],
        }
    )
    marker_fixtures._write_yaml(tmp_path, ".template-sync/manifest.yml", manifest)
    modules = ["baseline", "template-sync-support"] + (["agent-instructions"] if retained else [])
    marker_fixtures._write_marker(tmp_path, modules)
    marker_fixtures._write_text(tmp_path, "README.md")
    marker_fixtures._write_text(tmp_path, ".template-sync/scripts/validate_marker.py")
    if not retained:
        marker_fixtures._write_text(
            tmp_path, ".template-sync/instruction-contracts.yml", "version: 1\n"
        )
    result = marker_fixtures._run_validator(tmp_path)
    assert result.returncode == 1, result.stdout + result.stderr
    assert ".template-sync/instruction-contracts.yml" in result.stdout


@pytest.mark.parametrize("retained", [False, True])
def test_excluded_parent_contract_does_not_require_missing_guide(
    tmp_path: Path, retained: bool
) -> None:
    """A guide follows its parent module, while retained missing guides still fail."""
    catalog = fixtures._contracts(required_headings=["## Required"])
    catalog["protected_guide_section_obligations"] = [
        {
            "key": "claude-azure-section",
            "path": "CLAUDE.md",
            "target_modules": ["azure-devops-collaboration"],
            "stale_headings": ["## Azure"],
        }
    ]
    fixtures._write_common_contract_repo(tmp_path, catalog)
    modules = ["agent-instructions"] if retained else ["template-sync-support"]
    fixtures._write_yaml(tmp_path, ".template-sync/marker.yml", fixtures._marker(modules))
    result = fixtures._run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == (1 if retained else 0), result.stdout + result.stderr
