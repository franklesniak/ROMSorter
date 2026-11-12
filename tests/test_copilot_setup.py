"""Verify selected Copilot prerequisites, failure truth, and reviewed cleanup."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from typing import Any

from tests import test_materialize_downstream_adoption as lifecycle
from tests._pytest_compat import pytest
from tests.test_workflow_security_contract import ROOT, policy
from tests.test_workflow_security_lifecycle import snapshot, write_decisions

pytestmark = pytest.mark.upstream_template_only
WORKFLOW = ".github/workflows/copilot-setup-steps.yml"
GUIDE = "docs/copilot-setup.md"
COPILOT = ("agent-instructions", "agent-copilot", "github-actions")


def setup_steps(root: Path) -> list[dict[str, Any]]:
    """Read the supported single setup job from the selected repository."""
    document = policy.parse_yaml((root / WORKFLOW).read_text(encoding="utf-8"))
    assert set(document["jobs"]) == {"copilot-setup-steps"}
    job = document["jobs"]["copilot-setup-steps"]
    assert set(job) == {"permissions", "runs-on", "timeout-minutes", "steps"}
    assert job["permissions"] == {"contents": "read"}
    assert 0 < job["timeout-minutes"] <= 59
    assert all("continue-on-error" not in step for step in job["steps"])
    steps = job["steps"]
    assert isinstance(steps, list)
    return steps


@pytest.mark.parametrize(
    "modules",
    [
        COPILOT,
        (*COPILOT, "baseline", "markdown"),
        (*COPILOT, "python"),
        (*COPILOT, "terraform", "powershell"),
        (*COPILOT, "azure-pipelines"),
        (*COPILOT, "baseline", "template-sync-support"),
        ("agent-instructions", "github-actions", "baseline"),
        ("agent-instructions", "agent-codex", "github-actions", "baseline"),
        ("agent-instructions", "agent-copilot", "azure-pipelines"),
        ("github-actions",),
    ],
)
def test_setup_follows_actual_adoption(tmp_path: Path, modules: tuple[str, ...]) -> None:
    """Actual profiles retain only their prerequisites, even without Python scaffolding."""
    target = lifecycle.materialize_module_fixture(tmp_path, modules, authorize_protected_files=True)
    retained = set(COPILOT).issubset(modules)
    reference = lifecycle.materializer.remove_inline_blocks_for_modules(
        "<!-- template-sync: begin copilot-setup-reference-only -->\n"
        "Copilot setup reference\n"
        "<!-- template-sync: end copilot-setup-reference-only -->\n",
        modules,
        relative_path="GETTING_STARTED_NEW_REPO.md",
    )
    assert ("Copilot setup reference" in reference) is retained
    assert (target / WORKFLOW).is_file() is retained
    assert (target / GUIDE).is_file() is retained
    if "github-actions" in modules:
        assert policy.validate_repository(target) >= 1
        assert (WORKFLOW in policy.load_contract(target)["workflows"]) is retained
    if not retained:
        return
    steps = {step["name"]: step for step in setup_steps(target)}
    expected = {
        "Setup Python for retained validation tools": True,
        "Install shared validator dependencies": True,
        "Install pre-commit": "baseline" in modules,
        "Install Python project validation dependencies": "python" in modules,
        "Setup Node.js for Markdown": "markdown" in modules,
        "Install root Markdown dependencies": "markdown" in modules,
        "Setup Terraform": "terraform" in modules,
        "Setup TFLint": "terraform" in modules,
        "Install PowerShell validation modules": "powershell" in modules,
    }
    for name, present in expected.items():
        assert (name in steps) is present, name
    if "markdown" in modules:
        assert steps["Install root Markdown dependencies"]["run"] == "npm ci --ignore-scripts"
        assert (target / "package-lock.json").is_file()
    assert (target / "pyproject.toml").is_file() is ("python" in modules)
    before = snapshot(target)
    repeated = lifecycle.run_materialize(
        ROOT,
        target,
        "--decisions-file",
        "decisions.yml",
        *lifecycle.azure_provider_cli_args_for_modules(modules),
    )
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert snapshot(target) == before


@pytest.mark.parametrize("removed", ["agent-copilot", "github-actions"])
def test_setup_removal_requires_reviewed_cleanup(tmp_path: Path, removed: str) -> None:
    """Selection reports existing files; reviewed removal and repeat leave no setup assets."""
    target = lifecycle.materialize_module_fixture(tmp_path, COPILOT, authorize_protected_files=True)
    modules = tuple(module for module in COPILOT if module != removed)
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    removals = [
        relative
        for relative in snapshot(target)
        if (relation := lifecycle.materializer.selected_relation_for_path(relative, mappings))
        is not None
        and not relation.is_retained_by(modules)
    ]
    write_decisions(
        target,
        modules,
        protected_file_decisions=lifecycle.protected_take_decisions_for_modules(modules),
        local_overrides=(
            [
                {
                    "path": "docs/PR_REVIEW_PROMPTS.md",
                    "default_decision": "TAKE",
                    "reason": "Fixture owner approves removal of GitHub review instructions.",
                }
            ]
            if removed == "github-actions"
            else []
        ),
    )
    result = lifecycle.run_materialize(ROOT, target, "--decisions-file", "decisions.yml")
    assert result.returncode == 0, result.stdout + result.stderr
    assert (target / WORKFLOW).is_file()
    assert (target / GUIDE).is_file()
    assert WORKFLOW in result.stdout
    assert GUIDE in result.stdout
    # This scratch tree belongs to the fixture. Model review of all excluded paths,
    # including the old host contract when the host itself is removed.
    assert {WORKFLOW, GUIDE}.issubset(removals)
    for relative in removals:
        (target / relative).unlink()
    repeated = lifecycle.run_materialize(ROOT, target, "--decisions-file", "decisions.yml")
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert not (target / WORKFLOW).exists()
    assert not (target / GUIDE).exists()
    if "github-actions" in modules:
        assert WORKFLOW not in policy.load_contract(target)["workflows"]
        assert policy.validate_repository(target) >= 1

    clean_parent = tmp_path / "clean"
    clean_parent.mkdir()
    clean = lifecycle.materialize_module_fixture(
        clean_parent, modules, authorize_protected_files=True
    )
    ignored = {"decisions.yml", ".template-sync/marker.yml"}
    assert {p: b for p, b in snapshot(target).items() if p not in ignored} == {
        p: b for p, b in snapshot(clean).items() if p not in ignored
    }


def test_setup_uses_existing_toolchain_sources() -> None:
    """Existing CI remains the source for shared requirements and maintained selectors."""
    steps = {step["name"]: step for step in setup_steps(ROOT)}
    security = policy.parse_yaml(
        (ROOT / ".github/workflows/workflow-security.yml").read_text(encoding="utf-8")
    )
    install = next(
        step
        for step in security["jobs"]["validate"]["steps"]
        if step["name"] == "Install validator dependencies"
    )
    assert steps["Install shared validator dependencies"]["run"] == install["run"]
    precommit = policy.parse_yaml(
        (ROOT / ".github/workflows/precommit-ci.yml").read_text(encoding="utf-8")
    )
    existing = {step["name"]: step for step in precommit["jobs"]["pre-commit"]["steps"]}
    assert steps["Install pre-commit"]["run"] == existing["Install pre-commit"]["run"]
    assert steps["Setup Python for retained validation tools"]["with"]["python-version"] == (
        existing["Setup Python"]["with"]["python-version"]
    )
    assert steps["Setup Node.js for Markdown"]["with"]["node-version"] == (
        existing["Setup Node.js for nested Markdown"]["with"]["node-version"]
    )
    for name in ("Setup Terraform", "Setup TFLint"):
        assert steps[name]["with"] == existing[name]["with"]


def installer_script() -> str:
    """Return the actual pre-commit installation step, not a copied test implementation."""
    script = next(step["run"] for step in setup_steps(ROOT) if step["name"] == "Install pre-commit")
    assert isinstance(script, str)
    return script


def run_installer(tmp_path: Path, script: str, mode: str) -> subprocess.CompletedProcess[str]:
    """Execute installation control flow with deterministic external-process outcomes."""
    (tmp_path / "setup.py").write_text(script, encoding="utf-8")
    (tmp_path / "subprocess.py").write_text(
        "import os\n"
        "def run(command, check=False):\n"
        "    assert command[1:] == ['-m', 'pip', 'install', '--requirement', 'requirements-pre-commit.txt']\n"
        "    if os.environ['SETUP_TEST_MODE'] == 'pip-failure' and check:\n"
        "        raise RuntimeError('native pip failure')\n"
        "def check_output(command, text=False):\n"
        "    assert command == ['pre-commit', '--version']\n"
        "    if os.environ['SETUP_TEST_MODE'] == 'version-failure':\n"
        "        raise RuntimeError('native version failure')\n"
        "    if os.environ['SETUP_TEST_MODE'] == 'mismatch':\n"
        "        return 'pre-commit 0.0.0'\n"
        "    return 'pre-commit 9.8.7'\n",
        encoding="utf-8",
    )
    environment = dict(os.environ, SETUP_TEST_MODE=mode)
    return subprocess.run(
        [sys.executable, "setup.py"],
        cwd=tmp_path,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )


@pytest.mark.parametrize(
    ("requirement", "mode", "expected"),
    [
        ("pre-commit==9.8.7", "success", 0),
        ("pre-commit>=9", "success", 1),
        (None, "success", 1),
        ("pre-commit==9.8.7", "pip-failure", 1),
        ("pre-commit==9.8.7", "version-failure", 1),
        ("pre-commit==9.8.7", "mismatch", 1),
    ],
)
def test_setup_preserves_installer_failures(
    tmp_path: Path, requirement: str | None, mode: str, expected: int
) -> None:
    """Malformed/missing pins, installation, and verification errors remain native failures."""
    if requirement is not None:
        (tmp_path / "requirements-pre-commit.txt").write_text(requirement, encoding="utf-8")
    result = run_installer(tmp_path, installer_script(), mode)
    assert result.returncode == expected, result.stdout + result.stderr


def test_installer_failure_guard_mutant_is_detectable(tmp_path: Path) -> None:
    """Removing check=True demonstrably changes the independent failure expectation."""
    (tmp_path / "requirements-pre-commit.txt").write_text("pre-commit==9.8.7", encoding="utf-8")
    script = installer_script()
    assert script.count("check=True") == 1
    assert run_installer(tmp_path, script, "pip-failure").returncode == 1
    assert (
        run_installer(
            tmp_path, script.replace("check=True", "check=False"), "pip-failure"
        ).returncode
        == 0
    )


@pytest.mark.parametrize(
    ("before", "after"),
    [
        ("persist-credentials: false", "persist-credentials: true"),
        ("contents: read", "contents: write"),
        ("timeout-minutes: 30", "timeout-minutes: 60"),
        ("check=True", "check=False"),
        ("npm ci --ignore-scripts", "npm ci"),
    ],
)
def test_setup_contract_rejects_weakened_controls(tmp_path: Path, before: str, after: str) -> None:
    """Real candidate validation rejects altered setup controls and failure handling."""
    from tests.test_workflow_security_contract import copy_policy

    copy_policy(tmp_path)
    path = tmp_path / WORKFLOW
    text = path.read_text(encoding="utf-8")
    assert before in text
    path.write_text(text.replace(before, after), encoding="utf-8")
    assert policy.main(["--repo-root", str(tmp_path)]) == 1
