"""Verify the baseline-owned pre-commit runner consumers."""

from __future__ import annotations

import inspect
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import textwrap
from collections.abc import Callable
from pathlib import Path
from typing import Any, Literal

import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest

pytestmark = pytest.mark.upstream_template_only

REPO_ROOT = Path(__file__).resolve().parent.parent
RUNNER_REQUIREMENT_PATH = REPO_ROOT / "requirements-pre-commit.txt"
MARKDOWN_WORKFLOW_PATH = REPO_ROOT / ".github/workflows/markdownlint.yml"
UPSTREAM_TEMPLATE_CONDITION = "${{ github.repository == 'franklesniak/copilot-repo-template' }}"
PYTHON_INSTALLER_PATHS = (
    ".github/workflows/precommit-ci.yml",
    ".github/workflows/data-ci.yml",
    ".github/workflows/auto-fix-precommit.yml",
    ".azuredevops/pipelines/precommit.yml",
    ".azuredevops/pipelines/data-ci.yml",
)
GITHUB_CACHE_WORKFLOWS = (
    ".github/workflows/precommit-ci.yml",
    ".github/workflows/data-ci.yml",
)
CREDENTIAL_FREE_WORKFLOWS = PYTHON_INSTALLER_PATHS[:3] + (
    ".github/workflows/check-placeholders.yml",
    ".github/workflows/markdownlint.yml",
    ".github/workflows/powershell-ci.yml",
    ".github/workflows/python-ci.yml",
    ".github/workflows/terraform-ci.yml",
    ".github/workflows/toolchain-eol.yml",
)
HOOK_PATH = REPO_ROOT / ".claude/hooks/session-start.sh"
EXACT_REQUIREMENT_RE = re.compile(r"pre-commit==(?P<version>[0-9]+\.[0-9]+\.[0-9]+)\Z")


def runner_requirement() -> str:
    """Return the repository pin after independently validating its public shape."""
    requirement = RUNNER_REQUIREMENT_PATH.read_text(encoding="utf-8").strip()
    assert EXACT_REQUIREMENT_RE.fullmatch(requirement)
    return requirement


def assert_immutable_action_references(workflow: str) -> None:
    """Require full upstream-style references and same-line release annotations."""
    uses_lines = [
        line.strip() for line in workflow.splitlines() if line.strip().startswith("uses:")
    ]
    assert uses_lines
    for line in uses_lines:
        assert re.fullmatch(
            r"uses: [A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+@[0-9a-f]{40} # v[0-9]+\.[0-9]+\.[0-9]+",
            line,
        ), line


def assert_credential_free_checkouts(workflow_text: str) -> None:
    """Inspect every checkout and its effective permissions, including later steps."""
    workflow = yaml.safe_load(workflow_text)
    checkout_count = 0
    for job in workflow["jobs"].values():
        for step in job.get("steps", []):
            if not step.get("uses", "").startswith("actions/checkout@"):
                continue
            checkout_count += 1
            assert job.get("permissions", workflow.get("permissions")) == {"contents": "read"}
            inputs = step.get("with", {})
            assert isinstance(inputs, dict), "checkout inputs must be a mapping"
            value = inputs.get("persist-credentials")
            assert value is False or (
                isinstance(value, str) and value.strip(" \t\r\n").lower() == "false"
            )
    assert checkout_count, "the fixture must exercise checkout"


def assert_markdown_hook_regression_ci(workflow_text: str) -> None:
    """Require provisioned, blocking hook coverage and optional-profile conditions."""
    workflow = yaml.safe_load(workflow_text)
    steps = workflow["jobs"]["markdownlint"]["steps"]
    by_name = {step["name"]: step for step in steps}
    assert len(by_name) == len(steps)
    assert workflow_text.count("# template-sync: begin template-sync-support-only") == 2
    assert workflow_text.count("# template-sync: end template-sync-support-only") == 2
    assert all("template-sync:" not in str(step.get("if", "")) for step in steps)

    setup = by_name["Setup Python"]
    install = by_name["Install Python dependencies"]
    for step in (setup, install):
        condition = " ".join(step["if"].split())
        assert "github.repository == 'franklesniak/copilot-repo-template'" in condition
        assert "hashFiles('pyproject.toml') != ''" in condition
        assert "hashFiles('tests/test_materialize_downstream_adoption.py') != ''" in condition

    runner = by_name["Install pinned pre-commit runner"]
    assert runner["if"] == UPSTREAM_TEMPLATE_CONDITION
    assert runner["run"] == ("python -m pip install --requirement requirements-pre-commit.txt")

    generated = by_name["Lint materialized generated output (nested Markdown)"]
    generated_condition = " ".join(generated["if"].split())
    assert "hashFiles('pyproject.toml') != ''" in generated_condition
    assert "hashFiles('tests/test_materialize_downstream_adoption.py') != ''" in (
        generated_condition
    )
    assert generated["continue-on-error"] is True

    hook = by_name["Run actual nested Markdown hook regression"]
    assert hook["if"] == UPSTREAM_TEMPLATE_CONDITION
    assert hook["id"] == "test-nested-hook"
    assert hook["continue-on-error"] is True
    assert hook["run"].split() == [
        "pytest",
        (
            "tests/test_precommit_runner.py::"
            "test_nested_markdown_hook_actual_invocation_and_mutation"
        ),
        "-v",
    ]

    result = by_name["Check source-template regression results"]
    assert "steps.lint-generated.outcome == 'failure'" in result["if"]
    assert "steps.test-nested-hook.outcome == 'failure'" in result["if"]
    assert "exit 1" in result["run"]


def assert_candidate_hook_checkout(workflow_text: str) -> None:
    """Require credential-free checkout before actual candidate pre-commit steps."""
    assert_credential_free_checkouts(workflow_text)
    workflow = yaml.safe_load(workflow_text)
    assert workflow.get("permissions") == {"contents": "read"}
    hook_jobs = 0
    for job in workflow["jobs"].values():
        steps = job.get("steps", [])
        hook_indexes = [
            index
            for index, step in enumerate(steps)
            if re.search(r"\bpre-commit\s+run\b", step.get("run", ""))
        ]
        if not hook_indexes:
            continue
        hook_jobs += 1
        assert job.get("permissions") == {"contents": "read"}
        for hook_index in hook_indexes:
            checkouts = [
                step
                for step in steps[:hook_index]
                if step.get("uses", "").startswith("actions/checkout@")
            ]
            assert checkouts, "candidate hooks require a preceding checkout"
    assert hook_jobs, "the fixture must exercise an actual candidate-hook job"


@pytest.mark.parametrize("relative_path", PYTHON_INSTALLER_PATHS[:3])
def test_candidate_hook_workflows_do_not_persist_checkout_credentials(relative_path: str) -> None:
    """All three real hook entry points retain read-only, credential-free checkout."""
    assert_candidate_hook_checkout((REPO_ROOT / relative_path).read_text(encoding="utf-8"))


@pytest.mark.parametrize("relative_path", PYTHON_INSTALLER_PATHS[:3])
@pytest.mark.parametrize("mutation", ["omitted", "true", "dynamic", "write-permission"])
def test_candidate_hook_workflow_rejects_weakened_checkout(
    relative_path: str, mutation: str
) -> None:
    """Mutations of actual workflows must fail the independent structural oracle."""
    workflow = yaml.safe_load((REPO_ROOT / relative_path).read_text(encoding="utf-8"))
    checkout = next(
        step
        for job in workflow["jobs"].values()
        for step in job["steps"]
        if step.get("uses", "").startswith("actions/checkout@")
    )
    if mutation == "omitted":
        del checkout["with"]["persist-credentials"]
    elif mutation == "true":
        checkout["with"]["persist-credentials"] = True
    elif mutation == "dynamic":
        checkout["with"]["persist-credentials"] = "${{ inputs.persist }}"
    else:
        workflow["permissions"]["contents"] = "write"
    with pytest.raises(AssertionError):
        assert_candidate_hook_checkout(yaml.safe_dump(workflow))


@pytest.mark.parametrize("relative_path", PYTHON_INSTALLER_PATHS[:3])
def test_review_support_actions_are_immutable_and_release_annotated(relative_path: str) -> None:
    """Each changed GitHub workflow keeps direct immutable, updateable references."""
    workflow = (REPO_ROOT / relative_path).read_text(encoding="utf-8")
    assert_immutable_action_references(workflow)
    without_pin = re.sub(r"@[0-9a-f]{40}", "@v7", workflow, count=1)
    without_release = re.sub(r" # v[0-9]+\.[0-9]+\.[0-9]+", "", workflow, count=1)
    assert without_pin != workflow
    assert without_release != workflow
    for mutant in (without_pin, without_release):
        with pytest.raises(AssertionError):
            assert_immutable_action_references(mutant)


def expected_version_output(requirement: str | None = None) -> str:
    """Return the CLI identity selected by an exact runner requirement."""
    selected = requirement or runner_requirement()
    match = EXACT_REQUIREMENT_RE.fullmatch(selected)
    assert match is not None
    return f"pre-commit {match.group('version')}"


def github_python_block(path: Path) -> str:
    """Extract the one actual ``shell: python`` installer block from a workflow."""
    lines = path.read_text(encoding="utf-8").splitlines()
    shell_index = next(index for index, line in enumerate(lines) if line.strip() == "shell: python")
    run_index = next(
        index for index in range(shell_index + 1, len(lines)) if lines[index].strip() == "run: |"
    )
    run_indent = len(lines[run_index]) - len(lines[run_index].lstrip())
    block_lines: list[str] = []
    for line in lines[run_index + 1 :]:
        indent = len(line) - len(line.lstrip())
        if line.strip() and indent <= run_indent:
            break
        block_lines.append(line)
    return textwrap.dedent("\n".join(block_lines)).strip() + "\n"


def azure_python_block(path: Path) -> str:
    """Extract the actual Python heredoc installer from an Azure pipeline."""
    lines = path.read_text(encoding="utf-8").splitlines()
    start = next(index for index, line in enumerate(lines) if line.strip() == "python - <<'PY'")
    end = next(index for index in range(start + 1, len(lines)) if lines[index].strip() == "PY")
    return textwrap.dedent("\n".join(lines[start + 1 : end])).strip() + "\n"


def installer_source(relative_path: str) -> str:
    """Return one real installer source block without reimplementing it."""
    path = REPO_ROOT / relative_path
    if relative_path.startswith(".github/"):
        return github_python_block(path)
    return azure_python_block(path)


def execute_installer_source(
    source: str,
    tmp_path: Path,
    monkeypatch: Any,
    *,
    requirement_text: str | None,
    run: Callable[..., Any],
    check_output: Callable[..., str],
) -> None:
    """Execute an extracted installer with controlled filesystem and subprocesses."""
    if requirement_text is not None:
        (tmp_path / "requirements-pre-commit.txt").write_text(
            requirement_text,
            encoding="utf-8",
            newline="\n",
        )
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(subprocess, "run", run)
    monkeypatch.setattr(subprocess, "check_output", check_output)
    exec(compile(source, "<workflow pre-commit installer>", "exec"), {})  # noqa: S102


@pytest.mark.parametrize("relative_path", PYTHON_INSTALLER_PATHS)
def test_actual_ci_installer_uses_exact_runner_source_and_verifies_command(
    tmp_path: Path,
    monkeypatch: Any,
    relative_path: str,
) -> None:
    """Every real CI installer consumes the shared file and checks the runnable CLI."""
    requirement = runner_requirement()
    run_calls: list[tuple[list[str], dict[str, Any]]] = []
    version_calls: list[tuple[list[str], dict[str, Any]]] = []

    def fake_run(command: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
        run_calls.append((command, kwargs))
        return subprocess.CompletedProcess(command, 0)

    def fake_check_output(command: list[str], **kwargs: Any) -> str:
        version_calls.append((command, kwargs))
        return expected_version_output(requirement)

    execute_installer_source(
        installer_source(relative_path),
        tmp_path,
        monkeypatch,
        requirement_text=requirement + "\n",
        run=fake_run,
        check_output=fake_check_output,
    )

    assert run_calls == [
        (
            [
                sys.executable,
                "-m",
                "pip",
                "install",
                "--requirement",
                "requirements-pre-commit.txt",
            ],
            {"check": True},
        )
    ]
    assert version_calls == [(["pre-commit", "--version"], {"text": True})]


@pytest.mark.parametrize("relative_path", PYTHON_INSTALLER_PATHS)
@pytest.mark.parametrize(
    "requirement_text",
    [
        None,
        "pre-commit",
        "pre-commit>=4",
        f"{runner_requirement()}\nwheel==1",
    ],
    ids=("missing", "malformed", "range", "extra-requirement"),
)
def test_actual_ci_installer_rejects_missing_or_nonexact_runner_source(
    tmp_path: Path,
    monkeypatch: Any,
    relative_path: str,
    requirement_text: str | None,
) -> None:
    """Missing, ranged, malformed, and multi-requirement sources fail before install."""

    def unexpected(*_args: Any, **_kwargs: Any) -> Any:
        raise AssertionError("invalid input must fail before invoking subprocess")

    with pytest.raises((FileNotFoundError, SystemExit)):
        execute_installer_source(
            installer_source(relative_path),
            tmp_path,
            monkeypatch,
            requirement_text=requirement_text,
            run=unexpected,
            check_output=unexpected,
        )


@pytest.mark.parametrize("relative_path", PYTHON_INSTALLER_PATHS)
def test_actual_ci_installer_propagates_install_failure(
    tmp_path: Path,
    monkeypatch: Any,
    relative_path: str,
) -> None:
    """A failed pip install cannot continue to a version check."""

    def failed_run(command: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
        if kwargs.get("check"):
            raise subprocess.CalledProcessError(1, command)
        return subprocess.CompletedProcess(command, 1)

    def unexpected_version(*_args: Any, **_kwargs: Any) -> str:
        raise AssertionError("failed install must not reach version verification")

    with pytest.raises(subprocess.CalledProcessError):
        execute_installer_source(
            installer_source(relative_path),
            tmp_path,
            monkeypatch,
            requirement_text=runner_requirement(),
            run=failed_run,
            check_output=unexpected_version,
        )


@pytest.mark.parametrize("relative_path", PYTHON_INSTALLER_PATHS)
@pytest.mark.parametrize("failure", ["mismatch", "broken-command"])
def test_actual_ci_installer_rejects_wrong_or_broken_installed_command(
    tmp_path: Path,
    monkeypatch: Any,
    relative_path: str,
    failure: str,
) -> None:
    """Successful installation is insufficient without the expected executable identity."""

    def successful_run(command: list[str], **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 0)

    def failed_version(command: list[str], **_kwargs: Any) -> str:
        if failure == "broken-command":
            raise subprocess.CalledProcessError(1, command)
        return "pre-commit 0.0.0"

    expected_error = subprocess.CalledProcessError if failure == "broken-command" else SystemExit
    with pytest.raises(expected_error):
        execute_installer_source(
            installer_source(relative_path),
            tmp_path,
            monkeypatch,
            requirement_text=runner_requirement(),
            run=successful_run,
            check_output=failed_version,
        )


def test_installer_failure_oracle_detects_removed_check_guard(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """The failure oracle kills a mutant that stops asking pip to raise."""
    source = installer_source(PYTHON_INSTALLER_PATHS[0])
    mutant = source.replace("check=True", "check=False", 1)
    assert mutant != source

    def check_sensitive_run(command: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
        if kwargs.get("check"):
            raise subprocess.CalledProcessError(1, command)
        return subprocess.CompletedProcess(command, 1)

    def expected_output(*_args: Any, **_kwargs: Any) -> str:
        return expected_version_output()

    execute_installer_source(
        mutant,
        tmp_path,
        monkeypatch,
        requirement_text=runner_requirement(),
        run=check_sensitive_run,
        check_output=expected_output,
    )

    fresh = tmp_path / "production"
    fresh.mkdir()
    with pytest.raises(subprocess.CalledProcessError):
        execute_installer_source(
            source,
            fresh,
            monkeypatch,
            requirement_text=runner_requirement(),
            run=check_sensitive_run,
            check_output=expected_output,
        )


def test_github_runner_caches_follow_both_runner_inputs() -> None:
    """The two cached gates invalidate Python and hook caches on the shared source."""
    expected_hash = "hashFiles('.pre-commit-config.yaml', 'requirements-pre-commit.txt')"
    for relative_path in GITHUB_CACHE_WORKFLOWS:
        text = (REPO_ROOT / relative_path).read_text(encoding="utf-8")
        assert re.search(r"cache:\s*['\"]pip['\"]", text)
        assert "cache-dependency-path: requirements-pre-commit.txt" in text
        assert expected_hash in text

    auto_fix_text = (REPO_ROOT / ".github/workflows/auto-fix-precommit.yml").read_text(
        encoding="utf-8"
    )
    assert "cache: pip" not in auto_fix_text
    assert "hashFiles(" not in auto_fix_text


def hook_prefix() -> str:
    """Return the actual hook through the end of its baseline-only block."""
    text = HOOK_PATH.read_text(encoding="utf-8")
    end_marker = "# template-sync: end baseline-only"
    end = text.index(end_marker) + len(end_marker)
    return text[:end] + "\n"


def make_executable(path: Path, text: str) -> None:
    """Write an LF shell fixture and make it executable on POSIX hosts."""
    path.write_text(text, encoding="utf-8", newline="\n")
    path.chmod(0o755)


@pytest.mark.parametrize(
    "mode", ["success", "missing-node", "missing-npm", "install-failure", "mutant"]
)
def test_claude_markdown_bootstrap(tmp_path: Path, mode: str) -> None:
    """Actual setup installs locked packages from the root and preserves prerequisite failures."""
    bash = shutil.which("bash")
    if bash is None:
        pytest.skip("Bash is required to exercise the actual Claude hook")
    assert bash is not None
    text = HOOK_PATH.read_text(encoding="utf-8")
    block = text.split("# template-sync: begin markdown-only\n", 1)[1].split(
        "# template-sync: end markdown-only", 1
    )[0]
    if mode == "mutant":
        assert "npm ci --ignore-scripts" in block
        block = block.replace("npm ci --ignore-scripts", ":")
    repo = tmp_path / "repo"
    hook = repo / ".claude/hooks/session-start.sh"
    hook.parent.mkdir(parents=True)
    missing = mode.removeprefix("missing-") if mode.startswith("missing-") else "none"
    fixture = (
        "#!/bin/bash\nset -euo pipefail\n"
        f'command() {{ if [ "${{2:-}}" = {shlex.quote(missing)} ]; then return 1; fi; builtin command "$@"; }}\n'
        "node() { printf 'v22.0.0\\n'; }\n"
        'npm() { printf \'%s\\n\' "$PWD" "$*" > npm-invocation.txt; '
        f"return {23 if mode == 'install-failure' else 0}; }}\n" + block
    )
    make_executable(hook, fixture)
    result = subprocess.run(
        [bash, bash_path(bash, hook)], cwd=tmp_path, capture_output=True, text=True, check=False
    )
    log = repo / "npm-invocation.txt"
    if mode.startswith("missing-"):
        assert result.returncode == 1
        assert "requires Node.js 22 or newer and npm" in result.stderr
        assert not log.exists()
    elif mode == "mutant":
        assert result.returncode == 0
        assert not log.exists(), "Removing installation must violate the success-case oracle"
    else:
        assert result.returncode == (23 if mode == "install-failure" else 0), result.stderr
        lines = log.read_text(encoding="utf-8").splitlines()
        assert lines == [bash_path(bash, repo), "ci --ignore-scripts"]


def run_claude_node_fixture(
    tmp_path: Path,
    *,
    version: str,
    node_exit: int = 0,
    npm_exit: int = 0,
    hook_source: str | None = None,
) -> tuple[subprocess.CompletedProcess[str], Path, Path]:
    """Run the actual Markdown block with observable, deterministic runtime wrappers."""
    bash = shutil.which("bash")
    if bash is None:
        pytest.skip("Bash is required to exercise the actual Claude hook")
    assert bash is not None
    if hook_source is None:
        text = HOOK_PATH.read_text(encoding="utf-8")
        hook_source = text.split("# template-sync: begin markdown-only\n", 1)[1].split(
            "# template-sync: end markdown-only", 1
        )[0]
    repo = tmp_path / "repo"
    hook = repo / ".claude/hooks/session-start.sh"
    hook.parent.mkdir(parents=True, exist_ok=True)
    fixture = (
        "#!/bin/bash\nset -euo pipefail\n"
        "node() { printf '%s\\n' \"$*\" >> node-invocation.txt; "
        f"printf '%s\\n' {shlex.quote(version)}; return {node_exit}; }}\n"
        'npm() { printf \'%s\\n\' "$PWD" "$*" > npm-invocation.txt; '
        f"return {npm_exit}; }}\n" + hook_source
    )
    make_executable(hook, fixture)
    result = subprocess.run(
        [bash, bash_path(bash, hook)], cwd=tmp_path, capture_output=True, text=True, check=False
    )
    return result, repo / "npm-invocation.txt", tmp_path / "node-invocation.txt"


def assert_claude_node_rejected(result: subprocess.CompletedProcess[str], npm_log: Path) -> None:
    """Reject unsupported or unknown runtimes before any package installation."""
    assert result.returncode == 1, result.stdout + result.stderr
    assert (
        result.stderr.strip() == "Markdown validation requires Node.js 22 or newer and npm on PATH."
    )
    assert not npm_log.exists()


@pytest.mark.parametrize("version", ["v22.0.0", "v24.0.0"])
def test_claude_node_floor_accepts_supported_stable_runtime(tmp_path: Path, version: str) -> None:
    """Both the minimum and a newer stable runtime reach the unchanged locked install."""
    result, npm_log, node_log = run_claude_node_fixture(tmp_path, version=version)
    assert result.returncode == 0, result.stdout + result.stderr
    assert node_log.read_text(encoding="utf-8").splitlines() == ["--version"]
    bash = shutil.which("bash")
    assert bash is not None
    assert npm_log.read_text(encoding="utf-8").splitlines() == [
        bash_path(bash, tmp_path / "repo"),
        "ci --ignore-scripts",
    ]
    first_install_log = npm_log.read_bytes()
    # The same supported environment remains valid on repeated setup.
    repeated, repeated_npm, _ = run_claude_node_fixture(tmp_path, version=version)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert repeated_npm.read_bytes() == first_install_log


@pytest.mark.parametrize(
    ("version", "node_exit"),
    [
        ("v18.20.8", 0),
        ("v20.19.0", 0),
        ("v21.7.3", 0),
        ("", 0),
        ("not-a-version", 0),
        ("v22", 0),
        ("v22.0.0-rc.1", 0),
        ("v022.0.0", 0),
        ("v22.00.0", 0),
        ("v22.0.00", 0),
        ("v22.0.0\nextra", 0),
        ("v" + "9" * 80 + ".0.0", 0),
        ("v24.0.0", 7),
    ],
)
def test_claude_node_floor_rejects_unsupported_or_unknown_runtime(
    tmp_path: Path, version: str, node_exit: int
) -> None:
    """Fixed native negative cases do not depend on the implementation's predicate."""
    result, npm_log, node_log = run_claude_node_fixture(
        tmp_path, version=version, node_exit=node_exit
    )
    assert_claude_node_rejected(result, npm_log)
    assert node_log.read_text(encoding="utf-8").splitlines() == ["--version"]


def test_claude_node_floor_preserves_native_npm_failure(tmp_path: Path) -> None:
    """A supported Node runtime does not hide a subsequent installation failure."""
    result, npm_log, _ = run_claude_node_fixture(tmp_path, version="v22.0.0", npm_exit=23)
    assert result.returncode == 23, result.stdout + result.stderr
    assert npm_log.exists()
    assert result.stderr == ""


def test_claude_node_floor_oracle_detects_presence_only_mutant(tmp_path: Path) -> None:
    """Restoring the actual old gate defeats the unchanged Node20 rejection oracle."""
    text = HOOK_PATH.read_text(encoding="utf-8")
    block = text.split("# template-sync: begin markdown-only\n", 1)[1].split(
        "# template-sync: end markdown-only", 1
    )[0]
    version_clauses = (
        ' \\\n  || ! node_version="$(node --version 2>/dev/null)"'
        ' \\\n  || [[ ! "$node_version" =~ ^v(0|[1-9][0-9]*)\\.(0|[1-9][0-9]*)\\.(0|[1-9][0-9]*)$ ]]'
        ' \\\n  || ! [ "${BASH_REMATCH[1]}" -ge 22 ] 2>/dev/null'
    )
    assert block.count(version_clauses) == 1
    original, original_npm, _ = run_claude_node_fixture(
        tmp_path / "original", version="v20.19.0", hook_source=block
    )
    assert_claude_node_rejected(original, original_npm)
    mutant, mutant_npm, mutant_node = run_claude_node_fixture(
        tmp_path / "mutant",
        version="v20.19.0",
        hook_source=block.replace(version_clauses, ""),
    )
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    assert mutant_npm.exists()
    assert not mutant_node.exists()
    with pytest.raises(AssertionError):
        assert_claude_node_rejected(mutant, mutant_npm)


def test_claude_node_floor_matches_both_package_manifests() -> None:
    """A deliberate package support-floor change must update the bootstrap boundary."""
    package = json.loads((REPO_ROOT / "package.json").read_text(encoding="utf-8"))
    lock = json.loads((REPO_ROOT / "package-lock.json").read_text(encoding="utf-8"))
    assert package["engines"]["node"] == ">=22.0.0"
    assert lock["packages"][""]["engines"]["node"] == ">=22.0.0"


def test_claude_local_session_exits_before_runtime_checks(tmp_path: Path) -> None:
    """The full hook leaves local workstation runtime management unchanged."""
    source = "export CLAUDE_CODE_REMOTE=false\n" + HOOK_PATH.read_text(encoding="utf-8")
    result, npm_log, node_log = run_claude_node_fixture(
        tmp_path, version="v18.20.8", hook_source=source
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert not npm_log.exists()
    assert not node_log.exists()


@pytest.mark.parametrize("markdown", [True, False])
def test_claude_node_gate_follows_materialized_markdown_selection(
    tmp_path: Path, markdown: bool
) -> None:
    """Normal selection removes the version check along with all Markdown setup."""
    from tests.test_materialize_downstream_adoption import materialize_module_fixture

    modules = (
        ("agent-instructions", "agent-claude", "markdown")
        if markdown
        else ("agent-instructions", "agent-claude")
    )
    target = materialize_module_fixture(tmp_path, modules, authorize_protected_files=True)
    hook = (target / ".claude/hooks/session-start.sh").read_text(encoding="utf-8")
    for token in ("node --version", "node_version", "BASH_REMATCH", "npm ci --ignore-scripts"):
        assert (token in hook) is markdown
    assert "ensure_pre_commit" not in hook
    # Terraform is not selected, so its early-exit installer is stripped too.
    assert "TERRAFORM_VERSION" not in hook
    assert "# Idempotency:" not in hook
    if markdown:
        assert hook.index("node --version") < hook.index("npm ci --ignore-scripts")


def test_claude_terraform_setup_runs_after_every_other_hook_block() -> None:
    """The Terraform block can exit the hook early, so no other setup may follow it."""
    text = HOOK_PATH.read_text(encoding="utf-8")
    begin = text.index("# template-sync: begin terraform-only\n")
    end_marker = "# template-sync: end terraform-only\n"
    end = text.index(end_marker) + len(end_marker)

    assert "exit 0" in text[begin:end]
    assert text[end:].strip() == ""
    for marker_name in ("baseline-only", "markdown-only"):
        assert text.index(f"# template-sync: end {marker_name}\n") < begin
    assert text.index("npm ci --ignore-scripts") < text.index("# Idempotency:")


def bash_path(bash: str, path: Path) -> str:
    """Return a path usable by Bash on both native POSIX and Git for Windows."""
    if os.name != "nt":
        return str(path)
    if subprocess.check_output([bash, "-lc", "uname -s"], text=True).strip() == "Linux":
        drive, tail = os.path.splitdrive(str(path.resolve()))
        assert drive and drive.endswith(":")
        normalized_tail = tail.lstrip("\\/").replace("\\", "/")
        return f"/mnt/{drive[0].lower()}/{normalized_tail}"
    return subprocess.check_output(
        [bash, "-lc", 'cygpath -u "$1"', "bash", str(path)],
        text=True,
    ).strip()


def run_hook_fixture(
    tmp_path: Path,
    *,
    installer: str | None,
    initial_version: str | None,
    requirement_text: str | None = None,
    broken_precommit: bool = False,
    installer_exit: int = 0,
    installer_updates: bool = True,
) -> tuple[subprocess.CompletedProcess[str], str]:
    """Run the actual baseline hook block with local executable mocks."""
    bash = shutil.which("bash")
    if bash is None:
        pytest.skip("Bash is required to exercise the actual Claude hook")
    assert bash is not None

    repo = tmp_path / "repo"
    hook = repo / ".claude" / "hooks" / "session-start.sh"
    hook.parent.mkdir(parents=True)
    make_executable(hook, hook_prefix())
    if requirement_text is not None:
        (repo / "requirements-pre-commit.txt").write_text(
            requirement_text,
            encoding="utf-8",
            newline="\n",
        )

    fake_bin = tmp_path / "bin"
    fake_bin.mkdir()
    state = tmp_path / "pre-commit.state"
    install_log = tmp_path / "installer.log"
    env_file = tmp_path / "claude.env"
    env_file.write_text("", encoding="utf-8")
    if initial_version is not None:
        state.write_text(initial_version + "\n", encoding="utf-8", newline="\n")

    make_executable(
        fake_bin / "pre-commit",
        """#!/bin/sh
if [ "${BROKEN_PRECOMMIT:-}" = "1" ]; then
  exit 2
fi
if [ -f "$PRE_COMMIT_STATE" ]; then
  cat "$PRE_COMMIT_STATE"
else
  echo "pre-commit 0.0.0"
fi
""",
    )
    installer_script = """#!/bin/sh
printf '%s\n' "$*" >> "$INSTALL_LOG"
if [ "${INSTALLER_EXIT:-0}" -ne 0 ]; then
  exit "$INSTALLER_EXIT"
fi
if [ "${INSTALLER_UPDATES:-1}" = "1" ]; then
  printf '%s\n' "$EXPECTED_PRE_COMMIT" > "$PRE_COMMIT_STATE"
fi
"""
    if installer in {"uv", "pipx"}:
        make_executable(fake_bin / installer, installer_script)
    if installer == "python":
        make_executable(
            fake_bin / "python",
            """#!/bin/sh
if [ "$1" = "-m" ] && [ "$2" = "site" ] && [ "$3" = "--user-base" ]; then
  printf '%s\n' "$FAKE_USER_BASE"
  exit 0
fi
printf '%s\n' "$*" >> "$INSTALL_LOG"
if [ "${INSTALLER_EXIT:-0}" -ne 0 ]; then
  exit "$INSTALLER_EXIT"
fi
if [ "${INSTALLER_UPDATES:-1}" = "1" ]; then
  printf '%s\n' "$EXPECTED_PRE_COMMIT" > "$PRE_COMMIT_STATE"
fi
""",
        )

    requirement = requirement_text if requirement_text is not None else runner_requirement()
    expected = (
        expected_version_output(requirement)
        if EXACT_REQUIREMENT_RE.fullmatch(requirement.strip())
        else "invalid"
    )
    fake_bin_bash = bash_path(bash, fake_bin)
    shell_env = {
        "PATH": f"{fake_bin_bash}:/usr/bin:/bin",
        "CLAUDE_CODE_REMOTE": "true",
        "CLAUDE_ENV_FILE": bash_path(bash, env_file),
        "HOME": bash_path(bash, tmp_path / "home"),
        "UV_TOOL_BIN_DIR": fake_bin_bash,
        "PIPX_BIN_DIR": fake_bin_bash,
        "FAKE_USER_BASE": bash_path(bash, tmp_path),
        "PRE_COMMIT_STATE": bash_path(bash, state),
        "INSTALL_LOG": bash_path(bash, install_log),
        "EXPECTED_PRE_COMMIT": expected,
        "BROKEN_PRECOMMIT": "1" if broken_precommit else "0",
        "INSTALLER_EXIT": str(installer_exit),
        "INSTALLER_UPDATES": "1" if installer_updates else "0",
    }
    wrapper = tmp_path / "run-hook.sh"
    exports = "\n".join(f"export {name}={shlex.quote(value)}" for name, value in shell_env.items())
    make_executable(
        wrapper,
        f"#!/bin/sh\n{exports}\nexec /bin/bash {shlex.quote(bash_path(bash, hook))}\n",
    )
    result = subprocess.run(
        [bash, bash_path(bash, wrapper)],
        check=False,
        capture_output=True,
        text=True,
    )
    log = install_log.read_text(encoding="utf-8") if install_log.exists() else ""
    return result, log


def run_hook_path_precedence_fixture(
    tmp_path: Path,
    *,
    installer: Literal["uv", "pipx", "python"],
    selected_position: Literal["first", "later", "absent"],
    selected_version: str,
    hook_source: str | None = None,
) -> tuple[subprocess.CompletedProcess[str], str, str]:
    """Run the actual hook with separate stale, installer, and selected directories."""
    bash = shutil.which("bash")
    if bash is None:
        pytest.skip("Bash is required to exercise the actual Claude hook")
    assert bash is not None

    repo = tmp_path / "repo"
    hook = repo / ".claude" / "hooks" / "session-start.sh"
    hook.parent.mkdir(parents=True)
    make_executable(hook, hook_source or hook_prefix())
    (repo / "requirements-pre-commit.txt").write_text(
        runner_requirement(), encoding="utf-8", newline="\n"
    )

    stale_bin = tmp_path / "stale-bin"
    installer_bin = tmp_path / "installer-bin"
    selected_base = tmp_path / "selected-user" if installer == "python" else tmp_path
    selected_bin = selected_base / "bin" if installer == "python" else tmp_path / "selected-bin"
    for directory in (stale_bin, installer_bin, selected_bin):
        directory.mkdir(parents=True, exist_ok=True)

    selected_state = tmp_path / "selected-pre-commit.state"
    selected_state.write_text(selected_version + "\n", encoding="utf-8", newline="\n")
    install_log = tmp_path / "installer.log"
    env_file = tmp_path / "claude.env"
    env_file.write_text("", encoding="utf-8")

    make_executable(stale_bin / "pre-commit", "#!/bin/sh\necho 'pre-commit 0.0.0'\n")
    make_executable(
        selected_bin / "pre-commit",
        """#!/bin/sh
cat "$SELECTED_PRE_COMMIT_STATE"
""",
    )
    installer_script = """#!/bin/sh
printf '%s\n' "$*" >> "$INSTALL_LOG"
printf '%s\n' "$EXPECTED_PRE_COMMIT" > "$SELECTED_PRE_COMMIT_STATE"
"""
    if installer in {"uv", "pipx"}:
        make_executable(installer_bin / installer, installer_script)
    else:
        make_executable(
            installer_bin / "python",
            """#!/bin/sh
if [ "$1" = "-m" ] && [ "$2" = "site" ] && [ "$3" = "--user-base" ]; then
  printf '%s\n' "$FAKE_USER_BASE"
  exit 0
fi
printf '%s\n' "$*" >> "$INSTALL_LOG"
printf '%s\n' "$EXPECTED_PRE_COMMIT" > "$SELECTED_PRE_COMMIT_STATE"
""",
        )

    stale_bin_bash = bash_path(bash, stale_bin)
    installer_bin_bash = bash_path(bash, installer_bin)
    selected_bin_bash = bash_path(bash, selected_bin)
    if selected_position == "first":
        path_entries = [selected_bin_bash, stale_bin_bash, installer_bin_bash]
    elif selected_position == "later":
        path_entries = [stale_bin_bash, installer_bin_bash, selected_bin_bash]
    else:
        path_entries = [stale_bin_bash, installer_bin_bash]
    initial_path = ":".join([*path_entries, "/usr/bin", "/bin"])

    shell_env = {
        "PATH": initial_path,
        "CLAUDE_CODE_REMOTE": "true",
        "CLAUDE_ENV_FILE": bash_path(bash, env_file),
        "HOME": bash_path(bash, tmp_path / "home"),
        "UV_TOOL_BIN_DIR": selected_bin_bash,
        "PIPX_BIN_DIR": selected_bin_bash,
        "FAKE_USER_BASE": bash_path(bash, selected_base),
        "SELECTED_PRE_COMMIT_STATE": bash_path(bash, selected_state),
        "INSTALL_LOG": bash_path(bash, install_log),
        "EXPECTED_PRE_COMMIT": expected_version_output(),
    }
    exports = "\n".join(f"export {name}={shlex.quote(value)}" for name, value in shell_env.items())
    wrapper = tmp_path / "run-hook-and-probe.sh"
    make_executable(
        wrapper,
        (
            f"#!/bin/bash\n{exports}\n"
            f". {shlex.quote(bash_path(bash, hook))}\n"
            "printf 'FINAL_COMMAND=%s\\n' \"$(command -v pre-commit)\"\n"
            "printf 'FINAL_VERSION=%s\\n' \"$(pre-commit --version)\"\n"
            "printf 'FINAL_PATH=%s\\n' \"$PATH\"\n"
        ),
    )
    result = subprocess.run(
        [bash, bash_path(bash, wrapper)],
        check=False,
        capture_output=True,
        text=True,
    )
    log = install_log.read_text(encoding="utf-8") if install_log.exists() else ""
    return result, log, env_file.read_text(encoding="utf-8")


def assert_selected_runner_is_active(
    result: subprocess.CompletedProcess[str],
    *,
    bash: str,
    selected_bin: Path,
) -> None:
    """Require the ordinary command to resolve the selected exact runner."""
    selected_bin_bash = bash_path(bash, selected_bin)
    assert result.returncode == 0, result.stdout + result.stderr
    assert f"FINAL_COMMAND={selected_bin_bash}/pre-commit" in result.stdout
    assert f"FINAL_VERSION={expected_version_output()}" in result.stdout
    final_path = next(
        line.removeprefix("FINAL_PATH=")
        for line in result.stdout.splitlines()
        if line.startswith("FINAL_PATH=")
    )
    assert final_path.split(":", maxsplit=1)[0] == selected_bin_bash


def selected_runner_bin(tmp_path: Path, installer: str) -> Path:
    """Return the install directory chosen by one fixture installer branch."""
    return (
        tmp_path / "selected-user" / "bin" if installer == "python" else tmp_path / "selected-bin"
    )


def assert_expected_installer_call(installer: str, install_log: str) -> None:
    """Require one installer to receive the exact shared pre-commit pin."""
    requirement = runner_requirement()
    if installer == "uv":
        assert install_log == f"tool install --force {requirement}\n"
    elif installer == "pipx":
        assert install_log == f"install --force {requirement}\n"
    else:
        assert install_log == f"-m pip install --user {requirement}\n"


@pytest.mark.parametrize("installer", ["uv", "pipx", "python"])
def test_claude_hook_activates_selected_runner_ahead_of_stale_command(
    tmp_path: Path, installer: Literal["uv", "pipx", "python"]
) -> None:
    """Every installer branch makes its exact runner win ordinary PATH lookup."""
    result, install_log, env_text = run_hook_path_precedence_fixture(
        tmp_path,
        installer=installer,
        selected_position="later",
        selected_version="pre-commit 1.0.0",
    )
    bash = shutil.which("bash")
    assert bash is not None
    selected_bin = selected_runner_bin(tmp_path, installer)
    assert_selected_runner_is_active(result, bash=bash, selected_bin=selected_bin)
    assert_expected_installer_call(installer, install_log)

    selected_bin_bash = bash_path(bash, selected_bin)
    assert env_text.splitlines() == [f'export PATH="{selected_bin_bash}:$PATH"']
    final_path = next(
        line.removeprefix("FINAL_PATH=")
        for line in result.stdout.splitlines()
        if line.startswith("FINAL_PATH=")
    )
    assert final_path.split(":").count(selected_bin_bash) == 2


@pytest.mark.parametrize("installer", ["uv", "pipx", "python"])
def test_claude_hook_reuses_exact_runner_in_selected_later_directory(
    tmp_path: Path, installer: Literal["uv", "pipx", "python"]
) -> None:
    """Activation avoids reinstalling an exact runner hidden behind a stale command."""
    result, install_log, env_text = run_hook_path_precedence_fixture(
        tmp_path,
        installer=installer,
        selected_position="later",
        selected_version=expected_version_output(),
    )
    bash = shutil.which("bash")
    assert bash is not None
    selected_bin = selected_runner_bin(tmp_path, installer)
    assert_selected_runner_is_active(result, bash=bash, selected_bin=selected_bin)
    assert install_log == ""
    selected_bin_bash = bash_path(bash, selected_bin)
    assert env_text.splitlines() == [f'export PATH="{selected_bin_bash}:$PATH"']


@pytest.mark.parametrize("selected_position", ["first", "absent"])
def test_claude_hook_activation_preserves_first_and_absent_path_controls(
    tmp_path: Path, selected_position: Literal["first", "absent"]
) -> None:
    """Activation remains correct when the selected directory is first or absent."""
    result, install_log, env_text = run_hook_path_precedence_fixture(
        tmp_path,
        installer="uv",
        selected_position=selected_position,
        selected_version="pre-commit 1.0.0",
    )
    bash = shutil.which("bash")
    assert bash is not None
    selected_bin = selected_runner_bin(tmp_path, "uv")
    assert_selected_runner_is_active(result, bash=bash, selected_bin=selected_bin)
    assert_expected_installer_call("uv", install_log)
    selected_bin_bash = bash_path(bash, selected_bin)
    assert env_text.splitlines() == [f'export PATH="{selected_bin_bash}:$PATH"']
    final_path = next(
        line.removeprefix("FINAL_PATH=")
        for line in result.stdout.splitlines()
        if line.startswith("FINAL_PATH=")
    )
    assert final_path.split(":").count(selected_bin_bash) == 1


def test_claude_hook_activation_oracle_rejects_removed_force_front(tmp_path: Path) -> None:
    """The stale-command regression must return when only persistence remains."""
    source = hook_prefix()
    activation_call = 'activate_pre_commit_bin "$pre_commit_bin_dir"'
    assert source.count(activation_call) == 3
    mutant = source.replace(
        activation_call,
        'persist_path_prepend "$pre_commit_bin_dir"',
    )

    original_result, _original_log, _original_env = run_hook_path_precedence_fixture(
        tmp_path / "original",
        installer="uv",
        selected_position="later",
        selected_version="pre-commit 1.0.0",
        hook_source=source,
    )
    mutant_result, mutant_log, _mutant_env = run_hook_path_precedence_fixture(
        tmp_path / "mutant",
        installer="uv",
        selected_position="later",
        selected_version="pre-commit 1.0.0",
        hook_source=mutant,
    )

    assert original_result.returncode == 0, original_result.stdout + original_result.stderr
    assert mutant_result.returncode != 0, mutant_result.stdout + mutant_result.stderr
    assert_expected_installer_call("uv", mutant_log)
    assert "installation did not provide" in mutant_result.stderr


def test_claude_hook_reuses_the_exact_available_runner(tmp_path: Path) -> None:
    """An exact runnable command avoids every installer branch."""
    result, install_log = run_hook_fixture(
        tmp_path,
        installer="uv",
        initial_version=expected_version_output(),
        requirement_text=runner_requirement(),
    )

    assert result.returncode == 0, result.stdout + result.stderr
    assert "pre-commit already available" in result.stdout
    assert install_log == ""


@pytest.mark.parametrize("installer", ["uv", "pipx", "python"])
def test_claude_hook_installs_exact_runner_through_supported_branch(
    tmp_path: Path,
    installer: str,
) -> None:
    """Each supported installer receives the exact shared requirement."""
    result, install_log = run_hook_fixture(
        tmp_path,
        installer=installer,
        initial_version="pre-commit 0.0.0",
        requirement_text=runner_requirement(),
    )

    assert result.returncode == 0, result.stdout + result.stderr
    assert runner_requirement() in install_log
    if installer == "uv":
        assert f"tool install --force {runner_requirement()}" in install_log
    elif installer == "pipx":
        assert f"install --force {runner_requirement()}" in install_log
    else:
        assert f"-m pip install --user {runner_requirement()}" in install_log
    assert expected_version_output() in result.stdout


@pytest.mark.parametrize(
    ("case", "kwargs"),
    [
        pytest.param("missing", {"requirement_text": None}, id="missing-requirement"),
        pytest.param("range", {"requirement_text": "pre-commit>=4"}, id="ranged-requirement"),
        pytest.param("broken", {"broken_precommit": True}, id="broken-command"),
        pytest.param("installer", {"installer_exit": 7}, id="installer-failure"),
        pytest.param("wrong", {"installer_updates": False}, id="wrong-version-after-install"),
    ],
)
def test_claude_hook_fails_closed_for_invalid_or_unfulfilled_runner(
    tmp_path: Path,
    case: str,
    kwargs: dict[str, Any],
) -> None:
    """The hook rejects invalid input and cannot report an unfulfilled install."""
    call_kwargs = dict(kwargs)
    supplied_requirement = call_kwargs.pop("requirement_text", runner_requirement())
    if case == "missing":
        requirement_text = None
    else:
        requirement_text = supplied_requirement
    result, _install_log = run_hook_fixture(
        tmp_path,
        installer="uv",
        initial_version="pre-commit 0.0.0",
        requirement_text=requirement_text,
        **call_kwargs,
    )

    assert result.returncode != 0, result.stdout + result.stderr


@pytest.mark.parametrize("relative_path", CREDENTIAL_FREE_WORKFLOWS)
def test_every_live_checkout_explicitly_disables_credentials(relative_path: str) -> None:
    """The shipped workflows declare credential intent independently of action defaults."""
    assert_credential_free_checkouts((REPO_ROOT / relative_path).read_text(encoding="utf-8"))


@pytest.mark.parametrize("value", [True, "true", None, 0, 1, [], {}, "${{ inputs.persist }}", "no"])
def test_later_checkout_rejects_invalid_credential_inputs(value: Any) -> None:
    """A valid first checkout cannot conceal a later unsafe or ambiguous checkout."""
    workflow = yaml.safe_load(
        (REPO_ROOT / CREDENTIAL_FREE_WORKFLOWS[0]).read_text(encoding="utf-8")
    )
    job = next(iter(workflow["jobs"].values()))
    job["steps"].append({"uses": "actions/checkout@v7", "with": {"persist-credentials": value}})
    with pytest.raises(AssertionError):
        assert_credential_free_checkouts(yaml.safe_dump(workflow))


@pytest.mark.parametrize("value", ["false", "False", "FALSE", " false ", "\tfalse\t"])
def test_quoted_false_checkout_input_is_valid(value: str) -> None:
    """An action input string false has the same intended effect as YAML false."""
    workflow = (REPO_ROOT / CREDENTIAL_FREE_WORKFLOWS[0]).read_text(encoding="utf-8")
    document = yaml.safe_load(workflow)
    for job in document["jobs"].values():
        for step in job.get("steps", []):
            if step.get("uses", "").startswith("actions/checkout@"):
                step["with"]["persist-credentials"] = value
    quoted = yaml.safe_dump(document)
    assert quoted != workflow
    assert_credential_free_checkouts(quoted)


def test_checkout_oracle_detects_removed_credential_assertion() -> None:
    """Independent missing-input expectation fails when the credential assertion vanishes."""
    source = inspect.getsource(assert_credential_free_checkouts)
    assertion = (
        "            assert value is False or (\n"
        '                isinstance(value, str) and value.strip(" \\t\\r\\n").lower() == "false"\n'
        "            )"
    )
    assert source.count(assertion) == 1
    namespace: dict[str, Any] = {"yaml": yaml}
    # Execute only this module's owned oracle with one fixed guard-removal mutation.
    exec(source.replace(assertion, "            pass"), namespace)  # noqa: S102
    mutant = namespace["assert_credential_free_checkouts"]
    workflow = (REPO_ROOT / CREDENTIAL_FREE_WORKFLOWS[0]).read_text(encoding="utf-8")
    weakened = workflow.replace(
        "          persist-credentials: false\n", "          fetch-depth: 1\n"
    )
    assert weakened != workflow

    def rejects_missing(checker: Callable[[str], None]) -> None:
        try:
            checker(weakened)
        except AssertionError:
            return
        raise AssertionError("missing checkout credential protection was accepted")

    rejects_missing(assert_credential_free_checkouts)
    with pytest.raises(AssertionError, match="protection was accepted"):
        rejects_missing(mutant)


def test_markdown_ci_provisions_actual_hook_regression() -> None:
    """Markdown CI supplies every prerequisite and blocks on the actual hook regression."""
    assert_markdown_hook_regression_ci(MARKDOWN_WORKFLOW_PATH.read_text(encoding="utf-8"))


def test_markdown_ci_oracle_rejects_removed_hook_guards() -> None:
    """Independent mutations cannot remove runner setup, invocation, or failure blocking."""
    workflow = MARKDOWN_WORKFLOW_PATH.read_text(encoding="utf-8")
    mutations = (
        (
            "python -m pip install --requirement requirements-pre-commit.txt",
            "python -m pip install --requirement missing-runner.txt",
        ),
        (
            "test_nested_markdown_hook_actual_invocation_and_mutation",
            "test_nested_markdown_hook_launcher_preserves_arguments_and_failure",
        ),
        (
            "steps.test-nested-hook.outcome == 'failure'",
            "steps.test-nested-hook.outcome == 'cancelled'",
        ),
    )
    for selected, replacement in mutations:
        assert workflow.count(selected) == 1
        mutant = workflow.replace(selected, replacement, 1)
        with pytest.raises(AssertionError):
            assert_markdown_hook_regression_ci(mutant)


def test_nested_markdown_hook_launcher_preserves_arguments_and_failure(
    monkeypatch: Any,
    capsys: Any,
) -> None:
    """The launcher neither shells filenames nor converts a native lint failure to success."""
    import runpy

    namespace = runpy.run_path(str(REPO_ROOT / ".github/scripts/lint_nested_markdown_hook.py"))
    calls: list[tuple[list[str], dict[str, Any]]] = []

    def run(command: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
        calls.append((command, kwargs))
        return subprocess.CompletedProcess(command, 9)

    monkeypatch.setattr(shutil, "which", lambda _name: "chosen-node")
    monkeypatch.setattr(subprocess, "run", run)
    assert namespace["main"](["docs/file with spaces.md", "--odd.mdc"]) == 9
    assert calls == [
        (
            [
                "chosen-node",
                str(REPO_ROOT / ".github/scripts/lint-nested-markdown.js"),
                "docs/file with spaces.md",
                "--odd.mdc",
            ],
            {"cwd": REPO_ROOT, "check": False},
        )
    ]
    calls.clear()
    for args in ([], [""]):
        assert namespace["main"](args) == 1
    assert not calls
    monkeypatch.setattr(shutil, "which", lambda _name: None)
    assert namespace["main"](["readme.md"]) == 1
    assert "npm ci --ignore-scripts" in capsys.readouterr().err
    assert not calls


def test_nested_markdown_hook_os_error_diagnostics_hide_filenames(
    monkeypatch: Any,
    capsys: Any,
    tmp_path: Path,
) -> None:
    """Launcher diagnostics keep context and cause without exposing OSError paths."""
    import runpy

    launcher_path = REPO_ROOT / ".github/scripts/lint_nested_markdown_hook.py"
    source = launcher_path.read_text(encoding="utf-8")
    namespace = runpy.run_path(str(launcher_path))
    monkeypatch.setattr(shutil, "which", lambda _name: "chosen-node")

    private_paths = ("source-private-canary.md", "destination-private-canary.md")

    def require_safe_error(main: Callable[[list[str]], int], error: OSError) -> None:
        def fail_run(*_args: Any, **_kwargs: Any) -> subprocess.CompletedProcess[str]:
            raise error

        monkeypatch.setattr(subprocess, "run", fail_run)
        assert main(["README.md"]) == 1
        diagnostic = capsys.readouterr().err
        for private_path in private_paths:
            assert (
                private_path not in diagnostic
            ), f"diagnostic exposed private path: {private_path}"
        assert diagnostic.startswith("Cannot start nested Markdown lint: ")
        assert f"{type(error).__name__}: {error.strerror or 'I/O error'}" in diagnostic

    with_paths = OSError(5, "fixture I/O failure")
    with_paths.filename, with_paths.filename2 = private_paths
    require_safe_error(namespace["main"], with_paths)

    without_strerror = OSError()
    without_strerror.filename, without_strerror.filename2 = private_paths
    require_safe_error(namespace["main"], without_strerror)

    safe_summary = "f\"{type(error).__name__}: {error.strerror or 'I/O error'}\""
    assert source.count(safe_summary) == 1
    mutant_path = tmp_path / "raw_error_launcher.py"
    mutant_path.write_text(source.replace(safe_summary, "str(error)", 1), encoding="utf-8")
    mutant = runpy.run_path(str(mutant_path))
    with pytest.raises(
        AssertionError,
        match="diagnostic exposed private path: source-private-canary.md",
    ):
        require_safe_error(mutant["main"], with_paths)


def test_nested_markdown_hook_actual_invocation_and_mutation(tmp_path: Path) -> None:
    """Run the real selected hook, including spaces, deletion, no files and an invocation mutant."""
    executable = os.environ.get("PRE_COMMIT_EXECUTABLE") or shutil.which("pre-commit")
    if not executable or not shutil.which("node") or not (REPO_ROOT / "node_modules").is_dir():
        pytest.skip("Install the pinned pre-commit runner and run npm ci --ignore-scripts.")
    assert executable is not None
    version = subprocess.run([executable, "--version"], capture_output=True, text=True, check=False)
    assert version.returncode == 0, version.stderr
    assert version.stdout.strip() == expected_version_output()
    root = tmp_path / "hook repository with spaces"
    scripts = root / ".github/scripts"
    scripts.mkdir(parents=True)
    for name in ("lint-nested-markdown.js", "lint_nested_markdown_hook.py"):
        shutil.copy2(REPO_ROOT / ".github/scripts" / name, scripts / name)
    shutil.copy2(REPO_ROOT / ".markdownlint.jsonc", root / ".markdownlint.jsonc")
    config = yaml.safe_load((REPO_ROOT / ".pre-commit-config.yaml").read_text(encoding="utf-8"))
    hook = next(
        hook
        for repo in config["repos"]
        for hook in repo["hooks"]
        if hook["id"] == "lint-nested-markdown"
    )
    config_path = root / ".pre-commit-config.yaml"

    def write_config(selected_hook: dict[str, Any]) -> None:
        config_path.write_text(
            yaml.safe_dump({"repos": [{"repo": "local", "hooks": [selected_hook]}]}),
            encoding="utf-8",
        )

    write_config(hook)
    subprocess.run(["git", "init", "--quiet", str(root)], check=True, capture_output=True)
    bad = root / "bad example.mdc"
    bad.write_text("# Example\n\n```markdown\n# Bad heading.\n```\n", encoding="utf-8")
    good = root / "good example.md"
    good.write_text("# Example\n\nReadable without nested blocks.\n", encoding="utf-8")
    subprocess.run(
        ["git", "-C", str(root), "add", "--", bad.name, good.name], check=True, capture_output=True
    )
    environment = {**os.environ, "NODE_PATH": str(REPO_ROOT / "node_modules")}
    environment.pop("SKIP", None)

    def run_hook(*filenames: str) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [
                executable,
                "run",
                "lint-nested-markdown",
                "--config",
                str(config_path),
                "--files",
                *filenames,
            ],
            cwd=root,
            env=environment,
            capture_output=True,
            text=True,
            encoding="utf-8",
            check=False,
        )

    def require_rejected(result: subprocess.CompletedProcess[str]) -> None:
        assert result.returncode != 0, result.stdout + result.stderr
        assert "MD026" in result.stdout + result.stderr
        assert bad.name in result.stdout + result.stderr

    require_rejected(run_hook(bad.name))
    assert run_hook(good.name).returncode == 0
    # A removed invocation must fail the independent invalid-content oracle.
    mutant = dict(hook, entry='python -c "import sys; sys.exit(0)"')
    write_config(mutant)
    with pytest.raises(AssertionError):
        require_rejected(run_hook(bad.name))
    write_config(hook)
    bad.unlink()
    deleted = run_hook(bad.name)
    assert deleted.returncode == 0, deleted.stdout + deleted.stderr
    assert "Skipped" in deleted.stdout and "no files to check" in deleted.stdout
    (root / "unrelated.txt").write_text("No Markdown files selected.\n", encoding="utf-8")
    empty = run_hook("unrelated.txt")
    assert empty.returncode == 0, empty.stdout + empty.stderr
    assert "Skipped" in empty.stdout
    # Filtered deletion does not redefine explicit missing CLI inputs as success.
    direct = subprocess.run(
        [sys.executable, str(scripts / "lint_nested_markdown_hook.py"), bad.name],
        cwd=root,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )
    assert direct.returncode != 0
    assert bad.name in direct.stdout + direct.stderr
