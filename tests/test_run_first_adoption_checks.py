"""Exercise the first-adoption working-tree validation runner."""

from __future__ import annotations

import importlib
import io
import shutil
import subprocess
import sys
from collections.abc import Callable, Sequence
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any, cast

from tests._pytest_compat import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_PATH = REPO_ROOT / ".template-sync" / "scripts" / "run_first_adoption_checks.py"
SCRIPT_DIR = SCRIPT_PATH.parent

if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

first_adoption = cast(Any, importlib.import_module("run_first_adoption_checks"))


def test_decode_timeout_stream_handles_text_and_bytes_like_values() -> None:
    """Timeout stream decoding accepts every subprocess timeout stream shape."""
    assert first_adoption.decode_timeout_stream(None) == ""
    assert first_adoption.decode_timeout_stream("plain text") == "plain text"
    assert first_adoption.decode_timeout_stream(b"bytes text") == "bytes text"
    assert first_adoption.decode_timeout_stream(bytearray(b"bytearray text")) == "bytearray text"
    assert first_adoption.decode_timeout_stream(memoryview(b"memoryview text")) == "memoryview text"


def _run_git(repo_root: Path, *args: str) -> str:
    """Run a Git command in a fixture repository and return stdout."""
    result = subprocess.run(
        ["git", *args],
        cwd=repo_root,
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stderr
    return result.stdout.strip()


def _write_text(repo_root: Path, relative_path: str, text: str = "placeholder\n") -> None:
    """Write a UTF-8 fixture file."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _write_marker(repo_root: Path, text: str) -> None:
    """Write a downstream marker fixture."""
    _write_text(repo_root, ".template-sync/marker.yml", text)


def _write_manifest(repo_root: Path, text: str) -> None:
    """Write a template-sync manifest fixture."""
    _write_text(repo_root, ".template-sync/manifest.yml", text)


def _write_pytest_gate_manifest(repo_root: Path) -> None:
    """Write a manifest fixture with pytest configuration and test mappings."""
    _write_manifest(
        repo_root,
        "\n".join(
            [
                "template_manifest:",
                "  path_mappings:",
                "    - pattern: 'pyproject.toml'",
                "      requires_all:",
                "        - python",
                "    - pattern: 'tests/test_run_first_adoption_checks.py'",
                "      requires_all:",
                "        - template-sync-support",
                "    - pattern: 'tests/test_schema_examples.py'",
                "      requires_any:",
                "        - schema",
                "        - template-sync-support",
                "    - pattern: 'tests/test_template_manifest.py'",
                "      requires_all:",
                "        - template-sync-support",
                "    - pattern: 'tests/*.py'",
                "      requires_all:",
                "        - python",
                "",
            ]
        ),
    )


def _write_pytest_profile_root(tmp_path: Path) -> Path:
    """Create a temporary pytest root with the committed pytest configuration."""
    pytest_root = tmp_path / "pytest-root"
    pytest_root.mkdir()
    # PowerShell-only adoptions retain support tests without the Python project module.
    config_path = REPO_ROOT / "pytest.ini"
    if not config_path.is_file():
        config_path = REPO_ROOT / "pyproject.toml"
    shutil.copyfile(config_path, pytest_root / config_path.name)
    return pytest_root


def _recording_runner(records: list[tuple[str, ...]]) -> Callable[[Sequence[str], Path], int]:
    """Return a fake command runner that records commands."""

    def run(command: Sequence[str], _repo_root: Path) -> int:
        records.append(tuple(command))
        return 0

    return run


def _scripted_probe_executor(
    actions: dict[tuple[str, ...], Any | BaseException],
) -> tuple[
    Callable[[Sequence[str], Path, float], Any],
    list[tuple[tuple[str, ...], float]],
]:
    """Return a fake doctor probe executor that records commands and timeouts."""
    records: list[tuple[tuple[str, ...], float]] = []

    def execute(
        command: Sequence[str],
        _repo_root: Path,
        timeout_seconds: float,
    ) -> Any:
        command_tuple = tuple(command)
        records.append((command_tuple, timeout_seconds))
        action = actions.get(
            command_tuple,
            FileNotFoundError(2, "No such file or directory"),
        )
        if isinstance(action, BaseException):
            raise action
        return action

    return execute, records


def _queued_status_reader(*snapshots: tuple[str, ...]) -> Callable[[Path], tuple[str, ...]]:
    """Return a Git status reader that consumes a fixed sequence of snapshots."""
    remaining_snapshots = list(snapshots)

    def read_status(_repo_root: Path) -> tuple[str, ...]:
        assert remaining_snapshots, "No queued Git status snapshot remains."
        return remaining_snapshots.pop(0)

    return read_status


class _ExactStatusReader:
    """Git status reader that fails on over-consumption and exposes leftovers."""

    def __init__(self, *snapshots: tuple[str, ...]) -> None:
        self.remaining_snapshots = list(snapshots)
        self.read_count = 0

    def __call__(self, _repo_root: Path) -> tuple[str, ...]:
        assert self.remaining_snapshots, "No queued Git status snapshot remains."
        self.read_count += 1
        return self.remaining_snapshots.pop(0)

    def assert_consumed(self) -> None:
        """Assert that the runner consumed every queued status snapshot."""
        assert self.remaining_snapshots == []


def _failing_status_reader(
    *actions: tuple[str, ...] | BaseException,
) -> Callable[[Path], tuple[str, ...]]:
    """Return a status reader that returns snapshots or raises queued exceptions."""
    remaining_actions = list(actions)

    def read_status(_repo_root: Path) -> tuple[str, ...]:
        assert remaining_actions, "No queued Git status action remains."
        action = remaining_actions.pop(0)
        if isinstance(action, BaseException):
            raise action
        return action

    return read_status


def _output_section(output: str, header: str, stop_headers: Sequence[str]) -> str:
    """Return a named output section bounded by the next known header."""
    start = output.index(header)
    search_start = start + len(header)
    end_candidates = [
        output.index(stop_header, search_start)
        for stop_header in stop_headers
        if stop_header in output[search_start:]
    ]
    end = min(end_candidates) if end_candidates else len(output)
    return output[start:end]


def _transition_section(output: str) -> str:
    """Return the command-boundary transition section from runner output."""
    return _output_section(
        output,
        "Command-boundary status transitions:",
        ("Changed-file follow-up:", "First-adoption checks failed:", "Total elapsed time:"),
    )


def _changed_file_follow_up(output: str) -> str:
    """Return the changed-file follow-up block from runner output."""
    return _output_section(
        output,
        "Changed-file follow-up:",
        ("First-adoption checks failed:", "Total elapsed time:", "First-adoption checks passed."),
    )


def _utc_time(second: int) -> datetime:
    """Return a deterministic UTC timestamp for timing assertions."""
    return datetime(2026, 6, 3, 12, 0, second, tzinfo=UTC)


def _queued_time_source(*timestamps: datetime) -> Callable[[], datetime]:
    """Return a time source that consumes a fixed sequence of timestamps."""
    remaining_timestamps = list(timestamps)

    def now() -> datetime:
        assert remaining_timestamps, "No queued timestamp remains."
        return remaining_timestamps.pop(0)

    return now


def _doctor_result(
    group_label: str,
    name: str,
    command: tuple[str, ...],
    availability_state: str,
) -> Any:
    """Build a minimal doctor probe result for recommendation tests."""
    return first_adoption.DoctorProbeResult(
        group_label=group_label,
        name=name,
        command=command,
        availability_state=availability_state,
        exit_code=0 if availability_state == first_adoption.PROBE_AVAILABLE else 1,
        elapsed_time="0.000s",
        timed_out=False,
        stdout="",
        stderr="",
    )


def _incrementing_time_source() -> Callable[[], datetime]:
    """Return a deterministic time source that advances one second per read."""
    current_timestamp: list[datetime] = [_utc_time(0)]

    def now() -> datetime:
        timestamp = current_timestamp[0]
        current_timestamp[0] = timestamp + timedelta(seconds=1)
        return timestamp

    return now


def test_no_file_repository_exits_without_validation_commands(tmp_path: Path) -> None:
    """An empty Git repository exits cleanly with a useful no-file message."""
    _run_git(tmp_path, "init")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        stdout=stdout,
    )

    assert result == 0
    assert commands == []
    assert "Run mode: check" in stdout.getvalue()
    assert "No tracked or untracked non-ignored regular files found" in stdout.getvalue()
    assert "No first-adoption checks were available to run." in stdout.getvalue()
    assert "No Git status changes were detected" in stdout.getvalue()


def test_default_mode_is_explicit_check() -> None:
    """The CLI defaults to explicit check mode."""
    args = first_adoption.parse_args([])

    assert args.run_mode == first_adoption.CHECK_MODE


def test_fix_mode_is_explicit() -> None:
    """The CLI exposes an explicit fix mode for mutating hook/fixer runs."""
    args = first_adoption.parse_args(["--fix"])

    assert args.run_mode == first_adoption.FIX_MODE


def test_check_help_text_matches_changed_file_exit_contract(capsys: Any) -> None:
    """Check-mode help distinguishes final status changes from transient transitions."""
    with pytest.raises(SystemExit) as excinfo:
        first_adoption.parse_args(["--help"])

    assert excinfo.value.code == 0
    help_text = " ".join(capsys.readouterr().out.split())
    assert "final aggregate Git status change exits with the changed-file code" in help_text
    assert "Transient command-boundary status transitions are reported" in help_text
    assert "Any Git status change during the invocation exits nonzero" not in help_text


def test_fix_help_text_matches_fixer_and_failure_contract(capsys: Any) -> None:
    """Fix-mode help permits fixer arguments without swallowing command failures."""
    with pytest.raises(SystemExit) as excinfo:
        first_adoption.parse_args(["--help"])

    assert excinfo.value.code == 0
    help_text = " ".join(capsys.readouterr().out.split())
    assert "same validation surface" in help_text
    assert "intentional fixer behavior" in help_text
    assert "any planned command that exits nonzero still fails the run" in help_text
    assert "same validation commands run as in check mode" not in help_text
    assert "update files without failing the run" not in help_text


def test_doctor_mode_is_explicit() -> None:
    """The CLI exposes an explicit doctor mode with bounded diagnostic controls."""
    args = first_adoption.parse_args(
        [
            "--doctor",
            "--doctor-timeout",
            "1.5",
            "--doctor-max-bytes",
            "80",
            "--doctor-max-lines",
            "3",
        ]
    )

    assert args.run_mode == first_adoption.DOCTOR_MODE
    assert args.doctor_timeout == 1.5
    assert args.doctor_max_bytes == 80
    assert args.doctor_max_lines == 3


def test_retained_module_cli_option_is_repeatable() -> None:
    """Pre-marker retained module state can be supplied explicitly to the CLI."""
    args = first_adoption.parse_args(
        [
            "--retained-module",
            "python",
            "--retained-module",
            "template-sync-support",
        ]
    )

    assert args.retained_modules == ["python", "template-sync-support"]


def test_diagnostic_text_is_redacted_and_bounded() -> None:
    """Doctor diagnostics redact likely credentials before deterministic truncation."""
    diagnostic = "\n".join(
        [
            "token=abc123",
            "password: hunter2",
            "Authorization: Bearer secret-token",
            "https://user:secret@example.test/repo.git",
            "line five",
            "",
        ]
    )

    bounded = first_adoption.bound_diagnostic_text(
        diagnostic,
        max_bytes=200,
        max_lines=3,
    )

    assert "token=***" in bounded
    assert "password: ***" in bounded
    assert "secret-token" not in bounded
    assert "abc123" not in bounded
    assert "hunter2" not in bounded
    assert "[truncated: line limit 3 lines]" in bounded

    byte_bounded = first_adoption.bound_diagnostic_text(
        "safe-prefix-" + ("x" * 100),
        max_bytes=20,
        max_lines=20,
    )

    assert "[truncated: byte limit 20 bytes]" in byte_bounded


def test_diagnostic_text_truncation_markers_stay_within_caps() -> None:
    """Truncation markers are reserved within the caps, not appended past them."""
    many_lines = "\n".join(f"line {index}" for index in range(20))

    line_bounded = first_adoption.bound_diagnostic_text(
        many_lines,
        max_bytes=4096,
        max_lines=3,
    )

    rendered_lines = line_bounded.split("\n")
    assert len(rendered_lines) == 3
    assert rendered_lines[-1] == "[truncated: line limit 3 lines]"
    assert len(line_bounded.encode("utf-8")) <= 4096

    byte_bounded = first_adoption.bound_diagnostic_text(
        "x" * 500,
        max_bytes=80,
        max_lines=40,
    )

    assert len(byte_bounded.encode("utf-8")) <= 80
    assert "[truncated: byte limit 80 bytes]" in byte_bounded


def test_diagnostic_text_marker_emitted_when_cap_too_small() -> None:
    """A marker is still surfaced when the cap cannot hold both payload and marker."""
    bounded = first_adoption.bound_diagnostic_text(
        "x" * 200,
        max_bytes=10,
        max_lines=40,
    )

    assert bounded == "[truncated: byte limit 10 bytes]"


def test_diagnostic_text_collapses_markers_when_line_cap_too_small() -> None:
    """Both truncations with max_lines=1 collapse to a single in-cap marker line."""
    crowded = "\n".join("x" * 40 for _ in range(10))

    bounded = first_adoption.bound_diagnostic_text(
        crowded,
        max_bytes=120,
        max_lines=1,
    )

    rendered_lines = bounded.split("\n")
    assert len(rendered_lines) == 1
    assert rendered_lines[0] == "[truncated: byte limit 120 bytes; line limit 1 lines]"
    assert len(bounded.encode("utf-8")) <= 120


def test_pre_commit_prefix_falls_back_when_console_shim_fails(
    monkeypatch: Any,
) -> None:
    """Validation plans prefer a working module form when the console shim fails."""

    def fake_which(name: str) -> str | None:
        return "pre-commit" if name == "pre-commit" else None

    def fake_command_succeeds(
        command: Sequence[str],
        *,
        timeout_seconds: float = 5.0,
    ) -> bool:
        del timeout_seconds
        return tuple(command) == (sys.executable, "-m", "pre_commit", "--version")

    monkeypatch.setattr(first_adoption.shutil, "which", fake_which)
    monkeypatch.setattr(first_adoption, "command_succeeds", fake_command_succeeds)

    assert first_adoption.default_pre_commit_prefix() == (
        sys.executable,
        "-m",
        "pre_commit",
        "run",
        "--files",
    )


def test_doctor_mode_reports_probe_states_without_running_validation(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Doctor mode reports environment states without running hooks or fixers."""
    monkeypatch.setattr(first_adoption, "default_npm_executable", lambda: "npm")
    _write_text(
        tmp_path,
        ".pre-commit-config.yaml",
        "\n".join(
            [
                "repos:",
                "  - repo: local",
                "    hooks:",
                "      - id: yamllint",
                "      - id: black",
                "",
            ]
        ),
    )
    powershell_flags = first_adoption.POWERSHELL_PROBE_FLAGS
    sys_python = (sys.executable, "-m")
    actions: dict[tuple[str, ...], Any | BaseException] = {
        ("git", "--version"): first_adoption.ProbeExecution(0, "git version 2\n", ""),
        (*sys_python, first_adoption.PYTHON_DOCTOR_MODULE): first_adoption.ProbeExecution(
            0,
            "platform ok\n",
            "",
        ),
        ("pre-commit", "--version"): first_adoption.ProbeExecution(
            1,
            "",
            "broken shim token=abc123\n",
        ),
        (*sys_python, "pre_commit", "--version"): first_adoption.ProbeExecution(
            0,
            "pre-commit 4.0\n",
            "",
        ),
        ("pytest", "--version"): FileNotFoundError(2, "No such file or directory"),
        (*sys_python, "pytest", "--version"): first_adoption.ProbeExecution(
            0,
            "pytest 8.0\n",
            "",
        ),
        ("node", "--version"): subprocess.TimeoutExpired(
            ("node", "--version"),
            1.5,
            output="api_key=node-secret\n",
            stderr="",
        ),
        ("npm", "--version"): first_adoption.ProbeExecution(0, "11.0.0\n", ""),
        ("yamllint", "--version"): FileNotFoundError(2, "No such file or directory"),
        (*sys_python, "yamllint", "--version"): first_adoption.ProbeExecution(
            1,
            "",
            "ModuleNotFoundError: No module named yamllint\n",
        ),
        (
            "pwsh",
            *powershell_flags,
            first_adoption.POWERSHELL_VERSION_PROBE,
        ): first_adoption.ProbeExecution(0, "7.4.0\n", ""),
        (
            "pwsh",
            *powershell_flags,
            first_adoption.PSSCRIPTANALYZER_PROBE,
        ): first_adoption.ProbeExecution(0, "1.23.0\n", ""),
        (
            "powershell",
            *powershell_flags,
            first_adoption.POWERSHELL_VERSION_PROBE,
        ): first_adoption.ProbeExecution(0, "5.1\n", ""),
        (
            "powershell",
            *powershell_flags,
            first_adoption.PSSCRIPTANALYZER_PROBE,
        ): first_adoption.ProbeExecution(
            1,
            "",
            "PSScriptAnalyzer module not available\n",
        ),
    }
    probe_executor, records = _scripted_probe_executor(actions)
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_doctor(
        tmp_path,
        probe_timeout_seconds=1.5,
        max_output_bytes=120,
        max_output_lines=4,
        probe_executor=probe_executor,
        time_source=_incrementing_time_source(),
        stdout=stdout,
    )

    output = stdout.getvalue()
    assert result == 0
    assert records
    assert not any("run" in command and "--files" in command for command, _timeout in records)
    assert not any("--fix" in command for command, _timeout in records)
    assert {timeout for _command, timeout in records} == {1.5}
    assert "Run mode: doctor" in output
    assert "does not run validation hooks" in output
    assert "pre-commit console" in output
    assert "Availability: failed" in output
    assert "broken shim token=***" in output
    assert "abc123" not in output
    assert "Node.js executable" in output
    assert "Availability: timed-out" in output
    assert "Timed out after 1.500s." in output
    assert "api_key=***" in output
    assert "pytest console" in output
    assert "Availability: unavailable" in output
    assert "yamllint was not directly available" in output
    assert "PSScriptAnalyzer module (pwsh)" in output
    assert "PSScriptAnalyzer module (powershell)" in output
    assert "PSScriptAnalyzer module not available" in output
    assert "Use PSScriptAnalyzer host: pwsh" in output
    assert "pre_commit run --files" in output
    assert (
        f"Use pytest invocation: "
        f"{first_adoption.format_command(first_adoption.downstream_pytest_command())}"
    ) in output
    assert "Use pytest invocation: pytest" not in output
    assert "Doctor mode does not run pre-commit hooks" in output
    assert "Hook IDs detected in .pre-commit-config.yaml" in output
    assert "    - yamllint" in output
    assert "    - black" in output


def test_check_plan_assigns_stable_group_labels(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Planned validation commands carry stable group labels."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    monkeypatch.setattr(first_adoption, "default_npm_executable", lambda: "npm")
    _write_text(tmp_path, "README.md")
    _write_text(tmp_path, ".github/scripts/replace-template-placeholders.py")
    _write_marker(tmp_path, "template_sync:\n  included_modules:\n  - markdown\n")
    _write_text(tmp_path, ".template-sync/scripts/validate_marker.py")
    _write_text(
        tmp_path,
        "package.json",
        '{"scripts":{"lint:md":"markdownlint .","lint:md:nested":"markdownlint ."}}\n',
    )

    plan = first_adoption.build_check_plan(tmp_path, ("README.md",))

    assert [command.group_label for command in plan.commands] == [
        "pre-commit",
        "placeholder-scan",
        "marker-validation",
        "markdown-script",
        "markdown-script",
    ]
    assert plan.commands[0].command == ("pre-commit", "run", "--files", "README.md")
    assert plan.commands[-2].command == ("npm", "run", "lint:md")
    assert plan.commands[-1].command == ("npm", "run", "lint:md:nested")


def test_canonical_downstream_pytest_command_uses_runner_interpreter() -> None:
    """The pytest gate command is a single sys.executable module-form tuple."""
    assert first_adoption.downstream_pytest_command() == (
        sys.executable,
        "-m",
        "pytest",
        "-m",
        "not upstream_template_only",
    )


def test_downstream_pytest_candidate_paths_use_manifest_retention() -> None:
    """Candidate discovery applies the maintained candidate set and manifest state."""
    mappings = first_adoption.parse_manifest_path_mappings_text(
        "\n".join(
            [
                "template_manifest:",
                "  path_mappings:",
                "    - pattern: 'tests/test_run_first_adoption_checks.py'",
                "      requires_all: [template-sync-support]",
                "    - pattern: 'tests/test_template_manifest.py'",
                "      requires_all: [template-sync-support]",
                "    - pattern: 'tests/*.py'",
                "      requires_all: [python]",
                "",
            ]
        )
    )

    assert first_adoption.downstream_pytest_candidate_paths(
        (
            "tests/test_example.py",
            "tests/test_run_first_adoption_checks.py",
            "tests/test_template_manifest.py",
        ),
        manifest_mappings=mappings,
        retained_modules={"template-sync-support"},
    ) == ("tests/test_run_first_adoption_checks.py",)


def test_check_plan_includes_pytest_from_explicit_retained_modules(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Explicit pre-marker module state can make the pytest gate applicable."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _write_pytest_gate_manifest(tmp_path)
    _write_text(tmp_path, "pyproject.toml")
    _write_text(tmp_path, "tests/test_run_first_adoption_checks.py")

    plan = first_adoption.build_check_plan(
        tmp_path,
        ("pyproject.toml", "tests/test_run_first_adoption_checks.py"),
        retained_modules=("python", "template-sync-support"),
    )

    assert plan.commands[-1] == first_adoption.PlannedCommand(
        group_label=first_adoption.PYTEST_GROUP,
        command=first_adoption.downstream_pytest_command(),
    )
    assert not any("Pytest gate skipped" in note for note in plan.notes)


def test_pytest_ini_plans_support_tests_without_python_product(tmp_path: Path) -> None:
    """A real downstream configuration enables the retained support test command."""
    _write_pytest_gate_manifest(tmp_path)
    _write_text(tmp_path, "pytest.ini", "[pytest]\n")
    files = ("pytest.ini", "tests/test_run_first_adoption_checks.py")
    plan = first_adoption.pytest_gate_plan(
        tmp_path, files, explicit_retained_modules=("template-sync-support",)
    )
    assert [item.command for item in plan.commands] == [
        (sys.executable, "-m", "pytest", "-m", "not upstream_template_only")
    ]


def test_pytest_ini_does_not_enable_excluded_tests(tmp_path: Path) -> None:
    """Configuration presence does not override test ownership."""
    _write_pytest_gate_manifest(tmp_path)
    _write_text(tmp_path, "pytest.ini", "[pytest]\n")
    plan = first_adoption.pytest_gate_plan(
        tmp_path,
        ("pytest.ini", "tests/test_example.py"),
        explicit_retained_modules=("template-sync-support",),
    )
    assert not plan.commands


def test_pytest_ini_gate_propagates_test_failure(tmp_path: Path, monkeypatch: Any) -> None:
    """A scheduled support test failure remains a failed adoption check."""
    _write_pytest_gate_manifest(tmp_path)
    _write_text(tmp_path, "pytest.ini", "[pytest]\n")
    _write_text(tmp_path, "tests/test_run_first_adoption_checks.py")
    _run_git(tmp_path, "init")
    commands: list[tuple[str, ...]] = []

    def runner(command: Sequence[str], _root: Path) -> int:
        commands.append(tuple(command))
        return 7 if "pytest" in command else 0

    result = first_adoption.run_first_adoption_checks(
        tmp_path, retained_modules=("template-sync-support",),
        command_runner=runner, git_status_reader=lambda _root: (), stdout=io.StringIO(),
    )
    assert any("pytest" in command for command in commands)
    assert result != 0


def test_pytest_ini_gate_positive_control_detects_removed_guard(tmp_path: Path, monkeypatch: Any) -> None:
    """The expected command distinguishes a removed configuration predicate."""
    _write_pytest_gate_manifest(tmp_path)
    _write_text(tmp_path, "pytest.ini", "[pytest]\n")
    monkeypatch.setattr(first_adoption, "is_pytest_configuration_retained", lambda *a, **kw: False)
    plan = first_adoption.pytest_gate_plan(
        tmp_path, ("pytest.ini", "tests/test_run_first_adoption_checks.py"),
        explicit_retained_modules=("template-sync-support",),
    )
    with pytest.raises(AssertionError):
        assert [item.command for item in plan.commands] == [
            (sys.executable, "-m", "pytest", "-m", "not upstream_template_only")
        ]


def test_missing_pytest_ini_does_not_plan_support_tests(tmp_path: Path) -> None:
    """Support test retention still requires a configuration file."""
    _write_pytest_gate_manifest(tmp_path)
    plan = first_adoption.pytest_gate_plan(
        tmp_path,
        ("tests/test_run_first_adoption_checks.py",),
        explicit_retained_modules=("template-sync-support",),
    )
    assert not plan.commands


def test_explicitly_excluded_pytest_ini_stays_excluded(tmp_path: Path) -> None:
    """A manifest configuration restriction remains authoritative."""
    _write_manifest(tmp_path, "template_manifest:\n  path_mappings:\n    - pattern: 'pytest.ini'\n      requires_all: [python]\n")
    _write_text(tmp_path, "pytest.ini", "[pytest]\n")
    plan = first_adoption.pytest_gate_plan(
        tmp_path, ("pytest.ini", "tests/test_example.py"),
        explicit_retained_modules=("template-sync-support",),
    )
    assert not plan.commands
    assert any("configuration pruned" in note for note in plan.notes)


def test_marker_state_overrides_explicit_modules_and_file_presence(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Marker-derived module state prevents file presence from implying pytest."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _write_pytest_gate_manifest(tmp_path)
    _write_text(tmp_path, "pyproject.toml")
    _write_text(tmp_path, "tests/test_run_first_adoption_checks.py")
    _write_marker(
        tmp_path,
        "template_sync:\n  included_modules:\n    - template-sync-support\n",
    )
    _write_text(tmp_path, ".template-sync/scripts/validate_marker.py")

    plan = first_adoption.build_check_plan(
        tmp_path,
        ("pyproject.toml", "tests/test_run_first_adoption_checks.py"),
        retained_modules=("python", "template-sync-support"),
    )

    assert not any(command.group_label == first_adoption.PYTEST_GROUP for command in plan.commands)
    assert any("pytest configuration pruned" in note for note in plan.notes)


def test_check_plan_uses_file_presence_fallback_without_module_decision(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Without marker or explicit state, pyproject plus test files plans pytest."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _write_text(tmp_path, "pyproject.toml")
    _write_text(tmp_path, "tests/test_example.py")

    plan = first_adoption.build_check_plan(tmp_path, ("pyproject.toml", "tests/test_example.py"))

    assert any(command.group_label == first_adoption.PYTEST_GROUP for command in plan.commands)


def test_check_plan_skips_pytest_when_configuration_pruned_but_tests_remain(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """The pytest gate needs retained config even when candidate tests remain."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _write_pytest_gate_manifest(tmp_path)
    _write_text(tmp_path, "pyproject.toml")
    _write_text(tmp_path, "tests/test_run_first_adoption_checks.py")

    plan = first_adoption.build_check_plan(
        tmp_path,
        ("pyproject.toml", "tests/test_run_first_adoption_checks.py"),
        retained_modules=("template-sync-support",),
    )

    assert not any(command.group_label == first_adoption.PYTEST_GROUP for command in plan.commands)
    assert any("pytest configuration pruned" in note for note in plan.notes)


def test_check_plan_skips_pytest_when_no_candidate_tests_retained(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """A retained pytest config alone is not enough to plan the downstream gate."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _write_pytest_gate_manifest(tmp_path)
    _write_text(tmp_path, "pyproject.toml")

    plan = first_adoption.build_check_plan(
        tmp_path,
        ("pyproject.toml",),
        retained_modules=("python",),
    )

    assert not any(command.group_label == first_adoption.PYTEST_GROUP for command in plan.commands)
    assert any("no downstream pytest candidate paths" in note for note in plan.notes)


def test_plan_only_reports_pytest_command_without_availability_probe(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Plan-only prints the pytest gate and points environment checks to doctor."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_pytest_gate_manifest(tmp_path)
    _write_text(tmp_path, "pyproject.toml")
    _write_text(tmp_path, "tests/test_run_first_adoption_checks.py")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        plan_only=True,
        retained_modules=("python", "template-sync-support"),
        command_runner=_recording_runner(commands),
        git_status_reader=_queued_status_reader((), ()),
        stdout=stdout,
    )

    output = stdout.getvalue()
    assert result == 0
    assert commands == []
    assert (
        f"[pytest] {first_adoption.format_command(first_adoption.downstream_pytest_command())}"
        in output
    )
    assert "sys.executable -m pytest" not in output
    assert "pytest tool availability was not evaluated" in output
    assert "run --doctor" in output
    assert "Doctor probes executed" not in output


def test_doctor_recommends_exact_canonical_pytest_command() -> None:
    """Doctor recommendations render the same command tuple as the gate plan."""
    recommendations = first_adoption.doctor_recommendations(
        (
            _doctor_result(
                first_adoption.PYTEST_GROUP,
                "pytest console",
                ("pytest", "--version"),
                first_adoption.PROBE_AVAILABLE,
            ),
            _doctor_result(
                first_adoption.PYTEST_GROUP,
                "pytest module (sys.executable -m)",
                first_adoption.downstream_pytest_version_command(),
                first_adoption.PROBE_AVAILABLE,
            ),
        )
    )

    assert (
        f"Use pytest invocation: "
        f"{first_adoption.format_command(first_adoption.downstream_pytest_command())}"
    ) in recommendations
    assert "Use pytest invocation: pytest" not in recommendations
    assert not any("sys.executable -m" in recommendation for recommendation in recommendations)


def test_doctor_pytest_gate_verdict_ignores_alternate_working_forms() -> None:
    """A working alternate pytest form does not make the gate interpreter ready."""
    recommendations = first_adoption.doctor_recommendations(
        (
            _doctor_result(
                first_adoption.PYTEST_GROUP,
                "pytest console",
                ("pytest", "--version"),
                first_adoption.PROBE_AVAILABLE,
            ),
            _doctor_result(
                first_adoption.PYTEST_GROUP,
                "pytest module (sys.executable -m)",
                first_adoption.downstream_pytest_version_command(),
                first_adoption.PROBE_FAILED,
            ),
            _doctor_result(
                first_adoption.PYTEST_GROUP,
                "pytest module (python3 -m)",
                ("python3", "-m", "pytest", "--version"),
                first_adoption.PROBE_AVAILABLE,
            ),
        )
    )

    assert any(
        "No working pytest module invocation was detected for the runner's Python interpreter"
        in recommendation
        for recommendation in recommendations
    )
    assert not any(
        recommendation.startswith("Use pytest invocation:") for recommendation in recommendations
    )


def test_quality_reports_are_planned_before_fixers_when_helper_exists(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """The first-adoption runner inventories quality debt before fix-mode commands."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _write_text(tmp_path, ".template-sync/scripts/first_adoption_quality_reports.py")

    plan = first_adoption.build_check_plan(
        tmp_path,
        ("README.md",),
        run_mode=first_adoption.FIX_MODE,
    )

    assert [command.group_label for command in plan.commands[:5]] == [
        "quality-report",
        "quality-report",
        "quality-report",
        "markdown-fixer",
        "pre-commit",
    ]
    assert plan.commands[0].command[-1] == "line-endings"
    assert plan.commands[1].command[-1] == "path-references"
    assert plan.commands[2].command[-1] == "powershell"
    assert plan.commands[3].command[-2:] == ("markdown", "--fix")


def test_powershell_quality_report_is_skipped_when_marker_excludes_powershell(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Marker-derived module state suppresses stale PowerShell report commands."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _write_text(tmp_path, ".template-sync/scripts/first_adoption_quality_reports.py")
    _write_text(tmp_path, ".template-sync/scripts/validate_marker.py")
    _write_text(
        tmp_path,
        ".template-sync/marker.yml",
        "template_sync:\n  included_modules:\n    - baseline\n    - template-sync-support\n",
    )

    plan = first_adoption.build_check_plan(
        tmp_path,
        ("README.md",),
        run_mode=first_adoption.FIX_MODE,
    )

    quality_modes = [
        command.command[-1]
        for command in plan.commands
        if command.group_label == first_adoption.QUALITY_REPORT_GROUP
    ]
    assert quality_modes == ["line-endings", "path-references"]
    assert all("powershell" not in command.command for command in plan.commands)


def test_unsafe_candidate_exit_code_takes_precedence(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Unsafe PowerShell candidates propagate exit code 3 ahead of other failures."""
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    _write_text(tmp_path, ".template-sync/scripts/first_adoption_quality_reports.py")
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )

    def fake_runner(command: Sequence[str], _repo_root: Path) -> int:
        if command[-1] == "powershell":
            return cast(int, first_adoption.UNSAFE_CANDIDATE_EXIT_CODE)
        if command[0] == "pre-commit":
            return 1
        return 0

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=fake_runner,
        git_status_reader=lambda _repo_root: (),
        stdout=io.StringIO(),
    )

    assert result == first_adoption.UNSAFE_CANDIDATE_EXIT_CODE


def test_azure_marker_plans_host_setup_quality_report(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Azure DevOps module retention schedules the host setup report."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _write_text(tmp_path, ".template-sync/scripts/first_adoption_quality_reports.py")
    _write_text(tmp_path, ".template-sync/scripts/validate_marker.py")
    _write_text(
        tmp_path,
        ".template-sync/marker.yml",
        "template_sync:\n  included_modules:\n    - baseline\n    - azure-devops-platform\n",
    )

    plan = first_adoption.build_check_plan(
        tmp_path,
        ("README.md",),
        run_mode=first_adoption.FIX_MODE,
    )

    quality_modes = [
        command.command[-1]
        for command in plan.commands
        if command.group_label == first_adoption.QUALITY_REPORT_GROUP
    ]
    assert quality_modes == ["line-endings", "path-references", "host-setup"]


def test_marker_module_detection_ignores_sibling_string_lists(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Sibling marker lists like issue_labels are not read as retained modules."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _write_text(tmp_path, ".template-sync/scripts/first_adoption_quality_reports.py")
    _write_text(tmp_path, ".template-sync/scripts/validate_marker.py")
    # Block sequence items render at the key's own indentation under PyYAML's
    # default safe_dump style, and issue_labels can reuse module-like names.
    _write_text(
        tmp_path,
        ".template-sync/marker.yml",
        (
            "template_sync:\n"
            "  source_repo: https://github.com/franklesniak/copilot-repo-template\n"
            "  included_modules:\n"
            "  - baseline\n"
            "  - template-sync-support\n"
            "  issue_label_policy: custom\n"
            "  issue_labels:\n"
            "  - powershell\n"
            "  - markdown\n"
        ),
    )

    assert first_adoption.marker_included_modules(tmp_path) == frozenset(
        {"baseline", "template-sync-support"}
    )
    assert not first_adoption.marker_includes_markdown_module(tmp_path)

    plan = first_adoption.build_check_plan(
        tmp_path,
        ("README.md",),
        run_mode=first_adoption.FIX_MODE,
    )

    assert all("powershell" not in command.command for command in plan.commands)


def test_marker_module_detection_skips_comment_lines_inside_block(
    tmp_path: Path,
) -> None:
    """Full-line comments inside included_modules do not end the module scan."""
    # A hand-added comment at the block indentation must be skipped, not treated
    # as an end-of-block marker that drops the modules listed after it.
    _write_marker(
        tmp_path,
        (
            "template_sync:\n"
            "  included_modules:\n"
            "  - baseline\n"
            "  # downstream note kept after editing\n"
            "  - powershell\n"
            "  - template-sync-support\n"
            "  issue_labels:\n"
            "  - markdown\n"
        ),
    )

    assert first_adoption.marker_included_modules(tmp_path) == frozenset(
        {"baseline", "powershell", "template-sync-support"}
    )


def test_tracked_only_files_are_collected(tmp_path: Path) -> None:
    """Tracked files staged in the index are included in the pre-commit file list."""
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "tracked.txt")
    _run_git(tmp_path, "add", "tracked.txt")

    collection = first_adoption.collect_present_regular_files(tmp_path)

    assert collection.files == ("tracked.txt",)
    assert collection.skipped_non_regular_paths == ()


def test_untracked_only_files_are_collected(tmp_path: Path) -> None:
    """Untracked non-ignored files are included before the first adoption commit."""
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "untracked.txt")

    collection = first_adoption.collect_present_regular_files(tmp_path)

    assert collection.files == ("untracked.txt",)


def test_ignored_files_are_not_collected(tmp_path: Path) -> None:
    """Ignored files are excluded by the Git-visible working-tree query."""
    _run_git(tmp_path, "init")
    _write_text(tmp_path, ".gitignore", "ignored.txt\n")
    _write_text(tmp_path, "ignored.txt")

    collection = first_adoption.collect_present_regular_files(tmp_path)

    assert "ignored.txt" not in collection.files
    assert collection.files == (".gitignore",)


def test_plan_only_prints_collection_notes_and_plan_without_running(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Plan-only mode prints discovery and plan details without validation commands."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    _write_marker(tmp_path, "template_sync:\n  included_modules:\n  - markdown\n")
    _write_text(tmp_path, ".template-sync/scripts/validate_marker.py")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        plan_only=True,
        command_runner=_recording_runner(commands),
        time_source=_queued_time_source(_utc_time(0), _utc_time(2)),
        stdout=stdout,
    )

    output = stdout.getvalue()
    assert result == 0
    assert commands == []
    assert "Run mode: check" in output
    assert "Collected 3 tracked or untracked non-ignored regular file(s)." in output
    assert "Git-visible regular file(s):" in output
    assert "  - .template-sync/marker.yml" in output
    assert "  - .template-sync/scripts/validate_marker.py" in output
    assert "  - README.md" in output
    assert "Markdown module appears retained" in output
    assert "Planned validation commands (2):" in output
    assert "  1. [pre-commit] pre-commit run --files" in output
    assert "  2. [marker-validation]" in output
    assert "Plan-only mode: validation commands were not run." in output
    assert "No Git status changes were detected" in output
    assert "Total elapsed time: 2.000s" in output
    assert "Command 1/2 [pre-commit] start time" not in output


def test_fix_mode_plan_only_prints_fix_mode_without_running(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Plan-only mode can preview the explicit fix-mode command plan."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        run_mode=first_adoption.FIX_MODE,
        plan_only=True,
        command_runner=_recording_runner(commands),
        stdout=stdout,
    )

    output = stdout.getvalue()
    assert result == 0
    assert commands == []
    assert "Run mode: fix" in output
    assert "Plan-only mode: validation commands were not run." in output


def test_marker_absent_does_not_run_marker_validator(tmp_path: Path) -> None:
    """Marker validation is skipped when no downstream marker exists."""
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    commands: list[tuple[str, ...]] = []

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        stdout=io.StringIO(),
    )

    assert result == 0
    assert commands[0][-3:] == ("run", "--files", "README.md")
    assert not any("validate_marker.py" in arg for cmd in commands for arg in cmd)


def test_timing_output_uses_injected_time_source_and_prints_cold_start_guidance(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Command timing and total elapsed output are deterministic in tests."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    _write_text(tmp_path, ".github/scripts/replace-template-placeholders.py")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        time_source=_queued_time_source(
            _utc_time(0),
            _utc_time(5),
            _utc_time(9),
            _utc_time(10),
            _utc_time(13),
            _utc_time(20),
        ),
        stdout=stdout,
    )

    output = stdout.getvalue()
    assert result == 0
    assert len(commands) == 2
    assert "Planned validation commands (2):" in output
    assert "Cold pre-commit hook environment bootstrapping may take several minutes" in output
    assert "Command 1/2 [pre-commit] start time (UTC): 2026-06-03T12:00:05Z" in output
    assert "Command 1/2 [pre-commit] completed with exit code 0" in output
    assert "Command 1/2 [pre-commit] end time (UTC): 2026-06-03T12:00:09Z" in output
    assert "Command 1/2 [pre-commit] elapsed time: 4.000s" in output
    assert "Command 2/2 [placeholder-scan] start time (UTC): 2026-06-03T12:00:10Z" in output
    assert "Command 2/2 [placeholder-scan] elapsed time: 3.000s" in output
    assert "No Git status changes were detected" in output
    assert "Total elapsed time: 20.000s" in output


def test_unchanged_git_status_passes_after_validation(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """A stable dirty status from before the run is not treated as a new mutation."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        git_status_reader=_queued_status_reader(("?? README.md",), ("?? README.md",)),
        stdout=stdout,
    )

    assert result == 0
    assert commands == [("pre-commit", "run", "--files", "README.md")]
    assert "No Git status changes were detected" in stdout.getvalue()


def test_changed_git_status_exits_distinctly_after_validation(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """A command-boundary Git status change is reported and exits distinctly."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        git_status_reader=_queued_status_reader((), (" M README.md", "?? generated.txt")),
        stdout=stdout,
    )

    output = stdout.getvalue()
    transition_section = _transition_section(output)
    follow_up = _changed_file_follow_up(output)
    assert result == first_adoption.CHANGED_FILES_EXIT_CODE
    assert commands == [("pre-commit", "run", "--files", "README.md")]
    assert "Git changed-file summary:" in output
    assert "  Before invocation:" in output
    assert "    - <none>" in output
    assert "  After invocation:" in output
    assert "    -  M README.md" in output
    assert "    - ?? generated.txt" in output
    assert "Command 1/1 [pre-commit] pre-commit run --files README.md" in transition_section
    assert "  added status entries:" in transition_section
    assert "    -  M README.md" in transition_section
    assert "    - ?? generated.txt" in transition_section
    assert "  removed status entries:" in transition_section
    assert "    - <none>" in transition_section
    assert "Final Git status changed during this invocation" in follow_up
    assert "git status for changed-path inventory" in follow_up
    assert "git diff --cached" in follow_up
    assert "python .template-sync/scripts/run_first_adoption_checks.py --check" in follow_up
    assert "python .template-sync/scripts/run_first_adoption_checks.py --fix" in follow_up


def test_fix_mode_tolerates_mutations_without_failing(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Fix mode reports command-induced mutations but exits zero."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        run_mode=first_adoption.FIX_MODE,
        command_runner=_recording_runner(commands),
        git_status_reader=_queued_status_reader((), (" M README.md",)),
        stdout=stdout,
    )

    output = stdout.getvalue()
    transition_section = _transition_section(output)
    follow_up = _changed_file_follow_up(output)
    assert result == 0
    assert commands == [("pre-commit", "run", "--files", "README.md")]
    assert "Run mode: fix" in output
    assert "Git changed-file summary:" in output
    assert "    -  M README.md" in output
    assert "Command 1/1 [pre-commit] pre-commit run --files README.md" in transition_section
    assert "    -  M README.md" in transition_section
    assert "Fix mode observed final Git status changes" in follow_up
    assert "does not fail solely because files were modified" in follow_up
    assert "any planned command that exits nonzero still fails the run" in follow_up
    assert "python .template-sync/scripts/run_first_adoption_checks.py --check" in follow_up
    assert "Command failures are reported separately" not in follow_up


def test_check_mode_reports_command_failure_before_changed_files(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """A failing command takes exit-code precedence over a mutated work tree."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    stdout = io.StringIO()

    def failing_runner(command: Sequence[str], _repo_root: Path) -> int:
        del command
        return 1

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=failing_runner,
        git_status_reader=_queued_status_reader((), (" M README.md",)),
        stdout=stdout,
    )

    output = stdout.getvalue()
    follow_up = _changed_file_follow_up(output)
    assert result == 1
    assert "Git changed-file summary:" in output
    assert "    -  M README.md" in output
    assert "Command failures are reported separately" in follow_up
    assert "does not resolve a failed command" in follow_up
    assert "First-adoption checks failed:" in output
    assert "exited with 1" in output


def test_multi_command_status_transitions_report_status_entries_by_boundary(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Multiple command boundaries report only entries changed at each boundary."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    _write_marker(tmp_path, "template_sync:\n  included_modules:\n  - baseline\n")
    _write_text(tmp_path, ".template-sync/scripts/validate_marker.py")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        git_status_reader=_queued_status_reader(
            (" M existing.md",),
            (" M existing.md", " M status.txt", "?? generated.txt"),
            (" M existing.md", "A  status.txt"),
        ),
        stdout=stdout,
    )

    output = stdout.getvalue()
    transition_section = _transition_section(output)
    assert result == first_adoption.CHANGED_FILES_EXIT_CODE
    assert len(commands) == 2
    assert "Command 1/2 [pre-commit]" in transition_section
    assert "Command 2/2 [marker-validation]" in transition_section
    first_boundary = _output_section(
        transition_section,
        "Command 1/2 [pre-commit]",
        ("Command 2/2 [marker-validation]",),
    )
    second_boundary = _output_section(
        transition_section,
        "Command 2/2 [marker-validation]",
        (),
    )
    assert "    -  M status.txt" in first_boundary
    assert "    - ?? generated.txt" in first_boundary
    assert "  removed status entries:\n    - <none>" in first_boundary
    assert "    - A  status.txt" in second_boundary
    assert "  removed status entries:" in second_boundary
    assert "    -  M status.txt" in second_boundary
    assert "    - ?? generated.txt" in second_boundary
    assert "existing.md" not in transition_section


def test_repeated_status_entry_transitions_are_reported_each_time(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """The same opaque status entry can be added, removed, and added again."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    _write_text(tmp_path, ".github/scripts/replace-template-placeholders.py")
    _write_marker(tmp_path, "template_sync:\n  included_modules:\n  - baseline\n")
    _write_text(tmp_path, ".template-sync/scripts/validate_marker.py")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        git_status_reader=_queued_status_reader(
            (),
            ("?? generated.txt",),
            (),
            ("?? generated.txt",),
        ),
        stdout=stdout,
    )

    transition_section = _transition_section(stdout.getvalue())
    assert result == first_adoption.CHANGED_FILES_EXIT_CODE
    assert len(commands) == 3
    assert transition_section.count("    - ?? generated.txt") == 3
    assert "Command 1/3 [pre-commit]" in transition_section
    assert "Command 2/3 [placeholder-scan]" in transition_section
    assert "Command 3/3 [marker-validation]" in transition_section


def test_transient_status_transitions_report_state_b_without_exit_code(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Transient command-boundary transitions are reported without changing exit code."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    _write_marker(tmp_path, "template_sync:\n  included_modules:\n  - baseline\n")
    _write_text(tmp_path, ".template-sync/scripts/validate_marker.py")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        git_status_reader=_queued_status_reader((), ("?? transient.txt",), ()),
        stdout=stdout,
    )

    output = stdout.getvalue()
    transition_section = _transition_section(output)
    follow_up = _changed_file_follow_up(output)
    assert result == 0
    assert len(commands) == 2
    assert "Net Git status is unchanged at the final snapshot" in output
    assert "No Git status changes were detected during this invocation" not in output
    assert "Command 1/2 [pre-commit]" in transition_section
    assert "Command 2/2 [marker-validation]" in transition_section
    assert "    - ?? transient.txt" in transition_section
    assert "No net Git status changes remain at the final snapshot" in follow_up
    assert "keep or revert" not in follow_up
    assert "python .template-sync/scripts/run_first_adoption_checks.py --check" in follow_up
    assert "python .template-sync/scripts/run_first_adoption_checks.py --fix" in follow_up


def test_no_command_status_change_is_reported_without_attribution(
    tmp_path: Path,
) -> None:
    """Aggregate status changes without planned commands are reported as unattributed."""
    _run_git(tmp_path, "init")
    status_reader = _ExactStatusReader((), ("?? generated.txt",))
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        git_status_reader=status_reader,
        stdout=stdout,
    )

    output = stdout.getvalue()
    transition_section = _transition_section(output)
    assert result == first_adoption.CHANGED_FILES_EXIT_CODE
    assert commands == []
    assert status_reader.read_count == 2
    status_reader.assert_consumed()
    assert "  Before invocation:" in output
    assert "  After invocation:" in output
    assert "No validation commands were planned" in transition_section
    assert "Command 1/" not in transition_section


def test_plan_only_status_change_is_reported_without_executed_boundary(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Plan-only can print commands but cannot attribute status changes to execution."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    status_reader = _ExactStatusReader((), ("?? generated.txt",))
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        plan_only=True,
        command_runner=_recording_runner(commands),
        git_status_reader=status_reader,
        stdout=stdout,
    )

    output = stdout.getvalue()
    transition_section = _transition_section(output)
    assert result == first_adoption.CHANGED_FILES_EXIT_CODE
    assert commands == []
    assert status_reader.read_count == 2
    status_reader.assert_consumed()
    assert "Planned validation commands (1):" in output
    assert "Plan-only mode: validation commands were not run." in output
    assert "Plan-only mode printed planned validation commands" in transition_section
    assert "Command 1/1 [pre-commit] start time" not in output


def test_fix_mode_command_failure_still_fails_with_status_transition_summary(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Fix mode tolerates zero-exit mutations, not nonzero planned commands."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    stdout = io.StringIO()

    def failing_runner(command: Sequence[str], _repo_root: Path) -> int:
        del command
        return 1

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        run_mode=first_adoption.FIX_MODE,
        command_runner=failing_runner,
        git_status_reader=_queued_status_reader((), (" M README.md",)),
        stdout=stdout,
    )

    output = stdout.getvalue()
    transition_section = _transition_section(output)
    follow_up = _changed_file_follow_up(output)
    assert result == 1
    assert "Command 1/1 [pre-commit]" in transition_section
    assert "    -  M README.md" in transition_section
    assert "Fix mode observed final Git status changes" in follow_up
    assert "any planned command that exits nonzero still fails the run" in follow_up
    assert "Command failures are reported separately" in follow_up
    assert "First-adoption checks failed:" in output


def test_no_failure_status_change_omits_command_failure_status_scope(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """The command-failure status-scoping sentence appears only for failures."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner([]),
        git_status_reader=_queued_status_reader((), (" M README.md",)),
        stdout=stdout,
    )

    assert result == first_adoption.CHANGED_FILES_EXIT_CODE
    assert "Command failures are reported separately" not in _changed_file_follow_up(
        stdout.getvalue()
    )


def test_normal_loop_consumes_one_status_snapshot_per_executed_boundary(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Normal execution reads one initial snapshot plus one per planned command."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    _write_marker(tmp_path, "template_sync:\n  included_modules:\n  - baseline\n")
    _write_text(tmp_path, ".template-sync/scripts/validate_marker.py")
    status_reader = _ExactStatusReader((), ("?? one.txt",), ())
    commands: list[tuple[str, ...]] = []

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        git_status_reader=status_reader,
        stdout=io.StringIO(),
    )

    assert result == 0
    assert len(commands) == 2
    assert status_reader.read_count == 3
    status_reader.assert_consumed()


def test_no_command_and_plan_only_paths_consume_two_status_snapshots(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Non-executing branches keep their before/after status snapshot pair."""
    _run_git(tmp_path, "init")
    no_command_reader = _ExactStatusReader((), ())

    no_command_result = first_adoption.run_first_adoption_checks(
        tmp_path,
        git_status_reader=no_command_reader,
        stdout=io.StringIO(),
    )

    assert no_command_result == 0
    assert no_command_reader.read_count == 2
    no_command_reader.assert_consumed()

    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _write_text(tmp_path, "README.md")
    plan_only_reader = _ExactStatusReader((), ())

    plan_only_result = first_adoption.run_first_adoption_checks(
        tmp_path,
        plan_only=True,
        git_status_reader=plan_only_reader,
        stdout=io.StringIO(),
    )

    assert plan_only_result == 0
    assert plan_only_reader.read_count == 2
    plan_only_reader.assert_consumed()


def test_initial_snapshot_failure_propagates_before_commands(
    tmp_path: Path,
) -> None:
    """An initial status snapshot failure aborts through FirstAdoptionCheckError."""
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    with pytest.raises(first_adoption.FirstAdoptionCheckError):
        first_adoption.run_first_adoption_checks(
            tmp_path,
            command_runner=_recording_runner(commands),
            git_status_reader=_failing_status_reader(
                first_adoption.FirstAdoptionCheckError("status failed")
            ),
            stdout=stdout,
        )

    assert commands == []
    assert "Git changed-file summary:" not in stdout.getvalue()


def test_mid_loop_snapshot_failure_aborts_without_partial_transition_summary(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """A mid-loop status snapshot failure aborts before rendering partial attribution."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    _write_marker(tmp_path, "template_sync:\n  included_modules:\n  - baseline\n")
    _write_text(tmp_path, ".template-sync/scripts/validate_marker.py")
    commands: list[tuple[str, ...]] = []
    stdout = io.StringIO()

    with pytest.raises(first_adoption.FirstAdoptionCheckError):
        first_adoption.run_first_adoption_checks(
            tmp_path,
            command_runner=_recording_runner(commands),
            git_status_reader=_failing_status_reader(
                (),
                first_adoption.FirstAdoptionCheckError("status failed"),
            ),
            stdout=stdout,
        )

    output = stdout.getvalue()
    assert len(commands) == 1
    assert "Command 1/2 [pre-commit] completed with exit code 0" in output
    assert "Command-boundary status transitions:" not in output


def test_status_summary_output_is_deterministic_for_identical_inputs(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Identical commands and injected snapshots render identical sorted output."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")

    def run_once() -> str:
        stdout = io.StringIO()
        result = first_adoption.run_first_adoption_checks(
            tmp_path,
            command_runner=_recording_runner([]),
            git_status_reader=_queued_status_reader(
                ("",),
                ("?? z.txt", " M a.txt", ""),
            ),
            time_source=_queued_time_source(
                _utc_time(0),
                _utc_time(1),
                _utc_time(2),
                _utc_time(3),
            ),
            stdout=stdout,
        )
        assert result == first_adoption.CHANGED_FILES_EXIT_CODE
        return stdout.getvalue()

    first_output = run_once()
    second_output = run_once()

    assert first_output == second_output
    transition_section = _transition_section(first_output)
    assert transition_section.index("    -  M a.txt") < transition_section.index("    - ?? z.txt")


def test_git_status_lines_reports_oserror_as_check_error(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """A missing git executable surfaces as FirstAdoptionCheckError, not a raw OSError."""

    def raise_oserror(*_args: object, **_kwargs: object) -> object:
        raise FileNotFoundError(2, "No such file or directory")

    monkeypatch.setattr(first_adoption.subprocess, "run", raise_oserror)

    with pytest.raises(first_adoption.FirstAdoptionCheckError) as excinfo:
        first_adoption.git_status_lines(tmp_path)

    assert "Unable to inspect Git changed files" in str(excinfo.value)


def test_collect_present_regular_files_reports_oserror_as_check_error(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """A missing git executable surfaces as FirstAdoptionCheckError, not a raw OSError."""

    def raise_oserror(*_args: object, **_kwargs: object) -> object:
        raise FileNotFoundError(2, "No such file or directory")

    monkeypatch.setattr(first_adoption.subprocess, "run", raise_oserror)

    with pytest.raises(first_adoption.FirstAdoptionCheckError) as excinfo:
        first_adoption.collect_present_regular_files(tmp_path)

    assert "Unable to inspect Git-visible files" in str(excinfo.value)


def test_marker_present_runs_marker_validator(tmp_path: Path) -> None:
    """A present marker adds the marker validator after the pre-commit file check."""
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    _write_text(tmp_path, ".template-sync/marker.yml", "template_sync:\n")
    _write_text(tmp_path, ".template-sync/scripts/validate_marker.py", "print('ok')\n")
    commands: list[tuple[str, ...]] = []

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        stdout=io.StringIO(),
    )

    assert result == 0
    assert commands[-1] == (
        sys.executable,
        ".template-sync/scripts/validate_marker.py",
        "--require-marker",
    )
    assert ".template-sync/marker.yml" in commands[0]


def test_multiple_failures_are_reported_by_default(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """The runner continues through the plan and reports every failing command."""
    monkeypatch.setattr(
        first_adoption,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    _write_text(tmp_path, ".github/scripts/replace-template-placeholders.py")
    return_codes = [7, 3]
    records: list[tuple[str, ...]] = []

    def run(command: Sequence[str], _repo_root: Path) -> int:
        records.append(tuple(command))
        return return_codes.pop(0)

    stdout = io.StringIO()

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=run,
        time_source=_incrementing_time_source(),
        stdout=stdout,
    )

    output = stdout.getvalue()
    assert result == 1
    assert len(records) == 2
    assert "Command 1/2 [pre-commit] completed with exit code 7" in output
    assert "Command 2/2 [placeholder-scan] completed with exit code 3" in output
    assert "First-adoption checks failed:" in output
    assert "  - pre-commit: pre-commit run --files" in output
    assert "exited with 7" in output
    assert "  - placeholder-scan:" in output
    assert "exited with 3" in output


def test_placeholder_script_runs_when_present(tmp_path: Path) -> None:
    """A present placeholder helper adds the scan command."""
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    _write_text(tmp_path, ".github/scripts/replace-template-placeholders.py", "print('ok')\n")
    commands: list[tuple[str, ...]] = []

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        stdout=io.StringIO(),
    )

    assert result == 0
    assert (
        sys.executable,
        ".github/scripts/replace-template-placeholders.py",
        "scan",
    ) in commands


@pytest.mark.parametrize(
    "marker_text",
    [
        pytest.param(
            "\n".join(
                [
                    "template_sync:",
                    "  included_modules:",
                    "  - baseline",
                    "  - markdown",
                    "  issue_labels:",
                    "  - triage",
                    "",
                ]
            ),
            id="indentless-sequence",
        ),
        pytest.param(
            "\n".join(
                [
                    "template_sync:",
                    "  included_modules:",
                    "    - baseline",
                    "    - markdown",
                    "  issue_labels:",
                    "    - triage",
                    "",
                ]
            ),
            id="visually-indented-sequence",
        ),
        pytest.param(
            "\n".join(
                [
                    "template_sync:",
                    "    included_modules:",
                    "    - baseline",
                    "    - markdown",
                    "    issue_labels:",
                    "    - triage",
                    "",
                ]
            ),
            id="four-space-child-indent",
        ),
        pytest.param(
            "\n".join(
                [
                    "template_sync: # downstream marker",
                    "  included_modules: # retained modules",
                    "  - baseline",
                    '  - "markdown" # exact module',
                    "",
                ]
            ),
            id="key-comments",
        ),
    ],
)
def test_marker_includes_markdown_module_detects_scoped_module(
    tmp_path: Path,
    marker_text: str,
) -> None:
    """The marker scan detects markdown only in top-level retained modules."""
    _write_marker(tmp_path, marker_text)

    assert first_adoption.marker_includes_markdown_module(tmp_path) is True


@pytest.mark.parametrize(
    "item_line",
    [
        "- markdown",
        "- 'markdown'",
        '- "markdown"',
        "- markdown # note",
        "- 'markdown' # note",
        '- "markdown" # note',
    ],
)
def test_marker_includes_markdown_module_accepts_exact_scalar_items(
    tmp_path: Path,
    item_line: str,
) -> None:
    """Exact unquoted and matching-quoted markdown items are recognized."""
    _write_marker(tmp_path, f"template_sync:\n  included_modules:\n  {item_line}\n")

    assert first_adoption.marker_includes_markdown_module(tmp_path) is True


@pytest.mark.parametrize(
    "marker_text",
    [
        pytest.param(
            "\n".join(
                [
                    "template_sync:",
                    "  included_modules:",
                    "  - baseline",
                    "  issue_labels:",
                    "  - markdown",
                    "",
                ]
            ),
            id="issue-label-markdown",
        ),
        pytest.param(
            "template_sync:\n    included_modules:\n    - baseline\n",
            id="four-space-child-indent-without-markdown",
        ),
        pytest.param(
            "included_modules:\n- markdown\ntemplate_sync:\n  source_repo: example\n",
            id="top-level-included-modules",
        ),
        pytest.param(
            "root:\n  template_sync:\n    included_modules:\n    - markdown\n",
            id="nested-template-sync",
        ),
        pytest.param(
            "# template_sync:\n#   included_modules:\n#   - markdown\n",
            id="commented-template-sync",
        ),
        pytest.param(
            'notes: "template_sync:"\nincluded_modules:\n- markdown\n',
            id="scalar-template-sync",
        ),
        pytest.param(
            "template_sync:\n  source_repo: example\nother:\n  included_modules:\n  - markdown\n",
            id="included-modules-outside-template-sync",
        ),
        pytest.param(
            "template_sync:\n  some_nested_mapping:\n    included_modules:\n    - markdown\n",
            id="nested-included-modules",
        ),
        pytest.param(
            "template_sync:\n  # included_modules:\n  # - markdown\n",
            id="commented-included-modules",
        ),
        pytest.param(
            'template_sync:\n  notes: "included_modules:"\n  issue_labels:\n  - markdown\n',
            id="scalar-included-modules",
        ),
        pytest.param(
            "\n".join(
                [
                    "template_sync:",
                    "  included_modules:",
                    "  - markdown-extra",
                    "  - not-markdown",
                    "  - markdown#note",
                    "",
                ]
            ),
            id="substring-values",
        ),
        pytest.param(
            "template_sync:\n  included_modules:\n  - 'markdown\"\n  - \"markdown'\n",
            id="malformed-quotes",
        ),
        pytest.param(
            "source_repo: example\n",
            id="missing-template-sync",
        ),
        pytest.param(
            "template_sync:\n  source_repo: example\n",
            id="missing-included-modules",
        ),
        pytest.param(
            "template_sync:\n  included_modules: []\n  issue_labels:\n  - markdown\n",
            id="empty-list",
        ),
        pytest.param(
            "template_sync:\n  included_modules:\n  issue_labels:\n  - markdown\n",
            id="block-with-no-items",
        ),
    ],
)
def test_marker_includes_markdown_module_rejects_unscoped_or_inexact_matches(
    tmp_path: Path,
    marker_text: str,
) -> None:
    """Look-alike markdown tokens outside the module block are ignored."""
    _write_marker(tmp_path, marker_text)

    assert first_adoption.marker_includes_markdown_module(tmp_path) is False


def test_marker_includes_markdown_module_skips_blank_and_comment_lines(
    tmp_path: Path,
) -> None:
    """Blank and comment lines do not terminate the retained-module sequence."""
    _write_marker(
        tmp_path,
        "\n".join(
            [
                "template_sync:",
                "  included_modules:",
                "  # retained baseline module",
                "",
                "  - baseline",
                "",
                "  # retained markdown module",
                "  - markdown",
                "  issue_labels:",
                "  - triage",
                "",
            ]
        ),
    )

    assert first_adoption.marker_includes_markdown_module(tmp_path) is True


def test_markdown_note_ignores_markdown_issue_label(tmp_path: Path) -> None:
    """An issue label named markdown does not imply the Markdown module is retained."""
    _write_marker(
        tmp_path,
        "\n".join(
            [
                "template_sync:",
                "  included_modules:",
                "  - baseline",
                "  issue_labels:",
                "  - markdown",
                "",
            ]
        ),
    )

    plan = first_adoption.markdown_commands_and_notes(tmp_path)

    assert plan == first_adoption.CheckPlan(commands=(), notes=())


def test_package_markdown_scripts_run_when_present(tmp_path: Path) -> None:
    """Supported package scripts add optional Markdown validation commands."""
    _run_git(tmp_path, "init")
    _write_text(tmp_path, "README.md")
    _write_text(
        tmp_path,
        "package.json",
        '{"scripts":{"lint:md":"markdownlint .","lint:md:links":"remark ."}}\n',
    )
    commands: list[tuple[str, ...]] = []

    result = first_adoption.run_first_adoption_checks(
        tmp_path,
        command_runner=_recording_runner(commands),
        stdout=io.StringIO(),
    )

    assert result == 0
    assert any(command[-2:] == ("run", "lint:md") for command in commands)
    assert any(command[-2:] == ("run", "lint:md:links") for command in commands)


@pytest.mark.parametrize("available_form", [(sys.executable, "-m"), ("py", "-m")])
def test_pre_commit_prefix_falls_back_to_python_module(
    monkeypatch: Any,
    available_form: tuple[str, ...],
) -> None:
    """The first available module launcher works without ambient console/module state."""

    def missing_executable(_name: str) -> None:
        return None

    monkeypatch.setattr(first_adoption.shutil, "which", missing_executable)
    probes: list[tuple[str, ...]] = []

    def module_probe(command: Sequence[str]) -> bool:
        probes.append(tuple(command))
        return tuple(command) == (*available_form, "pre_commit", "--version")

    monkeypatch.setattr(first_adoption, "command_succeeds", module_probe)

    prefix = first_adoption.default_pre_commit_prefix()

    assert prefix == (*available_form, "pre_commit", "run", "--files")
    assert probes[-1] == (*available_form, "pre_commit", "--version")


def test_npm_executable_prefers_cmd_shim_on_windows_style_path(
    monkeypatch: Any,
) -> None:
    """The runner can use npm.cmd when bare npm is not directly executable."""

    def fake_which(name: str) -> str | None:
        if name == "npm.cmd":
            return "C:\\tools\\npm.cmd"
        return None

    monkeypatch.setattr(first_adoption.shutil, "which", fake_which)

    assert first_adoption.default_npm_executable() == "C:\\tools\\npm.cmd"


def test_downstream_pytest_selection_includes_unmarked_and_excludes_upstream_only(
    tmp_path: Path,
) -> None:
    """Negative marker selection keeps new unmarked tests in the downstream gate."""
    pytest_root = _write_pytest_profile_root(tmp_path)
    test_file = pytest_root / "test_profile_selection.py"
    test_file.write_text(
        "\n".join(
            [
                "import pytest",
                "",
                "def test_unmarked_is_selected():",
                "    pass",
                "",
                "@pytest.mark.upstream_template_only",
                "def test_upstream_only_is_excluded():",
                "    pass",
                "",
            ]
        ),
        encoding="utf-8",
    )

    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "pytest",
            "--collect-only",
            "-q",
            "-m",
            "not upstream_template_only",
            test_file.name,
        ],
        cwd=pytest_root,
        check=False,
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stdout + result.stderr
    assert "test_unmarked_is_selected" in result.stdout
    assert "test_upstream_only_is_excluded" not in result.stdout


def test_unregistered_pytest_marker_fails_collection(tmp_path: Path) -> None:
    """Committed strict marker configuration rejects marker typos during collection."""
    pytest_root = _write_pytest_profile_root(tmp_path)
    test_file = pytest_root / "test_bad_marker.py"
    test_file.write_text(
        "\n".join(
            [
                "import pytest",
                "",
                "@pytest.mark.not_registered_here",
                "def test_bad_marker():",
                "    pass",
                "",
            ]
        ),
        encoding="utf-8",
    )

    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "pytest",
            "--collect-only",
            test_file.name,
        ],
        cwd=pytest_root,
        check=False,
        capture_output=True,
        text=True,
    )

    assert result.returncode != 0
    assert "not_registered_here" in result.stdout + result.stderr
