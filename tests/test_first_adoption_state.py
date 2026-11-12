"""Exercise raw first-adoption state inspection."""

from __future__ import annotations

import importlib
import subprocess
import sys
from pathlib import Path
from typing import Any, cast

from tests._pytest_compat import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT_PATH = REPO_ROOT / ".template-sync" / "scripts" / "first_adoption_state.py"
SCRIPT_DIR = SCRIPT_PATH.parent

if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

first_adoption_state = cast(Any, importlib.import_module("first_adoption_state"))


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


def _write_text(repo_root: Path, relative_path: str, text: str = "content\n") -> None:
    """Write a text file below a fixture repository root."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _state_paths(repo_root: Path) -> tuple[Path, Path, Path]:
    """Return the default marker, checklist, and journal paths for a fixture."""
    return (
        repo_root / ".template-sync" / "marker.yml",
        repo_root / "_TODO-repo-init.md",
        repo_root / "_ADOPTION-DIFFICULTIES.md",
    )


def test_inspection_reports_git_state_empty_trees_and_missing_marker(
    tmp_path: Path,
) -> None:
    """Raw state separates Git state, empty directories, and marker evidence."""
    _git(tmp_path, "init")
    _git(tmp_path, "config", "user.name", "State Tests")
    _git(tmp_path, "config", "user.email", "state-tests@example.invalid")
    _write_text(tmp_path, ".gitignore", "node_modules/\n")
    _write_text(tmp_path, "README.md", "tracked\n")
    _git(tmp_path, "add", ".gitignore", "README.md")
    _git(tmp_path, "commit", "-m", "base")

    _write_text(tmp_path, "notes/raw.txt", "untracked\n")
    _write_text(tmp_path, "node_modules/pkg/generated.js", "ignored\n")
    (tmp_path / "empty" / "nested").mkdir(parents=True)
    (tmp_path / ".template-sync").mkdir()
    marker_path, todo_path, journal_path = _state_paths(tmp_path)

    state = first_adoption_state.inspect_first_adoption_state(
        repo_root=tmp_path,
        marker_path=marker_path,
        todo_path=todo_path,
        journal_path=journal_path,
    )

    assert "README.md" in state.tracked_files
    assert "notes/raw.txt" in state.untracked_files
    assert "node_modules/" in state.ignored_files
    assert state.empty_directory_trees == (".template-sync/", "empty/")
    assert state.marker_evidence == (
        "`.template-sync/` directory present",
        "`.template-sync/marker.yml` missing",
    )
    assert state.missing_state_files == (
        ".template-sync/marker.yml",
        "_ADOPTION-DIFFICULTIES.md",
        "_TODO-repo-init.md",
    )


def test_state_rendering_is_bounded_by_default_and_complete_in_full_mode() -> None:
    """Default rendering shows counts, stable samples, and explicit remainders."""
    state = first_adoption_state.FirstAdoptionState(
        marker_evidence=("`.template-sync/` directory present",),
        adoption_notes=(),
        tracked_files=("alpha.md", "beta.md", "gamma.md"),
        untracked_files=(),
        ignored_files=(),
        empty_directory_trees=(),
        missing_state_files=(),
        special_paths=(),
    )

    bounded = first_adoption_state.format_first_adoption_state(
        state,
        sample_limit=2,
    )
    full = first_adoption_state.format_first_adoption_state(
        state,
        full_state=True,
        sample_limit=2,
    )

    assert "### Tracked Files" in bounded
    assert "- Count: 3" in bounded
    assert "`alpha.md`" in bounded
    assert "`beta.md`" in bounded
    assert "`gamma.md`" not in bounded
    assert "Remainder: 1 entry not shown" in bounded
    assert "`gamma.md`" in full
    assert "Remainder:" not in full


def test_inspection_uses_read_only_git_commands_and_preserves_filesystem(
    tmp_path: Path,
) -> None:
    """The shared inspection path does not issue mutating Git or filesystem writes."""
    marker_path, todo_path, journal_path = _state_paths(tmp_path)
    commands: list[tuple[str, ...]] = []

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        commands.append(tuple(command))
        if command[:1] == ["ls-files"]:
            return subprocess.CompletedProcess(command, 0, stdout="", stderr="")
        if command[:1] == ["check-ignore"]:
            return subprocess.CompletedProcess(command, 1, stdout="", stderr="")
        return subprocess.CompletedProcess(command, 99, stdout="", stderr="unexpected command")

    before = sorted(path.relative_to(tmp_path).as_posix() for path in tmp_path.rglob("*"))
    first_adoption_state.inspect_first_adoption_state(
        repo_root=tmp_path,
        marker_path=marker_path,
        todo_path=todo_path,
        journal_path=journal_path,
        git_runner=runner,
    )
    after = sorted(path.relative_to(tmp_path).as_posix() for path in tmp_path.rglob("*"))

    assert before == after
    assert commands
    assert {command[0] for command in commands} == {"check-ignore", "ls-files"}


def test_unavailable_git_state_is_rendered_without_code_spans() -> None:
    """git_lines unavailable sentinels render as plain status lines, not code spans."""
    sentinel = "tracked files unavailable: fatal: not a git repository"
    state = first_adoption_state.FirstAdoptionState(
        marker_evidence=(),
        adoption_notes=(),
        tracked_files=(sentinel,),
        untracked_files=(),
        ignored_files=(),
        empty_directory_trees=(),
        missing_state_files=(),
        special_paths=(),
    )

    rendered = first_adoption_state.format_first_adoption_state(state)

    assert f"  - {sentinel}" in rendered
    assert f"`{sentinel}`" not in rendered


def test_directory_state_paths_are_reported_absent(tmp_path: Path) -> None:
    """Directories at note/state paths are treated as absent regular files."""
    marker_path, todo_path, journal_path = _state_paths(tmp_path)
    todo_path.mkdir(parents=True)
    journal_path.mkdir(parents=True)

    notes = first_adoption_state.existing_note_entries(todo_path, journal_path, tmp_path)
    missing = first_adoption_state.missing_state_entries(
        marker_path, todo_path, journal_path, tmp_path
    )

    assert notes == ()
    assert "_TODO-repo-init.md" in missing
    assert "_ADOPTION-DIFFICULTIES.md" in missing


def test_marker_evidence_matches_missing_state_for_symlinked_marker(tmp_path: Path) -> None:
    """A symlinked marker is reported consistently as not a regular file."""
    marker_path, todo_path, journal_path = _state_paths(tmp_path)
    marker_path.parent.mkdir(parents=True, exist_ok=True)
    target = tmp_path / "real-marker.yml"
    target.write_text("marker: true\n", encoding="utf-8")
    try:
        marker_path.symlink_to(target)
    except (OSError, NotImplementedError):
        pytest.skip("Filesystem does not support symlink creation")

    evidence = first_adoption_state.marker_evidence_entries(marker_path, tmp_path)
    missing = first_adoption_state.missing_state_entries(
        marker_path, todo_path, journal_path, tmp_path
    )

    # marker_evidence must agree with missing_state_entries: a symlinked marker
    # is not a regular file, so it is reported "missing", not "found".
    assert "`.template-sync/marker.yml` missing" in evidence
    assert "`.template-sync/marker.yml` found" not in evidence
    assert ".template-sync/marker.yml" in missing


def test_state_path_categories_are_lexicographically_ordered(tmp_path: Path) -> None:
    """Missing-state and found-note categories honor the documented lexicographic order."""
    marker_path, todo_path, journal_path = _state_paths(tmp_path)

    missing = first_adoption_state.missing_state_entries(
        marker_path, todo_path, journal_path, tmp_path
    )
    assert missing == tuple(sorted(missing))
    assert missing == (
        ".template-sync/marker.yml",
        "_ADOPTION-DIFFICULTIES.md",
        "_TODO-repo-init.md",
    )

    todo_path.write_text("todo\n", encoding="utf-8")
    journal_path.write_text("journal\n", encoding="utf-8")
    notes = first_adoption_state.existing_note_entries(todo_path, journal_path, tmp_path)
    assert notes == tuple(sorted(notes))
    assert notes == (
        "`_ADOPTION-DIFFICULTIES.md` found",
        "`_TODO-repo-init.md` found",
    )


def test_high_signal_inventory_reports_unknown_ignore_state_on_git_failure(
    tmp_path: Path,
) -> None:
    """A failing git check-ignore yields ignored: unknown, not a misleading no."""

    def runner(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
        if command[:1] == ["check-ignore"]:
            return subprocess.CompletedProcess(command, 128, stdout="", stderr="")
        return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

    marker_path, todo_path, journal_path = _state_paths(tmp_path)
    state = first_adoption_state.inspect_first_adoption_state(
        repo_root=tmp_path,
        marker_path=marker_path,
        todo_path=todo_path,
        journal_path=journal_path,
        git_runner=runner,
    )

    assert state.special_paths
    assert all(path_state.ignored is None for path_state in state.special_paths)
    summaries = [path_state.summary for path_state in state.special_paths]
    assert all("ignored: unknown" in summary for summary in summaries)
    assert not any("ignored: no" in summary for summary in summaries)


def test_special_path_summary_renders_unknown_for_none_ignored() -> None:
    """SpecialPathState renders ignored: unknown when the ignore state is None."""
    state = first_adoption_state.SpecialPathState(
        path=".github/", exists=True, kind="directory", ignored=None
    )
    assert state.summary == "`.github/`: present directory; ignored: unknown"
