"""Inspect raw first-adoption Git and filesystem state without mutating it."""

from __future__ import annotations

import subprocess
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from template_sync_materialization_helpers import (
    DEFAULT_MARKER_PATH,
    repository_relative_path,
)

DEFAULT_ADOPTION_TODO_PATH = "_TODO-repo-init.md"
DEFAULT_ADOPTION_JOURNAL_PATH = "_ADOPTION-DIFFICULTIES.md"
DEFAULT_STATE_SAMPLE_LIMIT = 8
UNAVAILABLE_STATE_MARKER = " unavailable: "
SPECIAL_STATE_PATHS = (
    ".template-sync/",
    DEFAULT_MARKER_PATH,
    ".github/",
    "AGENTS.md",
    "CLAUDE.md",
    "GEMINI.md",
    ".hermes.md",
    ".cursor/",
    ".claude/",
    ".codex/",
    ".vscode/",
    "node_modules/",
)
EMPTY_TREE_SKIP_DIRS = frozenset(
    {
        ".git",
        ".mypy_cache",
        ".pytest_cache",
        ".ruff_cache",
        ".terraform",
        ".venv",
        "__pycache__",
        "build",
        "dist",
        "node_modules",
        "venv",
    }
)


GitRunner = Callable[[Path, list[str]], subprocess.CompletedProcess[str]]


@dataclass(frozen=True)
class SpecialPathState:
    """Presence and ignore-state details for a path that matters during adoption."""

    path: str
    exists: bool
    kind: str
    ignored: bool | None

    @property
    def summary(self) -> str:
        """Return a deterministic Markdown-safe summary line."""
        presence = "present" if self.exists else "absent"
        if self.ignored is None:
            ignored = "unknown"
        else:
            ignored = "yes" if self.ignored else "no"
        return f"`{self.path}`: {presence} {self.kind}; ignored: {ignored}"


@dataclass(frozen=True)
class FirstAdoptionState:
    """Raw read-only state used by first-adoption preflight reporting."""

    marker_evidence: tuple[str, ...]
    adoption_notes: tuple[str, ...]
    tracked_files: tuple[str, ...]
    untracked_files: tuple[str, ...]
    ignored_files: tuple[str, ...]
    empty_directory_trees: tuple[str, ...]
    missing_state_files: tuple[str, ...]
    special_paths: tuple[SpecialPathState, ...]


def run_git(repo_root: Path, args: list[str]) -> subprocess.CompletedProcess[str]:
    """Run a read-only Git command and capture text output."""
    return subprocess.run(
        ["git", "-C", str(repo_root), *args],
        check=False,
        capture_output=True,
        text=True,
    )


def command_detail(result: subprocess.CompletedProcess[str]) -> str:
    """Return a compact diagnostic from a command result."""
    for stream_text in (result.stderr, result.stdout):
        for line in stream_text.splitlines():
            if line.strip():
                return line.strip()
    return f"exit {result.returncode}"


def git_lines(
    repo_root: Path,
    args: list[str],
    label: str,
    *,
    git_runner: GitRunner,
) -> tuple[str, ...]:
    """Return sorted unique Git output lines or an unavailable-state entry."""
    result = git_runner(repo_root, args)
    if result.returncode != 0:
        return (f"{label}{UNAVAILABLE_STATE_MARKER}{command_detail(result)}",)
    return tuple(sorted({line for line in result.stdout.splitlines() if line}))


def path_kind(repo_root: Path, relative_path: str) -> tuple[bool, str]:
    """Return presence and path kind for a repository-relative path."""
    path = repo_root / relative_path.rstrip("/")
    if path.is_symlink():
        return True, "symlink"
    if path.is_dir():
        return True, "directory"
    if path.is_file():
        return True, "file"
    if path.exists():
        return True, "path"
    return False, "path"


def is_ignored_path(repo_root: Path, relative_path: str, *, git_runner: GitRunner) -> bool | None:
    """Return whether Git treats a path as ignored, or ``None`` when unknown.

    ``git check-ignore -q`` exits 0 when the path is ignored and 1 when it is
    not. A higher exit code (for example 128 outside a Git repository, or when
    Git is otherwise unavailable) means the ignore status cannot be determined;
    that case returns ``None`` so the inventory reports ``ignored: unknown``
    instead of misreporting ``ignored: no``.
    """
    result = git_runner(
        repo_root,
        ["check-ignore", "-q", "--", relative_path.rstrip("/")],
    )
    if result.returncode == 0:
        return True
    if result.returncode == 1:
        return False
    return None


def collect_special_paths(
    repo_root: Path, *, git_runner: GitRunner
) -> tuple[SpecialPathState, ...]:
    """Return presence and ignore details for first-adoption high-signal paths."""
    states: list[SpecialPathState] = []
    for relative_path in SPECIAL_STATE_PATHS:
        exists, kind = path_kind(repo_root, relative_path)
        states.append(
            SpecialPathState(
                path=relative_path,
                exists=exists,
                kind=kind,
                ignored=is_ignored_path(repo_root, relative_path, git_runner=git_runner),
            )
        )
    return tuple(states)


def collect_empty_directory_trees(repo_root: Path) -> tuple[str, ...]:
    """Return physical empty directory-tree roots without following symlinks."""
    root = repo_root.resolve()
    empty_directories: set[Path] = set()

    def is_empty_tree(directory: Path) -> bool:
        try:
            children = sorted(directory.iterdir(), key=lambda child: child.name)
        except OSError:
            return False

        has_entry = False
        all_children_empty = True
        for child in children:
            if child.name in EMPTY_TREE_SKIP_DIRS and child.is_dir():
                has_entry = True
                all_children_empty = False
                continue
            if child.is_symlink():
                has_entry = True
                all_children_empty = False
                continue
            if child.is_dir():
                try:
                    child.resolve().relative_to(root)
                except (OSError, ValueError):
                    has_entry = True
                    all_children_empty = False
                    continue
                has_entry = True
                if not is_empty_tree(child):
                    all_children_empty = False
                continue
            has_entry = True
            all_children_empty = False

        is_empty = not has_entry or all_children_empty
        if directory != root and is_empty:
            empty_directories.add(directory)
        return is_empty

    is_empty_tree(root)

    empty_roots: list[Path] = []
    for directory in sorted(empty_directories):
        if any(parent in empty_directories for parent in directory.parents if parent != root):
            continue
        empty_roots.append(directory)

    return tuple(
        f"{repository_relative_path(directory, root)}/" for directory in sorted(empty_roots)
    )


def is_regular_file(path: Path) -> bool:
    """Return whether ``path`` is a regular file rather than a symlink or directory.

    Adoption note and state files are expected to be plain regular files. Using
    this stricter check instead of ``Path.exists()`` avoids misreporting a
    directory or symlink as a present state file.
    """
    return path.is_file() and not path.is_symlink()


def existing_note_entries(todo_path: Path, journal_path: Path, repo_root: Path) -> tuple[str, ...]:
    """Return found adoption note entries in lexicographic order."""
    entries: list[str] = []
    for path in (todo_path, journal_path):
        if is_regular_file(path):
            entries.append(f"`{repository_relative_path(path, repo_root)}` found")
    return tuple(sorted(entries))


def missing_state_entries(
    marker_path: Path,
    todo_path: Path,
    journal_path: Path,
    repo_root: Path,
) -> tuple[str, ...]:
    """Return missing adoption state files in lexicographic order."""
    entries: list[str] = []
    for path in (marker_path, todo_path, journal_path):
        if not is_regular_file(path):
            entries.append(repository_relative_path(path, repo_root))
    return tuple(sorted(entries))


def marker_evidence_entries(marker_path: Path, repo_root: Path) -> tuple[str, ...]:
    """Return marker directory and marker-file evidence as separate facts."""
    template_sync_dir = repo_root / ".template-sync"
    directory_status = "present" if template_sync_dir.is_dir() else "not found"
    marker_status = "found" if is_regular_file(marker_path) else "missing"
    return (
        f"`.template-sync/` directory {directory_status}",
        f"`{repository_relative_path(marker_path, repo_root)}` {marker_status}",
    )


def inspect_first_adoption_state(
    *,
    repo_root: Path,
    marker_path: Path,
    todo_path: Path,
    journal_path: Path,
    git_runner: GitRunner = run_git,
) -> FirstAdoptionState:
    """Return raw first-adoption Git and filesystem state without mutations."""
    return FirstAdoptionState(
        marker_evidence=marker_evidence_entries(marker_path, repo_root),
        adoption_notes=existing_note_entries(todo_path, journal_path, repo_root),
        tracked_files=git_lines(
            repo_root,
            ["ls-files", "--cached"],
            "tracked files",
            git_runner=git_runner,
        ),
        untracked_files=git_lines(
            repo_root,
            ["ls-files", "--others", "--exclude-standard"],
            "untracked files",
            git_runner=git_runner,
        ),
        ignored_files=git_lines(
            repo_root,
            ["ls-files", "--others", "--ignored", "--exclude-standard", "--directory"],
            "ignored files",
            git_runner=git_runner,
        ),
        empty_directory_trees=collect_empty_directory_trees(repo_root),
        missing_state_files=missing_state_entries(marker_path, todo_path, journal_path, repo_root),
        special_paths=collect_special_paths(repo_root, git_runner=git_runner),
    )


def format_state_items(
    *,
    title: str,
    items: tuple[str, ...],
    empty_text: str,
    full_state: bool,
    sample_limit: int,
) -> str:
    """Render one bounded or full raw-state category."""
    lines = [f"### {title}", "", f"- Count: {len(items)}"]
    if not items:
        lines.append(f"- {empty_text}")
        return "\n".join(lines)

    visible_items = items if full_state else items[:sample_limit]
    label = "Entries" if full_state or len(items) <= sample_limit else "Sample"
    lines.append(f"- {label}:")
    for item in visible_items:
        lines.append(f"  - {item}")
    if not full_state and len(items) > sample_limit:
        lines.append(
            f"  - Remainder: {len(items) - sample_limit} entr"
            f"{'y' if len(items) - sample_limit == 1 else 'ies'} not shown; "
            "rerun with `--full-state` to list every entry."
        )
    return "\n".join(lines)


def is_unavailable_entry(entry: str) -> bool:
    """Return whether ``entry`` is a ``git_lines`` unavailable status, not a path."""
    return UNAVAILABLE_STATE_MARKER in entry


def format_path_items(paths: tuple[str, ...]) -> tuple[str, ...]:
    """Return repository paths as Markdown code spans, leaving status lines plain.

    Entries produced by :func:`git_lines` when a Git command fails are
    human-readable ``<label> unavailable: <detail>`` status strings rather than
    repository paths, so they are returned without code-span formatting.
    """
    return tuple(entry if is_unavailable_entry(entry) else f"`{entry}`" for entry in paths)


def format_first_adoption_state(
    state: FirstAdoptionState,
    *,
    full_state: bool = False,
    sample_limit: int = DEFAULT_STATE_SAMPLE_LIMIT,
) -> str:
    """Render raw first-adoption state with bounded default output."""
    if sample_limit < 1:
        raise ValueError("sample_limit must be at least 1")

    categories = (
        (
            "Marker Evidence",
            state.marker_evidence,
            "No marker evidence was available.",
        ),
        (
            "Adoption Notes",
            state.adoption_notes,
            "No adoption note files were found.",
        ),
        (
            "Tracked Files",
            format_path_items(state.tracked_files),
            "No tracked files were reported by Git.",
        ),
        (
            "Untracked Files",
            format_path_items(state.untracked_files),
            "No untracked non-ignored files were reported by Git.",
        ),
        (
            "Ignored Files And Directories",
            format_path_items(state.ignored_files),
            "No ignored files or directories were reported by Git.",
        ),
        (
            "Physical Empty Directory Trees",
            format_path_items(state.empty_directory_trees),
            "No physically empty directory trees were found.",
        ),
        (
            "Missing State Files",
            format_path_items(state.missing_state_files),
            "No expected first-adoption state files are missing.",
        ),
        (
            "High-Signal Path Inventory",
            tuple(path_state.summary for path_state in state.special_paths),
            "No high-signal paths were configured for inspection.",
        ),
    )

    lines = [
        (
            "The categories below are read-only Git and filesystem evidence. "
            "Tracked, untracked, and ignored entries are Git state; empty "
            "directory trees are physical filesystem state that Git does not track."
        ),
        "",
        (
            "Default output is bounded and deterministic. Use `--full-state` with "
            "`--preflight` to list every entry."
        ),
    ]
    for title, items, empty_text in categories:
        lines.extend(
            (
                "",
                format_state_items(
                    title=title,
                    items=items,
                    empty_text=empty_text,
                    full_state=full_state,
                    sample_limit=sample_limit,
                ),
            )
        )
    return "\n".join(lines)
