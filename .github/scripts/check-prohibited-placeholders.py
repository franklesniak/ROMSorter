"""Check Markdown docs for prohibited placeholder markers.

The pre-commit hook calls this script with candidate Markdown paths. A named
path that does not exist fails the run, so a mistyped path in a hand run
cannot pass silently. The checker intentionally stays dependency-free so it
can run in the repo-local hook environment on Windows, macOS, Linux, and WSL.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PLACEHOLDER_PATTERN = re.compile(
    r"\(default[^)]*to\s+be\s+determined[^)]*\)"
    r"|\bTODO\b\s*:"
    r"|\bTBD\b"
    r"|\bFIXME\b"
    r"|\bXXX\b"
    r"|\bto\s+be\s+determined\b",
    re.IGNORECASE,
)
ALLOW_TBD_PATTERN = re.compile(r"<!--\s*ALLOW-TBD:\s*\S.*?-->", re.IGNORECASE)
ALLOWED_LABEL_PATTERN = re.compile(
    r"^\s*(?:(?:[-*+]|\d{1,9}[.)])\s+)?\*\*(?:Open Questions?|Assumption):\*\*",
    re.IGNORECASE,
)
FENCE_OPEN_PATTERN = re.compile(r"^ {0,3}(?P<marker>`{3,}|~{3,})")
BLOCK_QUOTE_PREFIX_PATTERN = re.compile(r"^ {0,3}> ?")
LIST_ITEM_PATTERN = re.compile(r"^(?P<indent> {0,3})(?P<marker>[-*+]|\d{1,9}[.)])(?P<spacing> +)")

REMEDIATION_HINT = (
    "replace with a measurable value, an **Open Question:** entry, an "
    "**Assumption:** entry, or a cross-reference to another requirement. "
    "To suppress with explicit justification, add <!-- ALLOW-TBD: <reason> --> "
    'on the same line. See .github/instructions/docs.instructions.md "Prohibited Patterns".'
)
MISSING_PATH_HINT = (
    "nothing is at this path, so there is nothing to check. "
    "Relative paths are read from the repository root, not the current folder. "
    "Check the path: a typo would otherwise pass silently, because a run that scans "
    "nothing finds no placeholders."
)


@dataclass(frozen=True)
class Violation:
    """A prohibited placeholder match in a Markdown file."""

    display_path: str
    line_number: int
    matched_text: str

    def format_message(self) -> str:
        """Return the hook failure message for this placeholder match."""
        return (
            f"{self.display_path}:{self.line_number}: prohibited placeholder "
            f"{json.dumps(self.matched_text)}; {REMEDIATION_HINT}"
        )


@dataclass(frozen=True)
class MissingPath:
    """A named path with nothing at it, so the hook had nothing to scan."""

    display_path: str

    def format_message(self) -> str:
        """Return the hook failure message for this missing path.

        An empty argument is shown as ``""``, so the message never starts with a bare colon.
        """
        shown_path = self.display_path or '""'
        return f"{shown_path}: {MISSING_PATH_HINT}"


CONTAINER_KIND_LIST = "list"
CONTAINER_KIND_BLOCK_QUOTE = "blockquote"


@dataclass(frozen=True)
class Container:
    """A peelable Markdown container prefix on a line.

    A ``list`` container consumes ``indent`` leading spaces from its parent
    container's interior. A ``blockquote`` container consumes a single
    ``BLOCK_QUOTE_PREFIX_PATTERN`` match (``>`` optionally followed by a
    space) at the start of its parent's interior; ``indent`` is unused.
    """

    kind: str
    indent: int = 0


@dataclass(frozen=True)
class FenceLine:
    """A Markdown line normalized for fenced code block detection."""

    content: str
    containment_path: tuple[Container, ...] = ()


@dataclass(frozen=True)
class ActiveFence:
    """The active fenced code block marker and containing Markdown context."""

    character: str
    minimum_length: int
    containment_path: tuple[Container, ...] = ()


@dataclass(frozen=True)
class ListContext:
    """An active Markdown list, identified by its full containment path from the document root."""

    containment_path: tuple[Container, ...]


class FileReadError(RuntimeError):
    """Raised when a candidate Markdown file cannot be read."""

    def __init__(self, display_path: str, error: OSError) -> None:
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        super().__init__(f"{display_path}: unable to read file ({error_summary})")


def is_changelog_file(path: Path) -> bool:
    """Return whether the file name matches the hook's changelog exemption."""
    name = path.name.lower()
    return name.startswith("changelog") and name.endswith(".md")


def is_scan_target(relative_path: Path) -> bool:
    """Return whether a repo-relative path is a docs Markdown scan target."""
    return (
        len(relative_path.parts) >= 2
        and relative_path.parts[0] == "docs"
        and relative_path.suffix.lower() == ".md"
        and not is_changelog_file(relative_path)
    )


def locate_path_argument(path_argument: str | Path, root: Path) -> Path:
    """Return the location a path argument names: itself if absolute, else under ``root``."""
    path = Path(path_argument)
    return path if path.is_absolute() else root / path


def is_missing_path(path_argument: str | Path, root: Path) -> bool:
    """Return whether no filesystem entry exists at a path argument.

    ``os.path.lexists`` is the filter pre-commit applies before it passes file
    names to a hook, so a name that pre-commit passes is never missing here. A
    symlink counts as present even when its target is gone; it keeps the skip
    that ``resolve_candidate_path`` gives every symlink. An empty string names
    nothing. It is checked before any ``Path`` conversion, because ``Path("")``
    is ``Path(".")``, which would name the root folder itself.
    """
    if path_argument == "":
        return True
    return not os.path.lexists(locate_path_argument(path_argument, root))


def resolve_candidate_path(path_argument: str | Path, root: Path) -> tuple[Path, str] | None:
    """Resolve a pre-commit path argument to a contained scan target."""
    root = root.resolve()
    candidate = locate_path_argument(path_argument, root)

    if candidate.is_symlink() or not candidate.is_file():
        return None

    resolved_candidate = candidate.resolve()
    try:
        relative_path = resolved_candidate.relative_to(root)
    except ValueError:
        return None

    if not is_scan_target(relative_path):
        return None

    return resolved_candidate, relative_path.as_posix()


def strip_html_comments(line: str, is_in_html_comment: bool) -> tuple[str, bool]:
    """Remove HTML comment spans from a Markdown line."""
    uncommented_parts: list[str] = []
    index = 0

    while index < len(line):
        if is_in_html_comment:
            comment_end = line.find("-->", index)
            if comment_end == -1:
                return "".join(uncommented_parts), True
            index = comment_end + len("-->")
            is_in_html_comment = False
            continue

        comment_start = line.find("<!--", index)
        if comment_start == -1:
            uncommented_parts.append(line[index:])
            break

        uncommented_parts.append(line[index:comment_start])
        comment_end = line.find("-->", comment_start + len("<!--"))
        if comment_end == -1:
            is_in_html_comment = True
            break
        index = comment_end + len("-->")

    return "".join(uncommented_parts), is_in_html_comment


def count_leading_spaces(line: str) -> int:
    """Return the number of leading space characters in a line."""
    return len(line) - len(line.lstrip(" "))


def peel_containers(line: str, path: tuple[Container, ...]) -> tuple[str, int]:
    """Peel container prefixes from ``line`` in order; return ``(remaining, peeled_count)``.

    A list container consumes ``container.indent`` leading spaces (or fails if
    the line has fewer). A blockquote container consumes one
    ``BLOCK_QUOTE_PREFIX_PATTERN`` match (or fails if the line does not start
    with one in its current coordinate system). Peeling stops at the first
    container that cannot be consumed; the caller can compare ``peeled_count``
    to ``len(path)`` to detect partial peels.
    """
    for index, container in enumerate(path):
        if container.kind == CONTAINER_KIND_LIST:
            if count_leading_spaces(line) < container.indent:
                return line, index
            line = line[container.indent :]
        else:
            match = BLOCK_QUOTE_PREFIX_PATTERN.match(line)
            if match is None:
                return line, index
            line = line[match.end() :]
    return line, len(path)


def prune_inactive_list_contexts(line: str, list_contexts: list[ListContext]) -> None:
    """Drop active list contexts that a nonblank Markdown line has outdented past."""
    if not line.strip():
        return

    while list_contexts:
        top = list_contexts[-1]
        remaining, peeled_count = peel_containers(line, top.containment_path)
        if peeled_count == len(top.containment_path):
            return
        # The line did not peel cleanly to this list's interior. Decide
        # whether the partial peel still keeps the list active.
        if not remaining.strip():
            # A blank line inside a partly-peeled container is a continuation.
            return
        failed_container = top.containment_path[peeled_count]
        if (
            failed_container.kind == CONTAINER_KIND_LIST
            and BLOCK_QUOTE_PREFIX_PATTERN.match(remaining) is not None
        ):
            # A deeper blockquote nested inside the list keeps the list active;
            # the list's content indent is not meaningful in that deeper
            # coordinate system.
            return
        list_contexts.pop()


def list_content_indent(match: re.Match[str]) -> int:
    """Return the list-item content indent relative to the marker's parent interior."""
    marker_end_column = match.end("marker")
    spacing_width = len(match.group("spacing"))
    content_padding = spacing_width if spacing_width <= 4 else 1
    return marker_end_column + content_padding


def normalize_for_fence_opening(line: str, list_contexts: list[ListContext]) -> FenceLine:
    """Return a line normalized to its current Markdown container content column."""
    prune_inactive_list_contexts(line, list_contexts)

    active_path = list_contexts[-1].containment_path if list_contexts else ()
    relative_line, peeled_count = peel_containers(line, active_path)
    effective_path = active_path[:peeled_count]

    extras: list[Container] = []
    while True:
        match = BLOCK_QUOTE_PREFIX_PATTERN.match(relative_line)
        if match is None:
            break
        extras.append(Container(kind=CONTAINER_KIND_BLOCK_QUOTE))
        relative_line = relative_line[match.end() :]

    list_match = LIST_ITEM_PATTERN.match(relative_line)
    if list_match is not None:
        content_indent_rel = list_content_indent(list_match)
        extras.append(Container(kind=CONTAINER_KIND_LIST, indent=content_indent_rel))
        relative_line = (
            relative_line[content_indent_rel:] if len(relative_line) >= content_indent_rel else ""
        )
        list_contexts.append(ListContext(containment_path=effective_path + tuple(extras)))

    return FenceLine(
        content=relative_line,
        containment_path=effective_path + tuple(extras),
    )


def normalize_for_fence_closing(line: str, active_fence: ActiveFence) -> str:
    """Return a fenced-block line normalized to the opening fence's container."""
    peeled, peeled_count = peel_containers(line, active_fence.containment_path)
    if peeled_count < len(active_fence.containment_path):
        return line
    return peeled


def build_active_fence(
    opening_fence: tuple[str, int],
    fence_line: FenceLine,
) -> ActiveFence:
    """Return active fenced code block state for a detected opening fence."""
    fence_character, minimum_length = opening_fence
    return ActiveFence(
        character=fence_character,
        minimum_length=minimum_length,
        containment_path=fence_line.containment_path,
    )


def parse_opening_fence(line: str) -> tuple[str, int] | None:
    """Return the opening fence marker character and length, if present."""
    match = FENCE_OPEN_PATTERN.match(line)
    if match is None:
        return None

    marker = match.group("marker")
    return marker[0], len(marker)


def is_closing_fence(line: str, fence_character: str, minimum_length: int) -> bool:
    """Return whether a line closes the active fenced code block."""
    closing_pattern = re.compile(rf"^ {{0,3}}{re.escape(fence_character)}{{{minimum_length},}}\s*$")
    return closing_pattern.match(line) is not None


def is_allowed_label_line(line: str) -> bool:
    """Return whether the line is an explicit Open Question or Assumption entry."""
    return ALLOWED_LABEL_PATTERN.match(line) is not None


def find_violations_in_text(text: str, display_path: str) -> list[Violation]:
    """Find prohibited placeholder markers in Markdown text."""
    violations: list[Violation] = []
    is_in_html_comment = False
    active_fence: ActiveFence | None = None
    list_contexts: list[ListContext] = []

    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        if active_fence is not None:
            closing_line = normalize_for_fence_closing(raw_line, active_fence)
            if is_closing_fence(
                closing_line,
                active_fence.character,
                active_fence.minimum_length,
            ):
                active_fence = None
            continue

        commentless_line, is_in_html_comment = strip_html_comments(raw_line, is_in_html_comment)

        opening_fence_line = normalize_for_fence_opening(commentless_line, list_contexts)
        opening_fence = parse_opening_fence(opening_fence_line.content)
        if opening_fence is not None:
            active_fence = build_active_fence(opening_fence, opening_fence_line)
            continue

        if ALLOW_TBD_PATTERN.search(raw_line):
            continue

        if is_allowed_label_line(commentless_line):
            continue

        for match in PLACEHOLDER_PATTERN.finditer(commentless_line):
            violations.append(
                Violation(
                    display_path=display_path,
                    line_number=line_number,
                    matched_text=match.group(0),
                )
            )

    return violations


def scan_files(
    path_arguments: Iterable[str | Path], root: Path = REPO_ROOT
) -> list[Violation | MissingPath]:
    """Find prohibited placeholder markers in candidate Markdown docs.

    A path that exists but is not a scan target, such as a directory, a
    symlink, a changelog, or a file outside ``docs/``, is skipped. A path that
    does not exist is recorded as a ``MissingPath``, and the scan continues
    with the remaining paths.
    """
    findings: list[Violation | MissingPath] = []
    for path_argument in path_arguments:
        candidate = resolve_candidate_path(path_argument, root)
        if candidate is None:
            if is_missing_path(path_argument, root):
                findings.append(MissingPath(display_path=str(path_argument)))
            continue

        path, display_path = candidate
        try:
            text = path.read_text(encoding="utf-8")
        except OSError as error:
            raise FileReadError(display_path, error) from error

        findings.extend(find_violations_in_text(text, display_path))

    return findings


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Check docs Markdown for prohibited placeholder markers."
    )
    parser.add_argument(
        "paths",
        nargs="*",
        help=(
            "Markdown files passed by pre-commit. Relative paths are read from the "
            "repository root. A path that does not exist fails the run."
        ),
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None, root: Path = REPO_ROOT) -> int:
    """Run the placeholder check."""
    args = parse_args(argv)

    try:
        findings = scan_files(args.paths, root=root)
    except FileReadError as error:
        print(error, file=sys.stderr)
        return 1

    for finding in findings:
        # A missing path is an input error, like an unreadable file, so it
        # goes to stderr; placeholder matches stay on stdout.
        output_stream = sys.stderr if isinstance(finding, MissingPath) else sys.stdout
        print(finding.format_message(), file=output_stream)

    return 1 if findings else 0


if __name__ == "__main__":
    raise SystemExit(main())
