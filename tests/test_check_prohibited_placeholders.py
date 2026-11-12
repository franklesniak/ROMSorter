"""Tests for the prohibited Markdown placeholder pre-commit hook."""

from __future__ import annotations

import importlib
import importlib.util
import os
import sys
from collections.abc import Iterable
from pathlib import Path
from typing import Any, Protocol, cast, runtime_checkable

from tests._pytest_compat import pytest

HOOK_PATH = (
    Path(__file__).resolve().parents[1] / ".github" / "scripts" / "check-prohibited-placeholders.py"
)
HOOK_SPEC = importlib.util.spec_from_file_location("check_prohibited_placeholders", HOOK_PATH)
if HOOK_SPEC is None or HOOK_SPEC.loader is None:
    raise RuntimeError(f"Unable to load placeholder hook module from {HOOK_PATH}")
_placeholder_hook = importlib.util.module_from_spec(HOOK_SPEC)
sys.modules[HOOK_SPEC.name] = _placeholder_hook
HOOK_SPEC.loader.exec_module(_placeholder_hook)


@runtime_checkable
class ViolationLike(Protocol):
    """Attributes exposed by a placeholder-hook violation."""

    display_path: str
    line_number: int
    matched_text: str


class MissingPathLike(Protocol):
    """Attributes exposed by a placeholder-hook refusal of a missing path."""

    display_path: str

    def format_message(self) -> str: ...


class PlaceholderHookModule(Protocol):
    """Typed public surface used from the file-path-loaded placeholder hook."""

    def scan_files(
        self,
        path_arguments: Iterable[str | Path],
        root: Path,
    ) -> list[ViolationLike | MissingPathLike]: ...

    def main(self, argv: Iterable[str] | None = None, root: Path = ...) -> int: ...


placeholder_hook: PlaceholderHookModule = cast(PlaceholderHookModule, _placeholder_hook)


def write_file(path: Path, content: str) -> Path:
    """Create a test file and its parent directories."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    return path


def scan_single_file(path: Path, root: Path) -> list[ViolationLike]:
    """Scan one file through the public hook path and return its violations.

    The test fails if a finding has no line number, such as a missing-path refusal.
    """
    violations: list[ViolationLike] = []
    for finding in placeholder_hook.scan_files([path], root=root):
        assert isinstance(finding, ViolationLike), f"expected a violation, got {finding!r}"
        violations.append(finding)
    return violations


@pytest.mark.parametrize(
    ("content", "matched_text"),
    [
        ("The value is TBD.\n", "TBD"),
        ("TODO: document the timeout.\n", "TODO:"),
        ("Fix this FIXME marker.\n", "FIXME"),
        ("Remove this XXX marker.\n", "XXX"),
        ("The limit is to be determined.\n", "to be determined"),
        (
            "The timeout uses (default duration to be determined).\n",
            "(default duration to be determined)",
        ),
    ],
)
def test_required_patterns_are_flagged(
    tmp_path: Path,
    content: str,
    matched_text: str,
) -> None:
    """Each required placeholder pattern is prohibited in docs Markdown."""
    path = write_file(tmp_path / "docs" / "spec" / "example.md", content)

    violations = scan_single_file(path, tmp_path)

    assert len(violations) == 1
    assert violations[0].display_path == "docs/spec/example.md"
    assert violations[0].line_number == 1
    assert violations[0].matched_text == matched_text


@pytest.mark.parametrize(
    "content",
    [
        "Keep this value TBD. <!-- ALLOW-TBD: upstream contract not finalized -->\n",
        "<!-- TODO: this single-line HTML comment is ignored. -->\n",
        "**Open Question:** Should the default be TBD?\n",
        "**Open Questions:** Is the timeout to be determined by callers?\n",
        "**Assumption:** FIXME is quoted from an upstream draft.\n",
        "- **Open Question:** Can this stay TODO: until the spec lands?\n",
        "- **Assumption:** The value is XXX in the source document.\n",
        "1. **Open Question:** Does the limit stay TBD until v2?\n",
        "12. **Assumption:** XXX represents a placeholder index.\n",
        "1) **Open Question:** Does the parenthesized marker stay TBD?\n",
        "12) **Assumption:** XXX in a parenthesized ordered list.\n",
    ],
)
def test_required_allowlist_contexts_are_not_flagged(tmp_path: Path, content: str) -> None:
    """Allowed suppression, comment, Open Question, and Assumption lines pass."""
    path = write_file(tmp_path / "docs" / "spec" / "example.md", content)

    violations = scan_single_file(path, tmp_path)

    assert violations == []


def test_ordered_allowlist_label_requires_marker_spacing(tmp_path: Path) -> None:
    """A malformed ordered-list marker does not suppress placeholder checks."""
    path = write_file(tmp_path / "docs" / "spec" / "example.md", "1.**Open Question:** TBD\n")

    violations = scan_single_file(path, tmp_path)

    assert len(violations) == 1
    assert violations[0].line_number == 1
    assert violations[0].matched_text == "TBD"


def test_ordered_allowlist_label_rejects_overlong_digit_marker(tmp_path: Path) -> None:
    """An ordered marker with more than 9 digits is not a CommonMark list marker."""
    path = write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "1234567890. **Open Question:** TBD with an overlong digit marker.\n",
    )

    violations = scan_single_file(path, tmp_path)

    assert len(violations) == 1
    assert violations[0].line_number == 1
    assert violations[0].matched_text == "TBD"


def test_changelog_markdown_files_are_not_flagged(tmp_path: Path) -> None:
    """CHANGELOG*.md files are exempt from the docs placeholder check."""
    path = write_file(tmp_path / "docs" / "CHANGELOG-2026.md", "The value is TBD.\n")

    violations = scan_single_file(path, tmp_path)

    assert violations == []


def test_multiline_fenced_code_block_is_not_flagged(tmp_path: Path) -> None:
    """A genuine multi-line fenced code block is excluded from scanning."""
    path = write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "\n".join(
            [
                "```text",
                "TODO: this is an example inside a fence.",
                "```",
                "The real value is measurable.",
            ]
        )
        + "\n",
    )

    violations = scan_single_file(path, tmp_path)

    assert violations == []


@pytest.mark.parametrize(
    "content",
    [
        "\n".join(
            [
                "- A bullet with a continuation fence:",
                "",
                "    ```text",
                "    TODO: example value",
                "    ```",
                "- Next bullet.",
            ]
        )
        + "\n",
        "\n".join(
            [
                "1. A numbered item with a continuation fence:",
                "",
                "    ```text",
                "    TBD inside a numbered-list fence.",
                "    ```",
                "2. Next numbered item.",
            ]
        )
        + "\n",
        "\n".join(
            [
                "> ```text",
                "> FIXME inside a blockquote fence.",
                "> ```",
            ]
        )
        + "\n",
        "\n".join(
            [
                "- Parent item:",
                "  - Child item with a continuation fence:",
                "",
                "      ```text",
                "      XXX inside a nested-list fence.",
                "      ```",
            ]
        )
        + "\n",
    ],
)
def test_container_fenced_code_blocks_are_not_flagged(
    tmp_path: Path,
    content: str,
) -> None:
    """Fenced code blocks indented under Markdown containers are excluded."""
    path = write_file(tmp_path / "docs" / "spec" / "example.md", content)

    violations = scan_single_file(path, tmp_path)

    assert violations == []


def test_blockquote_inside_list_item_preserves_list_context(tmp_path: Path) -> None:
    """A blockquote within a list item must not pop the active list context."""
    path = write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "\n".join(
            [
                "- Parent item with an aside:",
                "  > A quoted aside inside the list item.",
                "",
                "    ```text",
                "    TBD inside a continuation fence after a blockquote.",
                "    ```",
            ]
        )
        + "\n",
    )

    violations = scan_single_file(path, tmp_path)

    assert violations == []


def test_list_inside_blockquote_preserves_list_context_across_blank_line(
    tmp_path: Path,
) -> None:
    """A list nested inside a blockquote must not be pruned by a blank blockquote line."""
    path = write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "\n".join(
            [
                "> 1. Item with a continuation fence:",
                ">",
                ">     ```text",
                ">     TBD inside a blockquote-list continuation fence.",
                ">     ```",
            ]
        )
        + "\n",
    )

    violations = scan_single_file(path, tmp_path)

    assert violations == []


def test_blockquote_inside_nested_list_item_fence_is_not_flagged(tmp_path: Path) -> None:
    """A blockquote-wrapped fence inside a nested list item (> past column 3) is excluded."""
    path = write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "\n".join(
            [
                "- Parent item:",
                "  - Child item with a quoted fence:",
                "",
                "      > ```text",
                "      > XXX inside a nested-list blockquote fence.",
                "      > ```",
            ]
        )
        + "\n",
    )

    violations = scan_single_file(path, tmp_path)

    assert violations == []


def test_nested_blockquote_in_blockquote_list_preserves_continuation_fence(
    tmp_path: Path,
) -> None:
    """A nested blockquote aside inside a blockquote-contained list must not break later fences."""
    path = write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "\n".join(
            [
                "> 1. Item with a nested aside:",
                "> > Note: see the upstream draft.",
                ">",
                ">     ```text",
                ">     TBD inside a continuation fence after a nested blockquote.",
                ">     ```",
            ]
        )
        + "\n",
    )

    violations = scan_single_file(path, tmp_path)

    assert violations == []


def test_list_inside_blockquote_inside_list_item_preserves_fence(
    tmp_path: Path,
) -> None:
    """A list nested inside a blockquote that is itself nested inside another list item."""
    path = write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "\n".join(
            [
                "123. Outer item with content_indent=5.",
                "     > 1. Nested list inside blockquote.",
                "     >",
                "     >     ```text",
                "     >     TBD inside continuation fence.",
                "     >     ```",
            ]
        )
        + "\n",
    )

    violations = scan_single_file(path, tmp_path)

    assert violations == []


def test_top_level_four_space_fence_marker_is_not_treated_as_fence(tmp_path: Path) -> None:
    """A top-level 4-space-indented fence marker does not suppress later prose."""
    path = write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "\n".join(
            [
                "    ```text",
                "TODO: this line remains scannable outside a recognized fence.",
                "    ```",
            ]
        )
        + "\n",
    )

    violations = scan_single_file(path, tmp_path)

    assert len(violations) == 1
    assert violations[0].line_number == 2
    assert violations[0].matched_text == "TODO:"


def test_multiline_tilde_fenced_code_block_is_not_flagged(tmp_path: Path) -> None:
    """Tilde fences receive the same multi-line exclusion as backtick fences."""
    path = write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "\n".join(
            [
                "~~~text",
                "FIXME inside a tilde fence.",
                "~~~",
            ]
        )
        + "\n",
    )

    violations = scan_single_file(path, tmp_path)

    assert violations == []


def test_multiline_html_comment_is_not_flagged(tmp_path: Path) -> None:
    """A genuine multi-line HTML comment is excluded from scanning."""
    path = write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "\n".join(
            [
                "<!--",
                "The default is to be determined.",
                "-->",
                "The real value is measurable.",
            ]
        )
        + "\n",
    )

    violations = scan_single_file(path, tmp_path)

    assert violations == []


def test_allow_tbd_on_html_comment_closing_line_preserves_state(tmp_path: Path) -> None:
    """ALLOW-TBD on a line that also closes a multi-line HTML comment keeps state correct."""
    path = write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "\n".join(
            [
                "<!-- multi-line comment opening",
                "still inside the comment --> <!-- ALLOW-TBD: closing line -->",
                "The value is TBD on a regular line.",
            ]
        )
        + "\n",
    )

    violations = scan_single_file(path, tmp_path)

    assert len(violations) == 1
    assert violations[0].line_number == 3
    assert violations[0].matched_text == "TBD"


def test_allow_tbd_on_fence_opening_line_enters_fence(tmp_path: Path) -> None:
    """ALLOW-TBD on a line that also opens a fenced code block still enters the fence."""
    path = write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "\n".join(
            [
                "```text <!-- ALLOW-TBD: documented example -->",
                "TBD inside the fenced block should not be flagged.",
                "```",
                "The value is TBD outside the fenced block.",
            ]
        )
        + "\n",
    )

    violations = scan_single_file(path, tmp_path)

    assert len(violations) == 1
    assert violations[0].line_number == 4
    assert violations[0].matched_text == "TBD"


def test_todo_task_list_without_colon_is_not_flagged(tmp_path: Path) -> None:
    """GitHub issue-style TODO task items without a colon are allowed."""
    path = write_file(tmp_path / "docs" / "spec" / "example.md", "- [ ] TODO\n")

    violations = scan_single_file(path, tmp_path)

    assert violations == []


def test_markdown_outside_docs_is_not_scanned(tmp_path: Path) -> None:
    """The hook scope is limited to Markdown files under docs/."""
    path = write_file(tmp_path / "README.md", "The value is TBD.\n")

    violations = scan_single_file(path, tmp_path)

    assert violations == []


def test_main_reports_actionable_failure_message(
    tmp_path: Path,
    capsys: Any,
) -> None:
    """CLI output includes path, line, matched text, remediation, and suppression syntax."""
    write_file(
        tmp_path / "docs" / "spec" / "example.md",
        "\n".join(
            [
                "Measured value.",
                "The timeout uses (default duration to be determined).",
            ]
        )
        + "\n",
    )

    result = placeholder_hook.main(["docs/spec/example.md"], root=tmp_path)

    captured = capsys.readouterr()
    assert result == 1
    assert captured.err == ""
    assert (
        'docs/spec/example.md:2: prohibited placeholder "(default duration to be determined)"'
    ) in captured.out
    assert "replace with a measurable value" in captured.out
    assert "<!-- ALLOW-TBD: <reason> -->" in captured.out
    assert '.github/instructions/docs.instructions.md "Prohibited Patterns"' in captured.out


@pytest.mark.parametrize(
    ("missing_path", "is_absolute"),
    [
        ("docs/spec/exmaple.md", False),
        ("docs/spec/exmaple.md", True),
        ("docs/sepc/example.md", False),
        ("docs/spec/deleted.md", False),
        ("READMEE.md", False),
        ("", False),
    ],
    ids=[
        "file-typo",
        "absolute-file-typo",
        "folder-typo",
        "deleted-file",
        "outside-docs-typo",
        "empty-argument",
    ],
)
def test_main_refuses_a_missing_path(
    tmp_path: Path,
    capsys: Any,
    missing_path: str,
    is_absolute: bool,
) -> None:
    """A named path that does not exist fails the run by name instead of passing silently."""
    write_file(tmp_path / "docs" / "spec" / "example.md", "Measured value.\n")
    write_file(tmp_path / "docs" / "spec" / "deleted.md", "Measured value.\n").unlink()
    write_file(tmp_path / "README.md", "Measured value.\n")
    # The test decides "missing" itself, not through the hook's predicate. An empty
    # argument names nothing: os.path.lexists("") is False, while Path("") is the root.
    location: Path | str = tmp_path / missing_path if missing_path else missing_path
    assert not os.path.lexists(location)
    path_argument = str(tmp_path / missing_path) if is_absolute else missing_path
    shown_path = '""' if path_argument == "" else path_argument

    result = placeholder_hook.main([path_argument], root=tmp_path)

    captured = capsys.readouterr()
    assert result == 1
    assert captured.out == ""
    assert captured.err.startswith(f"{shown_path}: nothing is at this path, so")
    assert "Relative paths are read from the repository root, not the current folder." in (
        captured.err
    )
    assert "a typo would otherwise pass silently" in captured.err


@pytest.mark.parametrize(
    "is_missing_path_first", [False, True], ids=["missing-last", "missing-first"]
)
@pytest.mark.parametrize(
    ("real_file_text", "expected_findings"),
    [
        ("Measured value.\n", []),
        ("The value is TBD.\n", ['docs/spec/example.md:1: prohibited placeholder "TBD"']),
    ],
    ids=["clean-real-file", "real-file-with-placeholder"],
)
def test_main_refuses_a_missing_path_beside_a_real_file(
    tmp_path: Path,
    capsys: Any,
    real_file_text: str,
    expected_findings: list[str],
    is_missing_path_first: bool,
) -> None:
    """A missing path fails the run beside a real file, and the real file is still scanned."""
    write_file(tmp_path / "docs" / "spec" / "example.md", real_file_text)
    path_arguments = ["docs/spec/example.md", "docs/spec/exmaple.md"]
    if is_missing_path_first:
        path_arguments.reverse()

    result = placeholder_hook.main(path_arguments, root=tmp_path)

    captured = capsys.readouterr()
    assert result == 1
    assert [line.split(";")[0] for line in captured.out.splitlines()] == expected_findings
    assert captured.err.startswith("docs/spec/exmaple.md: nothing is at this path, so")


@pytest.mark.parametrize(
    "path_argument", ["docs/spec", "README.md", "docs/CHANGELOG.md", ".", "./"]
)
def test_main_skips_existing_paths_that_are_not_scan_targets(
    tmp_path: Path,
    capsys: Any,
    path_argument: str,
) -> None:
    """A path that exists but is not a scan target is skipped, not refused as missing."""
    write_file(tmp_path / "docs" / "spec" / "example.md", "Measured value.\n")
    write_file(tmp_path / "README.md", "The value is TBD.\n")
    write_file(tmp_path / "docs" / "CHANGELOG.md", "The value is TBD.\n")

    result = placeholder_hook.main([path_argument], root=tmp_path)

    captured = capsys.readouterr()
    assert result == 0
    assert captured.out == ""
    assert captured.err == ""


def test_main_skips_a_dangling_symlink(tmp_path: Path, capsys: Any) -> None:
    """A symlink exists even when its target is gone, so it keeps the symlink skip."""
    link = tmp_path / "docs" / "spec" / "link.md"
    link.parent.mkdir(parents=True)
    try:
        link.symlink_to(tmp_path / "docs" / "spec" / "gone.md")
    except OSError:
        pytest.skip("this platform or account cannot create symlinks")

    result = placeholder_hook.main(["docs/spec/link.md"], root=tmp_path)

    captured = capsys.readouterr()
    assert result == 0
    assert captured.out == ""
    assert captured.err == ""
