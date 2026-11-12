"""Exercise shared template-sync materialization planning helpers."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, cast

import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT_DIR = REPO_ROOT / ".template-sync" / "scripts"
MANIFEST_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-manifest.schema.json"
MARKER_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-marker.schema.json"
VALID_MARKER_EXAMPLE_PATHS = tuple(
    sorted((REPO_ROOT / "schemas" / "examples" / "template-sync-marker" / "valid").glob("*.yml"))
)

if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from template_sync_materialization_helpers import (  # noqa: E402
    EMBEDDED_MARKDOWN_FENCE_CONTEXT,
    MARKDOWN_FENCE_CONTEXT,
    LocalPathOwnership,
    ManifestCompatibilityAlternative,
    ManifestCompatibilityGroup,
    MismatchedInlineBlockError,
    MissingExpectedInlineBlockError,
    NestedInlineBlockError,
    RepositoryPathError,
    TemplateSyncMaterializationError,
    UnclosedInlineBlockError,
    UnmatchedInlineBlockEndError,
    classify_repository_file,
    ensure_regular_repository_file_target,
    format_marker_yaml,
    inline_block_families_to_prune,
    is_excluded_template_path,
    is_protected_instruction_path,
    is_protected_manifest_pattern,
    is_protected_prose_path,
    is_retained_template_path,
    is_template_managed_path,
    lines_outside_markdown_fences,
    load_validated_manifest,
    load_validated_marker_decision_data,
    markdown_lines,
    parse_manifest_compatibility_groups,
    parse_manifest_mappings,
    remove_inline_block_family,
    remove_inline_blocks_for_modules,
    resolve_safe_repository_target_path,
    selected_local_path_ownership_for_path,
    selected_relation_for_path,
    validate_inline_block_markers,
    validate_marker_yaml_text,
    validate_module_compatibility,
    write_repository_relative_file_bytes,
)

SOURCE_REPO = "https://github.com/franklesniak/copilot-repo-template.git"
FULL_SHA = "0123456789abcdef0123456789abcdef01234567"


def _manifest(path_mappings: list[dict[str, Any]]) -> dict[str, Any]:
    """Build a small schema-shaped manifest fixture."""
    return {
        "template_manifest": {
            "version": 2,
            "modules": [
                {"name": "baseline", "description": "Baseline files."},
                {"name": "python", "description": "Python files."},
                {"name": "yaml", "description": "YAML files."},
                {"name": "schema", "description": "Schema files."},
                {
                    "name": "template-sync-support",
                    "description": "Template sync support files.",
                },
            ],
            "path_mappings": path_mappings,
            "filtering": {
                "default_semantics": "AND",
                "requires_any_semantics": "OR",
                "path_matching": "most_specific_match_wins",
                "same_specificity_action": "union_modules",
                "unmapped_action": "surface_for_owner",
            },
            "notes": {
                "downstream_retention": "Downstream repositories keep marker data.",
            },
        }
    }


def _write_yaml(repo_root: Path, relative_path: str, data: dict[str, Any]) -> None:
    """Write YAML below a fixture repository root."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data, sort_keys=False), encoding="utf-8")


def _marker(included_modules: list[str]) -> dict[str, Any]:
    """Build a marker-shaped decision fixture."""
    return {
        "template_sync": {
            "source_repo": SOURCE_REPO,
            "last_reviewed_template_commit": FULL_SHA,
            "included_modules": included_modules,
            "local_overrides": [
                {
                    "path": "local/",
                    "default_decision": "SKIP",
                    "reason": "Local project owns this directory.",
                }
            ],
            "protected_file_decisions": [
                {
                    "path": "AGENTS.md",
                    "decision": "MERGE",
                    "adoption_mode": "minimal-preservation",
                    "authorization_basis": "Owner authorized protected edits.",
                    "authorized_scope": "AGENTS.md only.",
                }
            ],
        }
    }


def test_manifest_relation_selection_uses_specificity_union_and_module_logic() -> None:
    """Path mapping resolution keeps existing manifest filtering semantics."""
    _module_names, mappings = parse_manifest_mappings(
        _manifest(
            [
                {"pattern": "configs/**", "requires_all": ["baseline"]},
                {"pattern": "configs/app.yml", "requires_all": ["yaml"]},
                {"pattern": "configs/app.yml", "requires_all": ["schema"]},
                {"pattern": "shared.yml", "requires_all": ["baseline"]},
                {"pattern": "shared.yml", "requires_any": ["yaml", "schema"]},
            ]
        )
    )

    relation = selected_relation_for_path("configs/app.yml", mappings)
    assert relation is not None
    assert relation.patterns == ("configs/app.yml", "configs/app.yml")
    assert relation.requires_all == frozenset({"yaml", "schema"})
    assert relation.requires_any == frozenset()
    assert relation.is_retained_by({"yaml", "schema"})
    assert not relation.is_retained_by({"yaml"})

    shared_relation = selected_relation_for_path("shared.yml", mappings)
    assert shared_relation is not None
    assert shared_relation.is_retained_by({"baseline", "yaml"})
    assert shared_relation.is_retained_by({"baseline", "schema"})
    assert not shared_relation.is_retained_by({"baseline"})


def test_manifest_loader_schema_validates_real_manifest() -> None:
    """The helper can load and schema-validate the checked-in manifest."""
    modules, mappings = load_validated_manifest(
        REPO_ROOT,
        REPO_ROOT / ".template-sync" / "manifest.yml",
        MANIFEST_SCHEMA_PATH,
    )

    assert "template-sync-support" in modules
    assert selected_relation_for_path(".template-sync/scripts/validate_marker.py", mappings)


def test_marker_loader_schema_validates_decision_data(tmp_path: Path) -> None:
    """Marker-shaped decision data is loaded, schema-validated, and parsed."""
    marker = _marker(["baseline"])
    marker["template_sync"]["protected_guide_contract_waivers"] = [
        {
            "path": "AGENTS.md",
            "contract_key": "agents-azure-devops-pr-review-protocol",
            "target_module": "azure-devops-collaboration",
            "reason": "GitHub-only fixture keeps protected guide text visible.",
            "authorization_basis": "Fixture owner authorized this waiver.",
        }
    ]
    _write_yaml(tmp_path, ".template-sync/marker.yml", marker)

    marker_data = load_validated_marker_decision_data(
        tmp_path,
        tmp_path / ".template-sync" / "marker.yml",
        MARKER_SCHEMA_PATH,
        validate_protected_decision_integrity=True,
    )

    assert marker_data.included_modules == frozenset({"baseline"})
    assert marker_data.local_overrides[0].matches("local/file.txt")
    assert marker_data.protected_decisions[0].path == "AGENTS.md"
    assert marker_data.protected_guide_contract_waivers[0].contract_key == (
        "agents-azure-devops-pr-review-protocol"
    )


def test_format_marker_yaml_is_canonical_and_validates_serialized_text() -> None:
    """Canonical marker YAML has yamllint-compatible sequence indentation."""
    marker_document = _marker(["baseline", "template-sync-support"])

    marker_text = format_marker_yaml(marker_document)
    parsed = validate_marker_yaml_text(marker_text, repo_root=REPO_ROOT)

    assert parsed == marker_document
    assert not marker_text.startswith("---")
    assert marker_text.endswith("\n")
    assert not marker_text.endswith("\n\n")
    assert "  included_modules:\n    - baseline\n    - template-sync-support\n" in marker_text
    assert "  local_overrides:\n    - path: local/\n" in marker_text


@pytest.mark.parametrize("example_path", VALID_MARKER_EXAMPLE_PATHS, ids=lambda path: path.name)
def test_valid_marker_examples_match_canonical_formatter(example_path: Path) -> None:
    """Committed valid marker examples stay in the canonical generated format."""
    marker_text = example_path.read_text(encoding="utf-8")
    marker_document = yaml.safe_load(marker_text)

    assert isinstance(marker_document, dict)
    assert format_marker_yaml(cast(dict[str, Any], marker_document)) == marker_text


def test_repository_file_writer_creates_parent_directories(tmp_path: Path) -> None:
    """Shared file writes create parents and write bytes exactly."""
    target_path = tmp_path / ".template-sync" / "marker.yml"

    written_path = write_repository_relative_file_bytes(
        tmp_path,
        ".template-sync/marker.yml",
        b"marker\n",
        field_name="marker path",
    )

    assert written_path == target_path
    assert target_path.read_bytes() == b"marker\n"


def test_regular_target_guard_rejects_existing_symlink_target(
    tmp_path: Path,
) -> None:
    """Existing symlink targets are rejected even when they point at regular files."""
    real_target = tmp_path / "real-marker.yml"
    real_target.write_text("template_sync:\n  source_repo: x\n", encoding="utf-8")
    symlink_target = tmp_path / ".template-sync" / "marker.yml"
    symlink_target.parent.mkdir()
    try:
        symlink_target.symlink_to(real_target)
    except (OSError, NotImplementedError) as error:
        if isinstance(error, OSError):
            detail = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        else:
            detail = type(error).__name__
        pytest.skip(f"Filesystem does not support symlink creation: {detail}")

    with pytest.raises(TemplateSyncMaterializationError) as excinfo:
        ensure_regular_repository_file_target(symlink_target, ".template-sync/marker.yml")

    assert ".template-sync/marker.yml" in str(excinfo.value)
    assert "not a regular file" in str(excinfo.value)


def test_local_path_ownership_matching_uses_most_specific_record() -> None:
    """Parent and child local ownership records are permitted and deterministic."""
    parent = LocalPathOwnership(
        path="docs",
        reason="General documentation ownership.",
        overlap_exception_reason=None,
        is_directory=True,
    )
    child = LocalPathOwnership(
        path="docs/api",
        reason="API documentation ownership.",
        overlap_exception_reason=None,
        is_directory=True,
    )
    exact = LocalPathOwnership(
        path="docs/api/index.md",
        reason="Landing page ownership.",
        overlap_exception_reason=None,
        is_directory=False,
    )
    records = (parent, child, exact)

    assert selected_local_path_ownership_for_path("docs", records) == parent
    assert selected_local_path_ownership_for_path("docs/guide.md", records) == parent
    assert selected_local_path_ownership_for_path("docs/api/reference.md", records) == child
    assert selected_local_path_ownership_for_path("docs/api/index.md", records) == exact
    assert selected_local_path_ownership_for_path("src/app.py", records) is None


def test_mapping_classification_identifies_retained_excluded_and_unmapped_paths() -> None:
    """Shared helpers distinguish mapped, retained, excluded, and unmapped paths."""
    _module_names, mappings = parse_manifest_mappings(
        _manifest(
            [
                {"pattern": "README.md", "requires_all": ["baseline"]},
                {"pattern": "templates/python/**", "requires_all": ["python"]},
            ]
        )
    )
    included_modules = {"baseline"}

    assert is_template_managed_path("README.md", mappings)
    assert is_retained_template_path("README.md", mappings, included_modules)
    assert is_excluded_template_path("templates/python/app.py", mappings, included_modules)
    assert not is_template_managed_path("docs/unmapped.md", mappings)


def test_manifest_compatibility_groups_parse_and_validate_host_mixes() -> None:
    """Compatibility helpers allow declared mixes and report unsupported host mixes."""
    manifest = _manifest(
        [
            {"pattern": ".github/workflows/**", "requires_all": ["github-actions"]},
            {"pattern": ".azuredevops/pipelines/**", "requires_all": ["azure-pipelines"]},
        ]
    )
    manifest["template_manifest"]["modules"].extend(
        [
            {"name": "github-actions", "description": "GitHub Actions files."},
            {"name": "azure-pipelines", "description": "Azure Pipelines files."},
        ]
    )
    manifest["template_manifest"]["compatibility_groups"] = [
        {
            "name": "ci-host",
            "description": "CI host modules.",
            "default_modules": ["github-actions"],
            "alternatives": [
                {"host": "github", "modules": ["github-actions"]},
                {"host": "azure-devops", "modules": ["azure-pipelines"]},
            ],
            "mixed_hosts": "allowed",
        }
    ]

    groups = parse_manifest_compatibility_groups(manifest)

    assert groups[0].default_modules == frozenset({"github-actions"})
    assert not validate_module_compatibility({"github-actions", "azure-pipelines"}, groups)

    unsupported_group = ManifestCompatibilityGroup(
        name="ci-host",
        description="CI host modules.",
        default_modules=frozenset({"github-actions"}),
        alternatives=(
            ManifestCompatibilityAlternative(
                host="github",
                modules=frozenset({"github-actions"}),
            ),
            ManifestCompatibilityAlternative(
                host="azure-devops",
                modules=frozenset({"azure-pipelines"}),
            ),
        ),
        mixed_hosts="unsupported",
    )

    errors = validate_module_compatibility(
        {"github-actions", "azure-pipelines"},
        (unsupported_group,),
    )

    assert len(errors) == 1
    assert "Unsupported mixed-host selection in compatibility group ci-host" in errors[0]
    assert "Select one host family" in errors[0]


def test_protected_file_classification_uses_shared_rules() -> None:
    """Protected instruction path classification follows repository governance rules."""
    assert is_protected_instruction_path(".github/copilot-instructions.md")
    assert is_protected_instruction_path(".github/instructions/python.instructions.md")
    assert is_protected_instruction_path(".cursor/rules/repository-instructions.mdc")
    assert is_protected_instruction_path(".template-sync/instruction-contracts.yml")
    assert is_protected_manifest_pattern(".template-sync/instruction-contracts.yml")
    assert not is_protected_instruction_path(".template-sync/manifest.yml")
    assert not is_protected_instruction_path(
        "schemas/template-sync-instruction-contracts.schema.json"
    )
    assert not is_protected_instruction_path("README.md")
    for path in (
        ".github/copilot-instructions.md",
        ".hermes.md",
        "AGENTS.md",
        "CLAUDE.md",
        "GEMINI.md",
        ".github/instructions/python.instructions.md",
        ".cursor/rules/repository-instructions.mdc",
        ".github/instructions/nested/policy",
        ".cursor/rules/nested/policy",
    ):
        assert is_protected_instruction_path(path)
        assert is_protected_prose_path(path)
    assert not is_protected_prose_path(".template-sync/instruction-contracts.yml")
    assert not is_protected_prose_path(".template-sync/manifest.yml")
    assert not is_protected_prose_path("README.md")


@pytest.mark.parametrize(
    ("source", "without", "with_endings"),
    [
        ("", [], []),
        ("one", ["one"], ["one"]),
        ("one\n", ["one"], ["one\n"]),
        ("one\r", ["one"], ["one\r"]),
        ("one\r\n", ["one"], ["one\r\n"]),
        ("\n\r\r\n", ["", "", ""], ["\n", "\r", "\r\n"]),
        (
            "one\r\ntwo\rthree\nfour",
            ["one", "two", "three", "four"],
            ["one\r\n", "two\r", "three\n", "four"],
        ),
    ],
)
def test_markdown_physical_line_endings_are_exact(
    source: str, without: list[str], with_endings: list[str]
) -> None:
    """Only CR/LF/CRLF split source text, including exact terminal-ending behavior."""
    assert markdown_lines(source) == without
    assert markdown_lines(source, keepends=True) == with_endings
    assert "".join(markdown_lines(source, keepends=True)) == source


@pytest.mark.parametrize(
    "literal",
    ["\v", "\f", "\x1c", "\x1d", "\x1e", "\x85", "\u2028", "\u2029", "\u00a0", "\u2003", "\x1f"],
)
def test_non_markdown_separators_remain_literal_in_shared_consumers(literal: str) -> None:
    """Unicode/control whitespace cannot close a fence or turn text into a marker."""
    assert markdown_lines("left" + literal + "right") == ["left" + literal + "right"]
    assert markdown_lines(literal, keepends=True) == [literal]
    for fence in ("```", "~~~"):
        for context in (MARKDOWN_FENCE_CONTEXT, EMBEDDED_MARKDOWN_FENCE_CONTEXT):
            for ending in ("\r", "\n", "\r\n"):
                fenced = fence + ending + "example" + ending + fence + literal
                fenced += ending + "Must remain code"
                assert lines_outside_markdown_fences(fenced, fence_context=context) == ()
    for location in ("prefix", "suffix", "internal"):
        begin = "<!-- template-sync: begin python-reference-only -->"
        end = "<!-- template-sync: end python-reference-only -->"
        if location == "prefix":
            begin, end = literal + begin, literal + end
        elif location == "suffix":
            begin, end = begin + literal, end + literal
        else:
            begin, end = begin.replace(": ", ":" + literal), end.replace(": ", ":" + literal)
        text = begin + "\nPOLICY\n" + end + "\n"
        assert (
            remove_inline_blocks_for_modules(
                text, {"baseline", "markdown"}, relative_path="README.md"
            )
            == text
        )


@pytest.mark.parametrize("ending", ["\r", "\n", "\r\n"])
@pytest.mark.parametrize("spacing", ["", " ", "\t", " \t"])
def test_ascii_marker_spacing_and_real_line_endings_still_prune(ending: str, spacing: str) -> None:
    """Ordinary ASCII marker indentation retains exact supported pruning behavior."""
    text = (
        spacing
        + "<!-- template-sync: begin python-reference-only -->"
        + spacing
        + ending
        + "removed"
        + ending
        + spacing
        + "<!-- template-sync: end python-reference-only -->"
        + spacing
        + ending
        + "retained"
        + ending
    )
    assert (
        remove_inline_blocks_for_modules(text, {"baseline", "markdown"}, relative_path="README.md")
        == "retained" + ending
    )


def test_physical_line_slices_have_bounded_total_work() -> None:
    """The shared splitter copies each source character at most once per result."""

    class BoundedText(str):
        copied = 0

        def __getitem__(self, key: Any) -> str:
            assert isinstance(key, slice)
            self.copied += len(range(*key.indices(len(self))))
            assert self.copied <= len(self)
            return super().__getitem__(key)

    text = BoundedText(("first\r\nsecond\rthird\nfourth\v\u2028\n") * 10000)
    observed = markdown_lines(text, keepends=True)
    assert "".join(observed) == text
    assert text.copied == len(text)


def test_non_ascii_blank_hygiene_and_marker_guard_mutants() -> None:
    """Actual pruning preserves literal lines; restoring broad marker grammar breaks it."""
    import re

    import template_sync_materialization_helpers as helpers

    literal = "\u00a0"
    marker = "<!-- template-sync: {kind} python-reference-only -->"
    text = literal + marker.format(kind="begin") + "\nPOLICY\n"
    text += literal + marker.format(kind="end") + "\n"
    assert (
        remove_inline_blocks_for_modules(text, {"baseline", "markdown"}, relative_path="README.md")
        == text
    )
    original = helpers.INLINE_BLOCK_MARKER_RE
    try:
        helpers.INLINE_BLOCK_MARKER_RE = re.compile(original.pattern.replace(r"[ \t]", r"\s"))
        assert (
            remove_inline_blocks_for_modules(
                text, {"baseline", "markdown"}, relative_path="README.md"
            )
            == ""
        )
    finally:
        helpers.INLINE_BLOCK_MARKER_RE = original
    for literal in (
        "\v",
        "\f",
        "\x1c",
        "\x1d",
        "\x1e",
        "\x85",
        "\u2028",
        "\u2029",
        "\u00a0",
        "\u2003",
        "\x1f",
    ):
        content = "top\n" + (literal + "\n") * 4 + "bottom\n"
        assert helpers.apply_blank_line_hygiene(content, max_consecutive_blank_lines=1) == content


def test_inline_block_removal_uses_yaml_blank_line_limit() -> None:
    """Removing a YAML inline block collapses runs to yamllint's one-blank limit."""
    text = (
        "top\n"
        "\n"
        "# template-sync: begin python-only\n"
        "removed\n"
        "# template-sync: end python-only\n"
        "\n"
        "\n"
        "\n"
        "bottom\n"
    )

    assert (
        remove_inline_block_family(text, "python-only", relative_path="fixture.yml")
        == "top\n\nbottom\n"
    )


def test_inline_block_removal_preserves_indented_yaml_fenced_blanks() -> None:
    """Deeply-indented YAML block-scalar fenced examples keep their blank runs."""
    text = (
        "# template-sync: begin python-only\n"
        "removed\n"
        "# template-sync: end python-only\n"
        "\n"
        "        ```\n"
        "        first\n"
        "\n"
        "\n"
        "\n"
        "        second\n"
        "        ```\n"
    )

    expected = "".join(
        [
            "\n",
            "        ```\n",
            "        first\n",
            "\n",
            "\n",
            "\n",
            "        second\n",
            "        ```\n",
        ]
    )
    assert remove_inline_block_family(text, "python-only", relative_path="fixture.yml") == expected


def test_inline_block_removal_preserves_markdown_blank_line_allowance() -> None:
    """Markdown pruning can retain markdownlint's configured two-blank allowance."""
    text = (
        "top\n"
        "\n"
        "<!-- template-sync: begin python-reference-only -->\n"
        "removed\n"
        "<!-- template-sync: end python-reference-only -->\n"
        "\n"
        "<!-- template-sync: begin terraform-reference-only -->\n"
        "removed\n"
        "<!-- template-sync: end terraform-reference-only -->\n"
        "\n"
        "\n"
        "bottom\n"
    )

    assert (
        remove_inline_blocks_for_modules(
            text,
            {"markdown"},
            relative_path="fixture.md",
        )
        == "top\n\n\nbottom\n"
    )


def test_inline_block_removal_preserves_blank_runs_inside_fenced_code() -> None:
    """Blank-line runs inside fenced code examples survive inline-block pruning."""
    text = (
        "```python\n"
        "first = 1\n"
        "\n"
        "\n"
        "\n"
        "second = 2\n"
        "```\n"
        "\n"
        "<!-- template-sync: begin terraform-reference-only -->\n"
        "terraform note\n"
        "<!-- template-sync: end terraform-reference-only -->\n"
        "\n"
        "tail\n"
    )

    expected = "".join(
        [
            "```python\n",
            "first = 1\n",
            "\n",
            "\n",
            "\n",
            "second = 2\n",
            "```\n",
            "\n",
            "\n",
            "tail\n",
        ]
    )
    assert (
        remove_inline_block_family(
            text,
            "terraform-reference-only",
            relative_path="README.md",
        )
        == expected
    )


def test_inline_block_removal_preserves_trailing_blanks_in_markdown_fence() -> None:
    """A fenced Markdown example keeps its trailing blank lines after pruning."""
    text = (
        "```markdown\n"
        "Example heading\n"
        "\n"
        "\n"
        "```\n"
        "\n"
        "<!-- template-sync: begin terraform-reference-only -->\n"
        "terraform note\n"
        "<!-- template-sync: end terraform-reference-only -->\n"
        "\n"
        "tail\n"
    )

    expected = "".join(
        [
            "```markdown\n",
            "Example heading\n",
            "\n",
            "\n",
            "```\n",
            "\n",
            "\n",
            "tail\n",
        ]
    )
    assert (
        remove_inline_block_family(
            text,
            "terraform-reference-only",
            relative_path="README.md",
        )
        == expected
    )


def test_markdown_fenced_marker_examples_are_not_pruned() -> None:
    """Registered marker lines inside Markdown fences are documentation examples."""
    text = (
        "# Example\n\n"
        "```markdown\n"
        "<!-- template-sync: begin python-reference-only -->\n"
        "Example only.\n"
        "<!-- template-sync: end python-reference-only -->\n"
        "```\n"
        "live\n"
    )

    assert (
        remove_inline_blocks_for_modules(
            text,
            {"baseline", "markdown"},
            relative_path="README.md",
        )
        == text
    )
    assert (
        inline_block_families_to_prune(
            text,
            {"baseline", "markdown"},
            relative_path="README.md",
        )
        == ()
    )


@pytest.mark.parametrize("indent", ["", " ", "  ", "   "])
def test_markdown_fenced_marker_examples_may_be_indented_up_to_three_spaces(
    indent: str,
) -> None:
    """Fence-aware marker pruning follows GFM's 0-3 space fence indentation."""
    text = (
        f"{indent}```\n"
        "<!-- template-sync: begin python-reference-only -->\n"
        "Example only.\n"
        "<!-- template-sync: end python-reference-only -->\n"
        f"{indent}```\n"
    )

    assert (
        remove_inline_blocks_for_modules(
            text,
            {"baseline", "markdown"},
            relative_path="README.md",
        )
        == text
    )


def test_non_markdown_fenced_marker_lines_remain_live() -> None:
    """Non-Markdown text files keep existing marker semantics."""
    text = (
        "```\n"
        "# template-sync: begin python-only\n"
        "removed\n"
        "# template-sync: end python-only\n"
        "```\n"
    )

    assert (
        remove_inline_blocks_for_modules(
            text,
            {"baseline"},
            relative_path="fixture.txt",
        )
        == "```\n```\n"
    )


def test_blockquote_fenced_marker_examples_are_not_pruned() -> None:
    """Blockquote-prefixed Markdown fences suppress marker pruning inside them."""
    text = (
        "> ```\n"
        "> <!-- template-sync: begin python-reference-only -->\n"
        "> Example only.\n"
        "> <!-- template-sync: end python-reference-only -->\n"
        "> ```\n"
    )

    assert (
        remove_inline_blocks_for_modules(
            text,
            {"baseline", "markdown"},
            relative_path="README.md",
        )
        == text
    )


def test_blockquote_containing_block_closure_restores_live_markers() -> None:
    """A blockquote-contained unclosed fence does not hide later unquoted markers."""
    text = (
        "> ```\n"
        "> <!-- template-sync: begin python-reference-only -->\n"
        "> Example only.\n"
        "\n"
        "<!-- template-sync: begin python-reference-only -->\n"
        "Live block.\n"
        "<!-- template-sync: end python-reference-only -->\n"
    )

    assert (
        remove_inline_blocks_for_modules(
            text,
            {"baseline", "markdown"},
            relative_path="README.md",
        )
        == "> ```\n> <!-- template-sync: begin python-reference-only -->\n"
        "> Example only.\n\n"
    )


@pytest.mark.parametrize(
    ("list_opener", "continuation_indent"),
    [
        ("- ```", "  "),
        ("10. ```", "    "),
    ],
)
def test_list_item_containing_block_closure_restores_live_markers(
    list_opener: str,
    continuation_indent: str,
) -> None:
    """A list-item-contained unclosed fence does not hide following live markers."""
    text = (
        f"{list_opener}\n"
        f"{continuation_indent}<!-- template-sync: begin python-reference-only -->\n"
        f"{continuation_indent}Example only.\n"
        "\n"
        "<!-- template-sync: begin python-reference-only -->\n"
        "Live block.\n"
        "<!-- template-sync: end python-reference-only -->\n"
    )

    assert (
        remove_inline_blocks_for_modules(
            text,
            {"baseline", "markdown"},
            relative_path="README.md",
        )
        == f"{list_opener}\n"
        f"{continuation_indent}<!-- template-sync: begin python-reference-only -->\n"
        f"{continuation_indent}Example only.\n\n"
    )


@pytest.mark.parametrize(
    ("text", "expected_pruned"),
    [
        pytest.param(
            (
                "```\n"
                "~~~\n"
                "<!-- template-sync: begin python-reference-only -->\n"
                "Example only.\n"
                "<!-- template-sync: end python-reference-only -->\n"
                "```\n"
            ),
            False,
            id="mismatched-close-character",
        ),
        pytest.param(
            (
                "````\n"
                "```\n"
                "<!-- template-sync: begin python-reference-only -->\n"
                "Example only.\n"
                "<!-- template-sync: end python-reference-only -->\n"
                "````\n"
            ),
            False,
            id="too-short-close",
        ),
        pytest.param(
            (
                "```\n"
                "``` trailing\n"
                "<!-- template-sync: begin python-reference-only -->\n"
                "Example only.\n"
                "<!-- template-sync: end python-reference-only -->\n"
                "```\n"
            ),
            False,
            id="close-with-trailing-text",
        ),
        pytest.param(
            (
                "```bad`info\n"
                "<!-- template-sync: begin python-reference-only -->\n"
                "Live because the opener is invalid.\n"
                "<!-- template-sync: end python-reference-only -->\n"
            ),
            True,
            id="invalid-backtick-info",
        ),
    ],
)
def test_markdown_fence_rules_control_marker_liveness(
    text: str,
    expected_pruned: bool,
) -> None:
    """Marker liveness follows same-character close and valid-info GFM rules."""
    result = remove_inline_blocks_for_modules(
        text,
        {"baseline", "markdown"},
        relative_path="README.md",
    )

    assert ("python-reference-only" not in result) is expected_pruned


@pytest.mark.parametrize("context", [MARKDOWN_FENCE_CONTEXT, EMBEDDED_MARKDOWN_FENCE_CONTEXT])
@pytest.mark.parametrize("fence", ["```", "~~~"])
@pytest.mark.parametrize(
    ("suffix", "closed"),
    [("\t", True), (" \t ", True), ("text", False), ("\u00a0", False), ("\v", False)],
)
def test_shared_fence_close_whitespace_keeps_following_lines_live(
    context: str, fence: str, suffix: str, closed: bool
) -> None:
    """Only ASCII space and tab may follow a matching closing fence."""
    text = fence + "\nexample\n" + fence + suffix + "\nAfter fence"
    visible = lines_outside_markdown_fences(text, fence_context=context)
    assert any(line == "After fence" for _, line in visible) is closed


@pytest.mark.parametrize("context", [MARKDOWN_FENCE_CONTEXT, EMBEDDED_MARKDOWN_FENCE_CONTEXT])
@pytest.mark.parametrize("marker", ["١.", "１)", "1١.", "१२.", "1.", "2)"])
def test_shared_fence_list_markers_use_ascii_digits(context: str, marker: str) -> None:
    """Both normal and embedded Markdown keep Unicode-digit pseudo-items live."""
    indent = "        " if context == EMBEDDED_MARKDOWN_FENCE_CONTEXT else ""
    text = indent + marker + " ```\n" + indent + "   Visible unless a real list"
    visible = lines_outside_markdown_fences(text, fence_context=context)
    assert any("Visible unless" in line for _, line in visible) is (marker not in {"1.", "2)"})


def test_embedded_markdown_fence_context_preserves_yaml_block_scalar_examples() -> None:
    """Embedded-Markdown mode recognizes fences indented inside YAML scalars."""
    text = (
        "body:\n"
        "  - type: markdown\n"
        "    attributes:\n"
        "      value: |\n"
        "        ```\n"
        "        [inside](templates/python/example.py)\n"
        "        ```\n"
        "        [outside](templates/python/example.py)\n"
    )

    visible_lines = lines_outside_markdown_fences(
        text,
        fence_context=EMBEDDED_MARKDOWN_FENCE_CONTEXT,
    )

    assert (6, "        [inside](templates/python/example.py)") not in visible_lines
    assert (8, "        [outside](templates/python/example.py)") in visible_lines


def test_standard_markdown_context_keeps_yaml_indented_fences_visible() -> None:
    """Standard Markdown mode does not apply global YAML block-scalar rules."""
    text = "".join(
        [
            "        ```\n",
            "        [inside](templates/python/example.py)\n",
            "        ```\n",
        ]
    )

    visible_lines = lines_outside_markdown_fences(
        text,
        fence_context=MARKDOWN_FENCE_CONTEXT,
    )

    assert (2, "        [inside](templates/python/example.py)") in visible_lines


def test_embedded_fence_context_recognizes_blockquote_contained_fences() -> None:
    """Embedded-Markdown mode treats deeply-indented blockquote fences as opaque."""
    text = "".join(
        [
            "        > ```\n",
            "        > [inside](templates/python/example.py)\n",
            "        > ```\n",
            "        [outside](templates/python/example.py)\n",
        ]
    )

    visible_lines = lines_outside_markdown_fences(
        text,
        fence_context=EMBEDDED_MARKDOWN_FENCE_CONTEXT,
    )

    assert (2, "        > [inside](templates/python/example.py)") not in visible_lines
    assert (4, "        [outside](templates/python/example.py)") in visible_lines


def test_embedded_fence_context_recognizes_list_contained_fences() -> None:
    """Embedded-Markdown mode treats deeply-indented list-item fences as opaque."""
    text = "".join(
        [
            "        - ```\n",
            "          [inside](templates/python/example.py)\n",
            "          ```\n",
            "        [outside](templates/python/example.py)\n",
        ]
    )

    visible_lines = lines_outside_markdown_fences(
        text,
        fence_context=EMBEDDED_MARKDOWN_FENCE_CONTEXT,
    )

    assert (2, "          [inside](templates/python/example.py)") not in visible_lines
    assert (4, "        [outside](templates/python/example.py)") in visible_lines


@pytest.mark.parametrize(
    ("text", "error_type"),
    [
        (
            "# template-sync: end python-only\n",
            UnmatchedInlineBlockEndError,
        ),
        (
            "# template-sync: begin python-only\n",
            UnclosedInlineBlockError,
        ),
        (
            (
                "# template-sync: begin python-only\n"
                "# template-sync: begin terraform-only\n"
                "# template-sync: end terraform-only\n"
                "# template-sync: end python-only\n"
            ),
            NestedInlineBlockError,
        ),
        (
            "".join(
                [
                    "# template-sync: begin python-only\n",
                    "# template-sync: end terraform-only\n",
                ]
            ),
            MismatchedInlineBlockError,
        ),
    ],
)
def test_live_inline_marker_validation_reports_structural_errors(
    text: str,
    error_type: type[Exception],
) -> None:
    """Shared validation raises typed errors for malformed live markers."""
    with pytest.raises(error_type):
        validate_inline_block_markers(text, relative_path="fixture.yml")


def test_azure_devops_guide_reference_only_uses_or_retention() -> None:
    """Azure guide references strip only when every Azure host module is absent."""
    text = (
        "top\n"
        "<!-- template-sync: begin azure-devops-guide-reference-only -->\n"
        "See docs/azure-devops-support.md.\n"
        "<!-- template-sync: end azure-devops-guide-reference-only -->\n"
        "bottom\n"
    )

    assert (
        remove_inline_blocks_for_modules(
            text,
            {"baseline", "github-actions", "github-platform", "github-templates"},
            relative_path="README.md",
        )
        == "top\nbottom\n"
    )
    for module_name in (
        "azure-devops-platform",
        "azure-pipelines",
        "azure-devops-collaboration",
    ):
        assert (
            remove_inline_blocks_for_modules(
                text,
                {"baseline", module_name},
                relative_path="README.md",
            )
            == text
        )


def test_inline_block_removal_requires_expected_family() -> None:
    """Materialization callers get a typed error when an expected family is absent."""
    with pytest.raises(MissingExpectedInlineBlockError):
        remove_inline_block_family("plain\n", "python-only", relative_path="fixture.yml")


def test_text_versus_byte_only_classification_preserves_bytes(tmp_path: Path) -> None:
    """Only safe UTF-8 files without NUL bytes are transformable text."""
    text_path = tmp_path / "docs" / "note.txt"
    text_path.parent.mkdir(parents=True)
    text_path.write_bytes(b"hello\n")
    invalid_bytes = b"\xff\xfe\x00binary"
    byte_path = tmp_path / "assets" / "blob.bin"
    byte_path.parent.mkdir(parents=True)
    byte_path.write_bytes(invalid_bytes)

    text_classification = classify_repository_file(tmp_path, "docs/note.txt")
    byte_classification = classify_repository_file(tmp_path, "assets/blob.bin")

    assert text_classification.is_transformable_text
    assert text_classification.text == "hello\n"
    assert byte_classification.is_byte_only
    assert byte_classification.bytes_content == invalid_bytes
    assert byte_classification.text is None


def test_path_safety_rejects_traversal_absolute_backslash_and_symlink(
    tmp_path: Path,
) -> None:
    """Safe target resolution rejects paths that can escape the repository."""
    for raw_path in ("../outside.txt", "/absolute.txt", r"dir\file.txt"):
        with pytest.raises(RepositoryPathError):
            resolve_safe_repository_target_path(tmp_path, raw_path)

    outside = tmp_path.parent / "outside-target"
    outside.mkdir(exist_ok=True)
    symlink_path = tmp_path / "linked"
    try:
        symlink_path.symlink_to(outside, target_is_directory=True)
    except OSError as error:
        pytest.skip(f"Symlink creation unavailable in this environment: {error}")

    with pytest.raises(RepositoryPathError):
        resolve_safe_repository_target_path(tmp_path, "linked/file.txt")
