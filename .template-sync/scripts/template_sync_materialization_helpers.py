"""Shared planning helpers for template-sync materialization workflows."""

# ruff: noqa: E402

from __future__ import annotations

import fnmatch
import os
import re
import subprocess
import sys
from collections.abc import Collection, Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import Any, TypeGuard, cast

import yaml  # type: ignore[import-untyped]

_SHARED_SCRIPTS = Path(__file__).resolve().parents[2] / ".github" / "scripts"
if str(_SHARED_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SHARED_SCRIPTS))

from instruction_contract_support import (
    EMBEDDED_LIST_MARKER_RE,
    EMBEDDED_MARKDOWN_FENCE_CONTEXT,
    LIST_MARKER_RE,
    MARKDOWN_FENCE_CONTEXT,
    REMOVAL_DECISION,
    MarkdownFence,
    MarkdownLineState,
    ProtectedFileDecision,
    ProtectedGuideContractWaiver,
    RepositoryPathError,
    TemplateSyncMaterializationError,
    active_fence_content,
    consume_blockquote_prefix,
    line_body,
    line_is_outside_list_item,
    lines_outside_markdown_fences,
    load_json_mapping,
    load_yaml_mapping,
    markdown_line_states,
    markdown_lines,
    normalize_repository_path,
    os_error_summary,
    parse_fence_close_from_content,
    parse_fence_open_from_content,
    parse_json_mapping,
    parse_markdown_fence_open,
    parse_yaml_mapping,
    read_repository_text,
    repository_relative_path,
    resolve_repo_path,
    validate_schema,
)

# Explicit re-exports preserve the existing adapter API for sync callers.
__all__ = [
    "EMBEDDED_LIST_MARKER_RE",
    "EMBEDDED_MARKDOWN_FENCE_CONTEXT",
    "LIST_MARKER_RE",
    "MARKDOWN_FENCE_CONTEXT",
    "REMOVAL_DECISION",
    "MarkdownFence",
    "MarkdownLineState",
    "ProtectedFileDecision",
    "ProtectedGuideContractWaiver",
    "RepositoryPathError",
    "TemplateSyncMaterializationError",
    "active_fence_content",
    "consume_blockquote_prefix",
    "line_body",
    "line_is_outside_list_item",
    "lines_outside_markdown_fences",
    "load_json_mapping",
    "load_yaml_mapping",
    "markdown_line_states",
    "markdown_lines",
    "normalize_repository_path",
    "os_error_summary",
    "parse_fence_close_from_content",
    "parse_fence_open_from_content",
    "parse_json_mapping",
    "parse_markdown_fence_open",
    "parse_yaml_mapping",
    "read_repository_text",
    "repository_relative_path",
    "resolve_repo_path",
    "validate_schema",
]

DEFAULT_MARKER_PATH = ".template-sync/marker.yml"
DEFAULT_MANIFEST_PATH = ".template-sync/manifest.yml"
DEFAULT_MARKER_SCHEMA_PATH = "schemas/template-sync-marker.schema.json"
DEFAULT_MANIFEST_SCHEMA_PATH = "schemas/template-sync-manifest.schema.json"
DEFAULT_REMOVE_LOCAL_AUTHORIZATION_TOKENS = ("remov", "delet")
SKIPPED_DISCOVERY_DIRS = frozenset(
    {
        ".git",
        ".mypy_cache",
        ".pytest_cache",
        ".ruff_cache",
        ".venv",
        "build",
        "dist",
        "node_modules",
        "__pycache__",
    }
)
AGENT_INSTRUCTION_EXACT_PATHS = frozenset(
    {
        ".github/copilot-instructions.md",
        ".hermes.md",
        "AGENTS.md",
        "CLAUDE.md",
        "GEMINI.md",
    }
)
PROTECTED_EXACT_PATHS = (
    AGENT_INSTRUCTION_EXACT_PATHS
    | frozenset({".template-sync/instruction-contracts.yml"})
    | frozenset(
        {
            ".github/workflow-security-contract.yml",
            ".github/instruction-profile.yml",
            "schemas/instruction-contracts.schema.json",
            "schemas/instruction-profile.schema.json",
            ".github/instruction-contracts.yml",
        }
    )
)
PROTECTED_GLOB_PATTERNS = (
    ".github/instructions/**",
    ".cursor/rules/**",
)
INLINE_BLOCK_MARKER_RE = re.compile(
    r"^[ \t]*(?:#[ \t]*template-sync:|<!--[ \t]*template-sync:)[ \t]*"
    r"(?P<kind>begin|end)[ \t]+"
    r"(?P<name>[a-z0-9-]+-(?:reference-)?only)[ \t]*(?:-->)?[ \t]*(?:\r\n?|\n)?\Z"
)
MARKDOWN_RENDERED_SUFFIXES = frozenset({".md", ".mdc"})
YAML_SUFFIXES = frozenset({".yml", ".yaml"})
# Same as ``LIST_MARKER_RE`` but with no 0-3 space cap on the leading indent, so
# list-contained fences inside deeply-indented YAML block scalars are recognized
# under the embedded-Markdown fence context.
# AND-retention markers. A block in this family is retained only when *every*
# module it names is present in ``included_modules``; it is stripped when *any*
# named module is excluded (see ``remove_inline_blocks_for_modules``). This is
# the default inline-block semantics and covers both the ``*-only`` toolchain
# blocks and the single-module ``*-reference-only`` documentation blocks.
INLINE_BLOCK_MODULES = {
    "powershell-only": frozenset({"powershell"}),
    "instruction-enforcement-only": frozenset({"agent-instructions", "instruction-enforcement"}),
    "instruction-enforcement-reference-only": frozenset(
        {"agent-instructions", "instruction-enforcement"}
    ),
    "claude-review-command-reference-only": frozenset(
        {"agent-instructions", "agent-claude", "github-actions"}
    ),
    "copilot-setup-reference-only": frozenset(
        {"agent-instructions", "agent-copilot", "github-actions"}
    ),
    "baseline-only": frozenset({"baseline"}),
    "baseline-reference-only": frozenset({"baseline"}),
    "github-data-ci-reference-only": frozenset({"baseline", "github-actions"}),
    "git-lfs-only": frozenset({"git-lfs"}),
    "terraform-only": frozenset({"terraform"}),
    "markdown-only": frozenset({"markdown"}),
    "python-only": frozenset({"python"}),
    "yaml-only": frozenset({"yaml"}),
    "schema-only": frozenset({"schema"}),
    "template-sync-support-only": frozenset({"template-sync-support"}),
    "schema-template-sync-support-only": frozenset({"schema", "template-sync-support"}),
    "github-actions-only": frozenset({"github-actions"}),
    "github-platform-only": frozenset({"github-platform"}),
    "markdown-reference-only": frozenset({"markdown"}),
    "powershell-reference-only": frozenset({"powershell"}),
    "python-reference-only": frozenset({"python"}),
    "terraform-reference-only": frozenset({"terraform"}),
    "json-reference-only": frozenset({"json"}),
    "yaml-reference-only": frozenset({"yaml"}),
    "schema-reference-only": frozenset({"schema"}),
    "template-sync-support-reference-only": frozenset({"template-sync-support"}),
    "github-actions-reference-only": frozenset({"github-actions"}),
    "github-platform-reference-only": frozenset({"github-platform"}),
}
# OR-retention (ANY) markers. A block in this family is retained when *at least
# one* module it names is present in ``included_modules``; it is stripped only
# when the included modules are disjoint from the marker's module set (i.e.
# *none* of the named modules is included). This mirrors a manifest
# ``requires_any`` relation, so it can guard prose that documents a file which
# is itself materialized under OR semantics. The pip updater also uses this
# relation to cover baseline runner requirements or Python project metadata.
INLINE_BLOCK_ANY_MODULES = {
    "azure-devops-guide-reference-only": frozenset(
        {"azure-devops-platform", "azure-pipelines", "azure-devops-collaboration"}
    ),
    "pip-dependencies-only": frozenset({"baseline", "python"}),
}


class InlineBlockError(TemplateSyncMaterializationError):
    """Base class for typed inline-block marker errors."""

    def __init__(
        self,
        message: str,
        *,
        relative_path: str,
        line_number: int | None = None,
        marker_name: str | None = None,
    ) -> None:
        """Create an inline-block error with path, line, and marker context."""
        location = relative_path
        if line_number is not None:
            location = f"{location}:{line_number}"
        super().__init__(f"{location}: {message}")
        self.message = message
        self.relative_path = relative_path
        self.line_number = line_number
        self.marker_name = marker_name


class UnknownInlineBlockMarkerError(InlineBlockError):
    """Raised when text contains an unrecognized template-sync marker family."""


class NestedInlineBlockError(InlineBlockError):
    """Raised when an inline block begins before the previous one ended."""


class MismatchedInlineBlockError(InlineBlockError):
    """Raised when an end marker does not match the open begin marker."""


class UnclosedInlineBlockError(InlineBlockError):
    """Raised when an inline block begin marker has no closing end marker."""


class UnmatchedInlineBlockEndError(InlineBlockError):
    """Raised when an inline block end marker appears without a begin marker."""


class MissingExpectedInlineBlockError(InlineBlockError):
    """Raised when a requested inline marker family is absent from the text."""


@dataclass(frozen=True)
class LocalOverride:
    """A marker local override path that suppresses consistency checks."""

    path: str
    default_decision: str
    reason: str
    is_directory: bool

    def matches(self, relative_path: str) -> bool:
        """Return whether this override applies to ``relative_path``."""
        if relative_path == self.path:
            return True
        if self.is_directory:
            return relative_path.startswith(f"{self.path}/")
        return False


@dataclass(frozen=True)
class LocalPathOwnership:
    """A downstream-owned path family recorded in the sync marker."""

    path: str
    reason: str
    overlap_exception_reason: str | None
    is_directory: bool

    @property
    def display_path(self) -> str:
        """Return the marker path form, preserving directory-prefix notation."""
        return f"{self.path}/" if self.is_directory else self.path

    def matches(self, relative_path: str) -> bool:
        """Return whether this ownership record applies to ``relative_path``."""
        if relative_path == self.path:
            return True
        if self.is_directory:
            return relative_path.startswith(f"{self.path}/")
        return False


@dataclass(frozen=True)
class DeferredProtectedCandidate:
    """A protected path awaiting owner authorization for an upstream change."""

    path: str
    source_commit: str
    reason: str


@dataclass(frozen=True)
class MarkerPathOverlap:
    """Side-by-side marker records that apply to the same protected path."""

    path: str
    protected_decision: ProtectedFileDecision
    local_overrides: tuple[LocalOverride, ...]
    deferred_candidates: tuple[DeferredProtectedCandidate, ...]


@dataclass(frozen=True)
class ManifestMapping:
    """One path mapping row from the template sync manifest."""

    pattern: str
    requires_all: frozenset[str]
    requires_any: frozenset[str]
    notes: str | None = None
    unknown_modules: frozenset[str] = frozenset()

    @property
    def is_concrete(self) -> bool:
        """Return whether this mapping names a concrete path rather than a glob."""
        return not has_wildcard(self.pattern)


@dataclass(frozen=True)
class ManifestCompatibilityAlternative:
    """One host-family alternative in a manifest compatibility group."""

    host: str
    modules: frozenset[str]


@dataclass(frozen=True)
class ManifestCompatibilityGroup:
    """Host-family compatibility metadata for related optional modules."""

    name: str
    description: str
    default_modules: frozenset[str]
    alternatives: tuple[ManifestCompatibilityAlternative, ...]
    mixed_hosts: str


@dataclass(frozen=True)
class PathRelation:
    """The manifest module relation selected for a repository path."""

    patterns: tuple[str, ...]
    requires_all: frozenset[str]
    requires_any: frozenset[str]
    notes: tuple[str, ...] = ()
    unknown_modules: frozenset[str] = frozenset()

    @property
    def is_cross_module(self) -> bool:
        """Return whether the relation spans multiple module concerns."""
        return bool(self.requires_any) or len(self.requires_all | self.requires_any) > 1

    @property
    def description(self) -> str:
        """Return a compact human-readable relation summary."""
        parts: list[str] = []
        if self.requires_all:
            parts.append("requires all: " + ", ".join(sorted(self.requires_all)))
        if self.requires_any:
            parts.append("requires any: " + ", ".join(sorted(self.requires_any)))
        return "; ".join(parts) if parts else "no module relation"

    def is_retained_by(self, included_modules: Collection[str]) -> bool:
        """Return whether ``included_modules`` satisfies this path relation."""
        included_module_set = set(included_modules)
        if not self.requires_all.issubset(included_module_set):
            return False
        if not self.requires_any:
            return True
        return not self.requires_any.isdisjoint(included_module_set)


@dataclass(frozen=True)
class MarkerDecisionData:
    """Parsed marker data used by materialization and reporting helpers."""

    last_reviewed_template_commit: str | None
    included_modules: frozenset[str]
    local_overrides: tuple[LocalOverride, ...]
    local_path_ownership: tuple[LocalPathOwnership, ...]
    deferred_candidates: tuple[DeferredProtectedCandidate, ...]
    protected_decisions: tuple[ProtectedFileDecision, ...]
    protected_guide_contract_waivers: tuple[ProtectedGuideContractWaiver, ...]


@dataclass(frozen=True)
class InlineBlockMarker:
    """One parsed inline-block marker occurrence."""

    kind: str
    name: str
    line_number: int
    line: str


@dataclass(frozen=True)
class InlineBlockSpan:
    """One complete live template-sync inline block in a text file."""

    marker_name: str
    line_number: int
    content_lines: tuple[str, ...]


@dataclass(frozen=True)
class RepositoryFileClassification:
    """Text-versus-byte classification for a safe repository file."""

    relative_path: str
    path: Path
    is_transformable_text: bool
    text: str | None
    bytes_content: bytes

    @property
    def is_byte_only(self) -> bool:
        """Return whether the file is copyable bytes but not transformable text."""
        return not self.is_transformable_text


class _MarkerYamlDumper(yaml.SafeDumper):
    """PyYAML safe dumper that indents block sequences under mapping keys."""

    def increase_indent(self, flow: bool = False, indentless: bool = False) -> None:
        """Force block sequences to use ordinary nested indentation."""
        super().increase_indent(flow=flow, indentless=False)

    def ignore_aliases(self, data: object) -> bool:
        """Keep generated marker YAML free of anchors and aliases."""
        return True


def default_repo_root() -> Path:
    """Return the repository root implied by this script's committed location."""
    return Path(__file__).resolve().parents[2]


def resolve_repo_root(raw_repo_root: str | None) -> Path:
    """Resolve and validate the repository root argument."""
    repo_root = Path(raw_repo_root).expanduser() if raw_repo_root else default_repo_root()
    resolved = repo_root.resolve()
    if not resolved.is_dir():
        raise TemplateSyncMaterializationError(
            f"Repository root does not exist or is not a directory: {repo_root}"
        )
    return resolved


def dump_marker_yaml_object(marker_document: dict[str, Any]) -> object:
    """Return PyYAML output as an object for cross-checker runtime narrowing."""
    return yaml.dump(
        marker_document,
        Dumper=_MarkerYamlDumper,
        sort_keys=False,
        allow_unicode=False,
        default_flow_style=False,
        width=120,
    )


def format_marker_yaml(marker_document: dict[str, Any]) -> str:
    """Return canonical deterministic YAML for a template-sync marker document."""
    dumped_marker = dump_marker_yaml_object(marker_document)
    if not isinstance(dumped_marker, str):
        raise TemplateSyncMaterializationError("Formatted marker YAML must be text.")
    marker_text = dumped_marker.rstrip("\n") + "\n"
    if marker_text.startswith("---"):
        raise TemplateSyncMaterializationError(
            "Formatted marker YAML must not include a document-start marker."
        )
    return marker_text


def validate_marker_yaml_text(
    marker_text: str,
    *,
    repo_root: Path,
    marker_path: str = DEFAULT_MARKER_PATH,
    marker_schema_path: str = DEFAULT_MARKER_SCHEMA_PATH,
) -> dict[str, Any]:
    """Parse, schema-validate, and integrity-check formatted marker YAML."""
    marker_document_path = resolve_safe_repository_target_path(
        repo_root,
        marker_path,
        field_name="marker path",
    )
    marker_relative_path = repository_relative_path(marker_document_path, repo_root)
    try:
        parsed = yaml.safe_load(marker_text)
    except yaml.YAMLError as error:
        raise TemplateSyncMaterializationError(
            f"Invalid YAML in {marker_relative_path}: {error}"
        ) from error
    if not isinstance(parsed, dict):
        raise TemplateSyncMaterializationError(
            f"{marker_relative_path} must contain a YAML mapping."
        )
    marker_document = cast(dict[str, Any], parsed)

    schema_path = resolve_safe_repository_target_path(
        repo_root,
        marker_schema_path,
        field_name="marker schema path",
    )
    marker_schema = load_json_mapping(schema_path, repo_root)
    validate_schema(marker_document, marker_schema, marker_document_path, repo_root)
    parse_marker_decision_data(marker_document, validate_protected_decision_integrity=True)
    return marker_document


def ensure_regular_repository_file_target(target_path: Path, relative_path: str) -> None:
    """Raise when an existing write target is not a regular file.

    This guard inspects target_path only, rejecting an existing symlink,
    directory, or other non-regular file. It does not verify that target_path
    is contained within the repository root or that its ancestors are
    symlink-free. Callers must establish that precondition first by resolving
    the target through resolve_safe_repository_target_path (or by using
    write_repository_relative_file_bytes), which enforce repository containment
    and symlink-ancestor safety.
    """
    if target_path.is_symlink() or (target_path.exists() and not target_path.is_file()):
        raise TemplateSyncMaterializationError(
            f"Cannot reconcile {relative_path}: the target path exists but is not a "
            f"regular file (for example a directory or a symlink). Resolve the "
            f"conflict in the downstream repository, then rerun."
        )


def write_repository_file_bytes(
    target_path: Path,
    relative_path: str,
    bytes_content: bytes,
) -> None:
    """Write bytes to a target after enforcing regular-file safety.

    Like ensure_regular_repository_file_target, this writer only enforces that
    target_path is a regular file; it does not enforce repository containment
    or symlink-ancestor safety. Pass a target_path already resolved by
    resolve_safe_repository_target_path, or call
    write_repository_relative_file_bytes, when those guarantees are required.
    """
    ensure_regular_repository_file_target(target_path, relative_path)
    target_path.parent.mkdir(parents=True, exist_ok=True)
    target_path.write_bytes(bytes_content)


def write_repository_relative_file_bytes(
    repo_root: Path,
    relative_path: str,
    bytes_content: bytes,
    *,
    field_name: str = "path",
) -> Path:
    """Resolve and write a repository-relative file target safely."""
    target_path = resolve_safe_repository_target_path(
        repo_root,
        relative_path,
        field_name=field_name,
    )
    write_repository_file_bytes(target_path, relative_path, bytes_content)
    return target_path


def load_validated_manifest(
    repo_root: Path,
    manifest_path: Path,
    manifest_schema_path: Path,
    *,
    reject_unknown_modules: bool = True,
) -> tuple[frozenset[str], tuple[ManifestMapping, ...]]:
    """Load, schema-validate, and parse a template-sync manifest."""
    manifest = load_yaml_mapping(manifest_path, repo_root)
    manifest_schema = load_json_mapping(manifest_schema_path, repo_root)
    validate_schema(manifest, manifest_schema, manifest_path, repo_root)
    module_names, mappings = parse_manifest_mappings(
        manifest,
        reject_unknown_modules=reject_unknown_modules,
    )
    parse_manifest_compatibility_groups(
        manifest,
        module_names=module_names,
        reject_unknown_modules=reject_unknown_modules,
    )
    return module_names, mappings


def load_validated_marker_decision_data(
    repo_root: Path,
    marker_path: Path,
    marker_schema_path: Path,
    *,
    validate_protected_decision_integrity: bool = False,
) -> MarkerDecisionData:
    """Load, schema-validate, and parse marker-shaped decision data."""
    marker = load_yaml_mapping(marker_path, repo_root)
    marker_schema = load_json_mapping(marker_schema_path, repo_root)
    validate_schema(marker, marker_schema, marker_path, repo_root)
    return parse_marker_decision_data(
        marker,
        validate_protected_decision_integrity=validate_protected_decision_integrity,
    )


def normalize_manifest_pattern(raw_pattern: str) -> str:
    """Validate and normalize a manifest path pattern."""
    if "\\" in raw_pattern:
        raise RepositoryPathError(f"Manifest patterns must use POSIX separators: {raw_pattern}")
    if raw_pattern.startswith("/"):
        raise RepositoryPathError(f"Manifest patterns must be repository-relative: {raw_pattern}")
    parts = raw_pattern.split("/")
    if any(part in ("", ".", "..") for part in parts):
        raise RepositoryPathError(
            f"Manifest patterns must not contain empty or traversal segments: {raw_pattern}"
        )
    return raw_pattern


def is_string_list(value: object) -> TypeGuard[list[str]]:
    """Return whether ``value`` is a list containing only strings."""
    if not isinstance(value, list):
        return False
    return all(isinstance(item, str) for item in cast(list[object], value))


def relation_modules(mapping: dict[str, Any], relation_key: str) -> frozenset[str]:
    """Return module names for one manifest relation key."""
    modules = mapping.get(relation_key, [])
    if not isinstance(modules, list):
        pattern = mapping.get("pattern", "<unknown>")
        raise TemplateSyncMaterializationError(f"{pattern} {relation_key} must be a list.")
    module_items = cast(list[object], modules)
    if not all(isinstance(module, str) for module in module_items):
        pattern = mapping.get("pattern", "<unknown>")
        raise TemplateSyncMaterializationError(f"{pattern} {relation_key} values must be strings.")
    return frozenset(cast(list[str], module_items))


def parse_manifest_module_names(manifest: dict[str, Any]) -> frozenset[str]:
    """Extract declared module names from a manifest document."""
    template_manifest = manifest.get("template_manifest")
    if not isinstance(template_manifest, dict):
        raise TemplateSyncMaterializationError("Manifest must contain template_manifest mapping.")
    template_manifest = cast(dict[str, Any], template_manifest)

    raw_modules = template_manifest.get("modules")
    if not isinstance(raw_modules, list):
        raise TemplateSyncMaterializationError("template_manifest.modules must be a list.")
    module_names: set[str] = set()
    for module in cast(list[object], raw_modules):
        if not isinstance(module, dict):
            raise TemplateSyncMaterializationError(
                "Every manifest module must define a string name."
            )
        module_record = cast(dict[str, Any], module)
        module_name = module_record.get("name")
        if not isinstance(module_name, str):
            raise TemplateSyncMaterializationError(
                "Every manifest module must define a string name."
            )
        module_names.add(module_name)
    return frozenset(module_names)


def parse_manifest_mappings(
    manifest: dict[str, Any],
    *,
    reject_unknown_modules: bool = True,
) -> tuple[frozenset[str], tuple[ManifestMapping, ...]]:
    """Extract module definitions and path mappings from a validated manifest."""
    template_manifest = manifest.get("template_manifest")
    if not isinstance(template_manifest, dict):
        raise TemplateSyncMaterializationError("Manifest must contain template_manifest mapping.")
    template_manifest = cast(dict[str, Any], template_manifest)

    module_names = set(parse_manifest_module_names(manifest))

    raw_path_mappings = template_manifest.get("path_mappings")
    if not isinstance(raw_path_mappings, list):
        raise TemplateSyncMaterializationError("template_manifest.path_mappings must be a list.")

    mappings: list[ManifestMapping] = []
    for raw_mapping in cast(list[object], raw_path_mappings):
        if not isinstance(raw_mapping, dict):
            raise TemplateSyncMaterializationError("Every path mapping must be a mapping.")
        mapping = cast(dict[str, Any], raw_mapping)
        raw_pattern = mapping.get("pattern")
        if not isinstance(raw_pattern, str):
            raise TemplateSyncMaterializationError(
                "Every path mapping must define a string pattern."
            )
        notes = mapping.get("notes")
        if notes is not None and not isinstance(notes, str):
            raise TemplateSyncMaterializationError(f"{raw_pattern} notes must be a string.")

        requires_all = relation_modules(mapping, "requires_all")
        requires_any = relation_modules(mapping, "requires_any")
        if not requires_all and not requires_any:
            raise TemplateSyncMaterializationError(
                f"{raw_pattern} must reference at least one module."
            )
        unknown_modules = (requires_all | requires_any) - module_names
        if reject_unknown_modules and unknown_modules:
            raise TemplateSyncMaterializationError(
                f"{raw_pattern} references unknown manifest module(s): "
                + ", ".join(sorted(unknown_modules))
            )
        mappings.append(
            ManifestMapping(
                pattern=normalize_manifest_pattern(raw_pattern),
                requires_all=requires_all,
                requires_any=requires_any,
                notes=notes,
                unknown_modules=frozenset(unknown_modules),
            )
        )

    return frozenset(module_names), tuple(mappings)


def compatibility_modules(raw_modules: object, field_name: str) -> frozenset[str]:
    """Return a non-empty compatibility module set from a manifest field."""
    if not isinstance(raw_modules, list):
        raise TemplateSyncMaterializationError(f"{field_name} must be a list.")
    if not raw_modules:
        raise TemplateSyncMaterializationError(f"{field_name} must not be empty.")
    module_items = cast(list[object], raw_modules)
    if not all(isinstance(module, str) for module in module_items):
        raise TemplateSyncMaterializationError(f"{field_name} values must be strings.")
    return frozenset(cast(list[str], module_items))


def parse_manifest_compatibility_groups(
    manifest: dict[str, Any],
    *,
    module_names: Collection[str] | None = None,
    reject_unknown_modules: bool = True,
) -> tuple[ManifestCompatibilityGroup, ...]:
    """Extract host-family compatibility groups from a manifest document."""
    template_manifest = manifest.get("template_manifest")
    if not isinstance(template_manifest, dict):
        raise TemplateSyncMaterializationError("Manifest must contain template_manifest mapping.")
    template_manifest = cast(dict[str, Any], template_manifest)

    raw_groups = template_manifest.get("compatibility_groups")
    if raw_groups is None:
        return ()
    if not isinstance(raw_groups, list):
        raise TemplateSyncMaterializationError(
            "template_manifest.compatibility_groups must be a list."
        )

    known_modules = set(module_names or parse_manifest_module_names(manifest))
    group_names: set[str] = set()
    groups: list[ManifestCompatibilityGroup] = []
    for index, raw_group in enumerate(cast(list[object], raw_groups), 1):
        group_field = f"template_manifest.compatibility_groups[{index}]"
        if not isinstance(raw_group, dict):
            raise TemplateSyncMaterializationError(f"{group_field} must be a mapping.")
        group_record = cast(dict[str, Any], raw_group)

        name = group_record.get("name")
        description = group_record.get("description")
        mixed_hosts = group_record.get("mixed_hosts")
        raw_alternatives = group_record.get("alternatives")
        if not isinstance(name, str):
            raise TemplateSyncMaterializationError(f"{group_field}.name must be a string.")
        if not isinstance(description, str):
            raise TemplateSyncMaterializationError(f"{group_field}.description must be a string.")
        if not isinstance(mixed_hosts, str) or mixed_hosts not in {"allowed", "unsupported"}:
            raise TemplateSyncMaterializationError(
                f"{group_field}.mixed_hosts must be allowed or unsupported."
            )
        if name in group_names:
            raise TemplateSyncMaterializationError(f"Duplicate compatibility group: {name}")
        group_names.add(name)

        default_modules = compatibility_modules(
            group_record.get("default_modules"),
            f"{group_field}.default_modules",
        )
        if not isinstance(raw_alternatives, list) or not raw_alternatives:
            raise TemplateSyncMaterializationError(
                f"{group_field}.alternatives must be a non-empty list."
            )

        alternatives: list[ManifestCompatibilityAlternative] = []
        alternative_hosts: set[str] = set()
        alternative_modules: set[str] = set()
        module_hosts: dict[str, str] = {}
        for alternative_index, raw_alternative in enumerate(
            cast(list[object], raw_alternatives),
            1,
        ):
            alternative_field = f"{group_field}.alternatives[{alternative_index}]"
            if not isinstance(raw_alternative, dict):
                raise TemplateSyncMaterializationError(f"{alternative_field} must be a mapping.")
            alternative_record = cast(dict[str, Any], raw_alternative)
            host = alternative_record.get("host")
            if not isinstance(host, str):
                raise TemplateSyncMaterializationError(
                    f"{alternative_field}.host must be a string."
                )
            if host in alternative_hosts:
                raise TemplateSyncMaterializationError(
                    f"{group_field} defines duplicate host alternative: {host}"
                )
            alternative_hosts.add(host)

            modules = compatibility_modules(
                alternative_record.get("modules"),
                f"{alternative_field}.modules",
            )
            for module in modules:
                previous_host = module_hosts.get(module)
                if previous_host is not None:
                    raise TemplateSyncMaterializationError(
                        f"{group_field} maps module {module} to multiple hosts: "
                        f"{previous_host}, {host}"
                    )
                module_hosts[module] = host
            alternative_modules.update(modules)
            alternatives.append(ManifestCompatibilityAlternative(host=host, modules=modules))

        missing_default_modules = default_modules - alternative_modules
        if missing_default_modules:
            raise TemplateSyncMaterializationError(
                f"{group_field}.default_modules must be listed in an alternative: "
                + ", ".join(sorted(missing_default_modules))
            )

        unknown_modules = (default_modules | alternative_modules) - known_modules
        if reject_unknown_modules and unknown_modules:
            raise TemplateSyncMaterializationError(
                f"{group_field} references unknown manifest module(s): "
                + ", ".join(sorted(unknown_modules))
            )

        groups.append(
            ManifestCompatibilityGroup(
                name=name,
                description=description,
                default_modules=default_modules,
                alternatives=tuple(alternatives),
                mixed_hosts=mixed_hosts,
            )
        )

    return tuple(groups)


def validate_module_compatibility(
    included_modules: Collection[str],
    compatibility_groups: tuple[ManifestCompatibilityGroup, ...],
) -> tuple[str, ...]:
    """Return actionable errors for unsupported host-family module mixes."""
    included_module_set = set(included_modules)
    errors: list[str] = []
    if "instruction-enforcement" in included_module_set:
        if "agent-instructions" not in included_module_set:
            errors.append("instruction-enforcement requires agent-instructions.")
        if not included_module_set.intersection({"baseline", "github-actions", "azure-pipelines"}):
            errors.append("instruction-enforcement requires baseline or a selected host CI route.")
    if (
        any(
            module.startswith("agent-") and module != "agent-instructions"
            for module in included_module_set
        )
        and "agent-instructions" not in included_module_set
    ):
        errors.append("Agent selections require agent-instructions.")
    for group in compatibility_groups:
        if group.mixed_hosts == "allowed":
            continue
        selected_hosts: list[tuple[str, frozenset[str]]] = []
        for alternative in group.alternatives:
            selected_modules = alternative.modules & included_module_set
            if selected_modules:
                selected_hosts.append((alternative.host, selected_modules))
        if len(selected_hosts) <= 1:
            continue

        selected_summary = "; ".join(
            f"{host} ({', '.join(sorted(modules))})" for host, modules in selected_hosts
        )
        errors.append(
            f"Unsupported mixed-host selection in compatibility group {group.name}: "
            f"{selected_summary}. Select one host family or update "
            "template_manifest.compatibility_groups if this mixed-host shape is intentional."
        )
    return tuple(errors)


def optional_marker_string(
    raw_record: dict[str, Any],
    field_name: str,
    record_name: str,
) -> str | None:
    """Return an optional string field from a marker record."""
    value = raw_record.get(field_name)
    if value is None:
        return None
    if not isinstance(value, str):
        raise TemplateSyncMaterializationError(
            f"{record_name}.{field_name} must be a string when present."
        )
    return value


def marker_record_list(
    template_sync: dict[str, Any],
    field_name: str,
    record_name: str,
) -> list[dict[str, Any]]:
    """Return a marker field that must be absent or a list of mappings."""
    raw_records = template_sync.get(field_name, [])
    if not isinstance(raw_records, list):
        raise TemplateSyncMaterializationError(f"{record_name} must be a list.")
    records: list[dict[str, Any]] = []
    for raw_record in cast(list[object], raw_records):
        if not isinstance(raw_record, dict):
            raise TemplateSyncMaterializationError(
                f"Each {record_name.removeprefix('template_sync.')} record must be a mapping."
            )
        records.append(cast(dict[str, Any], raw_record))
    return records


def parse_protected_file_decisions(
    template_sync: dict[str, Any],
) -> tuple[ProtectedFileDecision, ...]:
    """Extract normalized protected-file decision records from ``template_sync``."""
    protected_decisions: list[ProtectedFileDecision] = []
    for raw_decision in marker_record_list(
        template_sync,
        "protected_file_decisions",
        "template_sync.protected_file_decisions",
    ):
        raw_path = raw_decision.get("path")
        decision = raw_decision.get("decision")
        if not isinstance(raw_path, str) or not isinstance(decision, str):
            raise TemplateSyncMaterializationError(
                "Each protected file decision must define string path and decision."
            )
        normalized_path, is_directory = normalize_repository_path(
            raw_path,
            "template_sync.protected_file_decisions[].path",
        )
        if is_directory:
            raise TemplateSyncMaterializationError(
                "template_sync.protected_file_decisions[].path must reference a file, "
                f"not a directory: {raw_path}"
            )
        protected_decisions.append(
            ProtectedFileDecision(
                path=normalized_path,
                decision=decision,
                adoption_mode=optional_marker_string(
                    raw_decision,
                    "adoption_mode",
                    "template_sync.protected_file_decisions[]",
                ),
                authorization_basis=optional_marker_string(
                    raw_decision,
                    "authorization_basis",
                    "template_sync.protected_file_decisions[]",
                ),
                authorized_scope=optional_marker_string(
                    raw_decision,
                    "authorized_scope",
                    "template_sync.protected_file_decisions[]",
                ),
                tailored_authorization_basis=optional_marker_string(
                    raw_decision,
                    "tailored_authorization_basis",
                    "template_sync.protected_file_decisions[]",
                ),
                reason=optional_marker_string(
                    raw_decision,
                    "reason",
                    "template_sync.protected_file_decisions[]",
                ),
            )
        )
    return tuple(protected_decisions)


def parse_protected_guide_contract_waivers(
    template_sync: dict[str, Any],
) -> tuple[ProtectedGuideContractWaiver, ...]:
    """Extract normalized protected-guide contract waiver records from ``template_sync``."""
    waivers: list[ProtectedGuideContractWaiver] = []
    seen_keys: set[tuple[str, str, str | None, str | None]] = set()
    duplicate_keys: set[tuple[str, str, str | None, str | None]] = set()
    for raw_waiver in marker_record_list(
        template_sync,
        "protected_guide_contract_waivers",
        "template_sync.protected_guide_contract_waivers",
    ):
        raw_path = raw_waiver.get("path")
        contract_key = raw_waiver.get("contract_key")
        reason = raw_waiver.get("reason")
        authorization_basis = raw_waiver.get("authorization_basis")
        if (
            not isinstance(raw_path, str)
            or not isinstance(contract_key, str)
            or not isinstance(reason, str)
            or not isinstance(authorization_basis, str)
        ):
            raise TemplateSyncMaterializationError(
                "Each protected guide contract waiver must define string path, "
                "contract_key, reason, and authorization_basis."
            )
        path, is_directory = normalize_repository_path(
            raw_path,
            "template_sync.protected_guide_contract_waivers[].path",
        )
        if is_directory:
            raise TemplateSyncMaterializationError(
                "template_sync.protected_guide_contract_waivers[].path must reference "
                f"a file, not a directory: {raw_path}"
            )

        target_path = optional_marker_string(
            raw_waiver,
            "target_path",
            "template_sync.protected_guide_contract_waivers[]",
        )
        if target_path is not None:
            target_path, target_is_directory = normalize_repository_path(
                target_path,
                "template_sync.protected_guide_contract_waivers[].target_path",
            )
            if target_is_directory:
                raise TemplateSyncMaterializationError(
                    "template_sync.protected_guide_contract_waivers[].target_path must "
                    f"reference a file, not a directory: {target_path}"
                )
        target_module = optional_marker_string(
            raw_waiver,
            "target_module",
            "template_sync.protected_guide_contract_waivers[]",
        )
        linked_local_override_path = optional_marker_string(
            raw_waiver,
            "linked_local_override_path",
            "template_sync.protected_guide_contract_waivers[]",
        )
        if linked_local_override_path is not None:
            linked_local_override_path, _linked_is_directory = normalize_repository_path(
                linked_local_override_path,
                "template_sync.protected_guide_contract_waivers[].linked_local_override_path",
            )

        waiver_key = (path, contract_key, target_path, target_module)
        if waiver_key in seen_keys:
            duplicate_keys.add(waiver_key)
        seen_keys.add(waiver_key)
        waivers.append(
            ProtectedGuideContractWaiver(
                path=path,
                contract_key=contract_key,
                target_path=target_path,
                target_module=target_module,
                linked_local_override_path=linked_local_override_path,
                reason=reason,
                authorization_basis=authorization_basis,
            )
        )

    if duplicate_keys:
        formatted_keys = ", ".join(
            (
                f"({path}, {contract_key}, "
                f"target_path={target_path or '<none>'}, "
                f"target_module={target_module or '<none>'})"
            )
            for path, contract_key, target_path, target_module in sorted(duplicate_keys)
        )
        raise TemplateSyncMaterializationError(
            "Duplicate protected_guide_contract_waivers key(s): " + formatted_keys
        )
    return tuple(waivers)


def parse_local_path_ownership(
    template_sync: dict[str, Any],
) -> tuple[LocalPathOwnership, ...]:
    """Extract normalized downstream-owned path records from ``template_sync``."""
    local_path_ownership: list[LocalPathOwnership] = []
    duplicate_paths: set[str] = set()
    seen_paths: set[str] = set()
    for raw_record in marker_record_list(
        template_sync,
        "local_path_ownership",
        "template_sync.local_path_ownership",
    ):
        raw_path = raw_record.get("path")
        reason = raw_record.get("reason")
        if not isinstance(raw_path, str) or not isinstance(reason, str):
            raise TemplateSyncMaterializationError(
                "Each local path ownership record must define string path and reason."
            )
        normalized_path, is_directory = normalize_repository_path(
            raw_path,
            "template_sync.local_path_ownership[].path",
        )
        if normalized_path in seen_paths:
            duplicate_paths.add(normalized_path)
        seen_paths.add(normalized_path)
        local_path_ownership.append(
            LocalPathOwnership(
                path=normalized_path,
                reason=reason,
                overlap_exception_reason=optional_marker_string(
                    raw_record,
                    "overlap_exception_reason",
                    "template_sync.local_path_ownership[]",
                ),
                is_directory=is_directory,
            )
        )
    if duplicate_paths:
        raise TemplateSyncMaterializationError(
            "Duplicate local_path_ownership path(s): " + ", ".join(sorted(duplicate_paths))
        )
    return tuple(local_path_ownership)


def parse_marker_decision_data(
    marker: dict[str, Any],
    *,
    validate_protected_decision_integrity: bool = False,
) -> MarkerDecisionData:
    """Extract marker fields used by sync validation and materialization planning."""
    template_sync = marker.get("template_sync")
    if not isinstance(template_sync, dict):
        raise TemplateSyncMaterializationError("Marker must contain template_sync mapping.")
    template_sync = cast(dict[str, Any], template_sync)

    last_reviewed = template_sync.get("last_reviewed_template_commit")
    if last_reviewed is not None and not isinstance(last_reviewed, str):
        raise TemplateSyncMaterializationError(
            "template_sync.last_reviewed_template_commit must be a string when present."
        )

    raw_included_modules = template_sync.get("included_modules")
    if not is_string_list(raw_included_modules):
        raise TemplateSyncMaterializationError(
            "template_sync.included_modules must be a list of strings."
        )
    included_modules = frozenset(raw_included_modules)

    local_overrides: list[LocalOverride] = []
    for raw_override in marker_record_list(
        template_sync,
        "local_overrides",
        "template_sync.local_overrides",
    ):
        raw_path = raw_override.get("path")
        default_decision = raw_override.get("default_decision")
        reason = raw_override.get("reason")
        if (
            not isinstance(raw_path, str)
            or not isinstance(default_decision, str)
            or not isinstance(reason, str)
        ):
            raise TemplateSyncMaterializationError(
                "Each local override must define string path, default_decision, and reason."
            )
        normalized_path, is_directory = normalize_repository_path(
            raw_path,
            "template_sync.local_overrides[].path",
        )
        local_overrides.append(
            LocalOverride(
                path=normalized_path,
                default_decision=default_decision,
                reason=reason,
                is_directory=is_directory,
            )
        )

    deferred_candidates: list[DeferredProtectedCandidate] = []
    for raw_candidate in marker_record_list(
        template_sync,
        "deferred_protected_candidates",
        "template_sync.deferred_protected_candidates",
    ):
        raw_path = raw_candidate.get("path")
        source_commit = raw_candidate.get("source_commit")
        reason = raw_candidate.get("reason")
        if (
            not isinstance(raw_path, str)
            or not isinstance(source_commit, str)
            or not isinstance(reason, str)
        ):
            raise TemplateSyncMaterializationError(
                "Each deferred protected candidate must define string path, source_commit, "
                "and reason."
            )
        normalized_path, is_directory = normalize_repository_path(
            raw_path,
            "template_sync.deferred_protected_candidates[].path",
        )
        if is_directory:
            raise TemplateSyncMaterializationError(
                "template_sync.deferred_protected_candidates[].path must reference a file, "
                f"not a directory: {raw_path}"
            )
        deferred_candidates.append(
            DeferredProtectedCandidate(
                path=normalized_path,
                source_commit=source_commit,
                reason=reason,
            )
        )

    protected_decisions = parse_protected_file_decisions(template_sync)
    protected_guide_contract_waivers = parse_protected_guide_contract_waivers(template_sync)
    local_path_ownership = parse_local_path_ownership(template_sync)
    local_override_tuple = tuple(local_overrides)
    deferred_candidate_tuple = tuple(deferred_candidates)
    if validate_protected_decision_integrity:
        validate_protected_file_decisions(
            protected_decisions,
            local_override_tuple,
            deferred_candidate_tuple,
        )

    return MarkerDecisionData(
        last_reviewed_template_commit=last_reviewed,
        included_modules=included_modules,
        local_overrides=local_override_tuple,
        local_path_ownership=local_path_ownership,
        deferred_candidates=deferred_candidate_tuple,
        protected_decisions=protected_decisions,
        protected_guide_contract_waivers=protected_guide_contract_waivers,
    )


def protected_decision_summary(protected_decision: ProtectedFileDecision) -> str:
    """Return a compact protected-file decision summary."""
    parts = [f"decision={protected_decision.decision}"]
    if protected_decision.adoption_mode is not None:
        parts.append(f"adoption_mode={protected_decision.adoption_mode}")
    if protected_decision.authorization_basis is not None:
        parts.append(f"authorization_basis={protected_decision.authorization_basis}")
    if protected_decision.authorized_scope is not None:
        parts.append(f"authorized_scope={protected_decision.authorized_scope}")
    if protected_decision.tailored_authorization_basis is not None:
        parts.append(
            f"tailored_authorization_basis={protected_decision.tailored_authorization_basis}"
        )
    if protected_decision.reason is not None:
        parts.append(f"reason={protected_decision.reason}")
    return "; ".join(parts)


def local_override_summary(local_override: LocalOverride) -> str:
    """Return a compact local override summary."""
    suffix = "/" if local_override.is_directory else ""
    return (
        f"path={local_override.path}{suffix}; "
        f"default_decision={local_override.default_decision}; "
        f"reason={local_override.reason}"
    )


def local_path_ownership_summary(local_path_ownership: LocalPathOwnership) -> str:
    """Return a compact local path ownership summary."""
    parts = [
        f"path={local_path_ownership.display_path}",
        f"reason={local_path_ownership.reason}",
    ]
    if local_path_ownership.overlap_exception_reason is not None:
        parts.append(f"overlap_exception_reason={local_path_ownership.overlap_exception_reason}")
    return "; ".join(parts)


def deferred_candidate_summary(candidate: DeferredProtectedCandidate) -> str:
    """Return a compact deferred protected candidate summary."""
    return f"source_commit={candidate.source_commit}; reason={candidate.reason}"


def marker_path_overlaps(
    protected_decisions: tuple[ProtectedFileDecision, ...],
    local_overrides: tuple[LocalOverride, ...],
    deferred_candidates: tuple[DeferredProtectedCandidate, ...],
) -> tuple[MarkerPathOverlap, ...]:
    """Return marker records that overlap with protected-file decisions."""
    overlaps: list[MarkerPathOverlap] = []
    for protected_decision in protected_decisions:
        matching_local_overrides = tuple(
            local_override
            for local_override in local_overrides
            if local_override.matches(protected_decision.path)
        )
        matching_deferred_candidates = tuple(
            candidate
            for candidate in deferred_candidates
            if candidate.path == protected_decision.path
        )
        if matching_local_overrides or matching_deferred_candidates:
            overlaps.append(
                MarkerPathOverlap(
                    path=protected_decision.path,
                    protected_decision=protected_decision,
                    local_overrides=matching_local_overrides,
                    deferred_candidates=matching_deferred_candidates,
                )
            )
    return tuple(overlaps)


def format_overlap_block(overlap: MarkerPathOverlap) -> str:
    """Return a side-by-side marker overlap block for diagnostics."""
    lines = [
        f"  - {overlap.path}",
        f"    protected_file_decisions: {protected_decision_summary(overlap.protected_decision)}",
    ]
    for local_override in overlap.local_overrides:
        lines.append(f"    local_overrides: {local_override_summary(local_override)}")
    for candidate in overlap.deferred_candidates:
        lines.append(f"    deferred_protected_candidates: {deferred_candidate_summary(candidate)}")
    return "\n".join(lines)


def validate_protected_file_decisions(
    protected_decisions: tuple[ProtectedFileDecision, ...],
    local_overrides: tuple[LocalOverride, ...],
    deferred_candidates: tuple[DeferredProtectedCandidate, ...],
    *,
    strict_remove_local_phrasing: bool = False,
    remove_local_tokens: tuple[str, ...] = DEFAULT_REMOVE_LOCAL_AUTHORIZATION_TOKENS,
) -> tuple[MarkerPathOverlap, ...]:
    """Validate protected-file decision integrity beyond JSON Schema checks."""
    seen_paths: set[str] = set()
    duplicate_paths: set[str] = set()
    for protected_decision in protected_decisions:
        if protected_decision.path in seen_paths:
            duplicate_paths.add(protected_decision.path)
        seen_paths.add(protected_decision.path)
    if duplicate_paths:
        raise TemplateSyncMaterializationError(
            "Duplicate protected_file_decisions path(s): " + ", ".join(sorted(duplicate_paths))
        )

    overlaps = marker_path_overlaps(
        protected_decisions,
        local_overrides,
        deferred_candidates,
    )
    contradictions: list[str] = []
    for overlap in overlaps:
        mismatched_local_overrides = tuple(
            local_override
            for local_override in overlap.local_overrides
            if local_override.path == overlap.path
            and local_override.default_decision != overlap.protected_decision.decision
        )
        if mismatched_local_overrides:
            contradictions.append(
                format_overlap_block(
                    MarkerPathOverlap(
                        path=overlap.path,
                        protected_decision=overlap.protected_decision,
                        local_overrides=mismatched_local_overrides,
                        deferred_candidates=(),
                    )
                )
                + "\n    conflict: protected decision and local override decisions differ."
            )
        if overlap.deferred_candidates:
            contradictions.append(
                format_overlap_block(
                    MarkerPathOverlap(
                        path=overlap.path,
                        protected_decision=overlap.protected_decision,
                        local_overrides=(),
                        deferred_candidates=overlap.deferred_candidates,
                    )
                )
                + "\n    conflict: protected decision asserts current authorization while "
                "the deferred candidate is awaiting authorization."
            )
    if contradictions:
        raise TemplateSyncMaterializationError(
            "Contradictory protected-file marker entries:\n" + "\n".join(contradictions)
        )

    if strict_remove_local_phrasing:
        normalized_tokens = tuple(token.lower() for token in remove_local_tokens if token)
        for protected_decision in protected_decisions:
            if protected_decision.decision != REMOVAL_DECISION:
                continue
            authorization_basis = protected_decision.authorization_basis or ""
            if not any(token in authorization_basis.lower() for token in normalized_tokens):
                raise TemplateSyncMaterializationError(
                    f"{protected_decision.path} REMOVE-LOCAL authorization_basis must "
                    "contain at least one configured removal token: " + ", ".join(normalized_tokens)
                )

    return overlaps


def has_wildcard(pattern: str) -> bool:
    """Return whether ``pattern`` contains shell-style wildcard syntax."""
    return any(wildcard in pattern for wildcard in "*?[")


def pattern_specificity(pattern: str) -> tuple[int, int, int]:
    """Return a sortable specificity rank for a manifest path pattern."""
    is_exact = not has_wildcard(pattern)
    literal_length = sum(1 for character in pattern if character not in "*?[]")
    return (int(is_exact), literal_length, pattern.count("/"))


def selected_relation_for_path(
    relative_path: str,
    mappings: tuple[ManifestMapping, ...],
) -> PathRelation | None:
    """Return the best manifest relation for ``relative_path``."""
    matches: list[tuple[tuple[int, int, int], ManifestMapping]] = []
    for mapping in mappings:
        if fnmatch.fnmatchcase(relative_path, mapping.pattern):
            matches.append((pattern_specificity(mapping.pattern), mapping))

    if not matches:
        return None

    best_specificity = max(specificity for specificity, _mapping in matches)
    selected = [mapping for specificity, mapping in matches if specificity == best_specificity]
    patterns: list[str] = []
    requires_all: set[str] = set()
    requires_any: set[str] = set()
    notes: list[str] = []
    unknown_modules: set[str] = set()
    for mapping in selected:
        patterns.append(mapping.pattern)
        requires_all.update(mapping.requires_all)
        requires_any.update(mapping.requires_any)
        unknown_modules.update(mapping.unknown_modules)
        if mapping.notes:
            notes.append(mapping.notes)

    return PathRelation(
        patterns=tuple(patterns),
        requires_all=frozenset(requires_all),
        requires_any=frozenset(requires_any),
        notes=tuple(notes),
        unknown_modules=frozenset(unknown_modules),
    )


def is_locally_overridden(
    relative_path: str,
    local_overrides: tuple[LocalOverride, ...],
) -> bool:
    """Return whether any marker local override covers ``relative_path``."""
    return any(local_override.matches(relative_path) for local_override in local_overrides)


def local_path_ownership_specificity(
    local_path_ownership: LocalPathOwnership,
) -> tuple[int, int]:
    """Return a deterministic specificity rank for local ownership matching."""
    return (
        local_path_ownership.path.count("/") + 1,
        int(not local_path_ownership.is_directory),
    )


def selected_local_path_ownership_for_path(
    relative_path: str,
    local_path_ownership: tuple[LocalPathOwnership, ...],
) -> LocalPathOwnership | None:
    """Return the most-specific local ownership record for ``relative_path``."""
    matches = [record for record in local_path_ownership if record.matches(relative_path)]
    if not matches:
        return None
    return max(
        matches,
        key=lambda record: (
            local_path_ownership_specificity(record),
            record.display_path,
        ),
    )


def is_locally_owned_path(
    relative_path: str,
    local_path_ownership: tuple[LocalPathOwnership, ...],
) -> bool:
    """Return whether any marker local ownership record covers ``relative_path``."""
    return selected_local_path_ownership_for_path(relative_path, local_path_ownership) is not None


def is_template_managed_path(
    relative_path: str,
    mappings: tuple[ManifestMapping, ...],
) -> bool:
    """Return whether a repository-relative path maps to the template manifest."""
    return selected_relation_for_path(relative_path, mappings) is not None


def manifest_covers_directory(
    directory: str,
    mappings: tuple[ManifestMapping, ...],
) -> bool:
    """Return whether any manifest pattern targets a path under ``directory``.

    A directory replaced by a symlink is recorded by Git as a single entry, so a
    glob mapping such as ``templates/python/**`` never matches the directory path
    itself (``templates/python``). ``selected_relation_for_path`` therefore returns
    ``None`` for such a directory even though the manifest manages its contents.
    This predicate recognizes the directory as template-managed when at least one
    manifest pattern falls under its prefix, so directory symlinks over managed
    trees are treated as managed rather than as unmanaged local paths.

    The trailing-slash prefix comparison avoids sibling false positives (for
    example, ``templates_other/**`` does not count as covering ``templates``).
    """
    normalized_directory = directory.rstrip("/")
    if not normalized_directory:
        return False
    prefix = f"{normalized_directory}/"
    return any(mapping.pattern.startswith(prefix) for mapping in mappings)


def directory_prefix_relation(
    directory: str,
    mappings: tuple[ManifestMapping, ...],
) -> PathRelation | None:
    """Return a relation aggregating manifest mappings under ``directory``.

    A directory path is never matched by a glob such as ``schemas/**`` because
    ``fnmatch`` does not match the directory path itself. Callers that need the
    module requirements "proximate" to a directory record therefore union every
    manifest mapping whose pattern falls under the directory prefix. Returns
    ``None`` when nothing falls under the directory. The trailing-slash prefix
    comparison avoids sibling false positives (``schemas_local/**`` does not count
    as covering ``schemas``).
    """
    normalized_directory = directory.rstrip("/")
    if not normalized_directory:
        return None
    prefix = f"{normalized_directory}/"
    proximate = [mapping for mapping in mappings if mapping.pattern.startswith(prefix)]
    if not proximate:
        return None

    patterns: list[str] = []
    requires_all: set[str] = set()
    requires_any: set[str] = set()
    notes: list[str] = []
    unknown_modules: set[str] = set()
    for mapping in proximate:
        patterns.append(mapping.pattern)
        requires_all.update(mapping.requires_all)
        requires_any.update(mapping.requires_any)
        unknown_modules.update(mapping.unknown_modules)
        if mapping.notes:
            notes.append(mapping.notes)

    return PathRelation(
        patterns=tuple(patterns),
        requires_all=frozenset(requires_all),
        requires_any=frozenset(requires_any),
        notes=tuple(notes),
        unknown_modules=frozenset(unknown_modules),
    )


def is_retained_template_path(
    relative_path: str,
    mappings: tuple[ManifestMapping, ...],
    included_modules: Collection[str],
) -> bool:
    """Return whether a path is template-managed and retained by selected modules."""
    relation = selected_relation_for_path(relative_path, mappings)
    return relation is not None and relation.is_retained_by(included_modules)


def is_excluded_template_path(
    relative_path: str,
    mappings: tuple[ManifestMapping, ...],
    included_modules: Collection[str],
) -> bool:
    """Return whether a path is template-managed but excluded by selected modules."""
    relation = selected_relation_for_path(relative_path, mappings)
    return relation is not None and not relation.is_retained_by(included_modules)


def git_visible_paths(repo_root: Path) -> tuple[str, ...]:
    """Return tracked and untracked non-ignored file paths according to Git."""
    result = subprocess.run(
        ["git", "ls-files", "--cached", "--others", "--exclude-standard"],
        cwd=repo_root,
        check=False,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        message = result.stderr.strip() or result.stdout.strip() or "git ls-files failed"
        raise TemplateSyncMaterializationError(f"Unable to inspect Git-visible files: {message}")
    return tuple(sorted({path for path in result.stdout.splitlines() if path}))


def git_present_paths(repo_root: Path) -> tuple[str, ...]:
    """Return Git-visible paths that are present in the working tree."""
    present_paths: list[str] = []
    for relative_path in git_visible_paths(repo_root):
        path = repo_root / relative_path
        if path.exists() or path.is_symlink():
            present_paths.append(relative_path)
    return tuple(present_paths)


def unresolved_concrete_manifest_patterns(
    mappings: tuple[ManifestMapping, ...],
    present_paths: Iterable[str],
    allowlist_paths: Collection[str] = (),
    included_modules: Collection[str] | None = None,
    local_overrides: tuple[LocalOverride, ...] = (),
) -> tuple[tuple[str, PathRelation], ...]:
    """Return concrete manifest mappings missing from the supplied present paths."""
    present_path_set = set(present_paths)
    missing_patterns: list[tuple[str, PathRelation]] = []

    for pattern in sorted({mapping.pattern for mapping in mappings if mapping.is_concrete}):
        if pattern in allowlist_paths or is_locally_overridden(pattern, local_overrides):
            continue

        relation = selected_relation_for_path(pattern, mappings)
        if relation is None:
            continue
        if included_modules is not None and not relation.is_retained_by(included_modules):
            continue
        if pattern not in present_path_set:
            missing_patterns.append((pattern, relation))

    return tuple(missing_patterns)


def iter_safe_repository_files(
    repo_root: Path,
    *,
    skipped_dirs: Collection[str] = SKIPPED_DISCOVERY_DIRS,
) -> tuple[tuple[str, ...], tuple[str, ...]]:
    """Return safely discovered regular files and skipped symlink paths."""
    discovered: list[str] = []
    skipped_symlinks: list[str] = []
    root = repo_root.resolve()

    for current_root, dir_names, file_names in os.walk(root, topdown=True, followlinks=False):
        current_path = Path(current_root)
        retained_dir_names: list[str] = []
        for dir_name in dir_names:
            candidate = current_path / dir_name
            if dir_name in skipped_dirs:
                continue
            if candidate.is_symlink():
                skipped_symlinks.append(f"{repository_relative_path(candidate, root)}/")
                continue
            retained_dir_names.append(dir_name)
        dir_names[:] = retained_dir_names

        for file_name in file_names:
            # A linked worktree stores ``.git`` as a gitlink *file*, which the
            # ``skipped_dirs`` directory pruning above cannot exclude. Skip only
            # that metadata file so legitimate regular files that merely share a
            # name with a skipped directory (such as ``build`` or ``dist``) are
            # still discovered.
            if file_name == ".git":
                continue
            file_path = current_path / file_name
            relative_path = repository_relative_path(file_path, root)
            if file_path.is_symlink():
                skipped_symlinks.append(relative_path)
                continue
            try:
                file_path.resolve().relative_to(root)
            except (OSError, ValueError):
                skipped_symlinks.append(relative_path)
                continue
            discovered.append(relative_path)

    return tuple(sorted(discovered)), tuple(sorted(skipped_symlinks))


def is_protected_instruction_path(relative_path: str) -> bool:
    """Return whether ``relative_path`` is a protected instruction/governance file."""
    if relative_path in PROTECTED_EXACT_PATHS:
        return True
    return any(fnmatch.fnmatchcase(relative_path, pattern) for pattern in PROTECTED_GLOB_PATTERNS)


def is_protected_prose_path(relative_path: str) -> bool:
    """Return whether a protected path holds human-facing instruction prose."""
    if relative_path in AGENT_INSTRUCTION_EXACT_PATHS:
        return True
    return any(fnmatch.fnmatchcase(relative_path, pattern) for pattern in PROTECTED_GLOB_PATTERNS)


def is_protected_manifest_pattern(pattern: str) -> bool:
    """Return whether a manifest pattern names protected instruction paths."""
    if not has_wildcard(pattern):
        return is_protected_instruction_path(pattern)
    if pattern in PROTECTED_GLOB_PATTERNS:
        return True
    if pattern.startswith((".github/instructions/", ".cursor/rules/")):
        return True
    return any(fnmatch.fnmatchcase(path, pattern) for path in PROTECTED_EXACT_PATHS)


def manifest_pattern_matches_path(pattern: str, relative_path: str) -> bool:
    """Return whether a manifest pattern could cover ``relative_path``."""
    if has_wildcard(pattern):
        return fnmatch.fnmatchcase(relative_path, pattern)
    return pattern == relative_path


def path_has_symlink_component(repo_root: Path, relative_path: str) -> bool:
    """Return whether any component of ``relative_path`` under ``repo_root`` is a symlink."""
    current = repo_root
    for part in Path(relative_path).parts:
        current = current / part
        if current.is_symlink():
            return True
    return False


def repository_path_exists(repo_root: Path, relative_path: str) -> bool:
    """Return whether a non-symlink path exists inside ``repo_root``."""
    if path_has_symlink_component(repo_root, relative_path):
        return False
    path = repo_root / relative_path
    try:
        path.resolve().relative_to(repo_root.resolve())
    except (OSError, ValueError):
        return False
    return path.exists()


def resolve_safe_repository_target_path(
    repo_root: Path,
    raw_path: str,
    *,
    field_name: str = "path",
) -> Path:
    """Resolve a write target path while rejecting unsafe repository paths."""
    normalized_path, is_directory = normalize_repository_path(raw_path, field_name)
    if is_directory:
        raise RepositoryPathError(
            f"{field_name} must reference a file, not a directory: {raw_path}"
        )
    if path_has_symlink_component(repo_root, normalized_path):
        raise RepositoryPathError(f"{field_name} must not traverse a symlink: {raw_path}")

    target_path = repo_root / normalized_path
    trusted_root = repo_root.resolve()
    try:
        target_path.resolve(strict=False).relative_to(trusted_root)
    except (OSError, ValueError) as error:
        raise RepositoryPathError(
            f"{field_name} escapes the repository root: {raw_path}"
        ) from error
    return target_path


def classify_repository_file(repo_root: Path, raw_path: str) -> RepositoryFileClassification:
    """Classify a safe regular file as transformable UTF-8 text or byte-only."""
    path = resolve_safe_repository_target_path(repo_root, raw_path)
    relative_path = path.relative_to(repo_root).as_posix()
    if not path.is_file():
        raise RepositoryPathError(f"path must reference a regular file: {raw_path}")
    try:
        bytes_content = path.read_bytes()
    except OSError as error:
        raise TemplateSyncMaterializationError(
            f"Unable to read {relative_path}: {os_error_summary(error)}"
        ) from error

    if b"\x00" in bytes_content:
        return RepositoryFileClassification(
            relative_path=relative_path,
            path=path,
            is_transformable_text=False,
            text=None,
            bytes_content=bytes_content,
        )
    try:
        text = bytes_content.decode("utf-8")
    except UnicodeDecodeError:
        return RepositoryFileClassification(
            relative_path=relative_path,
            path=path,
            is_transformable_text=False,
            text=None,
            bytes_content=bytes_content,
        )
    return RepositoryFileClassification(
        relative_path=relative_path,
        path=path,
        is_transformable_text=True,
        text=text,
        bytes_content=bytes_content,
    )


def is_known_inline_block_marker(marker_name: str) -> bool:
    """Return whether a marker family is recognized under AND or ANY retention."""
    return marker_name in INLINE_BLOCK_MODULES or marker_name in INLINE_BLOCK_ANY_MODULES


def inline_block_module_requirement(marker_name: str) -> frozenset[str] | None:
    """Return the module set for a known inline marker family, or ``None``.

    The returned set is the same for AND-retention (``INLINE_BLOCK_MODULES``) and
    ANY-retention (``INLINE_BLOCK_ANY_MODULES``) markers; callers that need to
    distinguish the two semantics consult ``INLINE_BLOCK_ANY_MODULES`` directly.
    """
    requirement = INLINE_BLOCK_MODULES.get(marker_name)
    if requirement is not None:
        return requirement
    return INLINE_BLOCK_ANY_MODULES.get(marker_name)


def markdown_fence_context_for_path(relative_path: str) -> str | None:
    """Return the fence context for live inline markers in ``relative_path``."""
    if Path(relative_path).suffix.lower() in MARKDOWN_RENDERED_SUFFIXES:
        return MARKDOWN_FENCE_CONTEXT
    return None


def fence_context_for_path(relative_path: str) -> str:
    """Return the fence-parsing context for opaque Markdown content in a file.

    Unlike :func:`markdown_fence_context_for_path` (which returns ``None`` for
    non-Markdown files because their inline markers are never fence-suppressed),
    this always returns a context for callers that treat fenced Markdown as
    opaque, such as blank-line hygiene and reference-link scanning. Markdown
    files (``.md``/``.mdc``) use the standard context, which recognizes
    blockquote- and list-contained fences indented up to three spaces. Other
    files (notably YAML issue forms) embed Markdown inside ``value: |`` block
    scalars whose fences are indented further, so they use the embedded context,
    which allows arbitrary leading indentation.
    """
    if Path(relative_path).suffix.lower() in MARKDOWN_RENDERED_SUFFIXES:
        return MARKDOWN_FENCE_CONTEXT
    return EMBEDDED_MARKDOWN_FENCE_CONTEXT


def blank_line_limit_for_path(relative_path: str) -> int:
    """Return the post-pruning blank-line limit for a retained text file."""
    if Path(relative_path).suffix.lower() in YAML_SUFFIXES:
        return 1
    return 2


def parse_inline_block_marker(
    line: str,
    *,
    line_number: int,
    relative_path: str,
) -> InlineBlockMarker | None:
    """Parse one template-sync inline-block marker line, if present."""
    match = INLINE_BLOCK_MARKER_RE.match(line)
    if match is None:
        return None
    marker_name = match.group("name")
    if not is_known_inline_block_marker(marker_name):
        raise UnknownInlineBlockMarkerError(
            f"Unknown template-sync inline marker family: {marker_name}",
            relative_path=relative_path,
            line_number=line_number,
            marker_name=marker_name,
        )
    return InlineBlockMarker(
        kind=match.group("kind"),
        name=marker_name,
        line_number=line_number,
        line=line,
    )


def live_inline_marker_for_line(
    line: str,
    *,
    line_number: int,
    relative_path: str,
    is_fenced: bool,
) -> InlineBlockMarker | None:
    """Parse a live inline marker line while ignoring fenced-code examples."""
    if is_fenced:
        return None
    return parse_inline_block_marker(
        line,
        line_number=line_number,
        relative_path=relative_path,
    )


def live_inline_marker_lines(
    text: str,
    *,
    relative_path: str,
) -> tuple[tuple[int, str, InlineBlockMarker | None], ...]:
    """Return text lines paired with live inline markers for ``relative_path``."""
    fence_context = markdown_fence_context_for_path(relative_path)
    raw_lines = markdown_lines(text, keepends=True)
    if fence_context is None:
        states = tuple(
            MarkdownLineState(
                line_number=line_number,
                line=line,
                is_fenced=False,
            )
            for line_number, line in enumerate(raw_lines, 1)
        )
    else:
        states = markdown_line_states(raw_lines, fence_context=fence_context)

    return tuple(
        (
            state.line_number,
            state.line,
            live_inline_marker_for_line(
                state.line,
                line_number=state.line_number,
                relative_path=relative_path,
                is_fenced=state.is_fenced,
            ),
        )
        for state in states
    )


def collect_live_inline_block_spans(
    text: str,
    *,
    relative_path: str,
) -> tuple[InlineBlockSpan, ...]:
    """Return complete live inline-block spans using shared marker semantics."""
    blocks: list[InlineBlockSpan] = []
    open_marker: tuple[InlineBlockMarker, list[str]] | None = None

    for _line_number, line, marker in live_inline_marker_lines(
        text,
        relative_path=relative_path,
    ):
        if marker is None:
            if open_marker is not None:
                open_marker[1].append(line.rstrip("\r\n"))
            continue

        if marker.kind == "begin":
            if open_marker is not None:
                raise NestedInlineBlockError(
                    f"Nested template-sync inline marker inside {open_marker[0].name}",
                    relative_path=relative_path,
                    line_number=marker.line_number,
                    marker_name=marker.name,
                )
            open_marker = (marker, [])
            continue

        if open_marker is None:
            raise UnmatchedInlineBlockEndError(
                "Unmatched template-sync inline marker end",
                relative_path=relative_path,
                line_number=marker.line_number,
                marker_name=marker.name,
            )
        begin_marker, content_lines = open_marker
        if begin_marker.name != marker.name:
            raise MismatchedInlineBlockError(
                f"End marker {marker.name!r} does not match begin marker "
                f"{begin_marker.name!r} from line {begin_marker.line_number}",
                relative_path=relative_path,
                line_number=marker.line_number,
                marker_name=marker.name,
            )
        blocks.append(
            InlineBlockSpan(
                marker_name=marker.name,
                line_number=begin_marker.line_number,
                content_lines=tuple(content_lines),
            )
        )
        open_marker = None

    if open_marker is not None:
        begin_marker = open_marker[0]
        raise UnclosedInlineBlockError(
            f"Unclosed template-sync inline marker: {begin_marker.name}",
            relative_path=relative_path,
            line_number=begin_marker.line_number,
            marker_name=begin_marker.name,
        )
    return tuple(blocks)


def validate_inline_block_markers(text: str, *, relative_path: str) -> None:
    """Validate live inline template-sync marker pairing for ``relative_path``."""
    collect_live_inline_block_spans(text, relative_path=relative_path)


def inline_block_families_to_prune(
    text: str,
    included_modules: Collection[str],
    *,
    relative_path: str,
) -> tuple[str, ...]:
    """Return live inline marker families that module pruning would remove."""
    included_module_set = set(included_modules)
    pruned_families: set[str] = set()
    for block in collect_live_inline_block_spans(text, relative_path=relative_path):
        required_modules = inline_block_module_requirement(block.marker_name)
        if required_modules is None:
            continue
        if block.marker_name in INLINE_BLOCK_ANY_MODULES:
            if included_module_set.isdisjoint(required_modules):
                pruned_families.add(block.marker_name)
            continue
        if not required_modules.issubset(included_module_set):
            pruned_families.add(block.marker_name)
    return tuple(sorted(pruned_families))


def hash_inline_marker(marker_name: str, kind: str) -> str:
    """Return a hash-comment inline marker string for a marker family."""
    return f"# template-sync: {kind} {marker_name}"


def html_inline_marker(marker_name: str, kind: str) -> str:
    """Return a Markdown-safe HTML inline marker string for a marker family."""
    return f"<!-- template-sync: {kind} {marker_name} -->"


def apply_blank_line_hygiene(
    text: str,
    *,
    max_consecutive_blank_lines: int = 2,
    fence_context: str = MARKDOWN_FENCE_CONTEXT,
) -> str:
    """Collapse excessive blank-line runs after inline-block removal.

    Fenced code blocks are treated as opaque: blank lines inside a fence are
    preserved verbatim. Fence-aware pruning never removes inline blocks inside
    fences, so any blank run within a fence is original example content that
    must not be altered (markdownlint MD012 already ignores blank lines inside
    fenced code). Excess blank runs only arise outside fences, where pruning
    removes live blocks, and those are collapsed to
    ``max_consecutive_blank_lines`` (the caller passes one for YAML, two
    otherwise, via ``blank_line_limit_for_path``).

    ``fence_context`` selects how fences are recognized; the caller passes the
    file-type-appropriate context (via ``fence_context_for_path``) so that
    deeply-indented fences inside YAML ``value: |`` block scalars are also
    treated as opaque.
    """
    hygienic_lines: list[str] = []
    blank_run = 0
    active_fence: MarkdownFence | None = None

    for line in markdown_lines(text, keepends=True):
        if line_body(line).strip(" \t"):
            body = line_body(line)
            if active_fence is not None:
                content = active_fence_content(body, active_fence)
                if content is not None:
                    if parse_fence_close_from_content(
                        content,
                        fence_character=active_fence.character,
                        minimum_length=active_fence.length,
                        allow_arbitrary_indent=active_fence.allow_arbitrary_indent,
                    ):
                        active_fence = None
                    hygienic_lines.append(line)
                    blank_run = 0
                    continue
                active_fence = None

            open_fence = parse_markdown_fence_open(body, fence_context)
            if open_fence is not None:
                active_fence = open_fence
            blank_run = 0
            hygienic_lines.append(line)
            continue

        if active_fence is not None:
            # Preserve blank lines inside fenced code blocks verbatim.
            hygienic_lines.append(line)
            continue

        blank_run += 1
        if blank_run <= max_consecutive_blank_lines:
            hygienic_lines.append(line)
    return "".join(hygienic_lines)


def remove_inline_block_family(
    text: str,
    marker_name: str,
    *,
    relative_path: str,
    require_seen: bool = True,
    preserve_blank_line_hygiene: bool = True,
) -> str:
    """Return ``text`` after removing one complete inline block marker family."""
    if not is_known_inline_block_marker(marker_name):
        raise UnknownInlineBlockMarkerError(
            f"Unknown template-sync inline marker family: {marker_name}",
            relative_path=relative_path,
            marker_name=marker_name,
        )

    output_lines: list[str] = []
    stack: list[InlineBlockMarker] = []
    removed_blocks = 0

    for _line_number, line, marker in live_inline_marker_lines(
        text,
        relative_path=relative_path,
    ):
        is_removing_target = bool(stack and stack[-1].name == marker_name)

        if marker is None:
            if not is_removing_target:
                output_lines.append(line)
            continue

        if marker.kind == "begin":
            if stack:
                open_marker = stack[-1]
                raise NestedInlineBlockError(
                    f"Nested template-sync inline marker inside {open_marker.name}",
                    relative_path=relative_path,
                    line_number=marker.line_number,
                    marker_name=marker.name,
                )
            stack.append(marker)
            if marker.name == marker_name:
                removed_blocks += 1
            else:
                output_lines.append(line)
            continue

        if not stack:
            raise UnmatchedInlineBlockEndError(
                "Unmatched template-sync inline marker end",
                relative_path=relative_path,
                line_number=marker.line_number,
                marker_name=marker.name,
            )

        begin_marker = stack.pop()
        if begin_marker.name != marker.name:
            raise MismatchedInlineBlockError(
                f"End marker {marker.name!r} does not match begin marker "
                f"{begin_marker.name!r} from line {begin_marker.line_number}",
                relative_path=relative_path,
                line_number=marker.line_number,
                marker_name=marker.name,
            )
        if marker.name != marker_name:
            output_lines.append(line)

    if stack:
        begin_marker = stack[-1]
        raise UnclosedInlineBlockError(
            f"Unclosed template-sync inline marker: {begin_marker.name}",
            relative_path=relative_path,
            line_number=begin_marker.line_number,
            marker_name=begin_marker.name,
        )
    if require_seen and removed_blocks == 0:
        raise MissingExpectedInlineBlockError(
            f"Missing inline block for {marker_name}",
            relative_path=relative_path,
            marker_name=marker_name,
        )
    if removed_blocks == 0:
        return text

    stripped_text = "".join(output_lines)
    if preserve_blank_line_hygiene:
        return apply_blank_line_hygiene(
            stripped_text,
            max_consecutive_blank_lines=blank_line_limit_for_path(relative_path),
            fence_context=fence_context_for_path(relative_path),
        )
    return stripped_text


def remove_inline_blocks_for_modules(
    text: str,
    included_modules: Collection[str],
    *,
    relative_path: str,
) -> str:
    """Remove inline marker families whose module requirements are excluded.

    AND-retention families (``INLINE_BLOCK_MODULES``) are stripped when *any*
    required module is excluded. ANY-retention families
    (``INLINE_BLOCK_ANY_MODULES``) are stripped only when the included modules are
    disjoint from the marker's module set, that is when *none* of the named
    modules is included.
    """
    result = text
    included_module_set = set(included_modules)
    for marker_name, required_modules in INLINE_BLOCK_MODULES.items():
        if required_modules.issubset(included_module_set):
            continue
        if marker_name not in result:
            continue
        result = remove_inline_block_family(
            result,
            marker_name,
            relative_path=relative_path,
            require_seen=False,
        )
    for marker_name, required_modules in INLINE_BLOCK_ANY_MODULES.items():
        if not included_module_set.isdisjoint(required_modules):
            continue
        if marker_name not in result:
            continue
        result = remove_inline_block_family(
            result,
            marker_name,
            relative_path=relative_path,
            require_seen=False,
        )
    return result
