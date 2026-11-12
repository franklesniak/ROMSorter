"""Bounded repository I/O and Markdown primitives shared by instruction checking."""

from __future__ import annotations

import importlib
import json
import re
from collections.abc import Hashable, Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import Any, cast

import yaml  # type: ignore[import-untyped]

REMOVAL_DECISION = "REMOVE-LOCAL"


MARKDOWN_FENCE_CONTEXT = "markdown"


EMBEDDED_MARKDOWN_FENCE_CONTEXT = "embedded-markdown"


LIST_MARKER_RE = re.compile(
    r"^(?P<indent> {0,3})(?P<marker>(?:[-+*]|[0-9]{1,9}[.)]))(?P<spaces> {1,4})(?P<rest>.*)$"
)


EMBEDDED_LIST_MARKER_RE = re.compile(
    r"^(?P<indent> *)(?P<marker>(?:[-+*]|[0-9]{1,9}[.)]))(?P<spaces> {1,4})(?P<rest>.*)$"
)


class TemplateSyncMaterializationError(Exception):
    """Raised when shared template-sync planning cannot continue safely."""


class RepositoryPathError(TemplateSyncMaterializationError):
    """Raised when a repository path is unsafe or malformed."""


@dataclass(frozen=True)
class ProtectedFileDecision:
    """A path-scoped protected-file decision recorded in the marker."""

    path: str
    decision: str
    adoption_mode: str | None
    authorization_basis: str | None
    authorized_scope: str | None
    tailored_authorization_basis: str | None
    reason: str | None


@dataclass(frozen=True)
class ProtectedGuideContractWaiver:
    """A protected-guide obligation waiver recorded in the marker."""

    path: str
    contract_key: str
    target_path: str | None
    target_module: str | None
    linked_local_override_path: str | None
    reason: str
    authorization_basis: str


@dataclass(frozen=True)
class MarkdownLineState:
    """One Markdown-rendered line classified by fenced-code visibility."""

    line_number: int
    line: str
    is_fenced: bool


@dataclass(frozen=True)
class MarkdownFence:
    """Active line-oriented Markdown fenced-code block state."""

    character: str
    length: int
    info: str
    quote_depth: int
    list_content_indent: int | None
    allow_arbitrary_indent: bool


def os_error_summary(error: OSError) -> str:
    """Return an OSError summary that avoids implicit filesystem paths."""
    return f"{type(error).__name__}: {error.strerror or 'I/O error'}"


def repository_relative_path(path: Path, repo_root: Path) -> str:
    """Return a POSIX-style path relative to the repository root."""
    return path.relative_to(repo_root).as_posix()


def resolve_repo_path(repo_root: Path, raw_path: str) -> Path:
    """Resolve ``raw_path`` inside ``repo_root`` and reject path traversal."""
    candidate = Path(raw_path)
    if candidate.is_absolute():
        path = candidate.resolve()
    else:
        path = (repo_root / candidate).resolve()

    try:
        path.relative_to(repo_root)
    except ValueError as error:
        raise RepositoryPathError(f"Path escapes the repository root: {raw_path}") from error
    return path


def read_repository_text(
    path: Path,
    repo_root: Path,
    *,
    encoding: str = "utf-8",
    maximum_bytes: int | None = None,
) -> str:
    """Read a caller-resolved repository file with strict decoding and safe errors.

    A supplied nonnegative byte limit bounds the read before decoding. Without
    a limit, retain the existing text-reader size behavior. Both paths preserve
    universal newlines. Callers retain responsibility for path containment.
    I/O, oversized input, and decoding failures raise the existing domain error.
    """
    if maximum_bytes is not None and maximum_bytes < 0:
        raise ValueError("maximum_bytes must be nonnegative.")
    try:
        if maximum_bytes is None:
            return path.read_text(encoding=encoding)
        with path.open("rb") as stream:
            data = stream.read(maximum_bytes + 1)
        if len(data) > maximum_bytes:
            relative_path = repository_relative_path(path, repo_root)
            raise TemplateSyncMaterializationError(
                f"{relative_path} exceeds the {maximum_bytes}-byte input limit."
            )
        return data.decode(encoding).replace("\r\n", "\n").replace("\r", "\n")
    except OSError as error:
        relative_path = repository_relative_path(path, repo_root)
        raise TemplateSyncMaterializationError(
            f"Unable to read {relative_path}: {os_error_summary(error)}"
        ) from error
    except UnicodeDecodeError as error:
        relative_path = repository_relative_path(path, repo_root)
        raise TemplateSyncMaterializationError(
            f"Invalid {encoding} in {relative_path}: {error.reason}."
        ) from error


def parse_json_mapping(text: str, display_path: str) -> dict[str, Any]:
    """Parse one JSON object, rejecting repeated names before dictionary creation."""

    def unique_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        """Preserve object order while rejecting ambiguous decoded names."""
        document: dict[str, Any] = {}
        for key, value in pairs:
            if key in document:
                raise TemplateSyncMaterializationError(
                    f"Invalid JSON in {display_path}: duplicate object key {key!r}."
                )
            document[key] = value
        return document

    try:
        parsed = json.loads(text, object_pairs_hook=unique_object)
    except json.JSONDecodeError as error:
        raise TemplateSyncMaterializationError(
            f"Invalid JSON in {display_path}: {error}"
        ) from error
    if not isinstance(parsed, dict):
        raise TemplateSyncMaterializationError(f"{display_path} must contain a JSON object.")
    return cast(dict[str, Any], parsed)


def mapping_source_display_path(path: Path, repo_root: Path) -> str:
    """Describe caller-resolved inputs without authorizing access or leaking roots."""
    if path.is_relative_to(repo_root):
        return repository_relative_path(path, repo_root)
    return path.name


def load_json_mapping(
    path: Path, repo_root: Path, *, maximum_bytes: int | None = None
) -> dict[str, Any]:
    """Load a unique-key JSON object with an optional pre-decode byte limit."""
    return parse_json_mapping(
        read_repository_text(path, repo_root, maximum_bytes=maximum_bytes),
        mapping_source_display_path(path, repo_root),
    )


class UniqueKeySafeLoader(yaml.SafeLoader):
    """Reject repeated explicit keys while preserving safe YAML merge behavior."""

    def __init__(self, stream: str) -> None:
        """Keep per-document mapping identities without changing global constructors."""
        super().__init__(stream)
        self.checked_mapping_nodes: set[int] = set()

    def flatten_mapping(self, node: Any) -> None:
        """Check original keys once, before merges create intentional overrides."""
        if id(node) in self.checked_mapping_nodes:
            return
        self.checked_mapping_nodes.add(id(node))
        explicit_keys: dict[Any, Any] = {}
        merge_key = object()
        for key_node, _ in node.value:
            if key_node.tag == "tag:yaml.org,2002:merge":
                key = merge_key
            elif key_node.tag == "tag:yaml.org,2002:value":
                # SafeConstructor normalizes the special '=' key to a string.
                key = self.construct_scalar(key_node)
            else:
                key = self.construct_object(key_node)
            if not isinstance(key, Hashable):
                raise yaml.constructor.ConstructorError(
                    "while constructing a mapping",
                    node.start_mark,
                    "found unhashable key",
                    key_node.start_mark,
                )
            if key in explicit_keys:
                raise yaml.constructor.ConstructorError(
                    "while constructing a mapping",
                    explicit_keys[key],
                    "found duplicate explicit mapping key",
                    key_node.start_mark,
                )
            explicit_keys[key] = key_node.start_mark
        # Alias nodes can refer to an already-flattened mapping. Checking that
        # mapping again would mistake inherited overrides for explicit duplicates.
        super().flatten_mapping(node)


def parse_yaml_mapping(text: str, display_path: str) -> dict[str, Any]:
    """Parse one safe mapping; reject duplicates and wrap YAML errors with context.

    Explicit keys must be unique before merge expansion. Safe aliases and
    inherited merge overrides retain PyYAML semantics. Callers own file access,
    decoding and any input-size limit.
    """
    try:
        parsed = yaml.load(text, Loader=UniqueKeySafeLoader)
    except yaml.YAMLError as error:
        raise TemplateSyncMaterializationError(
            f"Invalid YAML in {display_path}: {error}"
        ) from error
    if not isinstance(parsed, dict):
        raise TemplateSyncMaterializationError(f"{display_path} must contain a YAML mapping.")
    return cast(dict[str, Any], parsed)


def load_yaml_mapping(
    path: Path, repo_root: Path, *, maximum_bytes: int | None = None
) -> dict[str, Any]:
    """Load a unique-key YAML mapping with an optional pre-decode byte limit."""
    return parse_yaml_mapping(
        read_repository_text(path, repo_root, encoding="utf-8-sig", maximum_bytes=maximum_bytes),
        mapping_source_display_path(path, repo_root),
    )


def validate_local_schema_references(schema: dict[str, Any], specification: Any) -> None:
    """Inspect schema resources without treating annotation data as executable schemas."""
    pending = [specification.create_resource(schema)]
    while pending:
        resource = pending.pop()
        node = resource.contents
        if isinstance(node, dict):
            for keyword in ("$ref", "$dynamicRef"):
                if keyword in node:
                    reference = node[keyword]
                    if not isinstance(reference, str) or not reference.startswith("#"):
                        raise TemplateSyncMaterializationError(
                            "Only local instruction schema references are supported."
                        )
        pending.extend(resource.subresources())


def validate_schema(
    document: dict[str, Any], schema: dict[str, Any], document_path: Path, repo_root: Path
) -> None:
    """Validate a loaded document against a Draft 2020-12 JSON Schema."""
    try:
        jsonschema_module = cast(Any, importlib.import_module("jsonschema"))
    except ImportError as error:
        raise TemplateSyncMaterializationError(
            "jsonschema is unavailable. Install jsonschema, or run this through "
            "the pre-commit hook, which declares the validator dependency."
        ) from error
    try:
        jsonschema_module.Draft202012Validator.check_schema(schema)
    except jsonschema_module.exceptions.SchemaError as error:
        raise TemplateSyncMaterializationError("Invalid instruction validation schema.") from error
    referencing_module = cast(Any, importlib.import_module("referencing"))
    referencing_schema = cast(Any, importlib.import_module("referencing.jsonschema"))
    referencing_errors = cast(Any, importlib.import_module("referencing.exceptions"))
    try:
        validate_local_schema_references(schema, referencing_schema.DRAFT202012)
        validator = jsonschema_module.Draft202012Validator(
            schema, registry=referencing_module.Registry()
        )
        errors = sorted(validator.iter_errors(document), key=lambda error: error.json_path)
    except (
        referencing_errors.Unresolvable,
        referencing_errors.CannotDetermineSpecification,
        referencing_schema.UnknownDialect,
    ) as error:
        raise TemplateSyncMaterializationError(
            "Unable to resolve an instruction validation schema reference."
        ) from error
    if not errors:
        return

    relative_path = repository_relative_path(document_path, repo_root)
    messages = "\n".join(f"  - {error.json_path}: {error.message}" for error in errors[:10])
    remaining = len(errors) - 10
    if remaining > 0:
        messages += f"\n  - ... {remaining} more validation error(s)"
    raise TemplateSyncMaterializationError(
        f"Schema validation failed for {relative_path}:\n{messages}"
    )


def normalize_repository_path(raw_path: str, field_name: str) -> tuple[str, bool]:
    """Normalize a marker path and return ``(path, is_directory_prefix)``."""
    if "\\" in raw_path:
        raise RepositoryPathError(f"{field_name} must use POSIX separators: {raw_path}")
    if raw_path.startswith("/"):
        raise RepositoryPathError(f"{field_name} must be repository-relative: {raw_path}")

    is_directory = raw_path.endswith("/")
    stripped = raw_path.strip("/")
    if not stripped:
        raise RepositoryPathError(f"{field_name} must not be empty: {raw_path}")
    parts = stripped.split("/")
    if any(part in ("", ".", "..") for part in parts):
        raise RepositoryPathError(f"{field_name} must not contain traversal segments: {raw_path}")
    return stripped, is_directory


def line_body(line: str) -> str:
    """Return ``line`` without a trailing text line ending."""
    return line.rstrip("\r\n")


def markdown_lines(text: str, *, keepends: bool = False) -> list[str]:
    """Split only CR, LF and CRLF physical lines, preserving other characters.

    The fixed delimiter pattern scans once. Empty input and terminal endings
    follow splitlines behavior without promoting Unicode/control separators.
    """
    lines: list[str] = []
    start = 0
    for ending in re.finditer(r"\r\n?|\n", text):
        lines.append(text[start : ending.end() if keepends else ending.start()])
        start = ending.end()
    if start < len(text):
        lines.append(text[start:])
    return lines


def consume_blockquote_prefix(
    line: str,
    *,
    max_depth: int | None = None,
    allow_arbitrary_indent: bool = False,
) -> tuple[int, int]:
    """Return ``(depth, offset)`` after consuming Markdown blockquote prefixes.

    ``allow_arbitrary_indent`` lifts GFM's 0-3 space limit on the indentation
    preceding each ``>`` so blockquote-contained fences inside deeply-indented
    YAML ``value: |`` block scalars are recognized as well.
    """
    depth = 0
    offset = 0
    while max_depth is None or depth < max_depth:
        spaces = 0
        while offset < len(line) and line[offset] == " " and (allow_arbitrary_indent or spaces < 3):
            offset += 1
            spaces += 1
        if offset >= len(line) or line[offset] != ">":
            offset -= spaces
            break
        offset += 1
        depth += 1
        if offset < len(line) and line[offset] == " ":
            offset += 1
    return depth, offset


def line_is_outside_list_item(line: str, content_indent: int) -> bool:
    """Return whether ``line`` ends a conservative list-item containing block."""
    if not line.strip(" \t"):
        return False
    leading_spaces = len(line) - len(line.lstrip(" "))
    return leading_spaces < content_indent


def parse_fence_open_from_content(
    content: str,
    *,
    allow_arbitrary_indent: bool,
) -> tuple[str, int, str] | None:
    """Return opening fence ``(character, length, info)`` from Markdown content."""
    leading_spaces = len(content) - len(content.lstrip(" "))
    if not allow_arbitrary_indent and leading_spaces > 3:
        return None
    stripped = content[leading_spaces:]
    if not stripped.startswith(("```", "~~~")):
        return None

    fence_character = stripped[0]
    fence_length = 0
    while fence_length < len(stripped) and stripped[fence_length] == fence_character:
        fence_length += 1
    if fence_length < 3:
        return None

    info = stripped[fence_length:]
    if fence_character == "`" and "`" in info:
        return None
    return fence_character, fence_length, info


def parse_fence_close_from_content(
    content: str,
    *,
    fence_character: str,
    minimum_length: int,
    allow_arbitrary_indent: bool,
) -> bool:
    """Return whether Markdown content closes the active fenced-code block."""
    leading_spaces = len(content) - len(content.lstrip(" "))
    if not allow_arbitrary_indent and leading_spaces > 3:
        return False
    stripped = content[leading_spaces:]
    if not stripped.startswith(fence_character * minimum_length):
        return False

    fence_length = 0
    while fence_length < len(stripped) and stripped[fence_length] == fence_character:
        fence_length += 1
    if fence_length < minimum_length:
        return False
    return stripped[fence_length:].strip(" \t") == ""


def parse_markdown_fence_open(line: str, fence_context: str) -> MarkdownFence | None:
    """Return active fence state when ``line`` opens a fenced-code block.

    Both contexts recognize blockquote- and list-contained fences; the embedded
    context (``allow_arbitrary_indent``) lifts GFM's 0-3 space indentation caps
    so containing-block fences inside deeply-indented YAML ``value: |`` block
    scalars are recognized as opaque too.
    """
    allow_arbitrary_indent = fence_context == EMBEDDED_MARKDOWN_FENCE_CONTEXT

    quote_depth, quote_offset = consume_blockquote_prefix(
        line, allow_arbitrary_indent=allow_arbitrary_indent
    )
    content = line[quote_offset:]

    list_marker_re = EMBEDDED_LIST_MARKER_RE if allow_arbitrary_indent else LIST_MARKER_RE
    list_match = list_marker_re.match(content)
    if list_match is not None:
        list_content_indent = (
            len(list_match.group("indent"))
            + len(list_match.group("marker"))
            + len(list_match.group("spaces"))
        )
        opened = parse_fence_open_from_content(
            list_match.group("rest"),
            allow_arbitrary_indent=allow_arbitrary_indent,
        )
        if opened is not None:
            character, length, info = opened
            return MarkdownFence(
                character=character,
                length=length,
                info=info,
                quote_depth=quote_depth,
                list_content_indent=list_content_indent,
                allow_arbitrary_indent=allow_arbitrary_indent,
            )

    opened = parse_fence_open_from_content(content, allow_arbitrary_indent=allow_arbitrary_indent)
    if opened is None:
        return None
    character, length, info = opened
    return MarkdownFence(
        character=character,
        length=length,
        info=info,
        quote_depth=quote_depth,
        list_content_indent=None,
        allow_arbitrary_indent=allow_arbitrary_indent,
    )


def active_fence_content(line: str, active_fence: MarkdownFence) -> str | None:
    """Return line content inside ``active_fence`` or ``None`` if its container ended."""
    if active_fence.quote_depth:
        quote_depth, quote_offset = consume_blockquote_prefix(
            line,
            max_depth=active_fence.quote_depth,
            allow_arbitrary_indent=active_fence.allow_arbitrary_indent,
        )
        if quote_depth < active_fence.quote_depth:
            return None
        content = line[quote_offset:]
    else:
        content = line

    if active_fence.list_content_indent is None:
        return content
    if line_is_outside_list_item(content, active_fence.list_content_indent):
        return None
    if len(content) >= active_fence.list_content_indent:
        return content[active_fence.list_content_indent :]
    return ""


def markdown_line_states(
    lines: Iterable[str],
    *,
    fence_context: str,
) -> tuple[MarkdownLineState, ...]:
    """Classify Markdown-rendered lines as fenced-code or live content."""
    states: list[MarkdownLineState] = []
    active_fence: MarkdownFence | None = None

    for line_number, raw_line in enumerate(lines, 1):
        line = line_body(raw_line)
        if active_fence is not None:
            content = active_fence_content(line, active_fence)
            if content is not None:
                is_close = parse_fence_close_from_content(
                    content,
                    fence_character=active_fence.character,
                    minimum_length=active_fence.length,
                    allow_arbitrary_indent=active_fence.allow_arbitrary_indent,
                )
                states.append(
                    MarkdownLineState(
                        line_number=line_number,
                        line=raw_line,
                        is_fenced=True,
                    )
                )
                if is_close:
                    active_fence = None
                continue
            active_fence = None

        active_fence = parse_markdown_fence_open(line, fence_context)
        states.append(
            MarkdownLineState(
                line_number=line_number,
                line=raw_line,
                is_fenced=active_fence is not None,
            )
        )

    return tuple(states)


def lines_outside_markdown_fences(
    text: str,
    *,
    fence_context: str,
) -> tuple[tuple[int, str], ...]:
    """Return lines that are not inside Markdown fenced-code blocks."""
    return tuple(
        (state.line_number, state.line)
        for state in markdown_line_states(
            markdown_lines(text),
            fence_context=fence_context,
        )
        if not state.is_fenced
    )
