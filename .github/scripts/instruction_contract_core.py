"""Shared bounded static instruction-contract validation engine."""

from __future__ import annotations

import hashlib
import json
import os
import posixpath
import re
import stat
import subprocess
import sys
import threading
import time
from bisect import bisect_left, bisect_right
from dataclasses import dataclass, replace
from html import unescape
from html.entities import html5
from itertools import pairwise
from pathlib import Path
from typing import Any, NoReturn, cast
from urllib.parse import SplitResult, unquote, urlsplit

import instruction_contract_support as support
from instruction_contract_support import (
    LIST_MARKER_RE,
    MARKDOWN_FENCE_CONTEXT,
    MarkdownFence,
    active_fence_content,
    consume_blockquote_prefix,
    lines_outside_markdown_fences,
    markdown_lines,
    normalize_repository_path,
    parse_fence_close_from_content,
    parse_markdown_fence_open,
    read_repository_text,
)

MAXIMUM_INPUT_BYTES = 1024 * 1024


MAXIMUM_GIT_OUTPUT_BYTES = 1024 * 1024


MAXIMUM_CLAUDE_CONTAINER_STEPS = 64


GIT_TIMEOUT_SECONDS = 30


CLAUDE_IMPORT_PATTERN = re.compile(
    r"(?m)(?<!\S)@(?P<target>(?!(?:claude|codex)(?=$|\s))[^\s<>]+)(?=$|\s)"
)


POLICY_CELL_WORD = (
    r"[^|\u0009-\u000d\u001c-\u0020\u0085\u00a0\u1680"
    r"\u2000-\u200a\u2028-\u2029\u202f\u205f\u3000\ufeff]"
)


POLICY_CELL_PATTERN = re.compile(rf"^{POLICY_CELL_WORD}+(?: {POLICY_CELL_WORD}+)*(?![\s\S])")


POLICY_HEADING_PATTERN = re.compile(r"^#{1,6} [^\r\n]*[^ \t\r\n](?![\s\S])")


POLICY_HTML_LITERAL_START = re.compile(
    r"^ {0,3}<(?:pre|script|style|textarea)(?=[ \t>]|$)", re.IGNORECASE | re.ASCII
)


POLICY_HTML_LITERAL_END = re.compile(r"</(?:pre|script|style|textarea)>", re.IGNORECASE | re.ASCII)


POLICY_HTML_BLOCK_START = re.compile(
    r"^ {0,3}</?(?:address|article|aside|base|basefont|blockquote|body|caption|center|col|"
    r"colgroup|dd|details|dialog|dir|div|dl|dt|fieldset|figcaption|figure|footer|form|frame|"
    r"frameset|h[1-6]|head|header|hr|html|iframe|legend|li|link|main|menu|menuitem|nav|"
    r"noframes|ol|optgroup|option|p|param|search|section|summary|table|tbody|td|tfoot|th|"
    r"thead|title|tr|track|ul)(?=[ \t>]|/>|$)",
    re.IGNORECASE | re.ASCII,
)


POLICY_HTML_COMPLETE_TAG = re.compile(
    r"^ {0,3}(?:</[A-Za-z][A-Za-z0-9-]*[ \t]*>|"
    r"<(?![Pp][Rr][Ee](?=[ \t/>]))(?![Ss][Cc][Rr][Ii][Pp][Tt](?=[ \t/>]))"
    r"(?![Ss][Tt][Yy][Ll][Ee](?=[ \t/>]))(?![Tt][Ee][Xx][Tt][Aa][Rr][Ee][Aa](?=[ \t/>]))"
    r"[A-Za-z][A-Za-z0-9-]*(?:[ \t]+[A-Za-z_:][A-Za-z0-9_.:-]*"
    r"(?:[ \t]*=[ \t]*(?:[^ \t\"'=<>`]+|'[^']*'|\"[^\"]*\"))?)*[ \t]*/?>)[ \t]*$"
)


POLICY_HTML_BLANK_END = re.compile(r"^[ \t]*$")


POLICY_HTML_AMBIGUOUS_START = re.compile(
    r"^ {0,3}(?:<[Tt][Ee][Xx][Tt][Aa][Rr][Ee][Aa](?=[ \t>]|/>|$)|"
    r"</?(?:[Ss][Ee][Aa][Rr][Cc][Hh]|[Ss][Oo][Uu][Rr][Cc][Ee])(?=[ \t>]|/>|$)|<![a-z])"
)


POLICY_HTML_TEXTAREA_END = re.compile(r"</textarea>", re.IGNORECASE | re.ASCII)


POLICY_HTML_AMBIGUITY_PREFIX = "[unsupported HTML grammar] "


class InstructionContractValidationError(Exception):
    """Raised when instruction contract validation cannot produce a clean result."""


@dataclass(frozen=True)
class RequiredTable:
    """An exact ordered decision table in an operative Markdown section."""

    headers: tuple[str, ...]
    rows: tuple[tuple[str, ...], ...]


PolicyBlock = tuple[str, str | RequiredTable]


@dataclass(frozen=True)
class RequiredSection:
    """Complete ordered clauses and tables owned by one unique heading."""

    heading: str
    required_paragraphs: tuple[str, ...]
    required_tables: tuple[RequiredTable, ...]
    next_heading: str | None = None
    requires_modules: tuple[str, ...] = ()
    required_block_order: tuple[str, ...] = ()


@dataclass(frozen=True)
class PolicyLineIndex:
    """Positions and next-heading offsets for one operative document only."""

    positions: dict[str, list[int]]
    next_headings: list[int]


@dataclass(frozen=True)
class InstructionContract:
    """Required anchors for one protected instruction file."""

    path: str
    requires_modules: tuple[str, ...]
    required_headings: tuple[str, ...]
    required_phrases: tuple[str, ...]
    required_sections: tuple[RequiredSection, ...] = ()


@dataclass(frozen=True)
class ProtectedGuideSectionObligation:
    """Protected-guide headings or phrases that are stale when modules are excluded."""

    key: str
    path: str
    target_modules: tuple[str, ...]
    stale_headings: tuple[str, ...]
    stale_phrases: tuple[str, ...]


@dataclass(frozen=True)
class ProtectedGuideReferenceObligation:
    """Protected-guide references that are stale when modules are excluded."""

    key: str
    path: str
    reference_kind: str
    target_modules: tuple[str, ...]
    target_path: str | None
    tokens: tuple[str, ...]


@dataclass(frozen=True)
class ProtectedGuideReferenceFinding:
    """A declared reference that remains after its target module is excluded."""

    path: str
    line_number: int
    contract_key: str
    reference_kind: str
    target: str
    target_path: str | None
    target_modules: tuple[str, ...]


@dataclass(frozen=True)
class InstructionContractWaiver:
    """A marker waiver for one missing instruction-contract anchor."""

    path: str
    anchor: str
    reason: str
    authorization_basis: str


@dataclass(frozen=True)
class MissingAnchor:
    """A required heading or phrase that was not found in a retained file."""

    path: str
    anchor_type: str
    anchor: str


@dataclass(frozen=True)
class StaleProtectedGuideSection:
    """A protected-guide heading or phrase that still describes an excluded module."""

    path: str
    contract_key: str
    anchor_type: str
    anchor: str
    target_modules: tuple[str, ...]


@dataclass(frozen=True)
class MissingFile:
    """A retained instruction contract whose file is absent without removal authorization."""

    path: str


@dataclass(frozen=True)
class AuthorizedRemoval:
    """A missing protected file skipped because marker authorization removed it."""

    path: str
    authorization_basis: str
    authorized_scope: str
    reason: str


@dataclass(frozen=True)
class SkippedContract:
    """A downstream contract skipped because its required modules are not retained."""

    path: str
    requires_modules: tuple[str, ...]


@dataclass(frozen=True)
class ActiveClaudeImport:
    """One active Claude file import in a governed shared instruction file."""

    path: str
    line: int
    target: str


@dataclass(frozen=True)
class ClaudeImportLine:
    """One visible Claude-import line and its Markdown paragraph identity."""

    source: str
    visible: str
    paragraph_id: int | None = None


@dataclass(frozen=True)
class TrackedClaudeLocalMemory:
    """One exact-basename Claude local-memory path found in the Git index."""

    path: str


@dataclass(frozen=True)
class InstructionContractReport:
    """Instruction contract validation details to print for the operator."""

    mode: str
    contracts_checked: tuple[InstructionContract, ...]
    skipped_contracts: tuple[SkippedContract, ...]
    missing_files: tuple[MissingFile, ...]
    missing_anchors: tuple[MissingAnchor, ...]
    stale_protected_guide_sections: tuple[StaleProtectedGuideSection, ...]
    applied_waivers: tuple[InstructionContractWaiver, ...]
    applied_protected_guide_waivers: tuple[support.ProtectedGuideContractWaiver, ...]
    authorized_removals: tuple[AuthorizedRemoval, ...]
    warnings: tuple[str, ...]
    active_claude_imports: tuple[ActiveClaudeImport, ...] = ()
    tracked_claude_local_memory: tuple[TrackedClaudeLocalMemory, ...] = ()
    tracked_claude_local_memory_inventory_applicable: bool = True
    stale_protected_guide_references: tuple[ProtectedGuideReferenceFinding, ...] = ()

    @property
    def has_failures(self) -> bool:
        """Return whether validation found unwaived failures.

        A failure is any unwaived contract failure, active Claude import, or
        tracked Claude local-memory path.
        """
        return bool(
            self.missing_files
            or self.missing_anchors
            or self.stale_protected_guide_sections
            or self.stale_protected_guide_references
            or self.active_claude_imports
            or self.tracked_claude_local_memory
        )


def load_schema_validated_yaml(
    document_path: Path,
    schema_path: Path,
    repo_root: Path,
) -> dict[str, Any]:
    """Load a YAML mapping and validate it against a JSON Schema."""
    document = support.load_yaml_mapping(
        document_path, repo_root, maximum_bytes=MAXIMUM_INPUT_BYTES
    )
    schema = support.load_json_mapping(schema_path, repo_root, maximum_bytes=MAXIMUM_INPUT_BYTES)
    support.validate_schema(document, schema, document_path, repo_root)
    return document


def normalize_policy_paragraph(text: str) -> str:
    """Fold supported ASCII wrapping while retaining other literal characters."""
    return re.sub(r"[ \t\r\n]+", " ", text).strip(" ")


def _required_string_list(
    raw_contract: dict[str, object],
    field_name: str,
) -> tuple[str, ...]:
    """Return a tuple from an optional schema-validated string list."""
    values = raw_contract.get(field_name, [])
    if not isinstance(values, list):
        path_value = raw_contract.get("path")
        path = path_value if isinstance(path_value, str) else "<unknown>"
        raise InstructionContractValidationError(f"{path} {field_name} must be a string list.")
    string_values: list[str] = []
    for value in cast(list[object], values):
        if not isinstance(value, str):
            path_value = raw_contract.get("path")
            path = path_value if isinstance(path_value, str) else "<unknown>"
            raise InstructionContractValidationError(f"{path} {field_name} must be a string list.")
        string_values.append(value)
    return tuple(string_values)


def parse_required_sections(raw_contract: dict[str, object]) -> tuple[RequiredSection, ...]:
    """Reject ambiguous normalized paragraphs and table shapes in section contracts."""
    sections: list[RequiredSection] = []
    headings: set[str] = set()
    successor_owners: dict[str | None, str] = {}
    for raw in cast(list[dict[str, Any]], raw_contract.get("required_sections", [])):
        heading = cast(str, raw["heading"])
        next_heading = cast(str | None, raw["next_heading"])
        for field, value in (("heading", heading), ("next_heading", next_heading)):
            if (value is None and field == "heading") or (
                value is not None and POLICY_HEADING_PATTERN.fullmatch(value) is None
            ):
                raise InstructionContractValidationError(
                    f"Noncanonical contract {field}: {value!r}"
                )
        if heading in headings:
            raise InstructionContractValidationError(f"Duplicate required section: {heading}")
        headings.add(heading)
        if next_heading == heading:
            raise InstructionContractValidationError(f"Self-successor section: {heading}")
        if next_heading in successor_owners:
            raise InstructionContractValidationError(
                f"Shared section successor {next_heading!r}: "
                f"{successor_owners[next_heading]} and {heading}"
            )
        successor_owners[next_heading] = heading
        requires_modules = _required_string_list(raw, "requires_modules")
        if "requires_modules" in raw and (
            not requires_modules or len(set(requires_modules)) != len(requires_modules)
        ):
            raise InstructionContractValidationError(
                f"Section requires_modules must be nonempty and unique: {heading}"
            )
        paragraphs = _required_string_list(raw, "required_paragraphs")
        normalized_paragraphs = [normalize_policy_paragraph(value) for value in paragraphs]
        if any(not value or value.isspace() for value in normalized_paragraphs):
            raise InstructionContractValidationError(
                f"Empty normalized contract paragraph: {heading}"
            )
        if len(set(normalized_paragraphs)) != len(normalized_paragraphs):
            raise InstructionContractValidationError(
                f"Duplicate normalized contract paragraph: {heading}"
            )
        tables: list[RequiredTable] = []
        for table in cast(list[dict[str, Any]], raw.get("required_tables", [])):
            headers = tuple(cast(list[str], table["headers"]))
            rows = tuple(tuple(row) for row in cast(list[list[str]], table["rows"]))
            if any(
                POLICY_CELL_PATTERN.fullmatch(cell) is None
                for row in (headers, *rows)
                for cell in row
            ):
                raise InstructionContractValidationError(
                    f"Noncanonical contract table cell: {heading}"
                )
            if any(len(row) != len(headers) for row in rows):
                raise InstructionContractValidationError(f"Ragged contract table: {heading}")
            if len({row[0] for row in rows}) != len(rows):
                raise InstructionContractValidationError(f"Duplicate table condition: {heading}")
            tables.append(RequiredTable(headers, rows))
        if len(set(tables)) != len(tables):
            raise InstructionContractValidationError(f"Duplicate contract table: {heading}")
        block_order = _required_string_list(raw, "required_block_order")
        if "required_block_order" in raw and (
            any(kind not in {"paragraph", "table"} for kind in block_order)
            or block_order.count("paragraph") != len(paragraphs)
            or block_order.count("table") != len(tables)
        ):
            raise InstructionContractValidationError(
                f"Block order must match paragraph and table counts: {heading}"
            )
        sections.append(
            RequiredSection(
                heading,
                paragraphs,
                tuple(tables),
                next_heading,
                requires_modules,
                block_order,
            )
        )
    successors = {section.heading: section.next_heading for section in sections}
    completed: set[str] = set()
    for section in sections:
        visited: set[str] = set()
        successor: str | None = section.heading
        while successor in successors and successor not in completed:
            assert successor is not None
            if successor in visited:
                raise InstructionContractValidationError(
                    f"Cyclic section successor chain: {section.heading}"
                )
            visited.add(successor)
            successor = successors[successor]
        completed.update(visited)
    return tuple(sections)


def parse_contracts(
    contracts_document: dict[str, Any],
    manifest_modules: set[str],
) -> tuple[InstructionContract, ...]:
    """Extract normalized instruction contracts from a schema-validated document."""
    raw_contracts = contracts_document.get("instruction_contracts")
    if not isinstance(raw_contracts, list):
        raise InstructionContractValidationError("instruction_contracts must be a list.")
    raw_contracts = cast(list[object], raw_contracts)

    contracts: list[InstructionContract] = []
    seen_paths: set[str] = set()
    for raw_contract in raw_contracts:
        if not isinstance(raw_contract, dict):
            raise InstructionContractValidationError("Each instruction contract must be a mapping.")
        raw_contract = cast(dict[str, object], raw_contract)

        raw_path = raw_contract.get("path")
        if not isinstance(raw_path, str):
            raise InstructionContractValidationError("Each instruction contract must define path.")
        path, is_directory = normalize_repository_path(
            raw_path,
            "instruction_contracts[].path",
        )
        if is_directory:
            raise InstructionContractValidationError(
                f"instruction_contracts[].path must reference a file, not a directory: {raw_path}"
            )
        if path in seen_paths:
            raise InstructionContractValidationError(f"Duplicate instruction contract path: {path}")
        seen_paths.add(path)

        requires_modules = _required_string_list(raw_contract, "requires_modules")
        if not requires_modules:
            raise InstructionContractValidationError(f"{path} requires_modules must not be empty.")
        unknown_modules = set(requires_modules) - manifest_modules
        if unknown_modules:
            raise InstructionContractValidationError(
                f"{path} references unknown manifest module(s): "
                + ", ".join(sorted(unknown_modules))
            )

        required_headings = _required_string_list(raw_contract, "required_headings")
        required_phrases = _required_string_list(raw_contract, "required_phrases")
        required_sections = parse_required_sections(raw_contract)
        for section in required_sections:
            unknown_section_modules = set(section.requires_modules) - manifest_modules
            if unknown_section_modules:
                raise InstructionContractValidationError(
                    f"{path}: {section.heading} references unknown manifest module(s): "
                    + ", ".join(sorted(unknown_section_modules))
                )
        if not required_headings and not required_phrases and not required_sections:
            raise InstructionContractValidationError(
                f"{path} must define at least one required heading, phrase, or section."
            )

        contracts.append(
            InstructionContract(
                path=path,
                requires_modules=requires_modules,
                required_headings=required_headings,
                required_phrases=required_phrases,
                required_sections=required_sections,
            )
        )
    return tuple(contracts)


def validate_protected_guide_obligation_keys(contracts_document: dict[str, Any]) -> None:
    """Keep the type-less marker waiver identity unique across obligation families."""
    family_keys: list[set[tuple[str, str]]] = []
    for kind in ("section", "reference"):
        field = f"protected_guide_{kind}_obligations"
        raw_obligations = contracts_document.get(field, [])
        if not isinstance(raw_obligations, list):
            raise InstructionContractValidationError(f"{field} must be a list.")
        keys: set[tuple[str, str]] = set()
        for raw_obligation in raw_obligations:
            if not isinstance(raw_obligation, dict):
                raise InstructionContractValidationError(
                    f"Each protected guide {kind} obligation must be a mapping."
                )
            key = raw_obligation.get("key")
            raw_path = raw_obligation.get("path")
            if not isinstance(key, str) or not isinstance(raw_path, str):
                raise InstructionContractValidationError(
                    f"Each protected guide {kind} obligation must define key and path."
                )
            path, is_directory = normalize_repository_path(raw_path, f"{field}[].path")
            if is_directory:
                raise InstructionContractValidationError(
                    f"{field}[].path must reference a file, not a directory: {raw_path}"
                )
            keys.add((path, key))
        family_keys.append(keys)
    duplicate_keys = family_keys[0] & family_keys[1]
    if duplicate_keys:
        formatted_keys = ", ".join(f"({path}, {key})" for path, key in sorted(duplicate_keys))
        raise InstructionContractValidationError(
            "Duplicate protected-guide obligation (path, key) pair(s) across section "
            f"and reference types: {formatted_keys}"
        )


def parse_protected_guide_section_obligations(
    contracts_document: dict[str, Any],
    manifest_modules: set[str],
) -> tuple[ProtectedGuideSectionObligation, ...]:
    """Extract protected-guide stale-section obligations from a contract document."""
    raw_obligations = contracts_document.get("protected_guide_section_obligations", [])
    if not isinstance(raw_obligations, list):
        raise InstructionContractValidationError(
            "protected_guide_section_obligations must be a list."
        )

    obligations: list[ProtectedGuideSectionObligation] = []
    seen_keys: set[tuple[str, str]] = set()
    duplicate_keys: set[tuple[str, str]] = set()
    for raw_obligation in cast(list[object], raw_obligations):
        if not isinstance(raw_obligation, dict):
            raise InstructionContractValidationError(
                "Each protected guide section obligation must be a mapping."
            )
        raw_obligation = cast(dict[str, object], raw_obligation)

        key = raw_obligation.get("key")
        raw_path = raw_obligation.get("path")
        if not isinstance(key, str) or not isinstance(raw_path, str):
            raise InstructionContractValidationError(
                "Each protected guide section obligation must define key and path."
            )
        path, is_directory = normalize_repository_path(
            raw_path,
            "protected_guide_section_obligations[].path",
        )
        if is_directory:
            raise InstructionContractValidationError(
                "protected_guide_section_obligations[].path must reference a file, "
                f"not a directory: {raw_path}"
            )

        obligation_key = (path, key)
        if obligation_key in seen_keys:
            duplicate_keys.add(obligation_key)
        seen_keys.add(obligation_key)

        target_modules = _required_string_list(raw_obligation, "target_modules")
        if not target_modules:
            raise InstructionContractValidationError(f"{path} target_modules must not be empty.")
        unknown_modules = set(target_modules) - manifest_modules
        if unknown_modules:
            raise InstructionContractValidationError(
                f"{path} protected guide section obligation {key} references unknown "
                "manifest module(s): " + ", ".join(sorted(unknown_modules))
            )

        stale_headings = _required_string_list(raw_obligation, "stale_headings")
        stale_phrases = _required_string_list(raw_obligation, "stale_phrases")
        if not stale_headings and not stale_phrases:
            raise InstructionContractValidationError(
                f"{path} protected guide section obligation {key} must define at least "
                "one stale heading or phrase."
            )

        obligations.append(
            ProtectedGuideSectionObligation(
                key=key,
                path=path,
                target_modules=target_modules,
                stale_headings=stale_headings,
                stale_phrases=stale_phrases,
            )
        )

    if duplicate_keys:
        formatted_keys = ", ".join(f"({path}, {key})" for path, key in sorted(duplicate_keys))
        raise InstructionContractValidationError(
            "Duplicate protected_guide_section_obligations (path, key) pair(s): "
            f"{formatted_keys}"
        )
    validate_protected_guide_obligation_keys(contracts_document)
    return tuple(obligations)


def parse_protected_guide_reference_obligations(
    contracts_document: dict[str, Any],
    manifest_modules: set[str],
) -> tuple[ProtectedGuideReferenceObligation, ...]:
    """Extract protected-guide stale-reference obligations from a contract document."""
    raw_obligations = contracts_document.get("protected_guide_reference_obligations", [])
    if not isinstance(raw_obligations, list):
        raise InstructionContractValidationError(
            "protected_guide_reference_obligations must be a list."
        )

    obligations: list[ProtectedGuideReferenceObligation] = []
    seen_keys: set[tuple[str, str]] = set()
    duplicate_keys: set[tuple[str, str]] = set()
    for raw_obligation in cast(list[object], raw_obligations):
        if not isinstance(raw_obligation, dict):
            raise InstructionContractValidationError(
                "Each protected guide reference obligation must be a mapping."
            )
        raw_obligation = cast(dict[str, object], raw_obligation)

        key = raw_obligation.get("key")
        raw_path = raw_obligation.get("path")
        reference_kind = raw_obligation.get("reference_kind")
        if (
            not isinstance(key, str)
            or not isinstance(raw_path, str)
            or not isinstance(reference_kind, str)
        ):
            raise InstructionContractValidationError(
                "Each protected guide reference obligation must define key, path, "
                "and reference_kind."
            )
        path, is_directory = normalize_repository_path(
            raw_path,
            "protected_guide_reference_obligations[].path",
        )
        if is_directory:
            raise InstructionContractValidationError(
                "protected_guide_reference_obligations[].path must reference a file, "
                f"not a directory: {raw_path}"
            )

        obligation_key = (path, key)
        if obligation_key in seen_keys:
            duplicate_keys.add(obligation_key)
        seen_keys.add(obligation_key)

        target_modules = _required_string_list(raw_obligation, "target_modules")
        if not target_modules:
            raise InstructionContractValidationError(f"{path} target_modules must not be empty.")
        unknown_modules = set(target_modules) - manifest_modules
        if unknown_modules:
            raise InstructionContractValidationError(
                f"{path} protected guide reference obligation {key} references unknown "
                "manifest module(s): " + ", ".join(sorted(unknown_modules))
            )

        target_path = None
        raw_target_path = raw_obligation.get("target_path")
        if raw_target_path is not None:
            if not isinstance(raw_target_path, str):
                raise InstructionContractValidationError(
                    f"{path} protected guide reference obligation {key} target_path "
                    "must be a string."
                )
            target_path, target_is_directory = normalize_repository_path(
                raw_target_path,
                "protected_guide_reference_obligations[].target_path",
            )
            if target_is_directory:
                raise InstructionContractValidationError(
                    "protected_guide_reference_obligations[].target_path must reference "
                    f"a file, not a directory: {raw_target_path}"
                )

        tokens = _required_string_list(raw_obligation, "tokens")
        if reference_kind == "markdown-relative-link" and target_path is None:
            raise InstructionContractValidationError(
                f"{path} protected guide reference obligation {key} must define target_path."
            )
        if reference_kind in {"absolute-url", "prose-reference"} and not tokens:
            raise InstructionContractValidationError(
                f"{path} protected guide reference obligation {key} must define tokens."
            )

        obligations.append(
            ProtectedGuideReferenceObligation(
                key=key,
                path=path,
                reference_kind=reference_kind,
                target_modules=target_modules,
                target_path=target_path,
                tokens=tokens,
            )
        )

    if duplicate_keys:
        formatted_keys = ", ".join(f"({path}, {key})" for path, key in sorted(duplicate_keys))
        raise InstructionContractValidationError(
            "Duplicate protected_guide_reference_obligations (path, key) pair(s): "
            f"{formatted_keys}"
        )
    validate_protected_guide_obligation_keys(contracts_document)
    return tuple(obligations)


def read_instruction_file(repo_root: Path, relative_path: str) -> str | None:
    """Return instruction file text, or ``None`` when the file is absent."""
    path = support.resolve_repo_path(repo_root, relative_path)
    if not path.exists():
        return None
    if not path.is_file():
        raise InstructionContractValidationError(f"{relative_path} is not a regular file.")
    return read_repository_text(path, repo_root, maximum_bytes=MAXIMUM_INPUT_BYTES)


def find_waiver(
    waivers: tuple[InstructionContractWaiver, ...],
    path: str,
    anchor: str,
) -> InstructionContractWaiver | None:
    """Return a waiver that matches ``path`` and ``anchor``, if present."""
    for waiver in waivers:
        if waiver.path == path and waiver.anchor == anchor:
            return waiver
    return None


def find_protected_guide_waiver(
    waivers: tuple[support.ProtectedGuideContractWaiver, ...],
    *,
    path: str,
    contract_key: str,
    target_modules: tuple[str, ...],
    target_path: str | None = None,
) -> support.ProtectedGuideContractWaiver | None:
    """Return a matching protected-guide waiver, if present."""
    target_module_set = set(target_modules)
    for waiver in waivers:
        if waiver.path != path or waiver.contract_key != contract_key:
            continue
        if target_path is not None and waiver.target_path == target_path:
            return waiver
        if waiver.target_module is not None and waiver.target_module in target_module_set:
            return waiver
    return None


def authorized_removal_for(
    protected_decisions: tuple[support.ProtectedFileDecision, ...],
    path: str,
) -> AuthorizedRemoval | None:
    """Return authorized removal details for an absent protected file, if present."""
    for protected_decision in protected_decisions:
        if (
            protected_decision.path == path
            and protected_decision.decision == support.REMOVAL_DECISION
        ):
            return AuthorizedRemoval(
                path=path,
                authorization_basis=protected_decision.authorization_basis or "",
                authorized_scope=protected_decision.authorized_scope or "",
                reason=protected_decision.reason or "",
            )
    return None


def protected_guide_obligation_applies(
    obligation_target_modules: tuple[str, ...],
    included_modules: set[str],
) -> bool:
    """Return whether a protected-guide obligation is stale for retained modules."""
    return set(obligation_target_modules).isdisjoint(included_modules)


def strip_fenced_code_blocks(text: str) -> str:
    """Return ``text`` with CommonMark fenced code blocks removed.

    Required anchors must appear as live Markdown content; example anchors
    nested in a fenced code block should not satisfy the contract.
    """
    return "\n".join(
        line
        for _line_number, line in lines_outside_markdown_fences(
            text,
            fence_context=MARKDOWN_FENCE_CONTEXT,
        )
    )


def heading_is_present(text: str, heading: str) -> bool:
    """Return whether ``heading`` appears as a CommonMark ATX heading line.

    Per CommonMark, an ATX heading may have 0-3 leading spaces of indentation;
    lines with 4+ leading spaces or any leading tab character are indented code
    blocks rather than headings and cannot satisfy the contract.
    """
    for line in markdown_lines(text):
        leading_spaces = 0
        has_leading_tab = False
        for ch in line:
            if ch == " ":
                leading_spaces += 1
            elif ch == "\t":
                has_leading_tab = True
                break
            else:
                break
        if has_leading_tab or leading_spaces > 3:
            continue
        if line[leading_spaces:].rstrip(" \t") == heading:
            return True
    return False


def is_policy_heading(line: str) -> bool:
    """Recognize live ATX boundaries, including an empty heading or a tab separator."""
    return re.match(r"^ {0,3}#{1,6}(?:[ \t]|$)", line) is not None


def is_policy_thematic_break(line: str) -> bool:
    """Recognize a thematic break without promoting four-space paragraph content."""
    return (
        re.fullmatch(r" {0,3}(?:(?:\*[ \t]*){3,}|(?:-[ \t]*){3,}|(?:_[ \t]*){3,})", line)
        is not None
    )


def starts_policy_list(line: str) -> bool:
    """Recognize a list outside a quote, including empty items and arbitrary starts."""
    return re.match(r"^ {0,3}(?:[-+*]|[0-9]{1,9}[.)])(?:[ \t]|$)", line) is not None


def policy_list_content_indent(line: str) -> int | None:
    """Return a top-level item's content margin, including empty/tabbed items.

    Padding beyond four columns starts item code and uses a one-column margin.
    The scanner retains potential nested content; it does not parse that code.
    """
    expanded = line.expandtabs(4)
    marker = re.match(r"^ {0,3}(?:[-+*]|[0-9]{1,9}[.)])(?= |$)", expanded)
    if marker is None or is_policy_thematic_break(line):
        return None
    rest = expanded[marker.end() :]
    padding = len(rest) - len(rest.lstrip(" "))
    return marker.end() + (padding if rest.strip(" \t") and 1 <= padding <= 4 else 1)


def quoted_policy_paragraph(line: str, was_paragraph: bool) -> bool | None:
    """Classify a quoted leaf; None means unsupported syntax must fail closed.

    Paragraphs and simple list paragraphs can continue lazily. Structural leaves
    cannot. Ambiguous HTML, table, or mixed nested containers are not guessed.
    """
    _depth, source_offset = consume_blockquote_prefix(line)
    source_content = line[source_offset:]
    prefix = re.match(r"[ \t]*(?:[-+*]|[0-9]{1,9}[.)])?[ \t]*", source_content)
    assert prefix is not None
    if "\t" in prefix.group():
        # Nested container tab columns need a full block parser. Keep the
        # original ambiguous region visible instead of guessing laziness.
        return None
    expanded = line.expandtabs(4)
    _depth, offset = consume_blockquote_prefix(expanded)
    content = expanded[offset:]
    if is_policy_thematic_break(content):
        return False
    item = LIST_MARKER_RE.match(content)
    if item is not None:
        if item.group("rest").startswith(" "):
            # More than four padding columns begin item code, not a paragraph.
            return False
        content = item.group("rest")
        if consume_blockquote_prefix(content)[0] or starts_policy_list(content):
            return None
        was_paragraph = False
    elif starts_policy_list(content):
        return False if re.fullmatch(r" {0,3}(?:[-+*]|[0-9]{1,9}[.)])[ \t]*", content) else None
    if not content.strip(" \t") or is_policy_heading(content) or is_policy_thematic_break(content):
        return False
    if re.match(r"^(?: {4}| *\t)", content):
        return was_paragraph
    if was_paragraph and re.fullmatch(r" {0,3}(?:=+|-+)[ \t]*", content):
        return False
    if not was_paragraph and content.lstrip(" ").startswith("["):
        # A reference definition can span lines and cannot establish laziness.
        # Keep ambiguous new leaves visible instead of guessing their grammar.
        return None
    if content.lstrip(" ").startswith("<") or "|" in content:
        return None
    return True


def policy_code_span_ends(lines: list[str]) -> dict[tuple[int, int], tuple[int, int]]:
    """Index nearest equal-width backtick runs without rescanning paragraph tails.

    Each line and run is visited once in reverse order. Block starts prevent
    earlier paragraphs from consuming their delimiters; same-line spans remain
    supported. Unmatched runs have no entry and remain literal in the scanner.
    """
    ends: dict[tuple[int, int], tuple[int, int]] = {}
    next_runs: dict[int, tuple[int, int]] = {}
    for row in range(len(lines) - 1, -1, -1):
        line = lines[row]
        for run in reversed(list(re.finditer(r"`+", line))):
            width = run.end() - run.start()
            if width in next_runs:
                ends[(row, run.start())] = next_runs[width]
            # An escaped first backtick leaves the remaining run as an opener.
            # Inside a span, however, only the whole raw run can close it.
            if width > 1 and width - 1 in next_runs:
                ends[(row, run.start() + 1)] = next_runs[width - 1]
            next_runs[width] = (row, run.end())
        if (
            not line.strip(" \t")
            or re.match(
                r"^(?: {4}| *\t| {0,3}(?:>|#{1,6}(?:[ \t]|$)|<!--|[-+*] |[0-9]+[.)] ))", line
            )
            or parse_markdown_fence_open(line, MARKDOWN_FENCE_CONTEXT) is not None
            or policy_html_block_end(line, paragraph_can_continue=True) is not None
            or POLICY_HTML_AMBIGUOUS_START.match(line) is not None
        ):
            next_runs.clear()
    return ends


def policy_html_block_end(line: str, *, paragraph_can_continue: bool) -> re.Pattern[str] | None:
    """Return a non-comment HTML block's end condition, or None for live text.

    Literal tags, processing instructions, declarations and CDATA use explicit
    terminators. Block tags and standalone complete tags end at a blank line.
    Only the standalone complete-tag family cannot interrupt a paragraph.
    """
    if POLICY_HTML_LITERAL_START.match(line):
        return POLICY_HTML_LITERAL_END
    if re.match(r"^ {0,3}<\?", line):
        return re.compile(r"\?>")
    if re.match(r"^ {0,3}<![A-Za-z]", line):
        return re.compile(r">")
    if re.match(r"^ {0,3}<!\[CDATA\[", line):
        return re.compile(r"\]\]>")
    if POLICY_HTML_BLOCK_START.match(line) or (
        not paragraph_can_continue and POLICY_HTML_COMPLETE_TAG.fullmatch(line)
    ):
        return POLICY_HTML_BLANK_END
    return None


def operative_markdown_lines(text: str) -> list[str]:
    """Read the deliberately narrow live policy subset, preserving block boundaries.

    Section contracts accept top-level ATX headings, paragraphs/list items, and
    pipe tables. Fences, block quotes, indented code, and raw HTML blocks cannot
    supply an obligation. Existing loose heading/phrase contracts keep their
    compatibility behavior. This is a static policy guard, not an agent runner.
    """
    lines = markdown_lines(text)
    span_ends = policy_code_span_ends(lines)
    backtick_run = re.compile(r"`+")
    result: list[str] = []
    quoted_paragraph_can_continue = False
    quoted_paragraph_depth = 0
    unsupported_quote = False
    paragraph_can_continue = False
    list_content_indent: int | None = None
    in_comment = False
    active_fence: MarkdownFence | None = None
    active_html_end: re.Pattern[str] | None = None
    code_end: tuple[int, int] | None = None
    for row, line in enumerate(lines):
        continued_comment = in_comment
        if active_html_end is not None:
            if active_html_end is POLICY_HTML_LITERAL_END and POLICY_HTML_TEXTAREA_END.search(line):
                result.append(POLICY_HTML_AMBIGUITY_PREFIX + line)
            if active_html_end.search(line):
                active_html_end = None
            result.append("")
            continue
        if active_fence is not None:
            content = active_fence_content(line, active_fence)
            if content is not None:
                if parse_fence_close_from_content(
                    content,
                    fence_character=active_fence.character,
                    minimum_length=active_fence.length,
                    allow_arbitrary_indent=active_fence.allow_arbitrary_indent,
                ):
                    active_fence = None
                result.append("")
                continue
            active_fence = None
        if not in_comment and code_end is None:
            if POLICY_HTML_AMBIGUOUS_START.match(line):
                # These starts differ between CommonMark 0.31.2 and GFM 0.29.
                # Keep document-level uncertainty visible even outside a section.
                result.append(POLICY_HTML_AMBIGUITY_PREFIX + line)
            expanded = line.expandtabs(4)
            indent = len(expanded) - len(expanded.lstrip(" "))
            quote_depth, quote_offset = consume_blockquote_prefix(expanded)
            quote_content = expanded[quote_offset:]
            if (
                quoted_paragraph_can_continue
                and max(quoted_paragraph_depth, quote_depth) > 1
                and quote_depth != quoted_paragraph_depth
                and quote_content.strip(" ")
                and len(quote_content) - len(quote_content.lstrip(" ")) >= 4
            ):
                # Omitted or changed nested containers can expose an indented
                # block and later live text. Do not guess paragraph laziness.
                unsupported_quote = True
                quoted_paragraph_can_continue = False
                result.append("[unsupported quoted policy] " + line)
                continue
            html_end = policy_html_block_end(
                line,
                paragraph_can_continue=(paragraph_can_continue or quoted_paragraph_can_continue),
            )
            ordered = re.match(r"^ {0,3}([0-9]{1,9})[.)](?:[ \t]|$)", line)
            if (
                paragraph_can_continue
                and list_content_indent is None
                and ordered is not None
                and int(ordered.group(1)) != 1
            ):
                # A non-1 start cannot interrupt a direct paragraph. Do not
                # let a stateless fence parser hide this live continuation.
                result.append("[unsupported ordered continuation] " + line)
                continue
            outside_block = (
                is_policy_heading(line)
                or starts_policy_list(line)
                or is_policy_thematic_break(line)
                or consume_blockquote_prefix(line)[0] > 0
                or re.match(r"^ {0,3}<!--", line) is not None
                or html_end is not None
                or parse_markdown_fence_open(line, MARKDOWN_FENCE_CONTEXT) is not None
            )
            if line.strip(" \t") and list_content_indent is not None:
                if indent >= list_content_indent and (
                    indent >= 4
                    or is_policy_heading(line)
                    or consume_blockquote_prefix(line)[0]
                    or re.match(r"^ {0,3}<!--", line)
                    or html_end is not None
                    or parse_markdown_fence_open(line, MARKDOWN_FENCE_CONTEXT) is not None
                ):
                    result.append("[unsupported nested policy] " + line)
                    continue
                if indent < list_content_indent and (not paragraph_can_continue or outside_block):
                    list_content_indent = None
            if not line.strip(" \t") or outside_block:
                paragraph_can_continue = False
            if not line.strip(" \t") or is_policy_heading(line):
                quoted_paragraph_can_continue = False
                unsupported_quote = False
            if html_end is not None:
                # An unquoted interrupting block ends the preceding quote;
                # its contents must not become live at a later heading.
                quoted_paragraph_can_continue = False
                unsupported_quote = False
            if unsupported_quote:
                # Preserve the entire ambiguous region in the failure identity;
                # none of its lines can masquerade as a required live clause.
                result.append("[unsupported quoted policy] " + line)
                continue
            if not consume_blockquote_prefix(line)[0]:
                item_indent = policy_list_content_indent(line)
                if item_indent is not None and (
                    list_content_indent is None or indent < list_content_indent
                ):
                    list_content_indent = item_indent
                if re.match(r"^ {0,3}(?:[-+*]|[0-9]{1,9}[.)]) {5,}[^ ]", expanded):
                    # More than four visual padding columns begin item code.
                    # Preserve its source instead of normalizing it into policy.
                    paragraph_can_continue = False
                    quoted_paragraph_can_continue = False
                    quoted_paragraph_depth = 0
                    result.append("[unsupported list code] " + line)
                    continue
            active_fence = parse_markdown_fence_open(line, MARKDOWN_FENCE_CONTEXT)
            if active_fence is not None:
                quoted_paragraph_can_continue = False
                result.append("")
                continue
            if consume_blockquote_prefix(line)[0]:
                if not quoted_paragraph_can_continue:
                    quoted_paragraph_depth = 0
                leaf = quoted_policy_paragraph(
                    line,
                    quoted_paragraph_can_continue and quote_depth <= quoted_paragraph_depth,
                )
                unsupported_quote = leaf is None
                quoted_paragraph_can_continue = leaf is True
                if quoted_paragraph_can_continue:
                    quoted_paragraph_depth = max(quoted_paragraph_depth, quote_depth)
                else:
                    quoted_paragraph_depth = 0
                result.append("[unsupported quoted policy] " + line if unsupported_quote else "")
                continue
            if quoted_paragraph_can_continue:
                if (
                    starts_policy_list(line)
                    or is_policy_thematic_break(line)
                    or re.match(r"^ {0,3}<!--", line)
                ):
                    quoted_paragraph_can_continue = False
                elif re.match(r"^ {0,3}<", line) or "|" in line:
                    unsupported_quote = True
                    quoted_paragraph_can_continue = False
                    result.append("[unsupported quoted policy] " + line)
                    continue
                else:
                    result.append("")
                    continue
            if re.match(r"^(?: {4}| *\t)", line):
                result.append(
                    "[unsupported indented policy] " + line if paragraph_can_continue else ""
                )
                continue
            if html_end is not None:
                # A block start wins over inline delimiters on this whole line.
                if html_end is POLICY_HTML_LITERAL_END and POLICY_HTML_TEXTAREA_END.search(line):
                    result.append(POLICY_HTML_AMBIGUITY_PREFIX + line)
                active_html_end = None if html_end.search(line) else html_end
                paragraph_can_continue = False
                quoted_paragraph_can_continue = False
                list_content_indent = None
                result.append("")
                continue
        visible: list[str] = []
        elided_comment = in_comment
        column = 0
        while column < len(line):
            if code_end is not None:
                if row < code_end[0]:
                    visible.append(line[column:])
                    break
                visible.append(line[column : code_end[1]])
                column = code_end[1]
                code_end = None
            elif in_comment:
                end = line.find("-->", column)
                if end == -1:
                    break
                column = end + 3
                in_comment = False
                if line[column:].strip(" \t"):
                    visible.append("[unsupported comment tail] " + line)
                    break
            elif line[column] == "\\" and column + 1 < len(line):
                # Escaped delimiters are literal, including escaped backticks.
                visible.append(line[column : column + 2])
                column += 2
            elif line[column] == "`":
                run = backtick_run.match(line, column)
                assert run is not None
                width = len(run.group())
                code_end = span_ends.get((row, column))
                if code_end is None:
                    visible.append(run.group())
                    column += width
            elif line.startswith("<!--", column):
                end = line.find("-->", column + 2)
                block_comment = column <= 3 and not line[:column].strip(" ")
                if end == -1 and not block_comment:
                    # Inline comments need a complete delimiter. Wrapped forms
                    # are outside this subset and must never hide later policy.
                    visible.append("[unsupported inline comment] " + line[column:])
                    break
                comment_content = line[column + 4 : end]
                ambiguous_inline = end < column + 4 or (
                    comment_content.startswith((">", "->"))
                    or comment_content.endswith("-")
                    or "--" in comment_content
                )
                if end != -1 and not block_comment and ambiguous_inline:
                    # Use the shared CommonMark/GFM grammar. Renderer-specific
                    # inline forms must not hide a potentially visible clause.
                    visible.append("[unsupported inline comment] " + line[column:])
                    break
                if block_comment and end != -1 and line[end + 3 :].strip(" \t"):
                    visible.append("[unsupported comment tail] " + line)
                    break
                elided_comment = True
                if not block_comment and end != -1:
                    # A comment is an inline token boundary, not removable
                    # source text. Keep it before paragraph normalization so
                    # deleting it cannot synthesize links or other markup.
                    visible.append(line[column : end + 3])
                    column = end + 3
                    continue
                in_comment = end == -1
                column = column + 4 if in_comment else end + 3
            else:
                visible.append(line[column])
                column += 1
        observed = "".join(visible).strip(" \t")
        trailing_spaces = re.search(r" {2,}$", line)
        if (
            observed
            and trailing_spaces is not None
            and code_end is None
            and not in_comment
            and not is_policy_heading(line)
            and not observed.startswith("|")
        ):
            # Keep physical endings until the paragraph parser can distinguish
            # a hard break from harmless block-final padding.
            observed += trailing_spaces.group()
        if elided_comment and (
            (is_policy_heading(observed) and not is_policy_heading(line))
            or (starts_policy_list(observed) and not starts_policy_list(line))
            or observed.startswith("|")
        ):
            observed = "[unsupported comment structure] " + line
        result.append(observed)
        if observed:
            paragraph_can_continue = not (
                is_policy_heading(line)
                or is_policy_thematic_break(line)
                or observed.startswith("|")
                or re.match(r"^ {0,3}<!--", line)
                or re.fullmatch(r"(?:[-+*]|[0-9]{1,9}[.)])[ \t]*", observed)
            )
        elif not continued_comment and not in_comment:
            paragraph_can_continue = False
    return result


def claude_container_content(line: str) -> tuple[str, str]:
    """Return a line with Markdown container prefixes blanked and its content."""
    offset = 0
    steps = 0
    while offset < len(line) and steps < MAXIMUM_CLAUDE_CONTAINER_STEPS:
        advanced = False
        _depth, quote_offset = consume_blockquote_prefix(line[offset:])
        if quote_offset:
            offset += quote_offset
            steps += 1
            advanced = True
            if steps == MAXIMUM_CLAUDE_CONTAINER_STEPS:
                break
        list_match = LIST_MARKER_RE.match(line[offset:])
        if list_match is not None:
            offset += list_match.start("rest")
            steps += 1
            advanced = True
        if not advanced:
            break
    return " " * offset + line[offset:], line[offset:]


def claude_import_lines(text: str) -> list[ClaudeImportLine]:
    """Return visible lines with literals blanked and paragraph identities assigned."""
    lines: list[ClaudeImportLine] = []
    in_block_comment = False
    active_fence: MarkdownFence | None = None
    for line in markdown_lines(text):
        normalized, content = claude_container_content(line)
        if in_block_comment:
            comment_end = content.find("-->")
            if comment_end == -1:
                lines.append(ClaudeImportLine(source="", visible=""))
                continue
            in_block_comment = False
            tail_start = len(line) - len(content) + comment_end + 3
            visible = " " * tail_start + line[tail_start:]
            lines.append(ClaudeImportLine(source=visible, visible=visible))
            continue
        if active_fence is not None:
            fence_content = active_fence_content(line, active_fence)
            if fence_content is not None:
                if parse_fence_close_from_content(
                    fence_content,
                    fence_character=active_fence.character,
                    minimum_length=active_fence.length,
                    allow_arbitrary_indent=active_fence.allow_arbitrary_indent,
                ):
                    active_fence = None
                lines.append(ClaudeImportLine(source="", visible=""))
                continue
            active_fence = None
        comment_start = re.match(r"^ {0,3}<!--", content)
        if comment_start is not None:
            comment_end = content.find("-->", comment_start.end())
            if comment_end == -1:
                in_block_comment = True
                lines.append(ClaudeImportLine(source="", visible=""))
                continue
            tail_start = len(line) - len(content) + comment_end + 3
            visible = " " * tail_start + line[tail_start:]
            lines.append(ClaudeImportLine(source=visible, visible=visible))
            continue
        active_fence = parse_markdown_fence_open(line, MARKDOWN_FENCE_CONTEXT)
        if active_fence is not None:
            lines.append(ClaudeImportLine(source="", visible=""))
            continue
        lines.append(ClaudeImportLine(source=line, visible=normalized))
    return classify_claude_paragraphs(lines)


def claude_paragraph_text(content: str, *, paragraph_active: bool = False) -> bool:
    """Return whether content can continue an active Markdown paragraph."""
    if not content.strip(" \t"):
        return False
    if (
        is_policy_heading(content)
        or is_policy_thematic_break(content)
        or starts_policy_list(content)
        or parse_markdown_fence_open(content, MARKDOWN_FENCE_CONTEXT) is not None
        or policy_html_block_end(content, paragraph_can_continue=paragraph_active) is not None
        or POLICY_HTML_AMBIGUOUS_START.match(content) is not None
        or re.fullmatch(r" {0,3}(?:=+|-+)[ \t]*", content) is not None
        or re.match(r"^ {0,3}\[[^\]]+\]:", content) is not None
    ):
        return False
    return paragraph_active or re.match(r"^(?: {4}| *\t)", content) is None


ClaudeContainerStep = tuple[str, int]


@dataclass(frozen=True)
class ClaudeContinuation:
    """One active paragraph continuation resolved against its container path."""

    content_offset: int
    ordered_marker_is_text: bool = False


def claude_explicit_container(
    source: str,
) -> tuple[tuple[ClaudeContainerStep, ...], int, bool]:
    """Return the ordered visible container path, content offset, and overflow."""
    steps: list[ClaudeContainerStep] = []
    offset = 0
    while len(steps) < MAXIMUM_CLAUDE_CONTAINER_STEPS:
        quote_depth, quote_offset = consume_blockquote_prefix(source[offset:])
        if quote_depth:
            steps.append(("quote", quote_depth))
            offset += quote_offset
            continue
        list_match = LIST_MARKER_RE.match(source[offset:])
        if list_match is not None and not is_policy_thematic_break(source[offset:]):
            list_indent = list_match.start("rest")
            steps.append(("list", list_indent))
            offset += list_indent
            continue
        break
    quote_depth, _quote_offset = consume_blockquote_prefix(source[offset:])
    list_match = LIST_MARKER_RE.match(source[offset:])
    overflow = len(steps) == MAXIMUM_CLAUDE_CONTAINER_STEPS and (
        quote_depth > 0
        or (list_match is not None and not is_policy_thematic_break(source[offset:]))
    )
    return tuple(steps), offset, overflow


def claude_continuation_content_offset(
    source: str,
    active_steps: tuple[ClaudeContainerStep, ...],
    explicit_prefix_end: int,
) -> ClaudeContinuation | None:
    """Consume explicit or lazily omitted markers for one active container path."""
    offset = 0
    omitted_container_prefix = False
    omitted_list_indentation = False
    for kind, amount in active_steps:
        if kind == "quote":
            depth, consumed = consume_blockquote_prefix(source[offset:], max_depth=amount)
            # A quote reached only after omitting its owning list indentation is
            # an ancestor quote that interrupts the old list-item paragraph.
            if consumed and omitted_list_indentation:
                return None
            offset += consumed
            omitted_container_prefix = omitted_container_prefix or depth < amount
            continue
        remaining = source[offset:]
        leading_spaces = len(remaining) - len(remaining.lstrip(" "))
        if leading_spaces >= amount:
            offset += amount
        else:
            omitted_container_prefix = True
            omitted_list_indentation = True

    # Ordered markers whose start is not 1 cannot interrupt a paragraph at the
    # deepest container that stayed open. They are still boundaries after an
    # active list indentation was omitted, where they are sibling/ancestor
    # items rather than paragraph text.
    relative_list = LIST_MARKER_RE.match(source[offset:])
    if relative_list is not None and not omitted_container_prefix:
        marker = relative_list.group("marker")
        if marker[0].isdigit() and int(marker[:-1]) != 1:
            return ClaudeContinuation(offset, ordered_marker_is_text=True)
    # Any explicit marker left unconsumed starts a different block container.
    return ClaudeContinuation(offset) if offset >= explicit_prefix_end else None


def classify_claude_paragraphs(lines: list[ClaudeImportLine]) -> list[ClaudeImportLine]:
    """Assign bounded paragraph identities from raw list and quote structure.

    The identity is intentionally narrower than a complete CommonMark AST. It
    recognizes the explicit and lazy paragraph continuations that can carry a
    multi-line code span. Other block transitions start a fresh identity, so an
    unmatched delimiter cannot hide a later import across a structural boundary.
    """
    classified: list[ClaudeImportLine] = []
    next_paragraph_id = 0
    active_paragraph_id: int | None = None
    active_steps: tuple[ClaudeContainerStep, ...] = ()
    for line in lines:
        source = line.source
        if not source.strip(" \t"):
            classified.append(line)
            active_paragraph_id = None
            active_steps = ()
            continue

        explicit_steps, explicit_prefix_end, overflow = claude_explicit_container(source)
        continuation = None
        if active_paragraph_id is not None and not overflow:
            continuation = claude_continuation_content_offset(
                source, active_steps, explicit_prefix_end
            )
        same_container = continuation is not None
        content_offset = continuation.content_offset if continuation else explicit_prefix_end
        content = source[content_offset:]
        can_continue = not overflow and (
            (continuation is not None and continuation.ordered_marker_is_text)
            or claude_paragraph_text(content, paragraph_active=same_container)
        )
        if same_container and can_continue:
            paragraph_id = active_paragraph_id
        else:
            paragraph_id = next_paragraph_id
            next_paragraph_id += 1
        classified.append(replace(line, paragraph_id=paragraph_id))

        if same_container and can_continue:
            active_paragraph_id = paragraph_id
        elif can_continue:
            active_paragraph_id = paragraph_id
            active_steps = explicit_steps if not overflow else ()
        else:
            active_paragraph_id = None
            active_steps = ()
    return classified


def mask_matched_claude_code_spans(lines: list[ClaudeImportLine]) -> list[str]:
    """Blank spans matched within one paragraph; leave unmatched runs live."""
    span_ends: dict[tuple[int, int], tuple[int, int]] = {}
    next_runs: dict[tuple[int, int], tuple[int, int]] = {}
    for row in range(len(lines) - 1, -1, -1):
        line = lines[row]
        if line.paragraph_id is None:
            continue
        for run in reversed(list(re.finditer(r"`+", line.visible))):
            width = run.end() - run.start()
            key = (line.paragraph_id, width)
            if key in next_runs:
                span_ends[(row, run.start())] = next_runs[key]
            escaped_key = (line.paragraph_id, width - 1)
            if width > 1 and escaped_key in next_runs:
                span_ends[(row, run.start() + 1)] = next_runs[escaped_key]
            next_runs[key] = (row, run.end())
    offsets: list[int] = []
    total = 0
    for line in lines:
        offsets.append(total)
        total += len(line.visible) + 1
    flat = "\n".join(line.visible for line in lines)
    absolute_ends = {
        offsets[start_row] + start_column: offsets[end_row] + end_column
        for (start_row, start_column), (end_row, end_column) in span_ends.items()
    }
    masked = list(flat)
    position = 0
    while position < len(flat):
        if flat[position] == "\\" and position + 1 < len(flat):
            position += 2
            continue
        if flat[position] != "`":
            position += 1
            continue
        run_end = position + 1
        while run_end < len(flat) and flat[run_end] == "`":
            run_end += 1
        span_end = absolute_ends.get(position)
        if span_end is None:
            position = run_end
            continue
        for index in range(position, span_end):
            if masked[index] != "\n":
                masked[index] = " "
        position = span_end
    return "".join(masked).split("\n")


def active_claude_imports(path: str, text: str) -> tuple[ActiveClaudeImport, ...]:
    """Return active Claude imports from one retained governed document."""
    imports: list[ActiveClaudeImport] = []
    lines = mask_matched_claude_code_spans(claude_import_lines(text))
    for line_number, line in enumerate(lines, 1):
        for match in CLAUDE_IMPORT_PATTERN.finditer(line):
            imports.append(
                ActiveClaudeImport(
                    path=path,
                    line=line_number,
                    target=match.group("target"),
                )
            )
    return tuple(imports)


def sanitized_git_environment() -> dict[str, str]:
    """Return an environment without caller-controlled Git repository selectors."""
    exact_keys = {
        "GIT_ALTERNATE_OBJECT_DIRECTORIES",
        "GIT_CEILING_DIRECTORIES",
        "GIT_COMMON_DIR",
        "GIT_CONFIG",
        "GIT_DIR",
        "GIT_DISCOVERY_ACROSS_FILESYSTEM",
        "GIT_EXEC_PATH",
        "GIT_INDEX_FILE",
        "GIT_LITERAL_PATHSPECS",
        "GIT_GLOB_PATHSPECS",
        "GIT_NOGLOB_PATHSPECS",
        "GIT_ICASE_PATHSPECS",
        "GIT_NAMESPACE",
        "GIT_OBJECT_DIRECTORY",
        "GIT_PREFIX",
        "GIT_WORK_TREE",
    }
    environment = {
        key: value
        for key, value in os.environ.items()
        if key.upper() not in exact_keys and not key.upper().startswith("GIT_CONFIG_")
    }
    environment["GIT_CONFIG_GLOBAL"] = os.devnull
    environment["GIT_CONFIG_NOSYSTEM"] = "1"
    environment["GIT_CONFIG_SYSTEM"] = os.devnull
    environment["GIT_OPTIONAL_LOCKS"] = "0"
    environment["GIT_PAGER"] = "cat"
    return environment


def os_error_diagnostic(error: OSError) -> str:
    """Return an actionable OSError cause without either filename field."""
    return f"{type(error).__name__}: {error.strerror or 'I/O error'}"


def run_bounded_git(
    repo_root: Path, arguments: list[str], *, allowed_returncodes: tuple[int, ...] = (0,)
) -> bytes:
    """Run one shell-free Git query with bounded captured output."""
    try:
        process = subprocess.Popen(
            ["git", "-C", str(repo_root), *arguments],
            env=sanitized_git_environment(),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except FileNotFoundError as error:
        raise InstructionContractValidationError(
            "Git is required to inspect tracked Claude local memory for this worktree."
        ) from error
    except OSError as error:
        raise InstructionContractValidationError(
            f"Git inventory could not start: {os_error_diagnostic(error)}."
        ) from error

    stdout = bytearray()
    stderr = bytearray()
    output_exceeded = threading.Event()
    reader_errors: list[OSError] = []
    reader_error_lock = threading.Lock()

    def read_stream(stream: Any, destination: bytearray) -> None:
        try:
            while True:
                chunk = stream.read(64 * 1024)
                if not chunk:
                    return
                remaining = MAXIMUM_GIT_OUTPUT_BYTES + 1 - len(destination)
                destination.extend(chunk[:remaining])
                if len(destination) > MAXIMUM_GIT_OUTPUT_BYTES:
                    output_exceeded.set()
                    try:
                        process.kill()
                    except OSError:
                        pass
                    return
        except OSError as error:
            with reader_error_lock:
                reader_errors.append(error)
        finally:
            try:
                stream.close()
            except OSError as error:
                with reader_error_lock:
                    reader_errors.append(error)

    assert process.stdout is not None
    assert process.stderr is not None
    stdout_stream = process.stdout
    stderr_stream = process.stderr
    stdout_thread = threading.Thread(target=read_stream, args=(stdout_stream, stdout), daemon=True)
    stderr_thread = threading.Thread(target=read_stream, args=(stderr_stream, stderr), daemon=True)
    stdout_thread.start()
    stderr_thread.start()
    deadline = time.monotonic() + GIT_TIMEOUT_SECONDS

    try:
        process.wait(timeout=max(0.0, deadline - time.monotonic()))
    except subprocess.TimeoutExpired as error:
        try:
            process.kill()
        except OSError:
            pass
        try:
            process.wait(timeout=max(0.0, deadline - time.monotonic()))
        except subprocess.TimeoutExpired:
            pass
        raise InstructionContractValidationError(
            f"Git inventory timed out after {GIT_TIMEOUT_SECONDS} seconds."
        ) from error
    for thread in (stdout_thread, stderr_thread):
        thread.join(timeout=max(0.0, deadline - time.monotonic()))
    if stdout_thread.is_alive() or stderr_thread.is_alive():
        raise InstructionContractValidationError(
            f"Git inventory streams did not close within {GIT_TIMEOUT_SECONDS} seconds."
        )
    if reader_errors:
        reader_error = reader_errors[0]
        raise InstructionContractValidationError(
            f"Git inventory output could not be read: {os_error_diagnostic(reader_error)}."
        ) from reader_error
    if output_exceeded.is_set():
        raise InstructionContractValidationError(
            "Git inventory output exceeded the one-mebibyte per-stream limit."
        )
    if process.returncode not in allowed_returncodes:
        try:
            detail = bytes(stderr).decode("utf-8", errors="strict").strip()
        except UnicodeDecodeError as error:
            raise InstructionContractValidationError(
                "Git inventory failed and emitted non-UTF-8 diagnostics."
            ) from error
        if detail:
            detail = detail.splitlines()[-1][:500]
            raise InstructionContractValidationError(f"Git inventory failed: {detail}")
        raise InstructionContractValidationError(
            f"Git inventory failed with exit code {process.returncode}."
        )
    return bytes(stdout)


def decode_git_output(output: bytes, purpose: str) -> str:
    """Decode trusted-size Git output or fail closed on malformed bytes."""
    try:
        return output.decode("utf-8", errors="strict")
    except UnicodeDecodeError as error:
        raise InstructionContractValidationError(
            f"Git {purpose} output is not valid UTF-8."
        ) from error


def git_inventory_safe_prefix(repo_root: Path) -> list[str]:
    """Prevent repository-configured fsmonitor callbacks before reading the index."""
    version_text = decode_git_output(run_bounded_git(repo_root, ["--version"]), "version")
    match = re.match(r"^git version ([0-9]+)\.([0-9]+)", version_text.strip())
    if match is not None and tuple(map(int, match.groups())) >= (2, 36):
        return ["-c", "core.fsmonitor=false"]
    # Git <=2.35.1 treats even 'false' as a hook pathname, not a Boolean.
    # NUL termination distinguishes a configured empty value from an absent key.
    configured = run_bounded_git(
        repo_root, ["config", "-z", "--get-all", "core.fsmonitor"], allowed_returncodes=(0, 1)
    )
    if configured:
        raise InstructionContractValidationError(
            "Safe Git inventory requires Git detected as >=2.36 or an unconfigured "
            "core.fsmonitor setting for the validation root."
        )
    return []


def tracked_claude_local_memory(
    repo_root: Path,
) -> tuple[tuple[TrackedClaudeLocalMemory, ...], bool]:
    """Return matching Git-index paths and whether this root has its own worktree."""
    git_marker = repo_root / ".git"
    try:
        marker_status = git_marker.lstat()
    except FileNotFoundError:
        return (), False
    except OSError as error:
        raise InstructionContractValidationError(
            f"Cannot inspect the repository Git marker: {os_error_diagnostic(error)}."
        ) from error
    if not (stat.S_ISDIR(marker_status.st_mode) or stat.S_ISREG(marker_status.st_mode)):
        raise InstructionContractValidationError(
            "The repository .git marker is neither a directory nor a regular worktree file."
        )

    safe_prefix = git_inventory_safe_prefix(repo_root)
    top_level_output = run_bounded_git(repo_root, [*safe_prefix, "rev-parse", "--show-toplevel"])
    top_level_text = decode_git_output(top_level_output, "top-level discovery").rstrip("\r\n")
    if (
        not top_level_text
        or "\n" in top_level_text
        or "\r" in top_level_text
        or "\0" in top_level_text
    ):
        raise InstructionContractValidationError(
            "Git top-level discovery returned malformed output."
        )
    try:
        observed_top_level = Path(top_level_text).resolve(strict=True)
        expected_top_level = repo_root.resolve(strict=True)
    except OSError as error:
        raise InstructionContractValidationError(
            f"Cannot canonicalize the Git top level: {os_error_diagnostic(error)}."
        ) from error
    if observed_top_level != expected_top_level:
        raise InstructionContractValidationError(
            "The present root .git marker did not resolve to the validation root."
        )

    index_output = run_bounded_git(
        repo_root,
        [
            *safe_prefix,
            "ls-files",
            "-z",
            "--cached",
            "--",
            ":(icase)CLAUDE.local.md",
            ":(glob,icase)**/CLAUDE.local.md",
        ],
    )
    if index_output and not index_output.endswith(b"\0"):
        raise InstructionContractValidationError("Git index inventory returned malformed output.")
    inventory = decode_git_output(index_output, "index inventory")
    matches: set[str] = set()
    for path in inventory.split("\0") if inventory else ():
        if not path:
            continue
        parts = path.split("/")
        if any(part in {"", ".", ".."} for part in parts):
            raise InstructionContractValidationError("Git index inventory returned an unsafe path.")
        if parts[-1].casefold() == "claude.local.md":
            matches.add(path)
    return tuple(TrackedClaudeLocalMemory(path=path) for path in sorted(matches)), True


def index_policy_lines(lines: list[str]) -> PolicyLineIndex:
    """Index exact line occurrences and inclusive next-heading offsets once."""
    positions: dict[str, list[int]] = {}
    next_headings = [len(lines)] * (len(lines) + 1)
    for offset in range(len(lines) - 1, -1, -1):
        line = lines[offset]
        positions.setdefault(line, []).append(offset)
        next_headings[offset] = offset if is_policy_heading(line) else next_headings[offset + 1]
    for offsets in positions.values():
        offsets.reverse()
    return PolicyLineIndex(positions, next_headings)


def section_body(
    lines: list[str], heading: str, index: PolicyLineIndex | None = None
) -> list[str] | None:
    """Return one unique heading's direct body, ending at the next ATX heading."""
    if index is None:
        index = index_policy_lines(lines)
    matches = index.positions.get(heading, [])
    if len(matches) != 1:
        return None
    start = matches[0] + 1
    return lines[start : index.next_headings[start]]


def policy_inventory_digest(expected: object, observed: object) -> str:
    """Bind a waiver to both the expected contract and its observed deviation."""
    return hashlib.sha256(
        json.dumps([expected, observed], ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    ).hexdigest()


def section_boundary_failure(
    lines: list[str], section: RequiredSection, index: PolicyLineIndex | None = None
) -> str | None:
    """Reject unknown section boundaries without claiming ownership beyond the declared end.

    Non-null successors must be unique. The mismatch identity includes all live
    content up to the declared successor (or EOF), so waiving one added heading
    cannot silently authorize changed content below that heading.
    """
    if index is None:
        index = index_policy_lines(lines)
    starts = index.positions.get(section.heading, [])
    start = starts[0] + 1 if len(starts) == 1 else 0
    successors = (
        index.positions.get(section.next_heading, []) if section.next_heading is not None else []
    )
    next_offset = index.next_headings[start]
    actual_next = lines[next_offset] if next_offset < len(lines) else None
    unique_successor = len(starts) == 1 and (
        section.next_heading is None or (len(successors) == 1 and successors[0] >= start)
    )
    if actual_next != section.next_heading or not unique_successor:
        end = successors[0] if len(successors) == 1 and successors[0] >= start else len(lines)
        observed = [normalize_policy_paragraph(line) for line in lines[start:end] if line]
        identity = policy_inventory_digest(
            section.next_heading, [actual_next, len(successors), observed]
        )
        return f"section:{section.heading}:boundary:{identity}"
    return None


def parse_policy_body(
    lines: list[str],
    *,
    blocks: list[PolicyBlock] | None = None,
    hard_breaks: list[str] | None = None,
) -> tuple[list[str], list[RequiredTable]]:
    """Parse full paragraph/list-item clauses and strict, unescaped pipe tables.

    Inline pipes and multiline table cells are outside this policy subset.
    Malformed tables raise rather than quietly dropping decision rows.
    """
    paragraphs: list[str] = []
    tables: list[RequiredTable] = []
    paragraph: list[str] = []

    def finish_paragraph() -> None:
        if not paragraph:
            return
        value = normalize_policy_paragraph(" ".join(paragraph))
        paragraphs.append(value)
        if blocks is not None:
            blocks.append(("paragraph", value))
        if hard_breaks is not None and any(
            line.endswith("  ")
            and not is_policy_thematic_break(line)
            and not is_policy_thematic_break(following)
            for line, following in pairwise(paragraph)
        ):
            hard_breaks.append("\n".join(paragraph))
        paragraph.clear()

    index = 0
    while index < len(lines):
        line = lines[index]
        boundary = (
            not line or is_policy_heading(line) or line.startswith("|") or starts_policy_list(line)
        )
        if boundary and paragraph:
            finish_paragraph()
        if line.startswith("|"):
            cells: list[tuple[str, ...]] = []
            while index < len(lines) and lines[index].startswith("|"):
                row = lines[index]
                if not row.endswith("|") or "\\|" in row:
                    raise InstructionContractValidationError("Malformed policy table row.")
                row_cells = tuple(cell.strip(" \t") for cell in row[1:-1].split("|"))
                if any(POLICY_CELL_PATTERN.fullmatch(cell) is None for cell in row_cells):
                    raise InstructionContractValidationError("Noncanonical observed policy cell.")
                cells.append(row_cells)
                index += 1
            if (
                len(cells) < 3
                or len(cells[0]) < 2
                or any(len(row) != len(cells[0]) for row in cells)
                or any(not re.fullmatch(r":?-{3,}:?", cell) for cell in cells[1])
                or any(not cell for row in (cells[0], *cells[2:]) for cell in row)
                or len({row[0] for row in cells[2:]}) != len(cells[2:])
            ):
                raise InstructionContractValidationError("Malformed or duplicate policy table.")
            table = RequiredTable(cells[0], tuple(cells[2:]))
            tables.append(table)
            if blocks is not None:
                blocks.append(("table", table))
            continue
        if line and not is_policy_heading(line):
            paragraph.append(line)
        index += 1
    finish_paragraph()
    return paragraphs, tables


def expected_policy_blocks(section: RequiredSection) -> list[PolicyBlock]:
    """Expand explicit or default interleaving without duplicating clause content."""
    order = section.required_block_order or (
        ("paragraph",) * len(section.required_paragraphs)
        + ("table",) * len(section.required_tables)
    )
    paragraphs = iter(normalize_policy_paragraph(value) for value in section.required_paragraphs)
    tables = iter(section.required_tables)
    return [(kind, next(paragraphs) if kind == "paragraph" else next(tables)) for kind in order]


def serialized_policy_blocks(blocks: list[PolicyBlock]) -> list[object]:
    """Keep typed source inventories serializable in content-bound waiver identities."""
    return [
        [kind, value.headers, value.rows] if isinstance(value, RequiredTable) else [kind, value]
        for kind, value in blocks
    ]


def section_applies(section: RequiredSection, included_modules: set[str] | None) -> bool:
    """Apply all sections upstream and only retained section module sets downstream."""
    return included_modules is None or set(section.requires_modules).issubset(included_modules)


def applicable_section_boundary(
    section: RequiredSection,
    sections: dict[str, RequiredSection],
    live_lines: set[str],
    included_modules: set[str] | None,
    cache: dict[str, str | None] | None = None,
) -> RequiredSection:
    """Skip only absent, inapplicable declared successors in an acyclic catalog.

    A present optional heading remains the exact boundary, including when the
    separate protected-guide checks require an explicit stale-section waiver.
    Applicable or uncontracted successors cannot be skipped.
    """
    successor = section.next_heading
    if cache is None:
        cache = {}
    skipped: list[str] = []
    while successor in sections and successor not in live_lines:
        assert successor is not None
        if successor in cache:
            successor = cache[successor]
            break
        following = sections[successor]
        if section_applies(following, included_modules):
            break
        skipped.append(successor)
        successor = following.next_heading
    for heading in skipped:
        cache[heading] = successor
    return replace(section, next_heading=successor)


def section_failures(
    text: str,
    sections: tuple[RequiredSection, ...],
    included_modules: set[str] | None = None,
) -> list[str]:
    """Return stable, individually waivable failures for scoped policy contracts."""
    failures: list[str] = []
    lines = operative_markdown_lines(text)
    ambiguous_html = [line for line in lines if line.startswith(POLICY_HTML_AMBIGUITY_PREFIX)]
    document_identity = hashlib.sha256(text.encode("utf-8")).hexdigest() if ambiguous_html else ""
    sections_by_heading = {section.heading: section for section in sections}
    live_lines = set(lines)
    line_index = index_policy_lines(lines)
    successor_cache: dict[str, str | None] = {}
    for section in sections:
        if not section_applies(section, included_modules):
            continue
        prefix = f"section:{section.heading}"
        expected_blocks = expected_policy_blocks(section)
        if ambiguous_html:
            expected_section = (
                section.heading,
                section.next_heading,
                section.required_paragraphs,
                [(table.headers, table.rows) for table in section.required_tables],
                section.required_block_order,
            )
            identity = policy_inventory_digest(
                expected_section, [ambiguous_html, document_identity]
            )
            failures.append(f"{prefix}:html-grammar:{identity}")
        boundary = applicable_section_boundary(
            section, sections_by_heading, live_lines, included_modules, successor_cache
        )
        boundary_failure = section_boundary_failure(lines, boundary, line_index)
        if boundary_failure is not None:
            failures.append(boundary_failure)
        body = section_body(lines, section.heading, line_index)
        if body is None:
            failures.append(prefix)
            body = []
        hard_breaks: list[str] = []
        paragraphs, _ = parse_policy_body(
            ["" if line.startswith("|") else line for line in body],
            hard_breaks=hard_breaks,
        )
        if hard_breaks:
            identity = policy_inventory_digest(
                serialized_policy_blocks(expected_blocks), hard_breaks
            )
            failures.append(f"{prefix}:hard-break:{identity}")
        malformed = False
        observed_blocks: list[PolicyBlock] = []
        observed_tables: object
        try:
            _, tables = parse_policy_body(body, blocks=observed_blocks)
        except InstructionContractValidationError:
            malformed = True
            tables = []
            observed_tables = {
                "malformed": [line for line in body if line.startswith("|")],
                "body": body,
            }
        else:
            observed_tables = {
                "parsed": [(table.headers, table.rows) for table in tables],
                "blocks": serialized_policy_blocks(observed_blocks),
            }
        table_identity = policy_inventory_digest(
            serialized_policy_blocks(expected_blocks), observed_tables
        )
        table_anchor = f"{prefix}:tables:{table_identity}"
        # Ordered, complete paragraphs prevent scattered keywords or reordered
        # decision steps from substituting for the operative contract.
        cursor = 0
        paragraph_offsets: dict[str, list[int]] = {}
        for offset, paragraph in enumerate(paragraphs):
            paragraph_offsets.setdefault(paragraph, []).append(offset)
        for paragraph in section.required_paragraphs:
            expected = normalize_policy_paragraph(paragraph)
            offsets = paragraph_offsets.get(expected, [])
            offset_index = bisect_left(offsets, cursor)
            if offset_index < len(offsets):
                cursor = offsets[offset_index] + 1
            else:
                identity = hashlib.sha256(expected.encode("utf-8")).hexdigest()
                failures.append(f"{prefix}:paragraph:{identity}")
        expected_paragraphs = [
            normalize_policy_paragraph(value) for value in section.required_paragraphs
        ]
        expected_positions: dict[str, int] = {}
        for offset, value in enumerate(expected_paragraphs):
            expected_positions.setdefault(value, offset)
        positions = [
            expected_positions[value] for value in paragraphs if value in expected_positions
        ]
        if any(value not in expected_positions for value in paragraphs) or positions != sorted(
            set(positions)
        ):
            identity = policy_inventory_digest(
                serialized_policy_blocks(expected_blocks), observed_tables
            )
            failures.append(f"{prefix}:paragraphs:{identity}")
        if malformed or tables != list(section.required_tables):
            failures.append(table_anchor)
        if section.required_paragraphs and section.required_tables and not malformed:
            expected_ranks = {block: offset for offset, block in enumerate(expected_blocks)}
            observed_ranks = [
                expected_ranks[block] for block in observed_blocks if block in expected_ranks
            ]
            if observed_ranks != sorted(observed_ranks):
                identity = policy_inventory_digest(
                    serialized_policy_blocks(expected_blocks),
                    serialized_policy_blocks(observed_blocks),
                )
                failures.append(f"{prefix}:blocks:{identity}")
    return list(dict.fromkeys(failures))


MARKDOWN_LINK_LIST_MARKER = re.compile(r"(?:[-+*]|[0-9]{1,9}[.)])(?= |$)")


def markdown_html_token_ends(text: str) -> dict[int, int]:
    """Index raw HTML tokens without reparsing shared malformed tag suffixes.

    This lexical index does not remove source. The inline scanner consults it
    only outside accepted Markdown components and matched code spans.
    """
    positions = {
        token: [match.start() for match in re.finditer(re.escape(token), text)]
        for token in ('"', "'", ">", "-->", "?>", "]]>")
    }
    tails: dict[int, int] = {}
    space = re.compile(r"[ \t\n\f\r]+")
    attribute = re.compile(r"[A-Za-z_:][A-Za-z0-9_.:-]*")
    unquoted = re.compile(r"[^ \t\n\f\r\"'=<>`]+")

    def following(token: str, start: int) -> int:
        """Find a terminator using its once-built ordered index."""
        offsets = positions[token]
        index = bisect_left(offsets, start)
        return offsets[index] + len(token) if index < len(offsets) else -1

    def tag_tail(start: int) -> int:
        """Memoize deterministic attribute transitions, including invalid tails."""
        trail: list[int] = []
        index = start
        end = -1
        while index not in tails:
            trail.append(index)
            if text.startswith(">", index) or text.startswith("/>", index):
                end = index + (1 if text[index] == ">" else 2)
                break
            spacing = space.match(text, index)
            if spacing is None:
                break
            index = spacing.end()
            if text.startswith(">", index) or text.startswith("/>", index):
                end = index + (1 if text[index] == ">" else 2)
                break
            name = attribute.match(text, index)
            if name is None:
                break
            index = name.end()
            spacing = space.match(text, index)
            equals = spacing.end() if spacing is not None else index
            if text[equals : equals + 1] != "=":
                continue
            index = equals + 1
            spacing = space.match(text, index)
            if spacing is not None:
                index = spacing.end()
            if index < len(text) and text[index] in "\"'":
                index = following(text[index], index + 1)
                if index < 0:
                    break
            else:
                value = unquoted.match(text, index)
                if value is None:
                    break
                index = value.end()
        else:
            end = tails[index]
        for offset in trail:
            tails[offset] = end
        return end

    ends: dict[int, int] = {}
    tag = re.compile(r"</?[A-Za-z][A-Za-z0-9-]*")
    for opening in re.finditer("<", text):
        start = opening.start()
        end = -1
        if text.startswith("<!--", start):
            if text.startswith("<!-->", start):
                end = start + 5
            elif text.startswith("<!--->", start):
                end = start + 6
            else:
                end = following("-->", start + 4)
        elif text.startswith("<?", start):
            end = following("?>", start + 2)
        elif text.startswith("<![CDATA[", start):
            end = following("]]>", start + 9)
        elif re.match(r"<![A-Za-z]", text[start : start + 3]):
            end = following(">", start + 3)
        else:
            name = tag.match(text, start)
            if name is not None:
                if text.startswith("</", start):
                    spacing = space.match(text, name.end())
                    index = spacing.end() if spacing is not None else name.end()
                    end = index + 1 if text[index : index + 1] == ">" else -1
                else:
                    end = tag_tail(name.end())
        if end >= 0:
            ends[start] = end
    return ends


def markdown_link_html_block_end(content: str, *, paragraph_active: bool) -> re.Pattern[str] | None:
    """Recognize the seven raw block families without changing policy grammar."""
    if re.match(r"^ {0,3}<!--", content):
        return re.compile(r"-->")
    if POLICY_HTML_LITERAL_START.match(content):
        return POLICY_HTML_LITERAL_END
    for opening, closing in ((r"<\?", r"\?>"), (r"<![A-Za-z]", ">"), (r"<!\[CDATA\[", r"\]\]>")):
        if re.match(r"^ {0,3}" + opening, content):
            return re.compile(closing)
    if POLICY_HTML_BLOCK_START.match(content):
        return POLICY_HTML_BLANK_END
    start = markdown_prefix_end(content, 0)
    if paragraph_active or start > 3 or content[start : start + 1] != "<":
        return None
    name = re.match(r"</?([A-Za-z][A-Za-z0-9-]*)", content[start:])
    if name is None or (
        not content.startswith("</", start)
        and name.group(1).lower() in {"pre", "script", "style", "textarea"}
    ):
        return None
    end = markdown_html_token_ends(content).get(start)
    if end is not None and not content[end:].strip(" \t"):
        return POLICY_HTML_BLANK_END
    return None


def markdown_thematic_suffix(line: str) -> tuple[int, int]:
    """Index a possible final thematic run once, avoiding nested suffix rescans."""
    index = len(line) - 1
    while index >= 0 and line[index] == " ":
        index -= 1
    if index < 0 or line[index] not in "*_-":
        return len(line), -1
    marker = line[index]
    count = 0
    third_last = -1
    while index >= 0 and line[index] in (marker, " "):
        if line[index] == marker:
            count += 1
            if count == 3:
                third_last = index
        index -= 1
    return index + 1, third_last


def markdown_prefix_end(line: str, offset: int) -> int:
    """Count structural spaces without slicing or rescanning the remaining line."""
    while offset < len(line) and line[offset] == " ":
        offset += 1
    return offset


def markdown_link_container(
    line: str, offset: int, *, paragraph_active: bool = False
) -> tuple[str, int] | None:
    """Recognize one quote or list prefix in a column-expanded structural view."""
    start = markdown_prefix_end(line, offset)
    if start - offset > 3 or start >= len(line):
        return None
    if line[start] == ">":
        end = start + 1
        return "quote", end + (end < len(line) and line[end] == " ")
    marker = MARKDOWN_LINK_LIST_MARKER.match(line, start)
    if marker is None:
        return None
    end = marker.end()
    content = markdown_prefix_end(line, end)
    if paragraph_active and (
        content == len(line) or (marker.group()[0].isdigit() and int(marker.group()[:-1]) != 1)
    ):
        return None
    padding = content - end
    return "list", end + (padding if content < len(line) and 1 <= padding <= 4 else 1)


def markdown_indented_code_lines(
    text: str,
    definition_ends: dict[int, int] | None = None,
    *,
    include_indented: bool = True,
    active_definitions: set[int] | None = None,
    content_columns: dict[int, int] | None = None,
) -> set[int]:
    """Locate literal blocks using container margins and paragraph context.

    A structural view expands tabs to four-column stops; callers still scan the
    original lines, so destinations and their exception identities are untouched.
    Container steps consume prefixes monotonically. Literal fenced content never
    creates containers, but its opening line can establish a containing list.
    """
    code_lines: set[int] = set()
    containers: list[tuple[str, int]] = []
    quote_positions: list[int] = []
    empty_list: int | None = None
    paragraph_active = False
    definition_end = 0
    active_fence: MarkdownFence | None = None
    active_html_end: re.Pattern[str] | None = None
    for line_number, raw_line in enumerate(markdown_lines(text), 1):
        if line_number <= definition_end:
            continue
        if active_fence is not None:
            fence_content = active_fence_content(raw_line, active_fence)
            if fence_content is not None:
                code_lines.add(line_number)
                if parse_fence_close_from_content(
                    fence_content,
                    fence_character=active_fence.character,
                    minimum_length=active_fence.length,
                    allow_arbitrary_indent=active_fence.allow_arbitrary_indent,
                ):
                    active_fence = None
                paragraph_active = False
                continue
            active_fence = None
        opened_fence = parse_markdown_fence_open(raw_line, MARKDOWN_FENCE_CONTEXT)
        line = raw_line.expandtabs(4)
        offset = 0
        matched = 0
        space_end = 0
        for kind, margin in containers:
            if offset >= space_end:
                space_end = markdown_prefix_end(line, offset)
            start = space_end
            if kind == "quote":
                if start - offset > 3 or start == len(line) or line[start] != ">":
                    break
                offset = start + 1
                offset += offset < len(line) and line[offset] == " "
            elif start == len(line):
                # Blank lines may retain a list without its indentation.
                offset = start
                quote_index = bisect_left(quote_positions, matched)
                matched = (
                    quote_positions[quote_index]
                    if quote_index < len(quote_positions)
                    else len(containers)
                )
                break
            elif start - offset >= margin:
                offset += margin
            else:
                break
            matched += 1
        if content_columns is not None:
            content_columns[line_number] = offset
        content = line[offset:]
        blank = not content.strip(" ")
        if blank and empty_list is not None:
            # An item starting empty cannot consume another initial blank line.
            matched = min(matched, empty_list)
        empty_list = None
        if active_html_end is not None:
            if matched == len(containers):
                code_lines.add(line_number)
                if active_html_end.search(content):
                    active_html_end = None
                paragraph_active = False
                continue
            active_html_end = None
        html_end = markdown_link_html_block_end(content, paragraph_active=paragraph_active)
        heading = is_policy_heading(content)
        setext = (
            paragraph_active
            and matched == len(containers)
            and re.fullmatch(r" {0,3}(?:=+|-+)[ ]*", content) is not None
        )
        separator = is_policy_thematic_break(content) or (
            re.fullmatch(r" {0,3}(?:=+|-+)[ ]*", content) is not None
        )
        starts_container = markdown_link_container(
            line,
            offset,
            paragraph_active=paragraph_active and matched == len(containers),
        )
        if (
            paragraph_active
            and not blank
            and not heading
            and not separator
            and starts_container is None
            and opened_fence is None
            and html_end is None
        ):
            # Indentation alone cannot interrupt an open paragraph, including
            # lazy continuations that omit a quote or list prefix.
            continue
        del containers[matched:]
        while quote_positions and quote_positions[-1] >= matched:
            quote_positions.pop()
        paragraph_active = False
        if blank or setext:
            continue
        thematic_start, thematic_end = markdown_thematic_suffix(line)
        while not (
            thematic_start <= offset <= thematic_end
            and markdown_prefix_end(line, offset) - offset <= 3
        ):
            container = markdown_link_container(line, offset)
            if container is None:
                break
            kind, end = container
            if kind == "quote":
                quote_positions.append(len(containers))
            if kind == "list" and markdown_prefix_end(line, end) >= len(line):
                empty_list = len(containers)
            containers.append((kind, end - offset))
            offset = end
        if content_columns is not None:
            content_columns[line_number] = offset
        content = line[offset:]
        if markdown_prefix_end(line, offset) - offset >= 4:
            if include_indented:
                code_lines.add(line_number)
            continue
        if definition_ends is not None and line_number in definition_ends:
            definition_end = definition_ends[line_number]
            if active_definitions is not None:
                active_definitions.add(line_number)
            continue
        html_end = markdown_link_html_block_end(content, paragraph_active=False)
        if html_end is not None:
            code_lines.add(line_number)
            if not html_end.search(content):
                active_html_end = html_end
            continue
        active_fence = opened_fence
        if active_fence is not None:
            code_lines.add(line_number)
        paragraph_active = bool(content.strip(" ")) and not (
            active_fence is not None
            or is_policy_heading(content)
            or is_policy_thematic_break(content)
            or re.fullmatch(r" {0,3}(?:=+|-+)[ ]*", content) is not None
        )
    return code_lines


def markdown_reference_definition_ends(spans: list[list[tuple[int, str]]]) -> dict[int, int]:
    """Index complete definition boundaries with the existing component grammar."""
    ends: dict[int, int] = {}
    for span in spans:
        text = "\n".join(line for _, line in span)
        starts: list[int] = []
        offset = 0
        for _, line in span:
            starts.append(offset)
            offset += len(line) + 1
        labels, _closes = markdown_delimiter_pairs(text)
        definitions = {
            opening for opening, closing in labels if text[closing + 1 : closing + 2] == ":"
        }
        if not definitions:
            continue
        whitespace = [
            index
            for index, character in enumerate(text)
            if ord(character) <= 32 or ord(character) == 127
        ]
        whitespace.append(len(text))
        for opening, end, _target_start, _target_end in markdown_link_candidates(
            text, starts, whitespace
        ):
            if opening in definitions:
                first = bisect_right(starts, opening) - 1
                last = bisect_right(starts, end - 1) - 1
                ends[span[first][0]] = span[last][0]
    return ends


def markdown_link_spans(
    text: str,
    fence_context: str,
    reference_labels: set[str] | None = None,
    definition_lines: set[int] | None = None,
) -> list[list[tuple[int, str]]]:
    """Locate contextual definitions while removing literal blocks from link spans."""
    if fence_context != MARKDOWN_FENCE_CONTEXT:
        spans = markdown_live_link_spans(text, fence_context, set())
        active_definitions = set(markdown_reference_definition_ends(spans))
        if reference_labels is not None:
            reference_labels.update(markdown_definition_labels(spans, active_definitions))
        if definition_lines is not None:
            definition_lines.update(active_definitions)
        return spans
    content_columns: dict[int, int] = {}
    literal_lines = markdown_indented_code_lines(
        text, include_indented=False, content_columns=content_columns
    )
    spans = markdown_live_link_spans(
        text, fence_context, literal_lines, content_columns=content_columns
    )
    definition_ends = markdown_reference_definition_ends(spans)
    active_definitions = set()
    code_lines = markdown_indented_code_lines(
        text,
        definition_ends,
        active_definitions=active_definitions,
        content_columns=content_columns,
    )
    if reference_labels is not None:
        reference_labels.update(markdown_definition_labels(spans, active_definitions))
    if definition_lines is not None:
        definition_lines.update(active_definitions)
    return markdown_live_link_spans(
        text, fence_context, code_lines, content_columns=content_columns
    )


def markdown_source_content(line: str, column: int) -> str:
    """Remove structural columns while preserving raw content and partial tabs."""
    offset = 0
    consumed = 0
    while offset < len(line) and (consumed < column or line[offset] in " \t"):
        consumed += 4 - consumed % 4 if line[offset] == "\t" else 1
        offset += 1
    return " " * max(0, consumed - column) + line[offset:]


def markdown_live_link_spans(
    text: str,
    fence_context: str,
    code_lines: set[int],
    *,
    content_columns: dict[int, int] | None = None,
) -> list[list[tuple[int, str]]]:
    """Partition live inline text without bridging fences, blank lines or blocks."""
    spans: list[list[tuple[int, str]]] = []
    current: list[tuple[int, str]] = []
    previous_line = 0
    quote_depth = 0
    lines = (
        tuple(enumerate(markdown_lines(text), 1))
        if fence_context == MARKDOWN_FENCE_CONTEXT
        else lines_outside_markdown_fences(text, fence_context=fence_context)
    )
    for line_number, line in lines:
        if line_number in code_lines:
            continue
        depth, offset = consume_blockquote_prefix(
            line, allow_arbitrary_indent=fence_context != MARKDOWN_FENCE_CONTEXT
        )
        content = line[offset:]
        list_match = LIST_MARKER_RE.match(content)
        heading = re.match(r"^ {0,3}#{1,6}(?:[ \t]|$)", content) is not None
        separator = re.fullmatch(r" {0,3}(?:[-*_][ \t]*){3,}| {0,3}=+[ \t]*", content) is not None
        # A plain unprefixed line can lazily continue a quote/list paragraph.
        new_container = depth != quote_depth and depth != 0
        if current and (
            line_number != previous_line + 1 or new_container or list_match or heading or separator
        ):
            spans.append(current)
            current = []
        if not content.strip(" \t"):
            if current:
                spans.append(current)
                current = []
            quote_depth = 0
        else:
            if content_columns is not None:
                content = markdown_source_content(line, content_columns.get(line_number, 0))
            elif list_match:
                content = list_match.group("rest")
            current.append((line_number, content))
            if heading or separator:
                spans.append(current)
                current = []
            quote_depth = depth or quote_depth
        previous_line = line_number
    if current:
        spans.append(current)
    return spans


def markdown_delimiter_pairs(text: str) -> tuple[list[tuple[int, int]], dict[int, int]]:
    """Index bracket candidates and balanced parentheses in one bounded pass."""
    brackets: list[int] = []
    parentheses: list[int] = []
    labels: list[tuple[int, int]] = []
    closes: dict[int, int] = {}
    index = 0
    while index < len(text):
        character = text[index]
        if character == "\\" and index + 1 < len(text) and text[index + 1] in ASCII_PUNCTUATION:
            index += 2
            continue
        if character == "[":
            brackets.append(index)
        elif character == "]" and brackets:
            labels.append((brackets.pop(), index))
        elif character == "(":
            parentheses.append(index)
        elif character == ")" and parentheses:
            closes[parentheses.pop()] = index
        index += 1
    return labels, closes


ASCII_PUNCTUATION = r"!\"#$%&'()*+,-./:;<=>?@[\]^_`{|}~"


def markdown_link_space(text: str, index: int) -> int:
    """Skip component whitespace, returning -1 if it includes multiple newlines."""
    endings = 0
    while index < len(text) and text[index] in " \t\n":
        endings += text[index] == "\n"
        if endings > 1:
            return -1
        index += 1
    return index


def markdown_link_destination(
    text: str, index: int, closes: dict[int, int], whitespace: list[int]
) -> tuple[int, int, int] | None:
    """Read a destination without rescanning nested or unmatched parentheses."""
    if index >= len(text):
        return None
    start = index
    if text[index] == "<":
        index += 1
        while index < len(text):
            character = text[index]
            if character == "\\" and index + 1 < len(text) and text[index + 1] in ASCII_PUNCTUATION:
                index += 2
                continue
            if character == ">":
                return index + 1, start + 1, index
            if character in "<\n":
                return None
            index += 1
        return None
    next_space = whitespace[bisect_left(whitespace, start)]
    while index < next_space:
        character = text[index]
        if character == "\\" and index + 1 < next_space and text[index + 1] in ASCII_PUNCTUATION:
            index += 2
            continue
        if character == ")":
            break
        if character == "(":
            closing = closes.get(index)
            if closing is None or closing >= next_space:
                return None
            index = closing + 1
        else:
            index += 1
    return (index, start, index) if index > start else None


def markdown_link_title_end(text: str, index: int) -> int:
    """Read one nonblank multiline quoted title, returning -1 when incomplete."""
    if index >= len(text) or text[index] not in "\"'(":
        return -1
    opening = text[index]
    closing = ")" if opening == "(" else opening
    index += 1
    while index < len(text):
        character = text[index]
        if character == "\\" and index + 1 < len(text) and text[index + 1] in ASCII_PUNCTUATION:
            index += 2
            continue
        if character == closing:
            return index + 1
        if opening == "(" and character == "(":
            return -1
        index += 1
    return -1


def markdown_link_opening_is_image(text: str, opening: int) -> bool:
    """Return whether a bracket opening follows an unescaped image marker."""
    if opening == 0 or text[opening - 1] != "!":
        return False
    backslashes = 0
    index = opening - 2
    while index >= 0 and text[index] == "\\":
        backslashes += 1
        index -= 1
    # Escaped backslashes leave the following bang active; an odd run escapes it.
    return backslashes % 2 == 0


def markdown_code_span_ends(text: str) -> dict[int, int]:
    """Index equal-width backtick closers without repeated suffix searches.

    Closing runs remain eligible after a backslash inside code. Outside code,
    escaping the first tick of a longer run leaves a shorter opening run.
    """
    nearest: dict[int, int] = {}
    ends: dict[int, int] = {}
    for match in reversed(list(re.finditer(r"`+", text))):
        start, end = match.span()
        width = end - start
        if width in nearest:
            ends[start] = nearest[width]
        if width > 1 and width - 1 in nearest:
            ends[start + 1] = nearest[width - 1]
        nearest[width] = end
    return ends


def markdown_link_component(
    text: str,
    opening: int,
    closing: int,
    closes: dict[int, int],
    whitespace: list[int],
    line_start: int,
) -> tuple[int, int, int] | None:
    """Return consumed end and raw target indices for a complete link component."""
    component = closing + 1
    if component >= len(text) or text[component] not in "(:":
        return None
    is_reference = text[component] == ":"
    if is_reference and (
        opening - line_start > 3 or text[line_start:opening] not in ("", " ", "  ", "   ")
    ):
        return None
    index = markdown_link_space(text, component + 1)
    if index < 0:
        return None
    destination = markdown_link_destination(text, index, closes, whitespace)
    if destination is None:
        return None
    end, target_start, target_end = destination
    after_space = markdown_link_space(text, end)
    if after_space < 0:
        return None
    if not is_reference and after_space < len(text) and text[after_space] == ")":
        return after_space + 1, target_start, target_end
    title_end = markdown_link_title_end(text, after_space) if after_space > end else -1
    if is_reference:
        line_end = text.find("\n", end)
        line_end = len(text) if line_end < 0 else line_end
        if title_end >= 0:
            suffix_end = text.find("\n", title_end)
            suffix_end = len(text) if suffix_end < 0 else suffix_end
            if text[title_end:suffix_end].strip(" \t"):
                title_end = -1
        if title_end >= 0 or not text[end:line_end].strip(" \t"):
            return title_end if title_end >= 0 else end, target_start, target_end
    elif title_end >= 0:
        final = markdown_link_space(text, title_end)
        if final >= 0 and final < len(text) and text[final] == ")":
            return final + 1, target_start, target_end
    return None


def markdown_reference_label(text: str, opening: int, closing: int) -> str | None:
    """Normalize a bounded label, preserving literal escapes and rejecting brackets."""
    if not 1 <= closing - opening - 1 <= 999:
        return None
    index = opening + 1
    while index < closing:
        if text[index] == "\\" and index + 1 < closing and text[index + 1] in ASCII_PUNCTUATION:
            index += 2
            continue
        if text[index] in "[]":
            return None
        index += 1
    label = re.sub(r"[ \t\r\n]+", " ", text[opening + 1 : closing].casefold()).strip(" ")
    return label or None


def markdown_definition_labels(
    spans: list[list[tuple[int, str]]], active_lines: set[int]
) -> set[str]:
    """Collect lookup labels only from contextual definitions, not raw inventory."""
    labels: set[str] = set()
    for span in spans:
        if not any(line_number in active_lines for line_number, _line in span):
            continue
        text = "\n".join(line for _, line in span)
        starts: list[int] = []
        offset = 0
        for _, line in span:
            starts.append(offset)
            offset += len(line) + 1
        pairs, _closes = markdown_delimiter_pairs(text)
        for opening, closing in pairs:
            if text[closing + 1 : closing + 2] != ":":
                continue
            line_index = bisect_right(starts, opening) - 1
            line_start = starts[line_index]
            if (
                span[line_index][0] not in active_lines
                or opening - line_start > 3
                or text[line_start:opening] not in ("", " ", "  ", "   ")
            ):
                continue
            label = markdown_reference_label(text, opening, closing)
            if label is not None:
                labels.add(label)
    return labels


def markdown_reference_use_end(
    text: str,
    opening: int,
    closing: int,
    bracket_closes: dict[int, int],
    reference_labels: set[str],
) -> int | None:
    """Recognize reference activity without duplicating definition target emission."""
    if not reference_labels:
        return None
    following = closing + 1
    if text[following : following + 2] == "[]":
        label = markdown_reference_label(text, opening, closing)
        return following + 2 if label in reference_labels else None
    if following in bracket_closes:
        label_end = bracket_closes[following]
        label = markdown_reference_label(text, following, label_end)
        if label is not None:
            # A valid explicit label blocks shortcut fallback even when unknown.
            return label_end + 1 if label in reference_labels else None
    label = markdown_reference_label(text, opening, closing)
    return following if label in reference_labels else None


def markdown_link_candidates(
    text: str,
    starts: list[int],
    whitespace: list[int],
    reference_labels: set[str] | None = None,
    definition_starts: set[int] | None = None,
) -> list[tuple[int, int, int, int]]:
    """Scan inline labels while keeping destinations and definitions literal.

    Each accepted candidate records its opening, consumed end and raw target
    bounds. Raw parentheses retain their destination grammar. Code spans only
    control brackets in inline text, never backticks inside accepted components.
    """
    raw_labels, closes = markdown_delimiter_pairs(text)
    definitions = {
        opening: closing
        for opening, closing in raw_labels
        if text[closing + 1 : closing + 2] == ":"
    }
    code_ends = markdown_code_span_ends(text)
    html_ends = markdown_html_token_ends(text)
    bracket_closes = dict(raw_labels)
    # Each frame records its candidate and visible-link checkpoints. Successful
    # images contain label links; no child needs to rescan all ancestor frames.
    brackets: list[tuple[int, int, int, bool]] = []
    candidates: list[tuple[int, int, int, int]] = []
    visible_links = 0
    reference_labels = reference_labels or set()
    index = 0
    while index < len(text):
        character = text[index]
        if character == "\\" and index + 1 < len(text) and text[index + 1] in ASCII_PUNCTUATION:
            index += 2
            continue
        if character == "`":
            if index in code_ends:
                index = code_ends[index]
            else:
                index += 1
                while index < len(text) and text[index] == "`":
                    index += 1
            continue
        if character == "<" and index in html_ends:
            index = html_ends[index]
            continue
        opening = index
        closing = definitions.get(index) if character == "[" else None
        if closing is not None:
            line_start = starts[bisect_right(starts, opening) - 1]
            # Paragraph lookalikes stay inline; their apparent titles can
            # contain real links that definition consumption would hide.
            if definition_starts is None or line_start in definition_starts:
                component = markdown_link_component(
                    text, opening, closing, closes, whitespace, line_start
                )
                if component is not None:
                    end, target_start, target_end = component
                    candidates.append((opening, end, target_start, target_end))
                    index = end
                    continue
        if character == "[":
            brackets.append(
                (index, len(candidates), visible_links, markdown_link_opening_is_image(text, index))
            )
        elif character == "]" and brackets:
            opening, candidate_checkpoint, link_checkpoint, is_image = brackets.pop()
            if not is_image and visible_links != link_checkpoint:
                # A rendered link in this label leaves the outer syntax literal.
                index += 1
                continue
            component = None
            if text[index + 1 : index + 2] == "(":
                line_start = starts[bisect_right(starts, opening) - 1]
                component = markdown_link_component(
                    text, opening, index, closes, whitespace, line_start
                )
            reference_end = (
                markdown_reference_use_end(text, opening, index, bracket_closes, reference_labels)
                if component is None
                else None
            )
            if component is not None or reference_end is not None:
                if is_image:
                    del candidates[candidate_checkpoint:]
                    visible_links = link_checkpoint
                else:
                    visible_links += 1
                if component is not None:
                    end, target_start, target_end = component
                    candidates.append((opening, end, target_start, target_end))
                    index = end
                else:
                    assert reference_end is not None
                    index = reference_end
                continue
        index += 1
    return candidates


def markdown_span_link_targets(
    span: list[tuple[int, str]],
    reference_labels: set[str] | None = None,
    definition_lines: set[int] | None = None,
) -> list[tuple[int, str]]:
    """Extract the supported link surface from one contiguous live text span."""
    text = "\n".join(line for _, line in span)
    starts: list[int] = []
    offset = 0
    for _, line in span:
        starts.append(offset)
        offset += len(line) + 1
    whitespace = [
        index
        for index, character in enumerate(text)
        if ord(character) <= 32 or ord(character) == 127
    ]
    whitespace.append(len(text))
    definition_starts = (
        None
        if definition_lines is None
        else {start for start, (line, _) in zip(starts, span) if line in definition_lines}
    )
    targets: list[tuple[int, str]] = []
    consumed_until = 0
    for opening, end, target_start, target_end in sorted(
        markdown_link_candidates(text, starts, whitespace, reference_labels, definition_starts)
    ):
        if opening < consumed_until:
            continue
        if markdown_link_opening_is_image(text, opening):
            continue
        line_index = bisect_right(starts, opening) - 1
        targets.append((span[line_index][0], text[target_start:target_end]))
        consumed_until = end
    return targets


def markdown_link_targets_from_text(
    text: str, *, fence_context: str = MARKDOWN_FENCE_CONTEXT
) -> tuple[tuple[int, str], ...]:
    """Extract bounded multiline links, not a general CommonMark rendering tree."""
    reference_labels: set[str] = set()
    definition_lines: set[int] = set()
    spans = markdown_link_spans(text, fence_context, reference_labels, definition_lines)
    return tuple(
        target
        for span in spans
        for target in markdown_span_link_targets(span, reference_labels, definition_lines)
    )


def normalize_markdown_target(target: str) -> str:
    """Strip Markdown angle brackets from a link target."""
    if target.startswith("<") and target.endswith(">"):
        return target[1:-1]
    return target


MARKDOWN_DESTINATION_ESCAPE_RE = re.compile(
    r"\\["
    + re.escape(ASCII_PUNCTUATION)
    + r"]|&(?:#[xX][0-9a-fA-F]{1,6}|#[0-9]{1,7}|[A-Za-z][A-Za-z0-9]{1,31});"
)


def decode_markdown_destination(target: str) -> str:
    """Decode one layer of destination escapes without changing source identities.

    Match bounded, complete references before URL parsing. Substitutions are not
    rescanned: escaped ampersands, entity-produced backslashes and nested entities
    remain literal. Percent decoding belongs to the later URI/path layer.
    """

    def decode_match(match: re.Match[str]) -> str:
        """Decode one punctuation escape or strict character reference."""
        token = match.group(0)
        if token.startswith("\\"):
            return token[1:]
        if not token.startswith("&#"):
            return html5.get(token[1:], token)
        numeric = token[2:-1]
        codepoint = int(numeric[1:], 16) if numeric[:1] in {"x", "X"} else int(numeric)
        if codepoint == 0 or codepoint > 0x10FFFF or 0xD800 <= codepoint <= 0xDFFF:
            return "\ufffd"
        # HTML numeric replacements include C1 mappings. Python's HTML decoder
        # also drops some valid scalars; destinations must preserve those scalars.
        return unescape(token) or chr(codepoint)

    return MARKDOWN_DESTINATION_ESCAPE_RE.sub(decode_match, target)


def split_markdown_destination(target: str) -> SplitResult:
    """Parse one decoded destination without losing literal whitespace characters."""
    target = decode_markdown_destination(target)
    # urlsplit removes these literal characters. Protect them until path decoding
    # so a distinct destination cannot collapse to an ordinary repository path.
    protected_target = re.sub(
        r"[\x00-\x20\x7f]", lambda match: f"%{ord(match.group(0)):02X}", target
    )
    return urlsplit(protected_target)


def resolve_relative_markdown_target(source_path: str, target: str) -> str | None:
    """Resolve a Markdown link target to a repository-relative path when local."""
    parsed = split_markdown_destination(target)
    if parsed.scheme or parsed.netloc or target.startswith("#"):
        return None
    if parsed.path == "":
        return None

    decoded_path = unquote(parsed.path)
    if decoded_path.startswith("/"):
        return None

    source_dir = posixpath.dirname(source_path)
    normalized_path = posixpath.normpath(posixpath.join(source_dir, decoded_path))
    if normalized_path == "." or normalized_path.startswith("../") or normalized_path == "..":
        return None
    return normalized_path


def protected_guide_reference_findings(
    *, repo_root: Path, obligation: ProtectedGuideReferenceObligation
) -> tuple[ProtectedGuideReferenceFinding, ...]:
    """Match declared references with bounded input and the shared fence grammar."""
    try:
        path = support.resolve_repo_path(repo_root, obligation.path)
        text = read_repository_text(path, repo_root, maximum_bytes=MAXIMUM_INPUT_BYTES)
    except support.TemplateSyncMaterializationError as error:
        if not isinstance(error.__cause__, OSError):
            raise
        io_error = error.__cause__
        return (
            ProtectedGuideReferenceFinding(
                path=obligation.path,
                line_number=0,
                contract_key=obligation.key,
                reference_kind=obligation.reference_kind,
                target=f"Unable to read protected guide: {support.os_error_summary(io_error)}",
                target_path=obligation.target_path,
                target_modules=obligation.target_modules,
            ),
        )
    matches: list[tuple[int, str]] = []
    if obligation.reference_kind == "markdown-relative-link":
        for line_number, target in markdown_link_targets_from_text(text):
            if resolve_relative_markdown_target(obligation.path, target) == obligation.target_path:
                matches.append((line_number, target))
    else:
        for line_number, line in lines_outside_markdown_fences(
            text, fence_context=MARKDOWN_FENCE_CONTEXT
        ):
            matches.extend((line_number, token) for token in obligation.tokens if token in line)
    return tuple(
        ProtectedGuideReferenceFinding(
            path=obligation.path,
            line_number=line_number,
            contract_key=obligation.key,
            reference_kind=obligation.reference_kind,
            target=target,
            target_path=obligation.target_path,
            target_modules=obligation.target_modules,
        )
        for line_number, target in matches
    )


def validate_contracts(
    *,
    mode: str,
    repo_root: Path,
    contracts: tuple[InstructionContract, ...],
    protected_guide_section_obligations: tuple[ProtectedGuideSectionObligation, ...] = (),
    protected_guide_reference_obligations: tuple[ProtectedGuideReferenceObligation, ...] = (),
    included_modules: set[str] | None = None,
    protected_decisions: tuple[support.ProtectedFileDecision, ...] = (),
    waivers: tuple[InstructionContractWaiver, ...] = (),
    protected_guide_waivers: tuple[support.ProtectedGuideContractWaiver, ...] = (),
    warnings: tuple[str, ...] = (),
) -> InstructionContractReport:
    """Validate selected instruction contracts against the working tree."""
    checked_contracts: list[InstructionContract] = []
    skipped_contracts: list[SkippedContract] = []
    missing_files: list[MissingFile] = []
    missing_anchors: list[MissingAnchor] = []
    stale_protected_guide_sections: list[StaleProtectedGuideSection] = []
    stale_protected_guide_references: list[ProtectedGuideReferenceFinding] = []
    applied_waivers: list[InstructionContractWaiver] = []
    applied_protected_guide_waivers: list[support.ProtectedGuideContractWaiver] = []
    authorized_removals: list[AuthorizedRemoval] = []
    active_imports: list[ActiveClaudeImport] = []
    tracked_local_memory, tracked_inventory_applicable = tracked_claude_local_memory(repo_root)
    if not tracked_inventory_applicable:
        warnings = (
            *warnings,
            "Tracked Claude local memory: N/A; the validation root has no .git marker.",
        )

    for contract in contracts:
        if included_modules is not None and not set(contract.requires_modules).issubset(
            included_modules
        ):
            skipped_contracts.append(
                SkippedContract(
                    path=contract.path,
                    requires_modules=contract.requires_modules,
                )
            )
            continue

        text = read_instruction_file(repo_root, contract.path)
        if text is None:
            authorized_removal = authorized_removal_for(protected_decisions, contract.path)
            if authorized_removal is not None:
                authorized_removals.append(authorized_removal)
            else:
                missing_files.append(MissingFile(path=contract.path))
            continue

        checked_contracts.append(contract)
        if contract.path.rsplit("/", 1)[-1] == "CLAUDE.md":
            active_imports.extend(active_claude_imports(contract.path, text))
        scannable_text = strip_fenced_code_blocks(text)
        for heading in contract.required_headings:
            if heading_is_present(scannable_text, heading):
                continue
            waiver = find_waiver(waivers, contract.path, heading)
            if waiver is not None:
                applied_waivers.append(waiver)
            else:
                missing_anchors.append(
                    MissingAnchor(
                        path=contract.path,
                        anchor_type="heading",
                        anchor=heading,
                    )
                )
        for phrase in contract.required_phrases:
            if phrase in scannable_text:
                continue
            waiver = find_waiver(waivers, contract.path, phrase)
            if waiver is not None:
                applied_waivers.append(waiver)
            else:
                missing_anchors.append(
                    MissingAnchor(
                        path=contract.path,
                        anchor_type="phrase",
                        anchor=phrase,
                    )
                )

        for anchor in section_failures(text, contract.required_sections, included_modules):
            waiver = find_waiver(waivers, contract.path, anchor)
            if waiver is not None:
                applied_waivers.append(waiver)
            else:
                missing_anchors.append(MissingAnchor(contract.path, "section content", anchor))

    if included_modules is not None:
        excluded_contract_paths = {item.path for item in skipped_contracts}
        missing_file_paths = {missing_file.path for missing_file in missing_files}
        authorized_removal_paths = {
            authorized_removal.path for authorized_removal in authorized_removals
        }
        for obligation in protected_guide_section_obligations:
            if obligation.path in excluded_contract_paths:
                continue
            if not protected_guide_obligation_applies(
                obligation.target_modules,
                included_modules,
            ):
                continue

            text = read_instruction_file(repo_root, obligation.path)
            if text is None:
                authorized_removal = authorized_removal_for(
                    protected_decisions,
                    obligation.path,
                )
                if authorized_removal is not None:
                    if authorized_removal.path not in authorized_removal_paths:
                        authorized_removals.append(authorized_removal)
                        authorized_removal_paths.add(authorized_removal.path)
                elif obligation.path not in missing_file_paths:
                    missing_files.append(MissingFile(path=obligation.path))
                    missing_file_paths.add(obligation.path)
                continue

            scannable_text = strip_fenced_code_blocks(text)
            stale_sections: list[StaleProtectedGuideSection] = []
            for heading in obligation.stale_headings:
                if heading_is_present(scannable_text, heading):
                    stale_sections.append(
                        StaleProtectedGuideSection(
                            path=obligation.path,
                            contract_key=obligation.key,
                            anchor_type="heading",
                            anchor=heading,
                            target_modules=obligation.target_modules,
                        )
                    )
            for phrase in obligation.stale_phrases:
                if phrase in scannable_text:
                    stale_sections.append(
                        StaleProtectedGuideSection(
                            path=obligation.path,
                            contract_key=obligation.key,
                            anchor_type="phrase",
                            anchor=phrase,
                            target_modules=obligation.target_modules,
                        )
                    )
            if not stale_sections:
                continue
            protected_guide_waiver = find_protected_guide_waiver(
                protected_guide_waivers,
                path=obligation.path,
                contract_key=obligation.key,
                target_modules=obligation.target_modules,
            )
            if protected_guide_waiver is not None:
                applied_protected_guide_waivers.append(protected_guide_waiver)
            else:
                stale_protected_guide_sections.extend(stale_sections)

        for reference_obligation in protected_guide_reference_obligations:
            if reference_obligation.path in excluded_contract_paths:
                continue
            if not protected_guide_obligation_applies(
                reference_obligation.target_modules, included_modules
            ):
                continue
            if read_instruction_file(repo_root, reference_obligation.path) is None:
                removal = authorized_removal_for(protected_decisions, reference_obligation.path)
                if removal is not None:
                    if removal.path not in authorized_removal_paths:
                        authorized_removals.append(removal)
                        authorized_removal_paths.add(removal.path)
                elif reference_obligation.path not in missing_file_paths:
                    missing_files.append(MissingFile(path=reference_obligation.path))
                    missing_file_paths.add(reference_obligation.path)
                continue
            reference_findings = protected_guide_reference_findings(
                repo_root=repo_root, obligation=reference_obligation
            )
            if not reference_findings:
                continue
            reference_waiver = find_protected_guide_waiver(
                protected_guide_waivers,
                path=reference_obligation.path,
                contract_key=reference_obligation.key,
                target_modules=reference_obligation.target_modules,
                target_path=reference_obligation.target_path,
            )
            if reference_waiver is not None:
                applied_protected_guide_waivers.append(reference_waiver)
            else:
                stale_protected_guide_references.extend(reference_findings)

    return InstructionContractReport(
        mode=mode,
        contracts_checked=tuple(checked_contracts),
        skipped_contracts=tuple(skipped_contracts),
        missing_files=tuple(missing_files),
        missing_anchors=tuple(missing_anchors),
        stale_protected_guide_sections=tuple(stale_protected_guide_sections),
        stale_protected_guide_references=tuple(stale_protected_guide_references),
        applied_waivers=tuple(dict.fromkeys(applied_waivers)),
        applied_protected_guide_waivers=tuple(dict.fromkeys(applied_protected_guide_waivers)),
        authorized_removals=tuple(authorized_removals),
        warnings=warnings,
        active_claude_imports=tuple(active_imports),
        tracked_claude_local_memory=tracked_local_memory,
        tracked_claude_local_memory_inventory_applicable=tracked_inventory_applicable,
    )


def print_report(report: InstructionContractReport) -> None:
    """Print a human-readable instruction contract report."""
    for warning in report.warnings:
        print(f"WARNING: {warning}")

    if report.has_failures:
        print("Instruction-contract validation failed.")
    elif report.applied_waivers or report.applied_protected_guide_waivers:
        print("Instruction-contract validation passed with waivers.")
    else:
        print("Instruction-contract validation passed.")

    print(f"Mode: {report.mode}")
    print(f"Contracts checked: {len(report.contracts_checked)}")

    if report.skipped_contracts:
        print("\nContracts skipped by downstream module selection:")
        for skipped_contract in report.skipped_contracts:
            print(
                f"  - {skipped_contract.path} "
                f"(requires: {', '.join(skipped_contract.requires_modules)})"
            )

    if report.missing_files:
        print("\nRequired instruction files absent without authorized removal:")
        for missing_file in report.missing_files:
            print(f"  - {missing_file.path}")

    if report.missing_anchors:
        print("\nMissing required anchors:")
        for missing_anchor in report.missing_anchors:
            print(
                f"  - {missing_anchor.path}: missing required "
                f"{missing_anchor.anchor_type}: {missing_anchor.anchor}"
            )

    if report.stale_protected_guide_sections:
        print("\nStale protected-guide sections requiring owner review:")
        for stale_section in report.stale_protected_guide_sections:
            print(
                f"  - {stale_section.path}: {stale_section.contract_key}: stale "
                f"{stale_section.anchor_type}: {stale_section.anchor} "
                f"(target modules: {', '.join(stale_section.target_modules)})"
            )

    if report.stale_protected_guide_references:
        print("\nStale protected-guide references requiring owner review:")
        for reference in report.stale_protected_guide_references:
            print(
                f"  - {reference.path}:{reference.line_number}: {reference.contract_key}: "
                f"{reference.reference_kind}: {reference.target} "
                f"(target modules: {', '.join(reference.target_modules)})"
            )

    if report.active_claude_imports:
        print("\nActive Claude imports:")
        for active_import in report.active_claude_imports:
            print(f"  - {active_import.path}:{active_import.line}: @{active_import.target}")

    if report.tracked_claude_local_memory:
        print("\nTracked Claude local memory:")
        for local_memory in report.tracked_claude_local_memory:
            print(f"  - {local_memory.path}")

    if report.authorized_removals:
        print("\nAuthorized removals skipped:")
        for authorized_removal in report.authorized_removals:
            print(f"  - {authorized_removal.path}")
            print(f"    authorization_basis: {authorized_removal.authorization_basis}")
            print(f"    authorized_scope: {authorized_removal.authorized_scope}")
            print(f"    reason: {authorized_removal.reason}")

    if report.applied_waivers:
        print("\nInstruction contract waivers applied:")
        for waiver in report.applied_waivers:
            print(f"  - {waiver.path}: {waiver.anchor}")
            print(f"    reason: {waiver.reason}")
            print(f"    authorization_basis: {waiver.authorization_basis}")

    if report.applied_protected_guide_waivers:
        print("\nProtected guide contract waivers applied:")
        for protected_guide_waiver in report.applied_protected_guide_waivers:
            print(f"  - {protected_guide_waiver.path}: {protected_guide_waiver.contract_key}")
            if protected_guide_waiver.target_path is not None:
                print(f"    target_path: {protected_guide_waiver.target_path}")
            if protected_guide_waiver.target_module is not None:
                print(f"    target_module: {protected_guide_waiver.target_module}")
            if protected_guide_waiver.linked_local_override_path is not None:
                print(
                    "    linked_local_override_path: "
                    f"{protected_guide_waiver.linked_local_override_path}"
                )
            print(f"    reason: {protected_guide_waiver.reason}")
            print(f"    authorization_basis: {protected_guide_waiver.authorization_basis}")


def fail(message: str) -> NoReturn:
    """Print an error and exit non-zero."""
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(1)


def file_digest(root: Path, path: str) -> str:
    """Bind a declaration to exact normalized UTF-8 content or an absent file."""
    text = read_instruction_file(root, path)
    return "absent" if text is None else hashlib.sha256(text.encode("utf-8")).hexdigest()
