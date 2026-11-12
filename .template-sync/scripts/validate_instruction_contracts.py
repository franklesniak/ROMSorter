"""Marker-authoritative adapter for the shared instruction-contract engine."""

# ruff: noqa: E402

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any, cast

import validate_marker
from template_sync_materialization_helpers import normalize_repository_path

_SHARED_SCRIPTS = Path(__file__).resolve().parents[2] / ".github" / "scripts"
if str(_SHARED_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SHARED_SCRIPTS))

from instruction_contract_core import (
    CLAUDE_IMPORT_PATTERN,
    GIT_TIMEOUT_SECONDS,
    MAXIMUM_CLAUDE_CONTAINER_STEPS,
    MAXIMUM_GIT_OUTPUT_BYTES,
    MAXIMUM_INPUT_BYTES,
    POLICY_CELL_PATTERN,
    POLICY_CELL_WORD,
    POLICY_HEADING_PATTERN,
    POLICY_HTML_AMBIGUITY_PREFIX,
    POLICY_HTML_AMBIGUOUS_START,
    POLICY_HTML_BLANK_END,
    POLICY_HTML_BLOCK_START,
    POLICY_HTML_COMPLETE_TAG,
    POLICY_HTML_LITERAL_END,
    POLICY_HTML_LITERAL_START,
    POLICY_HTML_TEXTAREA_END,
    ActiveClaudeImport,
    AuthorizedRemoval,
    ClaudeContainerStep,
    ClaudeContinuation,
    ClaudeImportLine,
    InstructionContract,
    InstructionContractReport,
    InstructionContractValidationError,
    InstructionContractWaiver,
    MissingAnchor,
    MissingFile,
    PolicyBlock,
    PolicyLineIndex,
    ProtectedGuideReferenceObligation,
    ProtectedGuideSectionObligation,
    RequiredSection,
    RequiredTable,
    SkippedContract,
    StaleProtectedGuideSection,
    TrackedClaudeLocalMemory,
    _required_string_list,
    active_claude_imports,
    applicable_section_boundary,
    authorized_removal_for,
    classify_claude_paragraphs,
    claude_container_content,
    claude_continuation_content_offset,
    claude_explicit_container,
    claude_import_lines,
    claude_paragraph_text,
    decode_git_output,
    expected_policy_blocks,
    fail,
    find_protected_guide_waiver,
    find_waiver,
    git_inventory_safe_prefix,
    heading_is_present,
    index_policy_lines,
    is_policy_heading,
    is_policy_thematic_break,
    load_schema_validated_yaml,
    mask_matched_claude_code_spans,
    normalize_policy_paragraph,
    operative_markdown_lines,
    os_error_diagnostic,
    parse_contracts,
    parse_policy_body,
    parse_protected_guide_reference_obligations,
    parse_protected_guide_section_obligations,
    parse_required_sections,
    policy_code_span_ends,
    policy_html_block_end,
    policy_inventory_digest,
    policy_list_content_indent,
    print_report,
    protected_guide_obligation_applies,
    quoted_policy_paragraph,
    read_instruction_file,
    run_bounded_git,
    sanitized_git_environment,
    section_applies,
    section_body,
    section_boundary_failure,
    section_failures,
    serialized_policy_blocks,
    starts_policy_list,
    strip_fenced_code_blocks,
    tracked_claude_local_memory,
    validate_contracts,
)

# Explicit re-exports preserve the existing adapter API for sync callers.
__all__ = [
    "CLAUDE_IMPORT_PATTERN",
    "GIT_TIMEOUT_SECONDS",
    "MAXIMUM_CLAUDE_CONTAINER_STEPS",
    "MAXIMUM_GIT_OUTPUT_BYTES",
    "MAXIMUM_INPUT_BYTES",
    "POLICY_CELL_PATTERN",
    "POLICY_CELL_WORD",
    "POLICY_HEADING_PATTERN",
    "POLICY_HTML_AMBIGUITY_PREFIX",
    "POLICY_HTML_AMBIGUOUS_START",
    "POLICY_HTML_BLANK_END",
    "POLICY_HTML_BLOCK_START",
    "POLICY_HTML_COMPLETE_TAG",
    "POLICY_HTML_LITERAL_END",
    "POLICY_HTML_LITERAL_START",
    "POLICY_HTML_TEXTAREA_END",
    "ActiveClaudeImport",
    "AuthorizedRemoval",
    "ClaudeContainerStep",
    "ClaudeContinuation",
    "ClaudeImportLine",
    "InstructionContract",
    "InstructionContractReport",
    "InstructionContractValidationError",
    "InstructionContractWaiver",
    "MissingAnchor",
    "MissingFile",
    "PolicyBlock",
    "PolicyLineIndex",
    "ProtectedGuideReferenceObligation",
    "ProtectedGuideSectionObligation",
    "RequiredSection",
    "RequiredTable",
    "SkippedContract",
    "StaleProtectedGuideSection",
    "TrackedClaudeLocalMemory",
    "_required_string_list",
    "active_claude_imports",
    "applicable_section_boundary",
    "authorized_removal_for",
    "classify_claude_paragraphs",
    "claude_container_content",
    "claude_continuation_content_offset",
    "claude_explicit_container",
    "claude_import_lines",
    "claude_paragraph_text",
    "decode_git_output",
    "expected_policy_blocks",
    "fail",
    "find_protected_guide_waiver",
    "find_waiver",
    "git_inventory_safe_prefix",
    "heading_is_present",
    "index_policy_lines",
    "is_policy_heading",
    "is_policy_thematic_break",
    "load_schema_validated_yaml",
    "mask_matched_claude_code_spans",
    "normalize_policy_paragraph",
    "operative_markdown_lines",
    "os_error_diagnostic",
    "parse_contracts",
    "parse_policy_body",
    "parse_protected_guide_reference_obligations",
    "parse_protected_guide_section_obligations",
    "parse_required_sections",
    "policy_code_span_ends",
    "policy_html_block_end",
    "policy_inventory_digest",
    "policy_list_content_indent",
    "print_report",
    "protected_guide_obligation_applies",
    "quoted_policy_paragraph",
    "read_instruction_file",
    "run_bounded_git",
    "sanitized_git_environment",
    "section_applies",
    "section_body",
    "section_boundary_failure",
    "section_failures",
    "serialized_policy_blocks",
    "starts_policy_list",
    "strip_fenced_code_blocks",
    "tracked_claude_local_memory",
    "validate_contracts",
]

DEFAULT_CONTRACTS_PATH = ".template-sync/instruction-contracts.yml"


DEFAULT_CONTRACTS_SCHEMA_PATH = "schemas/template-sync-instruction-contracts.schema.json"


VALIDATION_MODES = ("upstream-template", "downstream")


def parse_args(argv: list[str]) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Validate required headings, phrases, and bounded policy sections in protected instruction files "
            "declared by .template-sync/instruction-contracts.yml."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Validation modes:\n"
            "  --mode upstream-template validates every contract against the template's "
            "own files, ignores marker-derived gating, and never fails merely because "
            ".template-sync/marker.yml is absent.\n"
            "  --mode downstream validates only contracts whose requires_modules are all "
            "listed in template_sync.included_modules from .template-sync/marker.yml.\n\n"
            "Warnings:\n"
            "  --mode upstream-template emits a non-blocking warning when the marker "
            "path exists, because that usually means the command is running against a "
            "downstream working tree that should use --mode downstream."
        ),
    )
    parser.add_argument(
        "--mode",
        choices=VALIDATION_MODES,
        required=True,
        help="Required validation mode. No implicit mode detection is performed.",
    )
    parser.add_argument(
        "--repo-root",
        default=None,
        help=(
            "Repository root to validate. Defaults to the parent of the .template-sync "
            "directory that contains this script."
        ),
    )
    parser.add_argument(
        "--contracts",
        default=DEFAULT_CONTRACTS_PATH,
        help=(
            "Instruction contract path relative to the repository root. "
            f"Default: {DEFAULT_CONTRACTS_PATH}"
        ),
    )
    parser.add_argument(
        "--contracts-schema",
        default=DEFAULT_CONTRACTS_SCHEMA_PATH,
        help=(
            "Instruction contract JSON Schema path relative to the repository root. "
            f"Default: {DEFAULT_CONTRACTS_SCHEMA_PATH}"
        ),
    )
    parser.add_argument(
        "--marker",
        default=validate_marker.DEFAULT_MARKER_PATH,
        help=(
            "Marker path relative to the repository root for --mode downstream. "
            f"Default: {validate_marker.DEFAULT_MARKER_PATH}"
        ),
    )
    parser.add_argument(
        "--marker-schema",
        default=validate_marker.DEFAULT_MARKER_SCHEMA_PATH,
        help=(
            "Marker JSON Schema path relative to the repository root for --mode downstream. "
            f"Default: {validate_marker.DEFAULT_MARKER_SCHEMA_PATH}"
        ),
    )
    parser.add_argument(
        "--manifest",
        default=validate_marker.DEFAULT_MANIFEST_PATH,
        help=(
            "Manifest path relative to the repository root. "
            f"Default: {validate_marker.DEFAULT_MANIFEST_PATH}"
        ),
    )
    parser.add_argument(
        "--manifest-schema",
        default=validate_marker.DEFAULT_MANIFEST_SCHEMA_PATH,
        help=(
            "Manifest JSON Schema path relative to the repository root. "
            f"Default: {validate_marker.DEFAULT_MANIFEST_SCHEMA_PATH}"
        ),
    )
    parser.add_argument(
        "--require-marker",
        action="store_true",
        help=(
            "In --mode downstream, fail when the marker file is absent instead of "
            "treating the run as a no-op."
        ),
    )
    parser.add_argument(
        "--skip-if-marker-present",
        action="store_true",
        help=(
            "In --mode upstream-template, exit 0 immediately without validating "
            "when the marker file exists. Intended for shared pre-commit and CI "
            "hooks that run upstream-template validation in the upstream template "
            "repository but should defer to --mode downstream in downstream forks."
        ),
    )
    return parser.parse_args(argv)


def load_manifest_modules(
    manifest_path: Path,
    manifest_schema_path: Path,
    repo_root: Path,
) -> set[str]:
    """Load the template manifest and return its declared module names."""
    manifest = load_schema_validated_yaml(manifest_path, manifest_schema_path, repo_root)
    module_names, _mappings = validate_marker.parse_manifest_mappings(manifest)
    return set(module_names)


def load_contracts(
    contracts_path: Path,
    contracts_schema_path: Path,
    manifest_path: Path,
    manifest_schema_path: Path,
    repo_root: Path,
) -> tuple[InstructionContract, ...]:
    """Load and validate instruction contracts against the manifest taxonomy."""
    manifest_modules = load_manifest_modules(manifest_path, manifest_schema_path, repo_root)
    contracts_document = load_schema_validated_yaml(
        contracts_path, contracts_schema_path, repo_root
    )
    return parse_contracts(contracts_document, manifest_modules)


def parse_instruction_contract_waivers(
    marker_document: dict[str, Any],
) -> tuple[InstructionContractWaiver, ...]:
    """Extract normalized instruction-contract waivers from a marker document."""
    template_sync = marker_document.get("template_sync")
    if not isinstance(template_sync, dict):
        raise InstructionContractValidationError("Marker must contain template_sync mapping.")
    template_sync = cast(dict[str, object], template_sync)

    waivers: list[InstructionContractWaiver] = []
    seen_pairs: set[tuple[str, str]] = set()
    duplicate_pairs: set[tuple[str, str]] = set()
    raw_waivers = template_sync.get("instruction_contract_waivers", [])
    if not isinstance(raw_waivers, list):
        raise InstructionContractValidationError("instruction_contract_waivers must be a list.")
    for raw_waiver in cast(list[object], raw_waivers):
        if not isinstance(raw_waiver, dict):
            raise InstructionContractValidationError(
                "Each instruction contract waiver must be a mapping."
            )
        raw_waiver = cast(dict[str, object], raw_waiver)
        raw_path = raw_waiver.get("path")
        anchor = raw_waiver.get("anchor")
        reason = raw_waiver.get("reason")
        authorization_basis = raw_waiver.get("authorization_basis")
        if (
            not isinstance(raw_path, str)
            or not isinstance(anchor, str)
            or not isinstance(reason, str)
            or not isinstance(authorization_basis, str)
        ):
            raise InstructionContractValidationError(
                "Each instruction contract waiver must define string path, anchor, "
                "reason, and authorization_basis."
            )
        path, is_directory = normalize_repository_path(
            raw_path,
            "template_sync.instruction_contract_waivers[].path",
        )
        if is_directory:
            raise InstructionContractValidationError(
                "template_sync.instruction_contract_waivers[].path must reference a file, "
                f"not a directory: {raw_path}"
            )
        waiver_pair = (path, anchor)
        if waiver_pair in seen_pairs:
            duplicate_pairs.add(waiver_pair)
        seen_pairs.add(waiver_pair)
        waivers.append(
            InstructionContractWaiver(
                path=path,
                anchor=anchor,
                reason=reason,
                authorization_basis=authorization_basis,
            )
        )
    if duplicate_pairs:
        formatted_pairs = ", ".join(
            f"({path}, {anchor})" for path, anchor in sorted(duplicate_pairs)
        )
        raise InstructionContractValidationError(
            "Duplicate template_sync.instruction_contract_waivers (path, anchor) "
            f"pair(s): {formatted_pairs}"
        )
    return tuple(waivers)


def load_marker_for_downstream(
    marker_path: Path,
    marker_schema_path: Path,
    manifest_modules: set[str],
    repo_root: Path,
) -> tuple[
    set[str],
    tuple[validate_marker.ProtectedFileDecision, ...],
    tuple[InstructionContractWaiver, ...],
    tuple[validate_marker.ProtectedGuideContractWaiver, ...],
]:
    """Load downstream marker state needed by instruction-contract validation."""
    marker = load_schema_validated_yaml(marker_path, marker_schema_path, repo_root)
    marker_data = validate_marker.parse_marker_decision_data(marker)
    included_modules = set(marker_data.included_modules)
    protected_decisions = marker_data.protected_decisions
    unknown_included_modules = included_modules - manifest_modules
    if unknown_included_modules:
        raise InstructionContractValidationError(
            "Marker includes module(s) that are not defined by the manifest: "
            + ", ".join(sorted(unknown_included_modules))
        )
    return (
        included_modules,
        protected_decisions,
        parse_instruction_contract_waivers(marker),
        marker_data.protected_guide_contract_waivers,
    )


def main(argv: list[str] | None = None) -> int:
    """Run instruction-contract validation."""
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        repo_root = validate_marker.resolve_repo_root(args.repo_root)
        profile_path = validate_marker.resolve_repo_path(
            repo_root, ".github/instruction-profile.yml"
        )
        if profile_path.exists():
            profile = validate_marker.load_yaml_mapping(
                profile_path, repo_root, maximum_bytes=MAXIMUM_INPUT_BYTES
            )
            if profile.get("mode") != "marker":
                raise InstructionContractValidationError(
                    "Conflicting modes: the marker adapter cannot run an active standalone profile."
                )
        contracts_path = validate_marker.resolve_repo_path(repo_root, args.contracts)
        contracts_schema_path = validate_marker.resolve_repo_path(repo_root, args.contracts_schema)
        marker_path = validate_marker.resolve_repo_path(repo_root, args.marker)
        marker_schema_path = validate_marker.resolve_repo_path(repo_root, args.marker_schema)
        manifest_path = validate_marker.resolve_repo_path(repo_root, args.manifest)
        manifest_schema_path = validate_marker.resolve_repo_path(repo_root, args.manifest_schema)

        if args.mode == "downstream" and not marker_path.exists():
            marker_relative_path = validate_marker.repository_relative_path(marker_path, repo_root)
            if args.require_marker:
                raise InstructionContractValidationError(
                    f"Marker is required but was not found at {marker_relative_path}."
                )
            print(
                f"No marker found at {marker_relative_path}; "
                "instruction-contract validation skipped."
            )
            return 0

        manifest_modules = load_manifest_modules(manifest_path, manifest_schema_path, repo_root)
        contracts_document = load_schema_validated_yaml(
            contracts_path,
            contracts_schema_path,
            repo_root,
        )
        contracts = parse_contracts(contracts_document, manifest_modules)
        protected_guide_section_obligations = parse_protected_guide_section_obligations(
            contracts_document,
            manifest_modules,
        )
        warnings: tuple[str, ...] = ()

        if args.mode == "upstream-template":
            marker_relative_path = validate_marker.repository_relative_path(marker_path, repo_root)
            if marker_path.exists():
                if args.skip_if_marker_present:
                    print(
                        f"--mode upstream-template skipped: {marker_relative_path} "
                        "is present and --skip-if-marker-present was supplied."
                    )
                    return 0
                warnings = (
                    (
                        "--mode upstream-template was invoked while "
                        f"{marker_relative_path} is present; use --mode downstream for "
                        "marker-gated downstream validation."
                    ),
                )
            report = validate_contracts(
                mode=args.mode,
                repo_root=repo_root,
                contracts=contracts,
                protected_guide_section_obligations=protected_guide_section_obligations,
                warnings=warnings,
            )
        else:
            (
                included_modules,
                protected_decisions,
                waivers,
                protected_guide_waivers,
            ) = load_marker_for_downstream(
                marker_path,
                marker_schema_path,
                manifest_modules,
                repo_root,
            )
            agent_modules = {
                module
                for contract in contracts
                for module in contract.requires_modules
                if module.startswith("agent-") and module != "agent-instructions"
            }
            if (
                "agent-instructions" in included_modules
                and agent_modules
                and not included_modules.intersection(agent_modules)
            ):
                retained_agent_paths = [
                    contract.path
                    for contract in contracts
                    if set(contract.requires_modules).intersection(agent_modules)
                    and (repo_root / contract.path).exists()
                ]
                if retained_agent_paths:
                    raise InstructionContractValidationError(
                        "Legacy agent selection requires explicit agent modules before validation: "
                        + ", ".join(retained_agent_paths)
                    )
            report = validate_contracts(
                mode=args.mode,
                repo_root=repo_root,
                contracts=contracts,
                protected_guide_section_obligations=protected_guide_section_obligations,
                included_modules=included_modules,
                protected_decisions=protected_decisions,
                waivers=waivers,
                protected_guide_waivers=protected_guide_waivers,
            )
    except (
        InstructionContractValidationError,
        validate_marker.MarkerValidationError,
    ) as error:
        fail(str(error))

    print_report(report)
    return 1 if report.has_failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
