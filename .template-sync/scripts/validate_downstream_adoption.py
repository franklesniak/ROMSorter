"""Validate a downstream repository's retained template adoption state."""

# ruff: noqa: E402

from __future__ import annotations

import argparse
import sys
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import NoReturn

_SHARED_SCRIPTS = Path(__file__).resolve().parents[2] / ".github" / "scripts"
if str(_SHARED_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SHARED_SCRIPTS))

import instruction_contract_core
import validate_instruction_contracts
import validate_marker
from template_sync_materialization_helpers import (
    INLINE_BLOCK_ANY_MODULES,
    InlineBlockError,
    ManifestMapping,
    PathRelation,
    collect_live_inline_block_spans,
    inline_block_module_requirement,
)

DEFAULT_CONTRACTS_PATH = validate_instruction_contracts.DEFAULT_CONTRACTS_PATH
DEFAULT_CONTRACTS_SCHEMA_PATH = validate_instruction_contracts.DEFAULT_CONTRACTS_SCHEMA_PATH
MARKDOWN_FILE_SUFFIXES = frozenset({".md", ".mdc"})
REQUIRED_TEMPLATE_SYNC_SUPPORT_FILES = (
    ".template-sync/scripts/validate_marker.py",
    ".template-sync/scripts/validate_instruction_contracts.py",
    ".template-sync/scripts/validate_downstream_adoption.py",
)
REQUIRED_TEMPLATE_SYNC_SCHEMA_FILES = (
    "schemas/template-sync-manifest.schema.json",
    "schemas/template-sync-marker.schema.json",
    "schemas/template-sync-instruction-contracts.schema.json",
)


class DownstreamAdoptionValidationError(Exception):
    """Raised when downstream adoption validation cannot continue."""


@dataclass(frozen=True)
class ValidationCommand:
    """A composed validator or in-process check considered by this helper."""

    name: str
    command: str


@dataclass(frozen=True)
class InlineBlockFailure:
    """A template-sync inline block inconsistency in a retained file."""

    path: str
    line_number: int
    message: str


@dataclass(frozen=True)
class MarkdownLinkFailure:
    """A retained Markdown file link to an excluded template-owned target."""

    path: str
    line_number: int
    target: str
    target_path: str
    relation: validate_marker.PathRelation


ProtectedGuideReferenceFinding = instruction_contract_core.ProtectedGuideReferenceFinding


@dataclass(frozen=True)
class DownstreamAdoptionReport:
    """Aggregate downstream adoption validation result."""

    retained_modules: tuple[str, ...]
    excluded_modules: tuple[str, ...]
    failures: tuple[str, ...]
    warnings: tuple[str, ...]
    waivers: tuple[str, ...]
    deferred_items: tuple[str, ...]
    protected_decisions: tuple[str, ...]
    commands_considered: tuple[ValidationCommand, ...]

    @property
    def has_failures(self) -> bool:
        """Return whether any composed validation check failed."""
        return bool(self.failures)

    @property
    def has_waivers(self) -> bool:
        """Return whether the report includes visible waiver-like skips."""
        return bool(self.waivers)


def parse_args(argv: list[str]) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Validate downstream partial template adoption using "
            ".template-sync/marker.yml as the source of truth."
        ),
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
        "--marker",
        default=validate_marker.DEFAULT_MARKER_PATH,
        help=(
            "Marker path relative to the repository root. "
            f"Default: {validate_marker.DEFAULT_MARKER_PATH}"
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
        "--marker-schema",
        default=validate_marker.DEFAULT_MARKER_SCHEMA_PATH,
        help=(
            "Marker JSON Schema path relative to the repository root. "
            f"Default: {validate_marker.DEFAULT_MARKER_SCHEMA_PATH}"
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
        "--require-marker",
        action="store_true",
        help="Fail when the marker file is absent instead of treating the run as a no-op.",
    )
    return parser.parse_args(argv)


def validation_commands(require_marker: bool) -> tuple[ValidationCommand, ...]:
    """Return composed command surfaces considered by this helper."""
    marker_suffix = " --require-marker" if require_marker else ""
    return (
        ValidationCommand(
            name="Marker-aware retained-file validation",
            command=f"python .template-sync/scripts/validate_marker.py{marker_suffix}",
        ),
        ValidationCommand(
            name="Downstream instruction-contract validation",
            command=(
                "python .template-sync/scripts/validate_instruction_contracts.py "
                f"--mode downstream{marker_suffix}"
            ),
        ),
        ValidationCommand(
            name="Inline-block consistency",
            command="in-process check: template-sync inline blocks match retained modules",
        ),
        ValidationCommand(
            name="Retained Markdown relative links",
            command="in-process check: retained Markdown links do not target excluded modules",
        ),
        ValidationCommand(
            name="Protected guide references",
            command=(
                "in-process check: protected-guide section and reference obligations "
                "are pruned or explicitly waived"
            ),
        ),
    )


def load_validated_manifest_context(
    repo_root: Path,
    manifest_path: Path,
    manifest_schema_path: Path,
) -> tuple[set[str], tuple[ManifestMapping, ...]]:
    """Load a schema-valid manifest and return its modules and mappings."""
    manifest = validate_marker.load_yaml_mapping(manifest_path, repo_root)
    manifest_schema = validate_marker.load_json_mapping(manifest_schema_path, repo_root)
    validate_marker.validate_schema(manifest, manifest_schema, manifest_path, repo_root)
    module_names, mappings = validate_marker.parse_manifest_mappings(manifest)
    return set(module_names), mappings


def load_validated_marker_context(
    repo_root: Path,
    marker_path: Path,
    marker_schema_path: Path,
    manifest_modules: set[str],
) -> tuple[
    set[str],
    tuple[validate_marker.LocalOverride, ...],
    tuple[validate_marker.DeferredProtectedCandidate, ...],
    tuple[validate_marker.ProtectedFileDecision, ...],
    tuple[validate_marker.ProtectedGuideContractWaiver, ...],
]:
    """Load a schema-valid marker and return parsed downstream state."""
    marker = validate_marker.load_yaml_mapping(marker_path, repo_root)
    marker_schema = validate_marker.load_json_mapping(marker_schema_path, repo_root)
    validate_marker.validate_schema(marker, marker_schema, marker_path, repo_root)
    marker_data = validate_marker.parse_marker_decision_data(marker)
    included_modules = set(marker_data.included_modules)
    local_overrides = marker_data.local_overrides
    deferred_candidates = marker_data.deferred_candidates
    protected_decisions = marker_data.protected_decisions
    unknown_modules = included_modules - manifest_modules
    if unknown_modules:
        raise DownstreamAdoptionValidationError(
            "Marker includes module(s) that are not defined by the manifest: "
            + ", ".join(sorted(unknown_modules))
        )
    return (
        included_modules,
        local_overrides,
        deferred_candidates,
        protected_decisions,
        marker_data.protected_guide_contract_waivers,
    )


def marker_report_failures(report: validate_marker.MarkerValidationReport) -> tuple[str, ...]:
    """Return human-readable failures from marker-aware validation."""
    failures: list[str] = []
    for relative_path in report.unsafe_managed_paths:
        failures.append(f"Template-managed path is a symlink or resolves unsafely: {relative_path}")
    for relative_path, relation in report.leftover_files:
        failures.append(
            f"Excluded-module leftover is present: {relative_path} ({relation.description})"
        )
    for relative_path, relation in report.missing_expected_files:
        failures.append(
            f"Retained concrete mapped file is missing: {relative_path} ({relation.description})"
        )
    return tuple(failures)


def marker_report_warnings(report: validate_marker.MarkerValidationReport) -> tuple[str, ...]:
    """Return human-readable warnings from marker-aware validation."""
    warnings: list[str] = []
    for relative_path in report.unsafe_local_paths:
        warnings.append(
            f"Git-visible local path is a symlink or resolves unsafely: {relative_path}"
        )
    # Bound the per-path warnings to the same sample size validate_marker.py
    # prints, so a downstream repo with a large unrecorded set cannot flood the
    # aggregate adoption report or CI logs.
    unrecorded_limit = validate_marker.UNRECORDED_LOCAL_PATH_LIMIT
    sampled_unrecorded = report.unrecorded_local_paths[:unrecorded_limit]
    for relative_path in sampled_unrecorded:
        warnings.append(
            f"Git-visible path is neither template-managed nor recorded in template_sync.local_path_ownership: {relative_path}"
        )
    remaining_unrecorded = len(report.unrecorded_local_paths) - len(sampled_unrecorded)
    if remaining_unrecorded > 0:
        warnings.append(
            f"... and {remaining_unrecorded} more unrecorded Git-visible local path(s) not shown."
        )
    return tuple(warnings)


def retained_support_file_failures(repo_root: Path, included_modules: set[str]) -> tuple[str, ...]:
    """Return missing retained template-sync helper and schema contract files."""
    required_paths: list[str] = []
    if "template-sync-support" in included_modules:
        required_paths.extend(REQUIRED_TEMPLATE_SYNC_SUPPORT_FILES)
        required_paths.extend(REQUIRED_TEMPLATE_SYNC_SCHEMA_FILES)

    failures: list[str] = []
    for relative_path in required_paths:
        path = validate_marker.resolve_repo_path(repo_root, relative_path)
        if not path.is_file():
            failures.append(f"Retained template-sync support file is missing: {relative_path}")
    return tuple(failures)


def instruction_contract_report_items(
    *,
    repo_root: Path,
    contracts_path: Path,
    contracts_schema_path: Path,
    marker_path: Path,
    marker_schema_path: Path,
    manifest_path: Path,
    manifest_schema_path: Path,
) -> tuple[tuple[str, ...], tuple[str, ...], tuple[str, ...]]:
    """Run downstream instruction-contract validation and return report items."""
    manifest_modules = validate_instruction_contracts.load_manifest_modules(
        manifest_path,
        manifest_schema_path,
        repo_root,
    )
    contracts_document = validate_instruction_contracts.load_schema_validated_yaml(
        contracts_path,
        contracts_schema_path,
        repo_root,
    )
    contracts = validate_instruction_contracts.parse_contracts(
        contracts_document,
        manifest_modules,
    )
    protected_guide_section_obligations = (
        validate_instruction_contracts.parse_protected_guide_section_obligations(
            contracts_document,
            manifest_modules,
        )
    )
    included_modules, protected_decisions, waivers, protected_guide_waivers = (
        validate_instruction_contracts.load_marker_for_downstream(
            marker_path,
            marker_schema_path,
            manifest_modules,
            repo_root,
        )
    )
    report = validate_instruction_contracts.validate_contracts(
        mode="downstream",
        repo_root=repo_root,
        contracts=contracts,
        protected_guide_section_obligations=protected_guide_section_obligations,
        included_modules=included_modules,
        protected_decisions=protected_decisions,
        waivers=waivers,
        protected_guide_waivers=protected_guide_waivers,
    )

    failures: list[str] = []
    for missing_file in report.missing_files:
        failures.append(
            "Required downstream instruction contract file is absent without "
            f"authorized removal: {missing_file.path}"
        )
    for missing_anchor in report.missing_anchors:
        failures.append(
            "Required downstream instruction contract anchor is missing: "
            f"{missing_anchor.path}: {missing_anchor.anchor_type}: {missing_anchor.anchor}"
        )
    for stale_section in report.stale_protected_guide_sections:
        failures.append(
            "Protected guide section requires owner review for excluded module(s): "
            f"{stale_section.path}: {stale_section.contract_key}: "
            f"{stale_section.anchor_type}: {stale_section.anchor} "
            f"(target modules: {', '.join(stale_section.target_modules)})"
        )

    waiver_items: list[str] = []
    for waiver in report.applied_waivers:
        waiver_items.append(
            "Instruction contract waiver: "
            f"{waiver.path}: {waiver.anchor} "
            f"(basis: {waiver.authorization_basis}; reason: {waiver.reason})"
        )
    for authorized_removal in report.authorized_removals:
        waiver_items.append(
            "Authorized protected instruction-file removal: "
            f"{authorized_removal.path} "
            f"(basis: {authorized_removal.authorization_basis}; "
            f"scope: {authorized_removal.authorized_scope}; "
            f"reason: {authorized_removal.reason})"
        )
    for protected_guide_waiver in report.applied_protected_guide_waivers:
        target_details = []
        if protected_guide_waiver.target_path is not None:
            target_details.append(f"target_path: {protected_guide_waiver.target_path}")
        if protected_guide_waiver.target_module is not None:
            target_details.append(f"target_module: {protected_guide_waiver.target_module}")
        target_suffix = "; " + "; ".join(target_details) if target_details else ""
        waiver_items.append(
            "Protected guide contract waiver: "
            f"{protected_guide_waiver.path}: {protected_guide_waiver.contract_key}{target_suffix} "
            f"(basis: {protected_guide_waiver.authorization_basis}; "
            f"reason: {protected_guide_waiver.reason})"
        )

    return tuple(failures), tuple(report.warnings), tuple(waiver_items)


def retained_relation(
    relative_path: str,
    mappings: tuple[ManifestMapping, ...],
    included_modules: set[str],
) -> PathRelation | None:
    """Return the selected retained relation for a path, if any."""
    relation = validate_marker.selected_relation_for_path(relative_path, mappings)
    if relation is None or not relation.is_retained_by(included_modules):
        return None
    return relation


def iter_present_safe_files(repo_root: Path) -> tuple[str, ...]:
    """Return Git-visible files safely discoverable in the working tree."""
    present_paths = set(validate_marker.git_present_paths(repo_root))
    repository_files, _skipped_symlinks = validate_marker.iter_safe_repository_files(repo_root)
    return tuple(path for path in repository_files if path in present_paths)


def validate_inline_blocks(
    repo_root: Path,
    relative_paths: Iterable[str],
    included_modules: set[str],
) -> tuple[InlineBlockFailure, ...]:
    """Validate template-sync inline block families against retained modules."""
    failures: list[InlineBlockFailure] = []
    for relative_path in relative_paths:
        path = validate_marker.resolve_repo_path(repo_root, relative_path)
        try:
            text = path.read_text(encoding="utf-8")
        except OSError as error:
            error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
            failures.append(
                InlineBlockFailure(
                    path=relative_path,
                    line_number=0,
                    message=f"Unable to read retained file for inline blocks: {error_summary}",
                )
            )
            continue

        try:
            blocks = collect_live_inline_block_spans(
                text,
                relative_path=relative_path,
            )
        except InlineBlockError as error:
            # Store the bare message: the renderer prepends "{path}:{line}: "
            # itself, so str(error) (which already embeds that prefix) would
            # duplicate the location in the final output.
            failures.append(
                InlineBlockFailure(
                    path=relative_path,
                    line_number=error.line_number or 0,
                    message=error.message,
                )
            )
            continue

        for block in blocks:
            marker_name = block.marker_name
            module_requirement = inline_block_module_requirement(marker_name)
            if module_requirement is None:
                failures.append(
                    InlineBlockFailure(
                        path=relative_path,
                        line_number=block.line_number,
                        message=f"Unknown template-sync inline marker family: {marker_name}",
                    )
                )
            elif marker_name in INLINE_BLOCK_ANY_MODULES:
                # OR-retention: the block is valid while at least one named module
                # is included, and is stale only when every named module is excluded.
                if module_requirement.isdisjoint(included_modules):
                    excluded = ", ".join(sorted(module_requirement))
                    failures.append(
                        InlineBlockFailure(
                            path=relative_path,
                            line_number=block.line_number,
                            message=(
                                f"Inline block {marker_name} remains but requires at "
                                f"least one of the excluded module(s): {excluded}"
                            ),
                        )
                    )
            elif not module_requirement.issubset(included_modules):
                excluded = ", ".join(sorted(module_requirement - included_modules))
                failures.append(
                    InlineBlockFailure(
                        path=relative_path,
                        line_number=block.line_number,
                        message=(
                            f"Inline block {marker_name} remains but requires "
                            f"excluded module(s): {excluded}"
                        ),
                    )
                )

    return tuple(failures)


def markdown_link_targets_outside_fences(path: Path) -> tuple[tuple[int, str], ...]:
    """Return Markdown link targets using the shared fence-aware matcher."""
    return instruction_contract_core.markdown_link_targets_from_text(
        path.read_text(encoding="utf-8")
    )


normalize_markdown_target = instruction_contract_core.normalize_markdown_target
resolve_relative_markdown_target = instruction_contract_core.resolve_relative_markdown_target


def validate_retained_markdown_links(
    repo_root: Path,
    relative_paths: Iterable[str],
    mappings: tuple[ManifestMapping, ...],
    included_modules: set[str],
    local_overrides: tuple[validate_marker.LocalOverride, ...],
) -> tuple[MarkdownLinkFailure, ...]:
    """Validate retained Markdown links do not point to excluded template files."""
    failures: list[MarkdownLinkFailure] = []
    for relative_path in relative_paths:
        if Path(relative_path).suffix not in MARKDOWN_FILE_SUFFIXES:
            continue
        if validate_marker.is_locally_overridden(relative_path, local_overrides):
            continue
        if retained_relation(relative_path, mappings, included_modules) is None:
            continue

        path = validate_marker.resolve_repo_path(repo_root, relative_path)
        try:
            targets = markdown_link_targets_outside_fences(path)
        except OSError as error:
            error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
            failures.append(
                MarkdownLinkFailure(
                    path=relative_path,
                    line_number=0,
                    target="<unreadable>",
                    target_path=f"Unable to read retained Markdown file: {error_summary}",
                    relation=validate_marker.PathRelation(
                        patterns=(),
                        requires_all=frozenset(),
                        requires_any=frozenset(),
                    ),
                )
            )
            continue

        for line_number, target in targets:
            target_path = resolve_relative_markdown_target(relative_path, target)
            if target_path is None:
                continue
            target_relation = validate_marker.selected_relation_for_path(target_path, mappings)
            if target_relation is None or target_relation.is_retained_by(included_modules):
                continue
            failures.append(
                MarkdownLinkFailure(
                    path=relative_path,
                    line_number=line_number,
                    target=target,
                    target_path=target_path,
                    relation=target_relation,
                )
            )
    return tuple(failures)


def protected_guide_reference_obligation_applies(
    obligation: validate_instruction_contracts.ProtectedGuideReferenceObligation,
    mappings: tuple[ManifestMapping, ...],
    included_modules: set[str],
) -> bool:
    """Return whether a protected-guide reference obligation is stale downstream."""
    if retained_relation(obligation.path, mappings, included_modules) is None:
        return False
    if obligation.target_path is not None:
        target_relation = validate_marker.selected_relation_for_path(
            obligation.target_path,
            mappings,
        )
        # An excluded target_path is stale on its own. A retained target_path does
        # not clear the obligation: fall through to the target_modules check so a
        # reference whose target modules are all excluded is still flagged (for
        # example, the path lingers via a local override while the modules it
        # documents are gone).
        if target_relation is not None and not target_relation.is_retained_by(included_modules):
            return True
    return validate_instruction_contracts.protected_guide_obligation_applies(
        obligation.target_modules,
        included_modules,
    )


def protected_guide_markdown_reference_findings(
    *,
    repo_root: Path,
    obligation: validate_instruction_contracts.ProtectedGuideReferenceObligation,
    mappings: tuple[ManifestMapping, ...],
    included_modules: set[str],
) -> tuple[ProtectedGuideReferenceFinding, ...]:
    """Preserve marker path applicability while sharing reference matching."""
    if obligation.target_path is None:
        return ()
    if Path(obligation.path).suffix not in MARKDOWN_FILE_SUFFIXES:
        return ()
    if retained_relation(obligation.path, mappings, included_modules) is None:
        return ()
    target_relation = validate_marker.selected_relation_for_path(obligation.target_path, mappings)
    if target_relation is None or target_relation.is_retained_by(included_modules):
        return ()
    return instruction_contract_core.protected_guide_reference_findings(
        repo_root=repo_root, obligation=obligation
    )


protected_guide_token_reference_findings = (
    instruction_contract_core.protected_guide_reference_findings
)


def protected_guide_reference_report_items(
    *,
    repo_root: Path,
    contracts_path: Path,
    contracts_schema_path: Path,
    manifest_modules: set[str],
    mappings: tuple[ManifestMapping, ...],
    included_modules: set[str],
    protected_guide_waivers: tuple[validate_marker.ProtectedGuideContractWaiver, ...],
) -> tuple[tuple[str, ...], tuple[str, ...]]:
    """Return failures and waiver items for protected-guide reference obligations."""
    contracts_document = validate_instruction_contracts.load_schema_validated_yaml(
        contracts_path,
        contracts_schema_path,
        repo_root,
    )
    obligations = validate_instruction_contracts.parse_protected_guide_reference_obligations(
        contracts_document,
        manifest_modules,
    )

    failures: list[str] = []
    waiver_items: list[str] = []
    applied_waivers: list[validate_marker.ProtectedGuideContractWaiver] = []
    for obligation in obligations:
        if not protected_guide_reference_obligation_applies(
            obligation,
            mappings,
            included_modules,
        ):
            continue
        if obligation.reference_kind == "markdown-relative-link":
            findings = protected_guide_markdown_reference_findings(
                repo_root=repo_root,
                obligation=obligation,
                mappings=mappings,
                included_modules=included_modules,
            )
        else:
            findings = protected_guide_token_reference_findings(
                repo_root=repo_root,
                obligation=obligation,
            )
        if not findings:
            continue

        waiver = validate_instruction_contracts.find_protected_guide_waiver(
            protected_guide_waivers,
            path=obligation.path,
            contract_key=obligation.key,
            target_modules=obligation.target_modules,
            target_path=obligation.target_path,
        )
        if waiver is not None:
            applied_waivers.append(waiver)
            continue

        for finding in findings:
            line = f":{finding.line_number}" if finding.line_number else ""
            target_path = f" -> {finding.target_path}" if finding.target_path is not None else ""
            failures.append(
                "Protected guide reference requires owner review for excluded module(s): "
                f"{finding.path}{line}: {finding.contract_key}: "
                f"{finding.reference_kind}: {finding.target}{target_path} "
                f"(target modules: {', '.join(finding.target_modules)})"
            )

    for waiver in dict.fromkeys(applied_waivers):
        target_details = []
        if waiver.target_path is not None:
            target_details.append(f"target_path: {waiver.target_path}")
        if waiver.target_module is not None:
            target_details.append(f"target_module: {waiver.target_module}")
        target_suffix = "; " + "; ".join(target_details) if target_details else ""
        waiver_items.append(
            "Protected guide contract waiver: "
            f"{waiver.path}: {waiver.contract_key}{target_suffix} "
            f"(basis: {waiver.authorization_basis}; reason: {waiver.reason})"
        )
    return tuple(failures), tuple(waiver_items)


def local_override_waiver_items(
    local_overrides: tuple[validate_marker.LocalOverride, ...],
) -> tuple[str, ...]:
    """Return visible waiver items for marker local overrides."""
    return tuple(
        "Local override: "
        f"{local_override.path}{'/' if local_override.is_directory else ''} "
        f"({local_override.default_decision}; reason: {local_override.reason})"
        for local_override in local_overrides
    )


def deferred_items(
    candidates: tuple[validate_marker.DeferredProtectedCandidate, ...],
) -> tuple[str, ...]:
    """Return visible deferred protected-file candidate items."""
    return tuple(
        f"{candidate.path} at {candidate.source_commit}: {candidate.reason}"
        for candidate in candidates
    )


def protected_decision_items(
    protected_decisions: tuple[validate_marker.ProtectedFileDecision, ...],
) -> tuple[str, ...]:
    """Return visible protected-file decision items."""
    items: list[str] = []
    for decision in protected_decisions:
        details = [f"{decision.path}: {decision.decision}"]
        if decision.adoption_mode is not None:
            details.append(f"adoption_mode={decision.adoption_mode}")
        if decision.authorization_basis is not None:
            details.append(f"authorization_basis={decision.authorization_basis}")
        if decision.authorized_scope is not None:
            details.append(f"authorized_scope={decision.authorized_scope}")
        if decision.tailored_authorization_basis is not None:
            details.append(f"tailored_authorization_basis={decision.tailored_authorization_basis}")
        if decision.reason is not None:
            details.append(f"reason={decision.reason}")
        items.append("; ".join(details))
    return tuple(items)


def build_report(
    *,
    repo_root: Path,
    marker_path: Path,
    manifest_path: Path,
    marker_schema_path: Path,
    manifest_schema_path: Path,
    contracts_path: Path,
    contracts_schema_path: Path,
    require_marker: bool,
) -> DownstreamAdoptionReport:
    """Build the aggregate downstream adoption validation report."""
    commands = validation_commands(require_marker)
    failures: list[str] = []
    warnings: list[str] = []
    waivers: list[str] = []
    deferred: list[str] = []
    protected_decisions: list[str] = []

    try:
        marker_report = validate_marker.validate_marker_state(
            repo_root=repo_root,
            marker_path=marker_path,
            manifest_path=manifest_path,
            marker_schema_path=marker_schema_path,
            manifest_schema_path=manifest_schema_path,
        )
        manifest_modules, mappings = load_validated_manifest_context(
            repo_root,
            manifest_path,
            manifest_schema_path,
        )
        (
            included_modules,
            local_overrides,
            deferred_candidates,
            parsed_protected_decisions,
            protected_guide_waivers,
        ) = load_validated_marker_context(
            repo_root,
            marker_path,
            marker_schema_path,
            manifest_modules,
        )
    except (validate_marker.MarkerValidationError, DownstreamAdoptionValidationError) as error:
        failures.append(f"Marker-aware validation: {error}")
        return DownstreamAdoptionReport(
            retained_modules=(),
            excluded_modules=(),
            failures=tuple(failures),
            warnings=(),
            waivers=(),
            deferred_items=(),
            protected_decisions=(),
            commands_considered=commands,
        )

    failures.extend(marker_report_failures(marker_report))
    warnings.extend(marker_report_warnings(marker_report))
    failures.extend(retained_support_file_failures(repo_root, included_modules))
    waivers.extend(local_override_waiver_items(local_overrides))
    deferred.extend(deferred_items(deferred_candidates))
    protected_decisions.extend(protected_decision_items(parsed_protected_decisions))

    try:
        contract_failures, contract_warnings, contract_waivers = instruction_contract_report_items(
            repo_root=repo_root,
            contracts_path=contracts_path,
            contracts_schema_path=contracts_schema_path,
            marker_path=marker_path,
            marker_schema_path=marker_schema_path,
            manifest_path=manifest_path,
            manifest_schema_path=manifest_schema_path,
        )
        failures.extend(contract_failures)
        warnings.extend(contract_warnings)
        waivers.extend(contract_waivers)
    except (
        validate_marker.MarkerValidationError,
        validate_instruction_contracts.InstructionContractValidationError,
    ) as error:
        failures.append(f"Downstream instruction-contract validation: {error}")

    present_files = iter_present_safe_files(repo_root)
    inline_failures = validate_inline_blocks(repo_root, present_files, included_modules)
    for failure in inline_failures:
        line = f":{failure.line_number}" if failure.line_number else ""
        failures.append(f"Inline-block consistency: {failure.path}{line}: {failure.message}")

    markdown_link_failures = validate_retained_markdown_links(
        repo_root,
        present_files,
        mappings,
        included_modules,
        local_overrides,
    )
    for link_failure in markdown_link_failures:
        failures.append(
            "Retained Markdown relative link targets excluded module(s): "
            f"{link_failure.path}:{link_failure.line_number}: {link_failure.target} -> "
            f"{link_failure.target_path} ({link_failure.relation.description})"
        )

    try:
        reference_failures, reference_waivers = protected_guide_reference_report_items(
            repo_root=repo_root,
            contracts_path=contracts_path,
            contracts_schema_path=contracts_schema_path,
            manifest_modules=manifest_modules,
            mappings=mappings,
            included_modules=included_modules,
            protected_guide_waivers=protected_guide_waivers,
        )
        failures.extend(reference_failures)
        waivers.extend(reference_waivers)
    except (
        validate_marker.MarkerValidationError,
        validate_instruction_contracts.InstructionContractValidationError,
    ) as error:
        failures.append(f"Protected guide reference validation: {error}")

    return DownstreamAdoptionReport(
        retained_modules=tuple(sorted(included_modules)),
        excluded_modules=tuple(sorted(manifest_modules - included_modules)),
        failures=tuple(failures),
        warnings=tuple(warnings),
        waivers=tuple(waivers),
        deferred_items=tuple(deferred),
        protected_decisions=tuple(protected_decisions),
        commands_considered=commands,
    )


def print_named_section(name: str, items: Iterable[str], *, empty_text: str = "None.") -> None:
    """Print one bullet-list report section."""
    print(f"\n{name}:")
    rendered_items = tuple(items)
    if not rendered_items:
        print(f"  {empty_text}")
        return
    for item in rendered_items:
        print(f"  - {item}")


def print_report(report: DownstreamAdoptionReport) -> None:
    """Print a human-readable aggregate validation report."""
    if report.has_failures:
        print("Downstream adoption validation failed.")
    elif report.has_waivers:
        print("Downstream adoption validation passed with waivers.")
    else:
        print("Downstream adoption validation passed.")

    print_named_section("Failures", report.failures)
    print_named_section("Warnings", report.warnings)
    print_named_section("Waivers", report.waivers)
    print_named_section("Deferred items", report.deferred_items)
    print_named_section("Protected file decisions", report.protected_decisions)
    print_named_section("Retained modules", report.retained_modules)
    print_named_section("Excluded modules", report.excluded_modules)
    print("\nValidation commands considered:")
    for command in report.commands_considered:
        print(f"  - {command.name}: {command.command}")


def fail(message: str) -> NoReturn:
    """Print an error and exit non-zero."""
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(1)


def main(argv: list[str] | None = None) -> int:
    """Run aggregate downstream adoption validation."""
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        repo_root = validate_marker.resolve_repo_root(args.repo_root)
        marker_path = validate_marker.resolve_repo_path(repo_root, args.marker)
        manifest_path = validate_marker.resolve_repo_path(repo_root, args.manifest)
        marker_schema_path = validate_marker.resolve_repo_path(repo_root, args.marker_schema)
        manifest_schema_path = validate_marker.resolve_repo_path(repo_root, args.manifest_schema)
        contracts_path = validate_marker.resolve_repo_path(repo_root, args.contracts)
        contracts_schema_path = validate_marker.resolve_repo_path(repo_root, args.contracts_schema)

        if not marker_path.exists():
            marker_relative_path = validate_marker.repository_relative_path(marker_path, repo_root)
            if args.require_marker:
                raise DownstreamAdoptionValidationError(
                    f"Marker is required but was not found at {marker_relative_path}."
                )
            print(
                f"No marker found at {marker_relative_path}; "
                "downstream adoption validation skipped."
            )
            return 0

        report = build_report(
            repo_root=repo_root,
            marker_path=marker_path,
            manifest_path=manifest_path,
            marker_schema_path=marker_schema_path,
            manifest_schema_path=manifest_schema_path,
            contracts_path=contracts_path,
            contracts_schema_path=contracts_schema_path,
            require_marker=args.require_marker,
        )
    except (validate_marker.MarkerValidationError, DownstreamAdoptionValidationError) as error:
        fail(str(error))

    print_report(report)
    return 1 if report.has_failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
