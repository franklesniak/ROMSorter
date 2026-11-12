"""Render an explicit local profile while preserving reviewed marker decisions."""

# ruff: noqa: E402

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Any

import yaml  # type: ignore[import-untyped]

_SHARED_SCRIPTS = Path(__file__).resolve().parents[2] / ".github" / "scripts"
if str(_SHARED_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SHARED_SCRIPTS))

import instruction_contract_core as core
from template_sync_materialization_helpers import (
    TemplateSyncMaterializationError,
    is_protected_instruction_path,
    parse_marker_decision_data,
    resolve_safe_repository_target_path,
)

PROFILE_PATH = ".github/instruction-profile.yml"
TRUSTED_TOOL_ROOT = Path(__file__).resolve().parents[2]
SOURCE_FIELDS = (
    "local_overrides",
    "local_path_ownership",
    "protected_file_decisions",
    "instruction_contract_waivers",
    "protected_guide_contract_waivers",
)


def load_reviewed_schema(relative_path: str) -> dict[str, Any]:
    """Load schemas from the installed tool bundle, never selected source data."""
    return core.support.load_json_mapping(
        core.support.resolve_repo_path(TRUSTED_TOOL_ROOT, relative_path),
        TRUSTED_TOOL_ROOT,
        maximum_bytes=core.MAXIMUM_INPUT_BYTES,
    )


def load_existing_profile(target_root: Path) -> dict[str, Any]:
    """Read bounded local data against the installed tool's reviewed schema."""
    path = core.support.resolve_repo_path(target_root, PROFILE_PATH)
    document = core.support.load_yaml_mapping(
        path, target_root, maximum_bytes=core.MAXIMUM_INPUT_BYTES
    )
    schema = load_reviewed_schema("schemas/instruction-profile.schema.json")
    core.support.validate_schema(document, schema, path, target_root)
    return document


def catalog_path_applicability(
    contracts: tuple[core.InstructionContract, ...],
    section_obligations: tuple[core.ProtectedGuideSectionObligation, ...],
    reference_obligations: tuple[core.ProtectedGuideReferenceObligation, ...],
    modules: set[str],
) -> dict[str, bool]:
    """Classify known catalog paths, including obligation-only source files."""
    applicability: dict[str, bool] = {}
    for section_obligation in section_obligations:
        path = section_obligation.path
        applicability[path] = applicability.get(
            path, False
        ) or core.protected_guide_obligation_applies(section_obligation.target_modules, modules)
    for reference_obligation in reference_obligations:
        path = reference_obligation.path
        applicability[path] = applicability.get(
            path, False
        ) or core.protected_guide_obligation_applies(reference_obligation.target_modules, modules)
    # Core skips all obligations attached to an excluded instruction-contract path.
    for contract in contracts:
        applicability[contract.path] = set(contract.requires_modules) <= modules
    return applicability


def declaration_applies(
    declaration: dict[str, str],
    contracts: tuple[core.InstructionContract, ...],
    section_obligations: tuple[core.ProtectedGuideSectionObligation, ...],
    modules: set[str],
    reference_obligations: tuple[core.ProtectedGuideReferenceObligation, ...] = (),
) -> bool:
    """Retire known excluded obligations, while leaving unknown anchors to fail.

    Applicability comes from catalog scope, never from a failure disappearing
    after a content edit. Direct anchors may themselves resemble section IDs.
    """
    path, anchor = declaration["path"], declaration["anchor"]
    if anchor == "file:absent":
        return catalog_path_applicability(
            contracts, section_obligations, reference_obligations, modules
        ).get(path, True)
    for contract in contracts:
        if contract.path != path:
            continue
        if not set(contract.requires_modules) <= modules:
            return False
        if (
            anchor == "file:absent"
            or anchor in contract.required_headings
            or anchor in contract.required_phrases
        ):
            return True
        # Longest headings distinguish names containing colons from prefixes.
        for section in sorted(
            contract.required_sections, key=lambda item: len(item.heading), reverse=True
        ):
            prefix = f"section:{section.heading}"
            if anchor == prefix or (
                anchor.startswith(prefix + ":")
                and re.fullmatch(
                    r"(?:boundary|html-grammar|hard-break|paragraph|paragraphs|tables|blocks):[0-9a-f]{64}",
                    anchor[len(prefix) + 1 :],
                )
            ):
                return core.section_applies(section, modules)
    for obligation in section_obligations:
        if obligation.path != path:
            continue
        anchors = {
            *(f"stale:{obligation.key}:heading:{heading}" for heading in obligation.stale_headings),
            *(f"stale:{obligation.key}:phrase:{phrase}" for phrase in obligation.stale_phrases),
        }
        if anchor in anchors:
            return core.protected_guide_obligation_applies(obligation.target_modules, modules)
    for reference_obligation in reference_obligations:
        prefix = f"reference:{reference_obligation.key}:{reference_obligation.reference_kind}:"
        if reference_obligation.path != path or not anchor.startswith(prefix):
            continue
        target = anchor[len(prefix) :]
        matches = (
            core.resolve_relative_markdown_target(path, target) == reference_obligation.target_path
            if reference_obligation.reference_kind == "markdown-relative-link"
            else target in reference_obligation.tokens
        )
        if matches:
            return core.protected_guide_obligation_applies(
                reference_obligation.target_modules, modules
            )
    return True


def selected_content_root(
    relative_path: str,
    staging_root: Path,
    target_root: Path,
    marker: dict[str, Any],
) -> tuple[Path, bool]:
    """Select reconciled bytes without granting authority to unresolved decisions."""
    decision = None
    if is_protected_instruction_path(relative_path):
        decision = next(
            (
                item["decision"]
                for item in marker.get("protected_file_decisions", [])
                if item["path"] == relative_path
            ),
            None,
        )
    else:
        local_overrides = parse_marker_decision_data(
            {
                "template_sync": {
                    "included_modules": marker["included_modules"],
                    "local_overrides": marker.get("local_overrides", []),
                }
            }
        ).local_overrides
        matches = [item for item in local_overrides if item.matches(relative_path)]
        if matches:
            # Match reconciliation: longest path, exact file, then last-listed tie.
            decision = max(
                reversed(matches), key=lambda item: (len(item.path), not item.is_directory)
            ).default_decision
    if decision == "SKIP":
        return target_root, False
    if decision == "TAKE" and core.support.resolve_repo_path(staging_root, relative_path).is_file():
        return staging_root, True
    local_path = core.support.resolve_repo_path(target_root, relative_path)
    return (target_root if local_path.exists() else staging_root), False


def report_failure_keys(report: core.InstructionContractReport) -> set[tuple[str, str]]:
    """Return the exact failure keys that a standalone declaration can address."""
    failures = {(item.path, item.anchor) for item in report.missing_anchors}
    failures.update((item.path, "file:absent") for item in report.missing_files)
    failures.update(
        (item.path, f"stale:{item.contract_key}:{item.anchor_type}:{item.anchor}")
        for item in report.stale_protected_guide_sections
    )
    failures.update(
        (item.path, f"reference:{item.contract_key}:{item.reference_kind}:{item.target}")
        for item in report.stale_protected_guide_references
    )
    return failures


def validate_selected_enforcement_inputs(
    staging_root: Path, target_root: Path, marker: dict[str, Any]
) -> None:
    """Require the selected runtime inputs without overriding local ownership."""
    paths = [
        PROFILE_PATH,
        "schemas/instruction-profile.schema.json",
        ".github/scripts/validate_instruction_profile.py",
        ".github/scripts/instruction_contract_core.py",
        ".github/scripts/instruction_contract_support.py",
    ]
    if "template-sync-support" not in marker["included_modules"]:
        paths.extend(
            [".github/instruction-contracts.yml", "schemas/instruction-contracts.schema.json"]
        )
    for relative_path in paths:
        content_root, _ = selected_content_root(relative_path, staging_root, target_root, marker)
        path = resolve_safe_repository_target_path(
            content_root, relative_path, field_name="selected enforcement input"
        )
        if not path.is_file():
            raise TemplateSyncMaterializationError(
                f"Selected enforcement input is missing or not a regular file: {relative_path}. "
                "SKIP requires preserved local input; review the file or its selection."
            )


def validate_skipped_profile_applicability(target_root: Path, marker: dict[str, Any]) -> None:
    """Reject contradictory preserved applicability without changing profile bytes."""
    if not any(
        item["path"] == PROFILE_PATH and item["decision"] == "SKIP"
        for item in marker.get("protected_file_decisions", [])
    ):
        return
    local = load_existing_profile(target_root)
    modules = set(marker["included_modules"])
    expected_mode = "marker" if "template-sync-support" in modules else "standalone"
    if local["mode"] != expected_mode:
        field = "mode"
    elif expected_mode == "marker" and local["context"] != "downstream":
        field = "context"
    elif expected_mode == "standalone" and set(local["modules"]) != modules:
        field = "modules"
    else:
        return
    raise TemplateSyncMaterializationError(
        f"Preserved {PROFILE_PATH} {field} conflicts with the selected materialization. "
        "Review the local profile or its SKIP decision; local bytes are not changed."
    )


def validate_retained_catalog_selection(
    staging_root: Path,
    target_root: Path,
    marker: dict[str, Any],
    document: dict[str, Any],
    reports: dict[Path, core.InstructionContractReport],
    path_applicability: dict[str, bool],
) -> None:
    """Reject incompatible retained catalogs without changing selected local bytes."""
    removed = {
        item["path"]
        for item in marker.get("protected_file_decisions", [])
        if item["decision"] == "REMOVE-LOCAL"
    }
    selections: dict[str, Path] = {}

    def selected_root(path: str) -> Path:
        if path not in selections:
            selections[path] = selected_content_root(path, staging_root, target_root, marker)[0]
        return selections[path]

    failures = {
        (path, anchor)
        for content_root, report in reports.items()
        for path, anchor in report_failure_keys(report)
        if path not in removed and selected_root(path) == content_root
    }
    failures.update(
        (path, "file:absent") for path in removed if path_applicability.get(path, False)
    )
    effective_profile = (
        load_existing_profile(target_root)
        if selected_root(PROFILE_PATH) == target_root
        else document
    )
    if effective_profile["mode"] != "standalone":
        raise TemplateSyncMaterializationError(
            "Selected preserved instruction profile mode conflicts with standalone migration. "
            "Review the profile or supply an explicit protected selection."
        )
    applied: set[tuple[str, str]] = set()
    digests: dict[str, str] = {}
    for declaration in effective_profile["exceptions"]:
        path, anchor = declaration["path"], declaration["anchor"]
        key = (path, anchor)
        if path not in digests:
            digests[path] = (
                "absent" if path in removed else core.file_digest(selected_root(path), path)
            )
        if key in applied or key not in failures or declaration["content_sha256"] != digests[path]:
            raise TemplateSyncMaterializationError(
                f"Retained instruction catalog conflicts with selected profile exception: {path}: {anchor}. "
                "Review the selected catalog, content, or profile declaration."
            )
        applied.add(key)
    remaining = sorted(failures - applied)
    if remaining:
        path, anchor = remaining[0]
        raise TemplateSyncMaterializationError(
            f"Retained instruction catalog conflicts with selected content: {path}: {anchor}. "
            "Review the selected content or supply an exact authorized declaration."
        )


def validate_selected_claude_state(
    staging_root: Path,
    target_root: Path,
    marker: dict[str, Any],
    reports: dict[Path, core.InstructionContractReport],
) -> None:
    """Reject non-exceptable selected imports and the preserved target Git inventory."""
    removed = {
        item["path"]
        for item in marker.get("protected_file_decisions", [])
        if item["decision"] == "REMOVE-LOCAL"
    }
    selections: dict[str, Path] = {}

    def selected_root(path: str) -> Path:
        if path not in selections:
            selections[path] = selected_content_root(path, staging_root, target_root, marker)[0]
        return selections[path]

    active_imports = [
        item
        for content_root, report in reports.items()
        for item in report.active_claude_imports
        if item.path not in removed and selected_root(item.path) == content_root
    ]
    if active_imports or reports[target_root].tracked_claude_local_memory:
        raise TemplateSyncMaterializationError(
            "Selected non-exceptable Claude instruction content conflicts with standalone migration. "
            "Review active imports and tracked local memory before migration."
        )


def render_instruction_profile(
    *,
    staging_root: Path,
    target_root: Path,
    marker_document: dict[str, Any],
) -> None:
    """Preserve the materializer's controlled diagnostic boundary for core failures."""
    try:
        _render_instruction_profile(
            staging_root=staging_root, target_root=target_root, marker_document=marker_document
        )
    except core.InstructionContractValidationError as error:
        raise TemplateSyncMaterializationError(str(error)) from error


def _render_instruction_profile(
    *,
    staging_root: Path,
    target_root: Path,
    marker_document: dict[str, Any],
) -> None:
    """Render a protected candidate; ordinary reconciliation still controls writes.

    Preserve source decisions as evidence, not authority. Convert only exact
    currently applicable failures; bind each declaration to the local file
    content so a later unrelated edit cannot reuse it silently.
    """
    marker = marker_document["template_sync"]
    modules = set(marker["included_modules"])
    if "instruction-enforcement" not in modules:
        return
    # Temporary roots can use OS aliases such as macOS /var -> /private/var.
    # Normalize the caller's roots, while keeping candidate containment checks.
    staging_root = staging_root.resolve()
    target_root = target_root.resolve()
    destination = staging_root / PROFILE_PATH
    if not destination.is_file():
        raise TemplateSyncMaterializationError("Selected enforcement profile candidate is missing.")
    validate_selected_enforcement_inputs(staging_root, target_root, marker)
    validate_skipped_profile_applicability(target_root, marker)
    previous = target_root / PROFILE_PATH
    if "template-sync-support" in modules:
        if previous.is_file():
            local = load_existing_profile(target_root)
            if local.get("mode") == "standalone" and (
                local.get("exceptions") or local.get("source_decisions")
            ):
                raise TemplateSyncMaterializationError(
                    "Reintroducing sync requires explicitly translating local declarations into "
                    "marker decisions and clearing the reviewed standalone declarations first."
                )
        document: dict[str, Any] = {"version": 1, "mode": "marker", "context": "downstream"}
    else:
        schema = load_reviewed_schema("schemas/instruction-profile.schema.json")
        catalog_root, _ = selected_content_root(
            ".github/instruction-contracts.yml", staging_root, target_root, marker
        )
        catalog_path = core.support.resolve_repo_path(
            catalog_root, ".github/instruction-contracts.yml"
        )
        catalog = core.support.load_yaml_mapping(
            catalog_path,
            catalog_root,
            maximum_bytes=core.MAXIMUM_INPUT_BYTES,
        )
        core.support.validate_schema(
            catalog,
            load_reviewed_schema("schemas/instruction-contracts.schema.json"),
            catalog_path,
            catalog_root,
        )
        known_modules = set(schema["$defs"]["moduleName"]["enum"])
        contracts = core.parse_contracts(catalog, known_modules)
        section_obligations = core.parse_protected_guide_section_obligations(catalog, known_modules)
        reference_obligations = core.parse_protected_guide_reference_obligations(
            catalog, known_modules
        )
        guide_waivers = tuple(
            core.support.ProtectedGuideContractWaiver(
                path=waiver["path"],
                contract_key=waiver["contract_key"],
                target_path=waiver.get("target_path"),
                target_module=waiver.get("target_module"),
                linked_local_override_path=waiver.get("linked_local_override_path"),
                reason=waiver["reason"],
                authorization_basis=waiver["authorization_basis"],
            )
            for waiver in marker.get("protected_guide_contract_waivers", [])
        )
        path_applicability = catalog_path_applicability(
            contracts, section_obligations, reference_obligations, modules
        )
        exceptions: list[dict[str, str]] = []
        source_decisions = {key: marker[key] for key in SOURCE_FIELDS if marker.get(key)}
        retired_exceptions: list[dict[str, str]] = []
        # Retain declarations from an existing local standalone profile. Protected
        # reconciliation refuses to replace locally changed profile content.
        if previous.is_file():
            local = load_existing_profile(target_root)
            if local.get("mode") == "standalone":
                source_decisions = {**local.get("source_decisions", {}), **source_decisions}
                retired_exceptions.extend(source_decisions.get("retired_exceptions", []))
                exceptions.extend(local.get("exceptions", []))
        reports = {}
        for content_root in (target_root, staging_root):
            reports[content_root] = core.validate_contracts(
                mode="migration",
                repo_root=content_root,
                contracts=contracts,
                included_modules=modules,
                protected_guide_section_obligations=section_obligations,
                protected_guide_reference_obligations=reference_obligations,
            )
        removed_paths = {
            item["path"]
            for item in marker.get("protected_file_decisions", [])
            if item["decision"] == "REMOVE-LOCAL"
        }
        for path in sorted(removed_paths):
            if not path_applicability.get(path, False):
                continue
            removal_target = resolve_safe_repository_target_path(
                target_root, path, field_name="selected protected removal"
            )
            if removal_target.exists() or removal_target.is_symlink():
                raise TemplateSyncMaterializationError(
                    f"Selected protected removal is not complete: {path}. "
                    "Complete the reviewed local removal before standalone migration."
                )
        failure_keys = {
            root: {
                (path, anchor)
                for path, anchor in report_failure_keys(report)
                if path not in removed_paths
            }
            | {
                (path, "file:absent")
                for path in removed_paths
                if path_applicability.get(path, False)
            }
            for root, report in reports.items()
        }
        content_digests: dict[tuple[Path, str], str] = {}
        for declaration in exceptions:
            if not declaration_applies(
                declaration, contracts, section_obligations, modules, reference_obligations
            ):
                continue
            path, anchor = declaration["path"], declaration["anchor"]
            content_root, _ = selected_content_root(path, staging_root, target_root, marker)
            content_key = (content_root, path)
            if content_key not in content_digests:
                content_digests[content_key] = (
                    "absent" if path in removed_paths else core.file_digest(content_root, path)
                )
            digest = content_digests[content_key]
            original_digest = declaration["content_sha256"]
            if (path, anchor) not in failure_keys[content_root] or original_digest != digest:
                raise TemplateSyncMaterializationError(
                    f"Existing standalone exception conflicts with selected content: {path}: {anchor}. "
                    "Review or remove the declaration; its content hash is not renewed automatically."
                )
        for waiver in marker.get("instruction_contract_waivers", []):
            content_root, taken = selected_content_root(
                waiver["path"], staging_root, target_root, marker
            )
            if (
                declaration_applies(
                    waiver, contracts, section_obligations, modules, reference_obligations
                )
                and (waiver["path"], waiver["anchor"]) not in failure_keys[content_root]
            ):
                raise TemplateSyncMaterializationError(
                    f"Instruction waiver conflicts with selected {'TAKE' if taken else 'preserved'} content: "
                    f"{waiver['path']}: {waiver['anchor']}. Review or remove the waiver."
                )
            exceptions.append(
                {
                    **waiver,
                    "content_sha256": (
                        "absent"
                        if waiver["path"] in removed_paths
                        else core.file_digest(content_root, waiver["path"])
                    ),
                }
            )
        for decision in marker.get("protected_file_decisions", []):
            if (
                decision["decision"] == "REMOVE-LOCAL"
                and path_applicability.get(decision["path"], False)
                and not any(
                    item["path"] == decision["path"] and item["anchor"] == "file:absent"
                    for item in exceptions
                )
            ):
                exceptions.append(
                    {
                        "path": decision["path"],
                        "anchor": "file:absent",
                        "content_sha256": "absent",
                        "reason": decision.get("reason") or "Reviewed marker removal declaration",
                        "authorization_basis": decision["authorization_basis"],
                    }
                )
        for content_root, report in reports.items():
            for stale in report.stale_protected_guide_sections:
                if (
                    stale.path in removed_paths
                    or content_root
                    != selected_content_root(stale.path, staging_root, target_root, marker)[0]
                ):
                    continue
                for waiver in marker.get("protected_guide_contract_waivers", []):
                    if waiver["path"] != stale.path or waiver["contract_key"] != stale.contract_key:
                        continue
                    if waiver.get("target_module") not in stale.target_modules:
                        continue
                    exceptions.append(
                        {
                            "path": stale.path,
                            "anchor": f"stale:{stale.contract_key}:{stale.anchor_type}:{stale.anchor}",
                            "content_sha256": core.file_digest(content_root, stale.path),
                            "reason": waiver["reason"],
                            "authorization_basis": waiver["authorization_basis"],
                        }
                    )
            for reference in report.stale_protected_guide_references:
                if (
                    reference.path in removed_paths
                    or content_root
                    != selected_content_root(reference.path, staging_root, target_root, marker)[0]
                ):
                    continue
                waiver = core.find_protected_guide_waiver(
                    guide_waivers,
                    path=reference.path,
                    contract_key=reference.contract_key,
                    target_modules=reference.target_modules,
                    target_path=reference.target_path,
                )
                if waiver is not None:
                    exceptions.append(
                        {
                            "path": reference.path,
                            "anchor": f"reference:{reference.contract_key}:{reference.reference_kind}:{reference.target}",
                            "content_sha256": core.file_digest(content_root, reference.path),
                            "reason": waiver.reason,
                            "authorization_basis": waiver.authorization_basis,
                        }
                    )
        unique: dict[tuple[str, str], dict[str, str]] = {}
        for exception in exceptions:
            if not declaration_applies(
                exception, contracts, section_obligations, modules, reference_obligations
            ):
                if exception not in retired_exceptions:
                    retired_exceptions.append(exception)
                continue
            key = (exception["path"], exception["anchor"])
            if key in unique and unique[key] != exception:
                raise TemplateSyncMaterializationError(
                    f"Conflicting migrated declarations require protected profile review: {key}"
                )
            unique[key] = exception
        document = {
            "version": 1,
            "mode": "standalone",
            "modules": sorted(modules),
            "exceptions": list(unique.values()),
            "source_decisions": source_decisions,
        }
        if retired_exceptions:
            source_decisions["retired_exceptions"] = retired_exceptions
        validate_selected_claude_state(staging_root, target_root, marker, reports)
        if catalog_root == target_root or any(
            item["path"] == PROFILE_PATH and item["decision"] == "SKIP"
            for item in marker.get("protected_file_decisions", [])
        ):
            validate_retained_catalog_selection(
                staging_root, target_root, marker, document, reports, path_applicability
            )
    core.support.validate_schema(
        document,
        load_reviewed_schema("schemas/instruction-profile.schema.json"),
        destination,
        staging_root,
    )
    rendered = yaml.safe_dump(document, sort_keys=False)
    if len(rendered.encode("utf-8")) > core.MAXIMUM_INPUT_BYTES:
        raise TemplateSyncMaterializationError(
            f"Rendered {PROFILE_PATH} exceeds the {core.MAXIMUM_INPUT_BYTES}-byte input limit."
        )
    destination.write_text(rendered, encoding="utf-8", newline="\n")
