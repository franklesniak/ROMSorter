"""Run explicit marker or standalone static instruction enforcement.

Standalone declarations report local exceptions; they do not establish that
their author obtained repository-owner authorization.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from dataclasses import replace
from pathlib import Path
from typing import Any

import instruction_contract_core as core
import instruction_contract_support as support

PROFILE_PATH = ".github/instruction-profile.yml"
PROFILE_SCHEMA = "schemas/instruction-profile.schema.json"
CATALOG_PATH = ".github/instruction-contracts.yml"
CATALOG_SCHEMA = "schemas/instruction-contracts.schema.json"


file_digest = core.file_digest


def validate_selection(modules: set[str]) -> None:
    """Reject enforcement without instructions or any selected execution route."""
    if "instruction-enforcement" not in modules or "agent-instructions" not in modules:
        raise core.InstructionContractValidationError(
            "Standalone enforcement requires agent-instructions and instruction-enforcement."
        )
    if not modules.intersection({"baseline", "github-actions", "azure-pipelines"}):
        raise core.InstructionContractValidationError(
            "Select baseline or a host CI route, or choose policy-only instructions."
        )
    if "template-sync-support" in modules:
        raise core.InstructionContractValidationError(
            "Conflicting modes: standalone modules cannot retain template-sync-support."
        )


def validate_standalone(root: Path, profile: dict[str, Any]) -> core.InstructionContractReport:
    """Apply exact path, failure anchor and content-bound local declarations."""
    modules = set(profile["modules"])
    validate_selection(modules)
    schema = support.load_json_mapping(
        support.resolve_repo_path(root, PROFILE_SCHEMA),
        root,
        maximum_bytes=core.MAXIMUM_INPUT_BYTES,
    )
    known_modules = set(schema["$defs"]["moduleName"]["enum"])
    catalog = core.load_schema_validated_yaml(
        support.resolve_repo_path(root, CATALOG_PATH),
        support.resolve_repo_path(root, CATALOG_SCHEMA),
        root,
    )
    contracts = core.parse_contracts(catalog, known_modules)
    report = core.validate_contracts(
        mode="standalone",
        repo_root=root,
        contracts=contracts,
        included_modules=modules,
        protected_guide_section_obligations=core.parse_protected_guide_section_obligations(
            catalog, known_modules
        ),
        protected_guide_reference_obligations=core.parse_protected_guide_reference_obligations(
            catalog, known_modules
        ),
    )
    failures = {(item.path, item.anchor) for item in report.missing_anchors}
    failures.update((item.path, "file:absent") for item in report.missing_files)
    failures.update(
        (item.path, f"stale:{item.contract_key}:{item.anchor_type}:{item.anchor}")
        for item in report.stale_protected_guide_sections
    )
    applied: set[tuple[str, str]] = set()
    failures.update(
        (item.path, f"reference:{item.contract_key}:{item.reference_kind}:{item.target}")
        for item in report.stale_protected_guide_references
    )
    for declaration in profile["exceptions"]:
        path, directory = support.normalize_repository_path(declaration["path"], "exception.path")
        if directory:
            raise core.InstructionContractValidationError("Exception path must identify one file.")
        key = (path, declaration["anchor"])
        if key in applied:
            raise core.InstructionContractValidationError(f"Duplicate exception: {key}")
        if key not in failures or declaration["content_sha256"] != file_digest(root, path):
            raise core.InstructionContractValidationError(
                f"Exception does not match a current failure and exact content: {key}"
            )
        applied.add(key)
        print(f"Applied local exception: {path}: {declaration['anchor']}")
        print(f"  reason: {declaration['reason']}")
        print(f"  declaration: {declaration['authorization_basis']}")
    if applied:
        print(
            "Local declarations are auditable data, not independent proof of owner authorization."
        )
    return replace(
        report,
        missing_anchors=tuple(
            item for item in report.missing_anchors if (item.path, item.anchor) not in applied
        ),
        missing_files=tuple(
            item for item in report.missing_files if (item.path, "file:absent") not in applied
        ),
        stale_protected_guide_sections=tuple(
            item
            for item in report.stale_protected_guide_sections
            if (item.path, f"stale:{item.contract_key}:{item.anchor_type}:{item.anchor}")
            not in applied
        ),
        stale_protected_guide_references=tuple(
            item
            for item in report.stale_protected_guide_references
            if (item.path, f"reference:{item.contract_key}:{item.reference_kind}:{item.target}")
            not in applied
        ),
    )


def main(argv: list[str] | None = None) -> int:
    """Validate explicit applicability and preserve native validator failures."""
    try:
        parser = argparse.ArgumentParser(description=__doc__)
        parser.add_argument("--repo-root", default=str(Path(__file__).resolve().parents[2]))
        args = parser.parse_args(argv)
        root = Path(args.repo_root).resolve()
        profile = core.load_schema_validated_yaml(
            support.resolve_repo_path(root, PROFILE_PATH),
            support.resolve_repo_path(root, PROFILE_SCHEMA),
            root,
        )
        if profile["mode"] == "marker":
            # Only the explicitly selected adapter uses sync-owned inputs.
            adapter = support.resolve_repo_path(
                root, ".template-sync/scripts/validate_instruction_contracts.py"
            )
            mode = profile["context"]
            command = [sys.executable, str(adapter), "--repo-root", str(root), "--mode", mode]
            if mode == "downstream":
                command.append("--require-marker")
            return subprocess.run(command, check=False).returncode
        # Presence is a conflict check only: standalone never reads, imports or
        # requires any sync-owned file, and works when the directory is absent.
        if (root / ".template-sync" / "marker.yml").exists():
            raise core.InstructionContractValidationError(
                "Conflicting modes: remove the reviewed marker during standalone migration."
            )
        report = validate_standalone(root, profile)
        core.print_report(report)
        return int(report.has_failures)
    except (
        core.InstructionContractValidationError,
        support.TemplateSyncMaterializationError,
    ) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1
    except OSError as error:
        print(f"ERROR: {support.os_error_summary(error)}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
