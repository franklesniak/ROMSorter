"""Materialize a selected template module set into a downstream work tree."""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import tempfile
from collections.abc import Callable, Collection, Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any, cast

SCRIPT_DIR = Path(__file__).resolve().parent
TRUSTED_TOOL_ROOT = SCRIPT_DIR.parent.parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from instruction_profile_migration import render_instruction_profile  # noqa: E402
from template_sync_materialization_helpers import (  # noqa: E402
    DEFAULT_MANIFEST_PATH,
    DEFAULT_MANIFEST_SCHEMA_PATH,
    DEFAULT_MARKER_PATH,
    DEFAULT_MARKER_SCHEMA_PATH,
    REMOVAL_DECISION,
    DeferredProtectedCandidate,
    InlineBlockError,
    LocalOverride,
    ManifestMapping,
    MarkerDecisionData,
    ProtectedFileDecision,
    TemplateSyncMaterializationError,
    classify_repository_file,
    ensure_regular_repository_file_target,
    format_marker_yaml,
    inline_block_families_to_prune,
    is_protected_instruction_path,
    is_string_list,
    iter_safe_repository_files,
    load_json_mapping,
    load_yaml_mapping,
    os_error_summary,
    parse_json_mapping,
    parse_manifest_mappings,
    parse_marker_decision_data,
    parse_yaml_mapping,
    remove_inline_blocks_for_modules,
    resolve_safe_repository_target_path,
    selected_relation_for_path,
    validate_inline_block_markers,
    validate_marker_yaml_text,
    validate_module_compatibility,
    validate_protected_file_decisions,
    validate_schema,
    write_repository_file_bytes,
    write_repository_relative_file_bytes,
)

EXIT_SUCCESS = 0
EXIT_RUNTIME_FAILURE = 1
EXIT_DECISIONS_REQUIRED = 2
EXIT_CLEANUP_FAILURE = 3
PLACEHOLDER_HELPER_PATH = ".github/scripts/replace-template-placeholders.py"
ADOPTION_MODES = ("minimal-preservation", "tailored")
TEMPLATE_TAKE_DECISION = "TAKE"
TEMPLATE_SKIP_DECISION = "SKIP"
FULL_LOWERCASE_SHA_RE = re.compile(r"^[0-9a-f]{40}$")
GIT_VERSION_RE = re.compile(r"^git version (?P<major>[0-9]+)\.(?P<minor>[0-9]+)")
SOURCE_GIT_ENV_OVERLAY = {
    "GIT_NO_LAZY_FETCH": "1",
    "GIT_OPTIONAL_LOCKS": "0",
}
PLACEHOLDER_DESTS = (
    "host_provider",
    "repository",
    "github_host",
    "azure_devops_organization",
    "azure_devops_organization_url",
    "azure_devops_project",
    "azure_devops_project_url",
    "azure_devops_repository",
    "azure_devops_repository_url",
    "azure_devops_clone_url",
    "azure_devops_default_branch",
    "azure_boards_policy",
    "azure_repos_pr_template_policy",
    "azure_branch_policy_reviewer_guidance",
    "azure_security_intake_policy",
    "azure_security_product_enablement",
    "azure_dependency_update_policy",
    "codeowners_owner",
    "conduct_contact",
    "conduct_contact_sentence",
    "security_contact",
    "security_contact_section",
    "security_reporting_mode",
    "issue_label_policy",
    "issue_labels",
    "discussions_policy",
    "collaboration_policy_follow_up_status",
    "vscode_title",
    "package_name",
    "package_description",
    "package_author",
    "package_version",
    "package_keywords",
)
HOST_PROVIDERS = (
    "github",
    "github-enterprise-server",
    "azure-devops-services",
    "dual",
)
SECURITY_REPORTING_MODES = ("github-private-only", "contact-only", "both")
ISSUE_LABEL_POLICIES = ("existing", "create-manual-follow-up", "omit", "custom")
DISCUSSIONS_POLICIES = (
    "enabled",
    "disabled",
    "deferred-planned-render",
    "deferred-not-rendered",
)
AZURE_DEVOPS_BOARDS_POLICIES = ("work-items", "disabled", "manual-follow-up")
AZURE_DEVOPS_PR_TEMPLATE_POLICIES = ("materialize", "disabled", "manual-follow-up")
AZURE_DEVOPS_BRANCH_POLICY_POLICIES = ("required-reviewers", "manual-follow-up", "none")
AZURE_DEVOPS_SECURITY_INTAKE_POLICIES = ("security-contact", "external-process", "manual-follow-up")
AZURE_DEVOPS_SECURITY_PRODUCT_POLICIES = (
    "none",
    "github-advanced-security",
    "github-secret-protection",
    "github-code-security",
    "github-secret-protection-and-code-security",
)
AZURE_DEVOPS_DEPENDENCY_UPDATE_POLICIES = ("none", "renovate", "manual-follow-up")
AZURE_DEVOPS_PLACEHOLDER_FIELDS = frozenset(
    {
        "host_provider",
        "azure_devops_organization",
        "azure_devops_organization_url",
        "azure_devops_project",
        "azure_devops_project_url",
        "azure_devops_repository",
        "azure_devops_repository_url",
        "azure_devops_clone_url",
        "azure_devops_default_branch",
        "azure_boards_policy",
        "azure_repos_pr_template_policy",
        "azure_branch_policy_reviewer_guidance",
        "azure_security_intake_policy",
        "azure_security_product_enablement",
        "azure_dependency_update_policy",
    }
)
ARGS_FILE_FORMATS = ("json", "yaml")
ARGS_FILE_EXTENSION_FORMATS = {
    ".json": "json",
    ".yaml": "yaml",
    ".yml": "yaml",
}
ARGS_FILE_FIELDS = frozenset(
    {
        "template_root",
        "template_ref",
        "template_revision",
        "target_root",
        "template_repo",
        "template_temp_root",
        "decisions_file",
        "included_modules",
        "included_modules_csv",
        "source_repo",
        "last_reviewed_template_commit",
        "stamp_resolved_source_as_reviewed",
        "default_adoption_mode",
        "host_provider",
        "repository",
        "github_host",
        "azure_devops_organization",
        "azure_devops_organization_url",
        "azure_devops_project",
        "azure_devops_project_url",
        "azure_devops_repository",
        "azure_devops_repository_url",
        "azure_devops_clone_url",
        "azure_devops_default_branch",
        "azure_boards_policy",
        "azure_repos_pr_template_policy",
        "azure_branch_policy_reviewer_guidance",
        "azure_security_intake_policy",
        "azure_security_product_enablement",
        "azure_dependency_update_policy",
        "codeowners_owner",
        "conduct_contact",
        "conduct_contact_sentence",
        "security_contact",
        "security_contact_section",
        "security_reporting_mode",
        "issue_label_policy",
        "issue_labels",
        "discussions_policy",
        "collaboration_policy_follow_up_status",
        "vscode_title",
        "package_name",
        "package_description",
        "package_author",
        "package_version",
        "package_keywords",
        "preserve_existing_license",
        "license_source_path",
        "allow_conflicts",
    }
)
STRING_ARGS_FILE_FIELDS = frozenset(
    {
        "template_root",
        "template_ref",
        "template_revision",
        "target_root",
        "template_repo",
        "template_temp_root",
        "decisions_file",
        "included_modules_csv",
        "source_repo",
        "last_reviewed_template_commit",
        "default_adoption_mode",
        "host_provider",
        "repository",
        "github_host",
        "azure_devops_organization",
        "azure_devops_organization_url",
        "azure_devops_project",
        "azure_devops_project_url",
        "azure_devops_repository",
        "azure_devops_repository_url",
        "azure_devops_clone_url",
        "azure_devops_default_branch",
        "azure_boards_policy",
        "azure_repos_pr_template_policy",
        "azure_branch_policy_reviewer_guidance",
        "azure_security_intake_policy",
        "azure_security_product_enablement",
        "azure_dependency_update_policy",
        "codeowners_owner",
        "conduct_contact",
        "conduct_contact_sentence",
        "security_contact",
        "security_contact_section",
        "security_reporting_mode",
        "issue_label_policy",
        "discussions_policy",
        "collaboration_policy_follow_up_status",
        "vscode_title",
        "package_name",
        "package_description",
        "package_author",
        "package_version",
        "license_source_path",
    }
)
LIST_STRING_ARGS_FILE_FIELDS = frozenset({"included_modules", "issue_labels", "package_keywords"})
BOOLEAN_ARGS_FILE_FIELDS = frozenset(
    {
        "stamp_resolved_source_as_reviewed",
        "preserve_existing_license",
        "allow_conflicts",
    }
)
CLI_FLAGS: dict[str, tuple[str, ...]] = {
    "template_root": ("--template-root",),
    "template_ref": ("--template-ref",),
    "template_revision": ("--template-revision",),
    "target_root": ("--target-root",),
    "template_repo": ("--template-repo",),
    "template_temp_root": ("--template-temp-root",),
    "decisions_file": ("--decisions-file",),
    "included_modules": ("--included-module", "--module"),
    "included_modules_csv": ("--included-modules",),
    "source_repo": ("--source-repo",),
    "last_reviewed_template_commit": ("--last-reviewed-template-commit",),
    "stamp_resolved_source_as_reviewed": (
        "--stamp-resolved-source-as-reviewed",
        "--no-stamp-resolved-source-as-reviewed",
    ),
    "default_adoption_mode": ("--default-adoption-mode",),
    "host_provider": ("--host-provider",),
    "repository": ("--repository",),
    "github_host": ("--github-host",),
    "azure_devops_organization": ("--azure-devops-organization",),
    "azure_devops_organization_url": ("--azure-devops-organization-url",),
    "azure_devops_project": ("--azure-devops-project",),
    "azure_devops_project_url": ("--azure-devops-project-url",),
    "azure_devops_repository": ("--azure-devops-repository",),
    "azure_devops_repository_url": ("--azure-devops-repository-url",),
    "azure_devops_clone_url": ("--azure-devops-clone-url",),
    "azure_devops_default_branch": ("--azure-devops-default-branch",),
    "azure_boards_policy": ("--azure-boards-policy",),
    "azure_repos_pr_template_policy": ("--azure-repos-pr-template-policy",),
    "azure_branch_policy_reviewer_guidance": ("--azure-branch-policy-reviewer-guidance",),
    "azure_security_intake_policy": ("--azure-security-intake-policy",),
    "azure_security_product_enablement": ("--azure-security-product-enablement",),
    "azure_dependency_update_policy": ("--azure-dependency-update-policy",),
    "codeowners_owner": ("--codeowners-owner",),
    "conduct_contact": ("--conduct-contact",),
    "conduct_contact_sentence": ("--conduct-contact-sentence",),
    "security_contact": ("--security-contact",),
    "security_contact_section": ("--security-contact-section",),
    "security_reporting_mode": ("--security-reporting-mode",),
    "issue_label_policy": ("--issue-label-policy",),
    "issue_labels": ("--issue-label",),
    "discussions_policy": ("--discussions-policy",),
    "collaboration_policy_follow_up_status": ("--collaboration-policy-follow-up-status",),
    "vscode_title": ("--vscode-title",),
    "package_name": ("--package-name",),
    "package_description": ("--package-description",),
    "package_author": ("--package-author",),
    "package_version": ("--package-version",),
    "package_keywords": ("--package-keyword",),
    "preserve_existing_license": ("--preserve-existing-license",),
    "license_source_path": ("--license-source-path",),
    "allow_conflicts": ("--allow-conflicts",),
}
MARKER_COPY_FIELDS = (
    "local_overrides",
    "local_path_ownership",
    "placeholder_waivers",
    "protected_file_decisions",
    "deferred_protected_candidates",
    "instruction_contract_waivers",
    "protected_guide_contract_waivers",
    "issue_label_policy",
    "issue_labels",
    "discussions_policy",
    "collaboration_policy_follow_up_status",
    "host_provider",
    "azure_devops_organization",
    "azure_devops_organization_url",
    "azure_devops_project",
    "azure_devops_project_url",
    "azure_devops_repository",
    "azure_devops_repository_url",
    "azure_devops_clone_url",
    "azure_devops_default_branch",
    "azure_boards_policy",
    "azure_repos_pr_template_policy",
    "azure_branch_policy_reviewer_guidance",
    "azure_security_intake_policy",
    "azure_security_product_enablement",
    "azure_dependency_update_policy",
)
LICENSE_TARGET_PATH = "LICENSE"
LICENSE_SOURCE_CANDIDATE_NAMES = frozenset(
    name.casefold()
    for name in (
        "LICENSE.txt",
        "LICENSE.md",
        "LICENSE.rst",
        "LICENCE",
        "LICENCE.txt",
        "LICENCE.md",
        "COPYING",
        "COPYING.txt",
        "COPYING.md",
    )
)


class MaterializationError(RuntimeError):
    """Raised when first-adoption materialization cannot proceed safely."""


@dataclass(frozen=True)
class GitVersion:
    """Parsed Git major/minor version used for capability gates."""

    major: int
    minor: int


@dataclass(frozen=True)
class LocalSourceDetection:
    """Best-effort stampability result for a local template source worktree."""

    observed_source_sha: str | None = None
    source_worktree_root: str | None = None
    stampable_sha: str | None = None
    not_stampable_reason: str | None = None


@dataclass
class SourceSummary:
    """Template source identity and cleanup information for one run."""

    target_root: str
    template_root: str
    source_mode: str
    source_value: str | None = None
    resolved_source_sha: str | None = None
    observed_source_sha: str | None = None
    source_worktree_root: str | None = None
    source_not_stampable_reason: str | None = None
    accepted_reviewed_commit: str | None = None
    reviewed_commit_source: str | None = None
    source_cross_check_note: str | None = None
    source_repository: str | None = None
    temporary_checkout_path: str | None = None
    cleanup_status: str = "not required"
    cleanup_failure: str | None = None
    manual_cleanup_command: str | None = None


@dataclass
class SourceCheckout:
    """A template source checkout owned by the materializer."""

    template_root: Path
    summary: SourceSummary
    template_repo: Path | None = None
    temporary_parent: Path | None = None
    temporary_checkout_path: Path | None = None


@dataclass(frozen=True)
class CleanupFailure:
    """A failed attempt to remove a tool-created temporary worktree."""

    detail: str
    residual_worktree_path: Path
    source_repository: Path
    manual_cleanup_command: str


@dataclass(frozen=True)
class Decisions:
    """Resolved scalar and marker-shaped adoption decisions."""

    source_repo: str
    last_reviewed_template_commit: str | None
    included_modules: frozenset[str]
    marker_data: MarkerDecisionData
    raw_marker_fields: dict[str, Any]


@dataclass(frozen=True)
class LicensePreservation:
    """A first-adoption license preservation operation."""

    source_path: str
    bytes_content: bytes
    local_override: LocalOverride


@dataclass(frozen=True)
class Conflict:
    """One target path that was not overwritten by materialization."""

    path: str
    classification: str
    resolution: str
    recorded: bool


def empty_string_set() -> set[str]:
    """Return a typed empty string set for dataclass defaults."""
    return set()


def empty_string_list() -> list[str]:
    """Return a typed empty string list for dataclass defaults."""
    return []


def empty_conflict_list() -> list[Conflict]:
    """Return a typed empty conflict list for dataclass defaults."""
    return []


@dataclass
class Summary:
    """Deterministic operation summary emitted after reconciliation."""

    retained_modules: list[str]
    excluded_modules: list[str]
    default_adoption_mode: str
    created: set[str] = field(default_factory=empty_string_set)
    updated: set[str] = field(default_factory=empty_string_set)
    unchanged: set[str] = field(default_factory=empty_string_set)
    skipped: set[str] = field(default_factory=empty_string_set)
    protected: set[str] = field(default_factory=empty_string_set)
    locally_overridden: set[str] = field(default_factory=empty_string_set)
    deferred: set[str] = field(default_factory=empty_string_set)
    recorded_removals: set[str] = field(default_factory=empty_string_set)
    unmapped: set[str] = field(default_factory=empty_string_set)
    excluded_paths: set[str] = field(default_factory=empty_string_set)
    byte_only: set[str] = field(default_factory=empty_string_set)
    placeholder_related: set[str] = field(default_factory=empty_string_set)
    placeholder_notes: list[str] = field(default_factory=empty_string_list)
    license_notes: list[str] = field(default_factory=empty_string_list)
    residual_cleanup_paths: set[str] = field(default_factory=empty_string_set)
    pruning_preview: set[str] = field(default_factory=empty_string_set)
    conflicts: list[Conflict] = field(default_factory=empty_conflict_list)
    marker_status: str = "preview-only"
    marker_reason: str = ""
    computed_marker_preview: str | None = None
    source: SourceSummary | None = None

    @property
    def unrecorded_conflicts(self) -> tuple[Conflict, ...]:
        """Return conflicts that still require a concrete path-scoped decision."""
        return tuple(conflict for conflict in self.conflicts if not conflict.recorded)

    @property
    def recorded_conflicts(self) -> tuple[Conflict, ...]:
        """Return conflicts already recorded as unresolved decisions."""
        return tuple(conflict for conflict in self.conflicts if conflict.recorded)


def args_file_format_for_path(path: Path, args_format: str | None) -> str:
    """Return the parser format selected by override or recognized file extension."""
    if args_format is not None:
        return args_format
    inferred_format = ARGS_FILE_EXTENSION_FORMATS.get(path.suffix.lower())
    if inferred_format is None:
        raise MaterializationError(
            "Unable to determine --args-file format from extension; use "
            "--args-format json or --args-format yaml, or name the file with "
            "a .json, .yaml, or .yml extension."
        )
    return inferred_format


def read_args_file_text(path: Path) -> str:
    """Read an explicit args file path."""
    try:
        # Tolerate a leading UTF-8 BOM (e.g. PowerShell `Set-Content -Encoding UTF8`).
        return path.read_text(encoding="utf-8-sig")
    except OSError as error:
        raise MaterializationError(
            f"--args-file: unable to read file ({os_error_summary(error)})."
        ) from error


def load_json_args_file(path: Path) -> dict[str, Any]:
    """Load a JSON args file that must contain an object."""
    try:
        return parse_json_mapping(read_args_file_text(path), "--args-file")
    except TemplateSyncMaterializationError as error:
        raise MaterializationError(str(error)) from error


def load_yaml_args_file(path: Path) -> dict[str, Any]:
    """Load a YAML args file through the retained YAML parser path."""
    try:
        return parse_yaml_mapping(read_args_file_text(path), "--args-file")
    except TemplateSyncMaterializationError as error:
        raise MaterializationError(str(error)) from error


def load_args_file_mapping(raw_path: str, args_format: str | None) -> dict[str, Any]:
    """Load an explicit JSON or YAML argument file."""
    path = Path(raw_path).expanduser()
    selected_format = args_file_format_for_path(path, args_format)
    if selected_format == "json":
        return load_json_args_file(path)
    if selected_format == "yaml":
        return load_yaml_args_file(path)
    raise AssertionError(f"Unhandled args-file format: {selected_format}")


def cli_supplied_fields(argv: Sequence[str]) -> set[str]:
    """Return argument destinations supplied directly on the command line."""
    supplied: set[str] = set()
    flag_to_field: dict[str, str] = {
        flag: field_name for field_name, flags in CLI_FLAGS.items() for flag in flags
    }
    for token in argv:
        flag = token.split("=", 1)[0]
        field_name = flag_to_field.get(flag)
        if field_name is not None:
            supplied.add(field_name)
    return supplied


# Conceptual inputs selectable through more than one destination. CLI precedence is
# per-family: when any member is supplied on the command line, every member is
# skipped from the args file, so a CLI selector fully overrides the args-file
# selector for that concept instead of conflicting with it (source selection) or
# merging with it (module selection).
CLI_FIELD_FAMILIES: tuple[frozenset[str], ...] = (
    frozenset({"template_root", "template_ref", "template_revision"}),
    frozenset({"included_modules", "included_modules_csv"}),
)


def cli_overridden_args_file_fields(direct_cli_fields: set[str]) -> set[str]:
    """Expand directly-supplied CLI fields to the full precedence family of each."""
    overridden = set(direct_cli_fields)
    for family in CLI_FIELD_FAMILIES:
        if direct_cli_fields & family:
            overridden |= family
    return overridden


def validate_args_file_value(field_name: str, value: Any) -> Any:
    """Validate one args-file value and return its normalized representation."""
    if value is None:
        return None
    if field_name in STRING_ARGS_FILE_FIELDS:
        if not isinstance(value, str):
            raise MaterializationError(f"--args-file field {field_name!r} must be a string.")
        return value
    if field_name in LIST_STRING_ARGS_FILE_FIELDS:
        if not is_string_list(value):
            raise MaterializationError(
                f"--args-file field {field_name!r} must be a list of strings."
            )
        return list(value)
    if field_name in BOOLEAN_ARGS_FILE_FIELDS:
        if not isinstance(value, bool):
            raise MaterializationError(f"--args-file field {field_name!r} must be a boolean.")
        return value
    raise AssertionError(f"Unhandled args-file field: {field_name}")


def validate_merged_arg_choices(args: argparse.Namespace) -> None:
    """Validate choices and mutual exclusions after args-file merging."""
    selected_sources = [
        name
        for name in ("template_root", "template_ref", "template_revision")
        if getattr(args, name) is not None
    ]
    if len(selected_sources) > 1:
        raise MaterializationError(
            "--template-root, --template-ref, and --template-revision are mutually exclusive."
        )
    if args.default_adoption_mode not in ADOPTION_MODES:
        quoted_modes = ", ".join(ADOPTION_MODES)
        raise MaterializationError(f"--default-adoption-mode must be one of: {quoted_modes}.")
    if (
        args.security_reporting_mode is not None
        and args.security_reporting_mode not in SECURITY_REPORTING_MODES
    ):
        quoted_modes = ", ".join(SECURITY_REPORTING_MODES)
        raise MaterializationError(f"--security-reporting-mode must be one of: {quoted_modes}.")
    if args.issue_label_policy is not None and args.issue_label_policy not in ISSUE_LABEL_POLICIES:
        quoted_policies = ", ".join(ISSUE_LABEL_POLICIES)
        raise MaterializationError(f"--issue-label-policy must be one of: {quoted_policies}.")
    if args.discussions_policy is not None and args.discussions_policy not in DISCUSSIONS_POLICIES:
        quoted_policies = ", ".join(DISCUSSIONS_POLICIES)
        raise MaterializationError(f"--discussions-policy must be one of: {quoted_policies}.")
    if args.host_provider is not None and args.host_provider not in HOST_PROVIDERS:
        quoted_providers = ", ".join(HOST_PROVIDERS)
        raise MaterializationError(f"--host-provider must be one of: {quoted_providers}.")
    azure_choice_fields = (
        ("azure_boards_policy", "--azure-boards-policy", AZURE_DEVOPS_BOARDS_POLICIES),
        (
            "azure_repos_pr_template_policy",
            "--azure-repos-pr-template-policy",
            AZURE_DEVOPS_PR_TEMPLATE_POLICIES,
        ),
        (
            "azure_branch_policy_reviewer_guidance",
            "--azure-branch-policy-reviewer-guidance",
            AZURE_DEVOPS_BRANCH_POLICY_POLICIES,
        ),
        (
            "azure_security_intake_policy",
            "--azure-security-intake-policy",
            AZURE_DEVOPS_SECURITY_INTAKE_POLICIES,
        ),
        (
            "azure_security_product_enablement",
            "--azure-security-product-enablement",
            AZURE_DEVOPS_SECURITY_PRODUCT_POLICIES,
        ),
        (
            "azure_dependency_update_policy",
            "--azure-dependency-update-policy",
            AZURE_DEVOPS_DEPENDENCY_UPDATE_POLICIES,
        ),
    )
    for attribute_name, flag_name, choices in azure_choice_fields:
        value = getattr(args, attribute_name)
        if value is None or value in choices:
            continue
        quoted_choices = ", ".join(choices)
        raise MaterializationError(f"{flag_name} must be one of: {quoted_choices}.")


def apply_args_file_values(
    args: argparse.Namespace,
    *,
    argv: Sequence[str],
) -> argparse.Namespace:
    """Merge args-file values into parsed args, with CLI flags taking precedence."""
    if args.args_file is not None:
        args_file_values = load_args_file_mapping(args.args_file, args.args_format)
        unknown_fields = sorted(set(args_file_values) - ARGS_FILE_FIELDS)
        if unknown_fields:
            raise MaterializationError(
                "Unknown --args-file field(s): " + ", ".join(unknown_fields) + "."
            )
        direct_cli_fields = cli_supplied_fields(argv)
        overridden_fields = cli_overridden_args_file_fields(direct_cli_fields)
        for field_name, raw_value in args_file_values.items():
            if field_name in overridden_fields:
                continue
            if raw_value is None:
                continue
            setattr(args, field_name, validate_args_file_value(field_name, raw_value))
    validate_merged_arg_choices(args)
    return args


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """Parse command-line arguments."""
    argv = tuple(sys.argv[1:] if argv is None else argv)
    parser = argparse.ArgumentParser(
        description=(
            "Materialize selected copilot-repo-template modules into a downstream "
            "working tree without overwriting unrecorded conflicts."
        )
    )
    parser.add_argument(
        "--args-file",
        default=None,
        help="JSON or YAML file containing shell-safe materializer argument values.",
    )
    parser.add_argument(
        "--args-format",
        choices=ARGS_FILE_FORMATS,
        default=None,
        help="Explicit args-file format; overrides the file extension.",
    )
    source_group = parser.add_mutually_exclusive_group()
    source_group.add_argument(
        "--template-root",
        default=None,
        help="Template repository root. Defaults to the repository containing this script.",
    )
    source_group.add_argument(
        "--template-ref",
        default=None,
        help=(
            "Local upstream template ref to materialize from. The ref is resolved "
            "without fetching, using --template-repo as the Git object database."
        ),
    )
    source_group.add_argument(
        "--template-revision",
        default=None,
        metavar="FULL_SHA",
        help=(
            "Full lowercase upstream template commit SHA to materialize from. "
            "Resolved without fetching, using --template-repo as the Git object database."
        ),
    )
    parser.add_argument(
        "--target-root",
        default=None,
        help="Downstream repository root that receives the staged materialization.",
    )
    parser.add_argument(
        "--template-repo",
        default=None,
        help=(
            "Git repository whose local object database resolves --template-ref or "
            "--template-revision. Defaults to --target-root."
        ),
    )
    parser.add_argument(
        "--template-temp-root",
        default=None,
        help=(
            "Directory for private temporary source worktrees. Defaults outside "
            "--target-root under the system temporary directory when possible."
        ),
    )
    parser.add_argument(
        "--decisions-file",
        default=None,
        help=(
            "Repository-relative marker-shaped decisions file. When supplied, "
            "overlapping CLI source, commit, and module values must match it."
        ),
    )
    parser.add_argument(
        "--included-module",
        "--module",
        dest="included_modules",
        action="append",
        default=[],
        metavar="MODULE",
        help="Retained module name. May be repeated.",
    )
    parser.add_argument(
        "--included-modules",
        dest="included_modules_csv",
        default=None,
        metavar="MODULES",
        help="Comma-separated retained module names.",
    )
    parser.add_argument(
        "--source-repo",
        default=None,
        help="Canonical upstream template repository URL.",
    )
    parser.add_argument(
        "--last-reviewed-template-commit",
        default=None,
        help="Full lowercase reviewed upstream template commit SHA.",
    )
    parser.add_argument(
        "--stamp-resolved-source-as-reviewed",
        action=argparse.BooleanOptionalAction,
        default=False,
        help=(
            "Assert that review of the resolved template source commit is complete "
            "and that the source is the intended upstream lineage; this does not "
            "select the materialization source."
        ),
    )
    parser.add_argument(
        "--default-adoption-mode",
        choices=ADOPTION_MODES,
        default="minimal-preservation",
        help=(
            "Default behavior posture used for reporting. This is not written as "
            "a marker-level default field."
        ),
    )
    parser.add_argument(
        "--host-provider",
        choices=HOST_PROVIDERS,
        default=None,
        help=(
            "Host-provider mode for placeholder materialization: github, "
            "github-enterprise-server, azure-devops-services, or dual."
        ),
    )
    parser.add_argument(
        "--repository",
        default=None,
        help="Replacement OWNER/REPO value for the existing placeholder helper.",
    )
    parser.add_argument(
        "--github-host",
        default=None,
        help="GitHub or GHES host for approved placeholder URL contexts.",
    )
    parser.add_argument(
        "--azure-devops-organization",
        default=None,
        help="Azure DevOps Services organization name.",
    )
    parser.add_argument(
        "--azure-devops-organization-url",
        default=None,
        help="Azure DevOps organization URL override.",
    )
    parser.add_argument(
        "--azure-devops-project",
        default=None,
        help="Azure DevOps project name.",
    )
    parser.add_argument(
        "--azure-devops-project-url",
        default=None,
        help="Azure DevOps project URL override.",
    )
    parser.add_argument(
        "--azure-devops-repository",
        default=None,
        help="Azure Repos repository name.",
    )
    parser.add_argument(
        "--azure-devops-repository-url",
        default=None,
        help="Azure Repos repository web URL override.",
    )
    parser.add_argument(
        "--azure-devops-clone-url",
        default=None,
        help="Azure Repos HTTPS clone URL override without embedded credentials.",
    )
    parser.add_argument(
        "--azure-devops-default-branch",
        default=None,
        help="Azure Repos default branch name; defaults to main in the placeholder helper.",
    )
    parser.add_argument(
        "--azure-boards-policy",
        choices=AZURE_DEVOPS_BOARDS_POLICIES,
        default=None,
        help="Azure Boards intake policy for first-adoption reporting.",
    )
    parser.add_argument(
        "--azure-repos-pr-template-policy",
        choices=AZURE_DEVOPS_PR_TEMPLATE_POLICIES,
        default=None,
        help="Azure Repos pull request template policy.",
    )
    parser.add_argument(
        "--azure-branch-policy-reviewer-guidance",
        choices=AZURE_DEVOPS_BRANCH_POLICY_POLICIES,
        default=None,
        help="Azure Repos branch policy reviewer-guidance status.",
    )
    parser.add_argument(
        "--azure-security-intake-policy",
        choices=AZURE_DEVOPS_SECURITY_INTAKE_POLICIES,
        default=None,
        help="Azure security intake policy for SECURITY.md and first-adoption reporting.",
    )
    parser.add_argument(
        "--azure-security-product-enablement",
        choices=AZURE_DEVOPS_SECURITY_PRODUCT_POLICIES,
        default=None,
        help="Azure DevOps Services security product enablement status.",
    )
    parser.add_argument(
        "--azure-dependency-update-policy",
        choices=AZURE_DEVOPS_DEPENDENCY_UPDATE_POLICIES,
        default=None,
        help="Azure Repos dependency update policy status.",
    )
    parser.add_argument(
        "--codeowners-owner",
        default=None,
        help="Replacement CODEOWNERS owner, e.g. @octocat or @org/team.",
    )
    parser.add_argument(
        "--conduct-contact",
        default=None,
        help="Replacement Code of Conduct contact method.",
    )
    parser.add_argument(
        "--conduct-contact-sentence",
        default=None,
        help="Replacement full Code of Conduct reporting sentence.",
    )
    parser.add_argument(
        "--security-contact",
        default=None,
        help="Replacement security contact method or intake address.",
    )
    parser.add_argument(
        "--security-contact-section",
        default=None,
        help="Replacement whole SECURITY.md contact section for contact-based modes.",
    )
    parser.add_argument(
        "--security-reporting-mode",
        choices=SECURITY_REPORTING_MODES,
        default=None,
        help=(
            "Security reporting mode: github-private-only, contact-only, or both. "
            "Omitting this while supplying --security-contact preserves the "
            "backward-compatible both mode."
        ),
    )
    parser.add_argument(
        "--issue-label-policy",
        choices=ISSUE_LABEL_POLICIES,
        default=None,
        help="Issue-template label policy: existing, create-manual-follow-up, omit, or custom.",
    )
    parser.add_argument(
        "--issue-label",
        dest="issue_labels",
        action="append",
        default=None,
        help="Custom issue label for --issue-label-policy custom. May be repeated.",
    )
    parser.add_argument(
        "--discussions-policy",
        choices=DISCUSSIONS_POLICIES,
        default=None,
        help=(
            "Issue-template Discussions contact-link policy: enabled, disabled, "
            "deferred-planned-render, or deferred-not-rendered."
        ),
    )
    parser.add_argument(
        "--collaboration-policy-follow-up-status",
        default=None,
        help=(
            "Single source status text from _TODO-repo-init.md for label or "
            "Discussions policies that leave manual setup open."
        ),
    )
    parser.add_argument(
        "--vscode-title",
        default=None,
        help="Replacement VS Code window title.",
    )
    parser.add_argument("--package-name", default=None, help="Replacement package name.")
    parser.add_argument(
        "--package-description",
        default=None,
        help="Replacement package description.",
    )
    parser.add_argument("--package-author", default=None, help="Replacement package author.")
    parser.add_argument(
        "--package-version",
        default=None,
        help="Replacement package version; updates package-lock root version fields.",
    )
    parser.add_argument(
        "--package-keyword",
        dest="package_keywords",
        action="append",
        default=None,
        help="Replacement package keyword. May be repeated.",
    )
    parser.add_argument(
        "--preserve-existing-license",
        action="store_true",
        help=(
            "Copy existing downstream license text to root LICENSE, suppress the "
            "template LICENSE, and leave the source license file for manual cleanup."
        ),
    )
    parser.add_argument(
        "--license-source-path",
        default=None,
        help=(
            "Repository-relative downstream license file to preserve. When omitted "
            "with --preserve-existing-license, the helper only auto-selects a single "
            "unambiguous root alternate such as LICENSE.txt or LICENSE.md."
        ),
    )
    parser.add_argument(
        "--allow-conflicts",
        action="store_true",
        help=(
            "Exit zero while still skipping conflicted paths and leaving marker "
            "advancement governed by unrecorded conflicts."
        ),
    )
    return apply_args_file_values(parser.parse_args(argv), argv=argv)


def default_template_root() -> Path:
    """Return the template root implied by this script's committed location."""
    return Path(__file__).resolve().parents[2]


def resolve_existing_root(raw_root: str | None, *, default: Path | None, name: str) -> Path:
    """Resolve and validate a directory root supplied by the operator."""
    root = Path(raw_root).expanduser() if raw_root is not None else default
    if root is None:
        raise MaterializationError(f"{name} is required.")
    resolved = root.resolve()
    if not resolved.is_dir():
        raise MaterializationError(f"{name} does not exist or is not a directory.")
    return resolved


def is_same_or_descendant(path: Path, root: Path) -> bool:
    """Return whether ``path`` is equal to or below ``root`` after resolution."""
    resolved_path = path.resolve(strict=False)
    resolved_root = root.resolve()
    if resolved_path == resolved_root:
        return True
    try:
        resolved_path.relative_to(resolved_root)
    except ValueError:
        return False
    return True


def filesystem_compare_key(path: Path) -> str:
    """Return a normalized path string for worktree-list comparisons."""
    resolved = str(path.resolve(strict=False))
    if os.name == "nt":
        return resolved.casefold()
    return resolved


def command_output_text(value: str | bytes) -> str:
    """Return subprocess output as displayable text for one-line diagnostics."""
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return value


def command_detail(result: subprocess.CompletedProcess[Any]) -> str:
    """Return the first useful diagnostic line from a completed command."""
    for stream_value in (result.stderr, result.stdout):
        stream_text = command_output_text(stream_value)
        for line in stream_text.splitlines():
            if line.strip():
                return line.strip()
    return f"exit {result.returncode}"


def run_git(
    repo_root: Path | None,
    args: Sequence[str],
    *,
    env_overlay: Mapping[str, str] | None = None,
    text: bool = True,
) -> subprocess.CompletedProcess[Any]:
    """Run a local Git command without invoking a shell."""
    command = ["git"]
    if repo_root is not None:
        command.extend(("-C", str(repo_root)))
    command.extend(args)
    env = None if env_overlay is None else {**os.environ, **env_overlay}
    try:
        return subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=text,
            env=env,
        )
    except OSError as error:
        raise MaterializationError(f"Unable to run git: {os_error_summary(error)}") from error


def run_git_with_options(
    repo_root: Path | None,
    args: Sequence[str],
    *,
    env_overlay: Mapping[str, str] | None = None,
    text: bool = True,
) -> subprocess.CompletedProcess[Any]:
    """Run Git with optional child-process environment overlay and output mode."""
    return run_git(repo_root, args, env_overlay=env_overlay, text=text)


def run_source_git(
    repo_root: Path,
    args: Sequence[str],
    *,
    text: bool = True,
) -> subprocess.CompletedProcess[Any]:
    """Run a source-inspection Git command without lazy fetches or optional locks."""
    return run_git_with_options(
        repo_root,
        args,
        env_overlay=SOURCE_GIT_ENV_OVERLAY,
        text=text,
    )


def parse_git_version(value: str) -> GitVersion | None:
    """Parse the leading major/minor Git version, if present."""
    match = GIT_VERSION_RE.match(value.strip())
    if match is None:
        return None
    return GitVersion(
        major=int(match.group("major")),
        minor=int(match.group("minor")),
    )


def detect_git_version() -> GitVersion | None:
    """Return the locally available Git version without requiring a repository."""
    try:
        result = run_git_with_options(None, ["--version"])
    except MaterializationError:
        return None
    if result.returncode != 0:
        return None
    return parse_git_version(command_output_text(result.stdout))


def git_version_at_least(version: GitVersion | None, major: int, minor: int) -> bool:
    """Return whether a parsed Git version is at least ``major.minor``."""
    if version is None:
        return False
    return (version.major, version.minor) >= (major, minor)


def effective_git_config_value(repo_root: Path, key: str) -> subprocess.CompletedProcess[Any]:
    """Read a single effective Git config value for ``key``."""
    return run_source_git(repo_root, ["config", "--get", key])


def local_git_config_value(repo_root: Path, key: str) -> subprocess.CompletedProcess[Any]:
    """Read a single repository-local Git config value for ``key``."""
    return run_source_git(repo_root, ["config", "--local", "--get", key])


def git_config_key_names(repo_root: Path, pattern: str) -> tuple[str, ...] | str:
    """Return effective Git config key names matching ``pattern`` or an error reason."""
    result = run_source_git(repo_root, ["config", "--name-only", "--get-regexp", pattern])
    if result.returncode == 1:
        return ()
    if result.returncode != 0:
        return f"unable to inspect Git config keys matching {pattern!r}: {command_detail(result)}"
    names = tuple(
        line.strip() for line in command_output_text(result.stdout).splitlines() if line.strip()
    )
    return names


def partial_promisor_guard_reason(repo_root: Path, git_version: GitVersion | None) -> str | None:
    """Return a blocking partial/promisor reason for Git versions without safe no-fetch."""
    if git_version_at_least(git_version, 2, 45):
        return None

    partial_result = local_git_config_value(repo_root, "extensions.partialClone")
    if partial_result.returncode == 0:
        return (
            "the template source is a partial clone via extensions.partialClone; "
            "use Git detected as >=2.45 so GIT_NO_LAZY_FETCH is honored, or use a "
            "complete local template source"
        )
    if partial_result.returncode != 1:
        return (
            "unable to determine partial-clone state from extensions.partialClone: "
            f"{command_detail(partial_result)}"
        )

    promisor_names = git_config_key_names(repo_root, r"^remote\..*\.promisor$")
    if isinstance(promisor_names, str):
        return promisor_names
    for key_name in promisor_names:
        result = run_source_git(repo_root, ["config", "--type=bool", "--get", key_name])
        if result.returncode == 1:
            continue
        if result.returncode != 0:
            return (
                f"unable to determine promisor remote state from {key_name}: "
                f"{command_detail(result)}"
            )
        if command_output_text(result.stdout).strip() == "true":
            return (
                f"the template source has {key_name}=true; use Git detected as "
                ">=2.45 so GIT_NO_LAZY_FETCH is honored, or use a complete local "
                "template source"
            )

    filter_names = git_config_key_names(repo_root, r"^remote\..*\.partialCloneFilter$")
    if isinstance(filter_names, str):
        return filter_names
    for key_name in filter_names:
        result = effective_git_config_value(repo_root, key_name)
        if result.returncode == 1:
            continue
        if result.returncode != 0:
            return (
                f"unable to determine partial-clone filter state from {key_name}: "
                f"{command_detail(result)}"
            )
        if command_output_text(result.stdout).strip():
            return (
                f"the template source has non-empty {key_name}; use Git detected as "
                ">=2.45 so GIT_NO_LAZY_FETCH is honored, or use a complete local "
                "template source"
            )
    return None


def source_completeness_reason(
    source_worktree: Path, *, git_args_prefix: Sequence[str]
) -> str | None:
    """Return a reason when tracked source content may be incomplete."""
    result = run_source_git(
        source_worktree,
        [*git_args_prefix, "ls-files", "-z", "-v", "--full-name"],
        text=False,
    )
    if result.returncode != 0:
        return f"unable to inspect tracked source completeness: {command_detail(result)}"
    assert isinstance(result.stdout, bytes)
    for record in result.stdout.split(b"\0"):
        if not record:
            continue
        status_byte = record[:1]
        if status_byte == b"S":
            return "the template source has tracked files marked skip-worktree"
        if status_byte and 97 <= status_byte[0] <= 122:
            return "the template source has tracked files marked assume-unchanged"

    stage_result = run_source_git(
        source_worktree,
        [*git_args_prefix, "ls-files", "-z", "--stage", "--full-name"],
        text=False,
    )
    if stage_result.returncode != 0:
        return f"unable to inspect source gitlinks: {command_detail(stage_result)}"
    assert isinstance(stage_result.stdout, bytes)
    for record in stage_result.stdout.split(b"\0"):
        if not record:
            continue
        metadata, separator, _path = record.partition(b"\t")
        if separator != b"\t":
            return "unable to parse source index metadata while checking for gitlinks"
        fields = metadata.split()
        if len(fields) < 3:
            return "unable to parse source index metadata while checking for gitlinks"
        try:
            mode = int(fields[0], 8)
        except ValueError:
            return "unable to parse source index mode while checking for gitlinks"
        if mode == 0o160000:
            return (
                "the template source records submodule gitlinks, which are not "
                "supported for stampability"
            )
    return None


def source_fsmonitor_safety(
    repo_root: Path,
    git_version: GitVersion | None,
) -> tuple[tuple[str, ...], str | None]:
    """Choose safe index-query options before completeness or status can run hooks."""
    if git_version_at_least(git_version, 2, 36):
        return ("-c", "core.fsmonitor=false"), None
    # Git <=2.35.1 interprets even Boolean 'false' as a hook pathname.
    fsmonitor_result = effective_git_config_value(repo_root, "core.fsmonitor")
    if fsmonitor_result.returncode == 0:
        return (), (
            "core.fsmonitor is configured; stamping this source requires Git to "
            "be detected as >=2.36 so index probes can force-disable "
            "fsmonitor, or an explicit --last-reviewed-template-commit FULL_SHA"
        )
    if fsmonitor_result.returncode != 1:
        return (), (
            "unable to determine core.fsmonitor state before index inspection: "
            f"{command_detail(fsmonitor_result)}"
        )
    return (), None


def status_probe_not_stampable_reason(
    repo_root: Path, *, git_args_prefix: Sequence[str]
) -> str | None:
    """Return a not-stampable reason using the preselected safe index-query options."""
    result = run_source_git(
        repo_root,
        [
            *git_args_prefix,
            "status",
            "--porcelain=v1",
            "--untracked-files=all",
            "--ignore-submodules=none",
        ],
    )
    if result.returncode != 0:
        return f"unable to inspect source status: {command_detail(result)}"
    if command_output_text(result.stdout):
        return "the template source has dirty tracked files or untracked non-ignored files"
    return None


def verify_source_worktree_stampable(
    source_worktree: Path,
    git_version: GitVersion | None,
    *,
    fatal: bool,
) -> bool:
    """Return whether ``source_worktree`` passes the source stampability gates.

    Runs the partial/promisor, tracked-content completeness, and clean-status gates in
    order. Returns ``True`` only when every gate passes. When a gate fails, raises
    ``MaterializationError`` if ``fatal`` is true, otherwise returns ``False``. This
    validates stampability only; it does not resolve or return a source commit SHA.
    """
    partial_reason = partial_promisor_guard_reason(source_worktree, git_version)
    if partial_reason is not None:
        if fatal:
            raise MaterializationError(partial_reason)
        return False

    git_args_prefix, fsmonitor_reason = source_fsmonitor_safety(source_worktree, git_version)
    if fsmonitor_reason is not None:
        if fatal:
            raise MaterializationError(fsmonitor_reason)
        return False

    completeness_reason = source_completeness_reason(
        source_worktree, git_args_prefix=git_args_prefix
    )
    if completeness_reason is not None:
        if fatal:
            raise MaterializationError(completeness_reason)
        return False

    status_reason = status_probe_not_stampable_reason(
        source_worktree, git_args_prefix=git_args_prefix
    )
    if status_reason is not None:
        if fatal:
            raise MaterializationError(status_reason)
        return False
    return True


def detect_local_template_source(
    template_root: Path,
    git_version: GitVersion | None,
) -> LocalSourceDetection:
    """Discover whether a local template root is a complete, clean Git checkout."""
    try:
        inside_result = run_source_git(
            template_root,
            ["rev-parse", "--is-inside-work-tree"],
        )
        if (
            inside_result.returncode != 0
            or command_output_text(inside_result.stdout).strip() != "true"
        ):
            return LocalSourceDetection(
                not_stampable_reason="the template root is not a Git worktree"
            )

        top_level_result = run_source_git(template_root, ["rev-parse", "--show-toplevel"])
        if top_level_result.returncode != 0:
            return LocalSourceDetection(
                not_stampable_reason=(
                    "unable to identify source worktree root: "
                    f"{command_detail(top_level_result)}"
                )
            )
        top_level_text = command_output_text(top_level_result.stdout).strip()
        if not top_level_text:
            return LocalSourceDetection(
                not_stampable_reason="Git did not report a source worktree root"
            )
        source_worktree_root_path = Path(top_level_text).resolve(strict=False)
        source_worktree_root = str(source_worktree_root_path)
        if filesystem_compare_key(template_root) != filesystem_compare_key(
            source_worktree_root_path
        ):
            return LocalSourceDetection(
                source_worktree_root=source_worktree_root,
                not_stampable_reason=(
                    "--template-root is not the Git worktree top level; pass the "
                    "full template worktree root instead"
                ),
            )

        partial_reason = partial_promisor_guard_reason(template_root, git_version)
        if partial_reason is not None:
            return LocalSourceDetection(
                source_worktree_root=source_worktree_root,
                not_stampable_reason=partial_reason,
            )

        head_result = run_source_git(
            template_root,
            ["rev-parse", "--verify", "--end-of-options", "HEAD^{commit}"],
        )
        if head_result.returncode != 0:
            return LocalSourceDetection(
                source_worktree_root=source_worktree_root,
                not_stampable_reason="template source HEAD is not resolvable to a commit",
            )
        head_lines = command_output_text(head_result.stdout).strip().splitlines()
        if not head_lines:
            return LocalSourceDetection(
                source_worktree_root=source_worktree_root,
                not_stampable_reason="template source HEAD is not resolvable to a commit",
            )
        observed_source_sha = head_lines[0].strip()
        if not FULL_LOWERCASE_SHA_RE.fullmatch(observed_source_sha):
            return LocalSourceDetection(
                observed_source_sha=observed_source_sha,
                source_worktree_root=source_worktree_root,
                not_stampable_reason=(
                    "the resolved template source commit is not a 40-character "
                    "lowercase SHA-1 value; SHA-256 repositories are not stampable"
                ),
            )

        git_args_prefix, fsmonitor_reason = source_fsmonitor_safety(template_root, git_version)
        if fsmonitor_reason is not None:
            return LocalSourceDetection(
                observed_source_sha=observed_source_sha,
                source_worktree_root=source_worktree_root,
                not_stampable_reason=fsmonitor_reason,
            )

        completeness_reason = source_completeness_reason(
            template_root, git_args_prefix=git_args_prefix
        )
        if completeness_reason is not None:
            return LocalSourceDetection(
                observed_source_sha=observed_source_sha,
                source_worktree_root=source_worktree_root,
                not_stampable_reason=completeness_reason,
            )

        status_reason = status_probe_not_stampable_reason(
            template_root, git_args_prefix=git_args_prefix
        )
        if status_reason is not None:
            return LocalSourceDetection(
                observed_source_sha=observed_source_sha,
                source_worktree_root=source_worktree_root,
                not_stampable_reason=status_reason,
            )
        return LocalSourceDetection(
            observed_source_sha=observed_source_sha,
            source_worktree_root=source_worktree_root,
            stampable_sha=observed_source_sha,
        )
    except MaterializationError as error:
        return LocalSourceDetection(not_stampable_reason=str(error).replace("\n", " "))


def resolve_template_commit(template_repo: Path, raw_ref: str, *, label: str) -> str:
    """Resolve a locally available ref or SHA to a commit SHA without fetching."""
    result = run_source_git(
        template_repo,
        ["rev-parse", "--verify", "--end-of-options", f"{raw_ref}^{{commit}}"],
    )
    if result.returncode != 0:
        raise MaterializationError(
            f"Unable to resolve {label} {raw_ref!r} to a local commit without fetching: "
            f"{command_detail(result)}"
        )
    resolved_sha = command_output_text(result.stdout).strip().splitlines()[0]
    if not FULL_LOWERCASE_SHA_RE.fullmatch(resolved_sha):
        raise MaterializationError(
            f"Git resolved {label} {raw_ref!r} to an unexpected commit value."
        )
    return resolved_sha


def resolve_template_temp_root(raw_temp_root: str | None, target_root: Path) -> Path:
    """Return an absolute temporary-worktree parent that stays outside the target."""
    if raw_temp_root is not None:
        temp_root = Path(raw_temp_root).expanduser().resolve(strict=False)
        if is_same_or_descendant(temp_root, target_root):
            raise MaterializationError("--template-temp-root must not be inside --target-root.")
        try:
            temp_root.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            raise MaterializationError(
                f"Unable to create --template-temp-root: {os_error_summary(error)}"
            ) from error
        if not temp_root.is_dir():
            raise MaterializationError("--template-temp-root does not name a directory.")
        return temp_root.resolve()

    temp_root = Path(tempfile.gettempdir()).resolve()
    if is_same_or_descendant(temp_root, target_root):
        temp_root = target_root.parent.resolve()
    if is_same_or_descendant(temp_root, target_root):
        raise MaterializationError(
            "Unable to choose a temporary checkout root outside --target-root; "
            "pass --template-temp-root."
        )
    return temp_root


def manual_cleanup_command(template_repo: Path, temporary_checkout_path: Path) -> str:
    """Return the worktree-removal command rendered for the host platform's shell.

    POSIX output is paste-safe via ``shlex.join``. Windows output mirrors
    Python's ``subprocess`` argv parsing via ``list2cmdline`` for display parity
    and is not guaranteed paste-safe for a specific Windows shell.
    """
    command_parts = [
        "git",
        "-C",
        str(template_repo),
        "worktree",
        "remove",
        "--force",
        str(temporary_checkout_path),
    ]
    # Render the command for the host platform: POSIX uses ``shlex.join``
    # (paste-safe for POSIX shells); Windows uses ``list2cmdline`` for parity
    # with Python's subprocess argv parsing: a display rendering, not a
    # paste-safe guarantee for a specific Windows shell.
    if os.name == "nt":
        return subprocess.list2cmdline(command_parts)
    return shlex.join(command_parts)


def worktree_list_contains_path(template_repo: Path, worktree_path: Path) -> bool:
    """Return whether ``git worktree list --porcelain`` still lists a path."""
    result = run_source_git(template_repo, ["worktree", "list", "--porcelain"])
    if result.returncode != 0:
        raise MaterializationError(
            f"Unable to verify temporary worktree cleanup: {command_detail(result)}"
        )
    expected_key = filesystem_compare_key(worktree_path)
    for line in command_output_text(result.stdout).splitlines():
        if not line.startswith("worktree "):
            continue
        listed_path = Path(line.removeprefix("worktree "))
        if filesystem_compare_key(listed_path) == expected_key:
            return True
    return False


def create_temporary_source_checkout(
    *,
    template_repo: Path,
    target_root: Path,
    source_mode: str,
    source_value: str,
    resolved_source_sha: str,
    temp_root: Path,
    git_version: GitVersion | None,
) -> SourceCheckout:
    """Create a detached temporary worktree for a resolved upstream source commit."""
    try:
        temporary_parent = Path(
            tempfile.mkdtemp(prefix="template-source-", dir=str(temp_root))
        ).resolve()
    except OSError as error:
        raise MaterializationError(
            f"Unable to create temporary source checkout parent: {os_error_summary(error)}"
        ) from error

    temporary_checkout_path = temporary_parent / "checkout"
    if is_same_or_descendant(temporary_checkout_path, target_root):
        shutil.rmtree(temporary_parent, ignore_errors=True)
        raise MaterializationError(
            "Temporary source checkout path would be inside --target-root; "
            "pass --template-temp-root outside the target repository."
        )

    partial_reason = partial_promisor_guard_reason(template_repo, git_version)
    if partial_reason is not None:
        shutil.rmtree(temporary_parent, ignore_errors=True)
        raise MaterializationError(partial_reason)

    result = run_source_git(
        template_repo,
        [
            "-c",
            "core.sparseCheckout=false",
            "worktree",
            "add",
            "--detach",
            str(temporary_checkout_path),
            resolved_source_sha,
        ],
    )
    if result.returncode != 0:
        shutil.rmtree(temporary_parent, ignore_errors=True)
        raise MaterializationError(
            f"Unable to create temporary source checkout: {command_detail(result)}"
        )

    try:
        verify_source_worktree_stampable(
            temporary_checkout_path,
            git_version,
            fatal=True,
        )
    except MaterializationError:
        run_git(template_repo, ["worktree", "remove", "--force", str(temporary_checkout_path)])
        shutil.rmtree(temporary_parent, ignore_errors=True)
        raise

    command = manual_cleanup_command(template_repo, temporary_checkout_path)
    return SourceCheckout(
        template_root=temporary_checkout_path,
        template_repo=template_repo,
        temporary_parent=temporary_parent,
        temporary_checkout_path=temporary_checkout_path,
        summary=SourceSummary(
            target_root=str(target_root),
            template_root=str(temporary_checkout_path),
            source_mode=source_mode,
            source_value=source_value,
            resolved_source_sha=resolved_source_sha,
            source_repository=str(template_repo),
            temporary_checkout_path=str(temporary_checkout_path),
            cleanup_status="pending",
            manual_cleanup_command=command,
        ),
    )


def resolve_template_source(args: argparse.Namespace, target_root: Path) -> SourceCheckout:
    """Resolve the template source root, creating a private worktree when requested."""
    git_version = detect_git_version()
    if args.template_revision is not None and not FULL_LOWERCASE_SHA_RE.fullmatch(
        args.template_revision
    ):
        raise MaterializationError("--template-revision must be a full lowercase commit SHA.")

    if args.template_ref is None and args.template_revision is None:
        template_root = resolve_existing_root(
            args.template_root,
            default=default_template_root(),
            name="--template-root",
        )
        detection = detect_local_template_source(template_root, git_version)
        resolved_source_sha = (
            detection.stampable_sha if args.stamp_resolved_source_as_reviewed else None
        )
        return SourceCheckout(
            template_root=template_root,
            summary=SourceSummary(
                target_root=str(target_root),
                template_root=str(template_root),
                source_mode="template-root",
                source_value=str(template_root),
                resolved_source_sha=resolved_source_sha,
                observed_source_sha=detection.observed_source_sha,
                source_worktree_root=detection.source_worktree_root,
                source_not_stampable_reason=detection.not_stampable_reason,
            ),
        )

    raw_source_value = (
        args.template_ref if args.template_ref is not None else args.template_revision
    )
    assert raw_source_value is not None
    source_mode = "template-ref" if args.template_ref is not None else "template-revision"
    template_repo = resolve_existing_root(
        args.template_repo,
        default=target_root,
        name="--template-repo",
    )
    partial_reason = partial_promisor_guard_reason(template_repo, git_version)
    if partial_reason is not None:
        raise MaterializationError(partial_reason)
    label = "--template-ref" if args.template_ref is not None else "--template-revision"
    resolved_source_sha = resolve_template_commit(template_repo, raw_source_value, label=label)
    temp_root = resolve_template_temp_root(args.template_temp_root, target_root)
    return create_temporary_source_checkout(
        template_repo=template_repo,
        target_root=target_root,
        source_mode=source_mode,
        source_value=raw_source_value,
        resolved_source_sha=resolved_source_sha,
        temp_root=temp_root,
        git_version=git_version,
    )


def cleanup_source_checkout(source_checkout: SourceCheckout) -> CleanupFailure | None:
    """Remove a tool-created temporary source checkout and verify it is unregistered."""
    temporary_checkout_path = source_checkout.temporary_checkout_path
    template_repo = source_checkout.template_repo
    if temporary_checkout_path is None or template_repo is None:
        source_checkout.summary.cleanup_status = "not required"
        return None

    remove_args = ["worktree", "remove", "--force", str(temporary_checkout_path)]
    detail = ""
    try:
        first_result = run_git(template_repo, remove_args)
        retried = False
        final_result = first_result
        if first_result.returncode != 0:
            retried = True
            final_result = run_git(template_repo, remove_args)
    except MaterializationError as error:
        retried = False
        final_result = None
        detail = str(error)

    if final_result is not None and final_result.returncode == 0:
        try:
            if worktree_list_contains_path(template_repo, temporary_checkout_path):
                detail = "temporary worktree path still appears in git worktree list"
            else:
                source_checkout.summary.cleanup_status = (
                    "removed after retry" if retried else "removed"
                )
                if source_checkout.temporary_parent is not None:
                    try:
                        source_checkout.temporary_parent.rmdir()
                    except OSError:
                        pass
                return None
        except MaterializationError as error:
            detail = str(error)
    elif final_result is not None:
        detail = command_detail(final_result)

    command = source_checkout.summary.manual_cleanup_command or manual_cleanup_command(
        template_repo,
        temporary_checkout_path,
    )
    source_checkout.summary.cleanup_status = "failed"
    source_checkout.summary.cleanup_failure = detail
    source_checkout.summary.manual_cleanup_command = command
    return CleanupFailure(
        detail=detail,
        residual_worktree_path=temporary_checkout_path,
        source_repository=template_repo,
        manual_cleanup_command=command,
    )


def format_cleanup_failure_diagnostic(
    failure: CleanupFailure,
    *,
    materialization_succeeded: bool,
) -> str:
    """Return an actionable cleanup-failure diagnostic."""
    if materialization_succeeded:
        lead = (
            "Target tree was materialized successfully, but temporary source "
            "checkout cleanup failed. The target tree is intact."
        )
    else:
        lead = "Temporary source checkout cleanup also failed after the materialization error."
    return "\n".join(
        [
            lead,
            f"Cleanup detail: {failure.detail}",
            f"Residual worktree path: {failure.residual_worktree_path}",
            f"Source repository: {failure.source_repository}",
            f"Manual cleanup command: {failure.manual_cleanup_command}",
        ]
    )


def validate_reviewed_commit_matches_source(
    decisions: Decisions,
    source_checkout: SourceCheckout,
) -> None:
    """Reject an already-reviewed SHA that disagrees with the resolved source commit."""
    resolved_source_sha = source_checkout.summary.resolved_source_sha
    reviewed_commit = decisions.last_reviewed_template_commit
    if resolved_source_sha is None or reviewed_commit is None:
        return
    if reviewed_commit == resolved_source_sha:
        return
    raise MaterializationError(
        "template_sync.last_reviewed_template_commit / --last-reviewed-template-commit "
        f"({reviewed_commit}) does not match the resolved source SHA "
        f"({resolved_source_sha}). Omit the reviewed value until review is complete "
        "or supply the matching SHA."
    )


def decisions_with_reviewed_commit(decisions: Decisions, reviewed_commit: str) -> Decisions:
    """Return decisions with top-level and marker-shaped reviewed state aligned."""
    marker_data = MarkerDecisionData(
        last_reviewed_template_commit=reviewed_commit,
        included_modules=decisions.marker_data.included_modules,
        local_overrides=decisions.marker_data.local_overrides,
        local_path_ownership=decisions.marker_data.local_path_ownership,
        deferred_candidates=decisions.marker_data.deferred_candidates,
        protected_decisions=decisions.marker_data.protected_decisions,
        protected_guide_contract_waivers=(decisions.marker_data.protected_guide_contract_waivers),
    )
    return Decisions(
        source_repo=decisions.source_repo,
        last_reviewed_template_commit=reviewed_commit,
        included_modules=decisions.included_modules,
        marker_data=marker_data,
        raw_marker_fields=decisions.raw_marker_fields,
    )


def apply_reviewed_source_assertion(
    args: argparse.Namespace,
    decisions: Decisions,
    source_checkout: SourceCheckout,
) -> Decisions:
    """Apply the explicit source-as-reviewed assertion to decision state."""
    summary = source_checkout.summary
    had_explicit_reviewed_commit = decisions.last_reviewed_template_commit is not None
    if args.stamp_resolved_source_as_reviewed and not had_explicit_reviewed_commit:
        trusted_source_sha = summary.resolved_source_sha
        if trusted_source_sha is None:
            reason = summary.source_not_stampable_reason or "no trusted source SHA is available"
            raise MaterializationError(
                f"--stamp-resolved-source-as-reviewed requires a trusted source SHA; {reason}."
            )
        decisions = decisions_with_reviewed_commit(decisions, trusted_source_sha)
        summary.reviewed_commit_source = "--stamp-resolved-source-as-reviewed"
    elif decisions.last_reviewed_template_commit is not None:
        summary.reviewed_commit_source = "explicit reviewed commit"

    summary.accepted_reviewed_commit = decisions.last_reviewed_template_commit
    if summary.source_mode != "template-root":
        return decisions

    reviewed_commit = decisions.last_reviewed_template_commit
    if reviewed_commit is None:
        return decisions
    if summary.resolved_source_sha is not None:
        return decisions
    if summary.source_not_stampable_reason is not None:
        summary.source_cross_check_note = (
            f"source HEAD could not be cross-checked: {summary.source_not_stampable_reason}"
        )
    elif (
        not args.stamp_resolved_source_as_reviewed
        and summary.observed_source_sha is not None
        and reviewed_commit != summary.observed_source_sha
    ):
        summary.source_cross_check_note = (
            "explicit reviewed commit differs from the observed source HEAD; "
            "rerun with --stamp-resolved-source-as-reviewed after review only when "
            "the local source commit is the intended upstream lineage"
        )
    return decisions


def resolve_target_decisions_file(target_root: Path, raw_path: str) -> Path:
    """Resolve a repository-relative decisions file inside the target root."""
    return resolve_safe_repository_target_path(
        target_root,
        raw_path,
        field_name="--decisions-file",
    )


def load_validated_manifest_context(
    template_root: Path,
) -> tuple[dict[str, Any], tuple[str, ...], tuple[ManifestMapping, ...]]:
    """Load, schema-validate, and parse the template manifest."""
    manifest_path = resolve_safe_repository_target_path(
        template_root,
        DEFAULT_MANIFEST_PATH,
        field_name="manifest path",
    )
    manifest_schema_path = resolve_safe_repository_target_path(
        template_root,
        DEFAULT_MANIFEST_SCHEMA_PATH,
        field_name="manifest schema path",
    )
    manifest = load_yaml_mapping(manifest_path, template_root)
    manifest_schema = load_json_mapping(manifest_schema_path, template_root)
    validate_schema(manifest, manifest_schema, manifest_path, template_root)
    module_names, mappings = parse_manifest_mappings(manifest)
    template_manifest = manifest.get("template_manifest")
    if not isinstance(template_manifest, dict):
        raise MaterializationError("Manifest must contain template_manifest mapping.")
    template_manifest = cast(dict[str, Any], template_manifest)
    raw_modules = template_manifest.get("modules")
    if not isinstance(raw_modules, list):
        raise MaterializationError("Manifest module definitions must be a list.")
    module_order_parts: list[str] = []
    for module in cast(list[object], raw_modules):
        if not isinstance(module, dict):
            continue
        module_record = cast(dict[str, Any], module)
        module_name = module_record.get("name")
        if isinstance(module_name, str):
            module_order_parts.append(module_name)
    module_order = tuple(module_order_parts)
    if set(module_order) != set(module_names):
        raise MaterializationError("Manifest module order does not match parsed modules.")
    return manifest, module_order, mappings


def split_cli_modules(args: argparse.Namespace) -> frozenset[str] | None:
    """Return CLI-supplied modules, or ``None`` when no module flag was supplied."""
    raw_modules: list[str] = []
    raw_modules.extend(args.included_modules)
    if args.included_modules_csv:
        raw_modules.extend(
            module.strip() for module in args.included_modules_csv.split(",") if module.strip()
        )
    if not raw_modules:
        return None
    return frozenset(raw_modules)


def scalar_from_marker(
    marker_template_sync: dict[str, Any],
    field_name: str,
) -> str | None:
    """Return an optional string scalar from marker-shaped decision data."""
    value = marker_template_sync.get(field_name)
    if value is None:
        return None
    if not isinstance(value, str):
        raise MaterializationError(f"template_sync.{field_name} must be a string.")
    return value


def validate_cli_marker_overlap(
    *,
    cli_value: str | None,
    marker_value: str | None,
    name: str,
) -> str | None:
    """Return the resolved scalar after checking CLI and decisions-file agreement."""
    if cli_value is not None and marker_value is not None and cli_value != marker_value:
        raise MaterializationError(
            f"Conflicting {name}: CLI value does not match --decisions-file."
        )
    return marker_value if marker_value is not None else cli_value


def validate_module_overlap(
    *,
    cli_modules: frozenset[str] | None,
    marker_modules: frozenset[str] | None,
) -> frozenset[str] | None:
    """Return modules after checking CLI and decisions-file agreement."""
    if cli_modules is not None and marker_modules is not None and cli_modules != marker_modules:
        raise MaterializationError(
            "Conflicting included modules: CLI values do not match --decisions-file."
        )
    return marker_modules if marker_modules is not None else cli_modules


def load_decisions(
    args: argparse.Namespace,
    *,
    template_root: Path,
    target_root: Path,
    module_order: Sequence[str],
) -> Decisions:
    """Resolve CLI and marker-shaped decision data into one deterministic object."""
    marker_document: dict[str, Any] | None = None
    marker_data: MarkerDecisionData | None = None
    marker_template_sync: dict[str, Any] = {}
    if args.decisions_file is not None:
        decisions_path = resolve_target_decisions_file(target_root, args.decisions_file)
        marker_schema_path = resolve_safe_repository_target_path(
            template_root,
            DEFAULT_MARKER_SCHEMA_PATH,
            field_name="marker schema path",
        )
        marker_document = load_yaml_mapping(decisions_path, target_root)
        marker_schema = load_json_mapping(marker_schema_path, template_root)
        validate_schema(marker_document, marker_schema, decisions_path, target_root)
        parsed = parse_marker_decision_data(
            marker_document,
            validate_protected_decision_integrity=True,
        )
        marker_data = parsed
        raw_template_sync = marker_document.get("template_sync")
        if not isinstance(raw_template_sync, dict):
            raise MaterializationError("--decisions-file must contain template_sync mapping.")
        marker_template_sync = cast(dict[str, Any], raw_template_sync)

    cli_modules = split_cli_modules(args)
    marker_modules = marker_data.included_modules if marker_data is not None else None
    included_modules = validate_module_overlap(
        cli_modules=cli_modules,
        marker_modules=marker_modules,
    )
    if included_modules is None:
        raise MaterializationError(
            "At least one included module is required through CLI flags or --decisions-file."
        )

    source_repo = validate_cli_marker_overlap(
        cli_value=args.source_repo,
        marker_value=scalar_from_marker(marker_template_sync, "source_repo"),
        name="source repo",
    )
    if source_repo is None:
        raise MaterializationError("--source-repo is required unless --decisions-file supplies it.")

    reviewed_commit = validate_cli_marker_overlap(
        cli_value=args.last_reviewed_template_commit,
        marker_value=scalar_from_marker(
            marker_template_sync,
            "last_reviewed_template_commit",
        ),
        name="last reviewed template commit",
    )
    module_set = set(included_modules)
    unknown_modules = module_set - set(module_order)
    if unknown_modules:
        raise MaterializationError(
            "Selected module(s) are not defined by the manifest: "
            + ", ".join(sorted(unknown_modules))
        )
    selection_errors = validate_module_compatibility(module_set, ())
    if selection_errors:
        raise MaterializationError(" ".join(selection_errors))
    validate_agent_selection_migration(target_root, module_set, marker_data)

    raw_marker_fields: dict[str, Any] = {}
    if marker_document is not None:
        for field_name in MARKER_COPY_FIELDS:
            if field_name in marker_template_sync:
                raw_marker_fields[field_name] = marker_template_sync[field_name]

    resolved_marker_data = marker_data or MarkerDecisionData(
        last_reviewed_template_commit=reviewed_commit,
        included_modules=frozenset(included_modules),
        local_overrides=(),
        local_path_ownership=(),
        deferred_candidates=(),
        protected_decisions=(),
        protected_guide_contract_waivers=(),
    )
    validate_protected_file_decisions(
        resolved_marker_data.protected_decisions,
        resolved_marker_data.local_overrides,
        resolved_marker_data.deferred_candidates,
    )

    return Decisions(
        source_repo=source_repo,
        last_reviewed_template_commit=reviewed_commit,
        included_modules=frozenset(included_modules),
        marker_data=resolved_marker_data,
        raw_marker_fields=raw_marker_fields,
    )


MARKER_PLACEHOLDER_FIELDS = (
    "issue_label_policy",
    "issue_labels",
    "discussions_policy",
    "collaboration_policy_follow_up_status",
) + tuple(sorted(AZURE_DEVOPS_PLACEHOLDER_FIELDS))


def validate_agent_selection_migration(
    target_root: Path, modules: set[str], marker_data: MarkerDecisionData | None
) -> None:
    """Require an explicit retention/removal decision for legacy all-agent profiles."""
    if marker_data is None or "agent-instructions" not in modules:
        return
    agents = {
        "AGENTS.md": "agent-codex",
        "CLAUDE.md": "agent-claude",
        "GEMINI.md": "agent-gemini",
        ".hermes.md": "agent-hermes",
        ".cursor/rules/repository-instructions.mdc": "agent-cursor",
    }
    if any(name in marker_data.included_modules for name in agents.values()):
        return
    removals = {
        item.path for item in marker_data.protected_decisions if item.decision == REMOVAL_DECISION
    }
    undecided = [
        path
        for path, module in agents.items()
        if (target_root / path).exists() and module not in modules and path not in removals
    ]
    if undecided:
        raise MaterializationError(
            "Legacy agent selection requires explicit agent modules or reviewed protected "
            "removal decisions before cleanup: " + ", ".join(undecided)
        )


def apply_marker_placeholder_values(args: argparse.Namespace, decisions: Decisions) -> None:
    """Apply marker-recorded placeholder policy fields when CLI values are absent."""
    for field_name in MARKER_PLACEHOLDER_FIELDS:
        if field_name not in decisions.raw_marker_fields:
            continue
        if getattr(args, field_name) is not None:
            continue
        value = decisions.raw_marker_fields[field_name]
        if field_name == "issue_labels":
            value = list(value)
        setattr(args, field_name, value)


def license_preservation_local_override(source_path: str) -> LocalOverride:
    """Return the durable local override that preserves downstream license text."""
    return LocalOverride(
        path=LICENSE_TARGET_PATH,
        default_decision=TEMPLATE_SKIP_DECISION,
        reason=(
            f"Preserve downstream license text from {source_path} during first adoption; "
            "suppress template LICENSE. Source file remains for manual cleanup after review."
        ),
        is_directory=False,
    )


def raw_local_override_record(local_override: LocalOverride) -> dict[str, str]:
    """Return a marker-shaped local override record."""
    path = f"{local_override.path}/" if local_override.is_directory else local_override.path
    return {
        "path": path,
        "default_decision": local_override.default_decision,
        "reason": local_override.reason,
    }


def exact_local_override_for_path(
    local_overrides: tuple[LocalOverride, ...],
    relative_path: str,
) -> LocalOverride | None:
    """Return an exact local override for ``relative_path`` if one exists."""
    for local_override in local_overrides:
        if not local_override.is_directory and local_override.path == relative_path:
            return local_override
    return None


def append_license_preservation_override(
    decisions: Decisions,
    preservation: LicensePreservation,
) -> Decisions:
    """Add the license-preservation local override to marker-shaped decisions."""
    existing_override = exact_local_override_for_path(
        decisions.marker_data.local_overrides,
        LICENSE_TARGET_PATH,
    )
    local_overrides = decisions.marker_data.local_overrides
    raw_marker_fields = {
        field_name: value for field_name, value in decisions.raw_marker_fields.items()
    }

    if existing_override is not None:
        if existing_override.default_decision != TEMPLATE_SKIP_DECISION:
            raise MaterializationError(
                "License preservation requires template_sync.local_overrides for "
                f"{LICENSE_TARGET_PATH} to be absent or SKIP; found "
                f"{existing_override.default_decision}."
            )
    else:
        local_overrides = (*local_overrides, preservation.local_override)
        raw_records = list(raw_marker_fields.get("local_overrides", []))
        raw_records.append(raw_local_override_record(preservation.local_override))
        raw_marker_fields["local_overrides"] = raw_records

    marker_data = MarkerDecisionData(
        last_reviewed_template_commit=decisions.marker_data.last_reviewed_template_commit,
        included_modules=decisions.marker_data.included_modules,
        local_overrides=local_overrides,
        local_path_ownership=decisions.marker_data.local_path_ownership,
        deferred_candidates=decisions.marker_data.deferred_candidates,
        protected_decisions=decisions.marker_data.protected_decisions,
        protected_guide_contract_waivers=(decisions.marker_data.protected_guide_contract_waivers),
    )
    validate_protected_file_decisions(
        marker_data.protected_decisions,
        marker_data.local_overrides,
        marker_data.deferred_candidates,
    )
    return Decisions(
        source_repo=decisions.source_repo,
        last_reviewed_template_commit=decisions.last_reviewed_template_commit,
        included_modules=decisions.included_modules,
        marker_data=marker_data,
        raw_marker_fields=raw_marker_fields,
    )


def detected_license_source_candidates(target_root: Path) -> tuple[str, ...]:
    """Return unambiguous root-level alternate license source candidates."""
    try:
        children = tuple(target_root.iterdir())
    except OSError as error:
        raise MaterializationError(
            f"Unable to inspect target root for license candidates: {os_error_summary(error)}"
        ) from error

    candidates = [
        child.name for child in children if child.name.casefold() in LICENSE_SOURCE_CANDIDATE_NAMES
    ]
    return tuple(sorted(candidates, key=str.casefold))


def resolve_license_source_argument(args: argparse.Namespace, target_root: Path) -> str | None:
    """Resolve the requested or auto-detected license source path."""
    if args.license_source_path is not None and not args.preserve_existing_license:
        raise MaterializationError(
            "--license-source-path requires --preserve-existing-license so the "
            "owner's license-preservation intent is explicit."
        )
    if not args.preserve_existing_license:
        return None
    if args.license_source_path is not None:
        return cast(str, args.license_source_path)

    candidates = detected_license_source_candidates(target_root)
    if not candidates:
        raise MaterializationError(
            "--preserve-existing-license requires --license-source-path because no "
            "alternate root license candidate was found."
        )
    if len(candidates) > 1:
        raise MaterializationError(
            "Multiple candidate license source paths found: "
            + ", ".join(candidates)
            + ". Pass --license-source-path with the owner-approved source path."
        )
    return next(iter(candidates))


def read_license_source_bytes(
    source_path: Path,
    relative_path: str,
    *,
    read_bytes: Callable[[], bytes] | None = None,
) -> bytes:
    """Read and validate a preserved license as UTF-8 text bytes."""
    try:
        bytes_content = (read_bytes or source_path.read_bytes)()
    except OSError as error:
        raise MaterializationError(
            f"Unable to read license source {relative_path}: {os_error_summary(error)}"
        ) from error
    if b"\x00" in bytes_content:
        raise MaterializationError(
            f"License source {relative_path} must be a text file; NUL bytes were found."
        )
    try:
        bytes_content.decode("utf-8")
    except UnicodeDecodeError as error:
        raise MaterializationError(f"License source {relative_path} must be UTF-8 text.") from error
    return bytes_content


def resolve_license_preservation(
    args: argparse.Namespace,
    *,
    target_root: Path,
) -> LicensePreservation | None:
    """Resolve and validate first-adoption license preservation inputs."""
    raw_source_path = resolve_license_source_argument(args, target_root)
    if raw_source_path is None:
        return None

    source_path = resolve_safe_repository_target_path(
        target_root,
        raw_source_path,
        field_name="--license-source-path",
    )
    source_relative_path = source_path.relative_to(target_root).as_posix()
    if source_relative_path == LICENSE_TARGET_PATH:
        raise MaterializationError(
            "--license-source-path already names root LICENSE. For same-path "
            "license preservation, keep the downstream LICENSE and record "
            "template_sync.local_overrides with default_decision SKIP instead of "
            "using --preserve-existing-license."
        )

    target_license_path = resolve_safe_repository_target_path(
        target_root,
        LICENSE_TARGET_PATH,
        field_name="license target path",
    )
    if target_license_path.exists():
        raise MaterializationError(
            "Cannot preserve downstream license text to root LICENSE because root "
            "LICENSE already exists. Use the existing same-path local_overrides "
            "SKIP decision for root LICENSE, or resolve the conflicting license "
            "paths before rerunning."
        )
    if not source_path.exists():
        raise MaterializationError(f"License source path does not exist: {source_relative_path}.")
    if not source_path.is_file():
        raise MaterializationError(
            f"License source path must reference a regular text file: {source_relative_path}."
        )

    return LicensePreservation(
        source_path=source_relative_path,
        bytes_content=read_license_source_bytes(source_path, source_relative_path),
        local_override=license_preservation_local_override(source_relative_path),
    )


def ordered_modules(module_order: Sequence[str], selected_modules: Collection[str]) -> list[str]:
    """Return selected modules in manifest display order."""
    selected = set(selected_modules)
    return [module for module in module_order if module in selected]


def computed_marker_document(
    *,
    decisions: Decisions,
    module_order: Sequence[str],
) -> dict[str, Any]:
    """Build the schema-valid marker document computed for this run."""
    template_sync: dict[str, Any] = {
        "source_repo": decisions.source_repo,
    }
    if decisions.last_reviewed_template_commit is not None:
        template_sync["last_reviewed_template_commit"] = decisions.last_reviewed_template_commit
    template_sync["included_modules"] = ordered_modules(
        module_order,
        decisions.included_modules,
    )
    for field_name in MARKER_COPY_FIELDS:
        if field_name in decisions.raw_marker_fields:
            template_sync[field_name] = decisions.raw_marker_fields[field_name]
    return {"template_sync": template_sync}


def marker_yaml(marker_document: dict[str, Any]) -> str:
    """Return deterministic marker YAML."""
    return format_marker_yaml(marker_document)


def validate_computed_marker(
    marker_document: dict[str, Any],
    *,
    template_root: Path,
) -> None:
    """Validate the computed marker document and formatted marker YAML."""
    validate_marker_yaml_text(marker_yaml(marker_document), repo_root=template_root)


def marker_target_path(target_root: Path) -> Path:
    """Return the safe target path for the repository marker."""
    return resolve_safe_repository_target_path(
        target_root,
        DEFAULT_MARKER_PATH,
        field_name="marker path",
    )


def validate_inline_markers(text: str, *, relative_path: str) -> None:
    """Validate inline template-sync marker pairing before module pruning."""
    try:
        validate_inline_block_markers(text, relative_path=relative_path)
    except InlineBlockError as error:
        raise MaterializationError(str(error)) from error


def write_staged_candidate(
    *,
    template_root: Path,
    staging_root: Path,
    mappings: tuple[ManifestMapping, ...],
    included_modules: Collection[str],
    summary: Summary,
) -> tuple[str, ...]:
    """Populate the staging tree with retained, transformed template-managed files."""
    template_paths, skipped_symlinks = iter_safe_repository_files(template_root)
    if skipped_symlinks:
        managed_symlinks = [
            path.rstrip("/")
            for path in skipped_symlinks
            if selected_relation_for_path(path.rstrip("/"), mappings) is not None
        ]
        if managed_symlinks:
            raise MaterializationError(
                "Template-managed symlink path(s) cannot be materialized: "
                + ", ".join(sorted(managed_symlinks))
            )

    expected_workflows = retained_workflow_paths(template_paths, mappings, included_modules)
    contract_path = ".github/workflow-security-contract.yml"
    contract_relation = selected_relation_for_path(contract_path, mappings)
    contract_retained = contract_relation is not None and contract_relation.is_retained_by(
        included_modules
    )
    rendered_contract = None
    if expected_workflows or contract_retained:
        if not expected_workflows or not contract_retained or contract_path not in template_paths:
            raise MaterializationError(
                "Retained workflow inventory requires a present, retained workflow contract "
                "and at least one retained workflow; review the source contract and manifest."
            )
        rendered_contract = render_workflow_contract(
            template_root,
            mappings,
            included_modules,
            template_paths=template_paths,
            expected_workflows=expected_workflows,
        )

    staged_paths: list[str] = []
    for relative_path in template_paths:
        relation = selected_relation_for_path(relative_path, mappings)
        if relation is None:
            summary.unmapped.add(relative_path)
            continue
        if not relation.is_retained_by(included_modules):
            summary.excluded_paths.add(relative_path)
            continue

        classification = classify_repository_file(template_root, relative_path)
        destination = resolve_safe_repository_target_path(
            staging_root,
            relative_path,
            field_name="staged path",
        )
        destination.parent.mkdir(parents=True, exist_ok=True)
        if classification.is_byte_only:
            destination.write_bytes(classification.bytes_content)
            summary.byte_only.add(relative_path)
        else:
            assert classification.text is not None
            validate_inline_markers(classification.text, relative_path=relative_path)
            pruned_families = inline_block_families_to_prune(
                classification.text,
                included_modules,
                relative_path=relative_path,
            )
            if pruned_families:
                summary.pruning_preview.add(f"{relative_path}: {', '.join(pruned_families)}")
            filtered_text = remove_inline_blocks_for_modules(
                classification.text,
                included_modules,
                relative_path=relative_path,
            )
            if relative_path == contract_path:
                assert rendered_contract is not None
                filtered_text = rendered_contract
            destination.write_bytes(filtered_text.encode("utf-8"))
        staged_paths.append(relative_path)
    return tuple(sorted(staged_paths))


def trusted_tool_path(relative_path: str) -> Path:
    """Resolve executable helpers only from the running materializer's tool bundle."""
    path = resolve_safe_repository_target_path(
        TRUSTED_TOOL_ROOT, relative_path, field_name="trusted tool path"
    )
    if not path.is_file():
        raise MaterializationError(
            f"Running materializer's trusted helper is unavailable: {relative_path}; "
            "review and update the installed tool bundle."
        )
    return path


def retained_workflow_paths(
    template_paths: Collection[str],
    mappings: tuple[ManifestMapping, ...],
    included_modules: Collection[str],
) -> set[str]:
    """Select present top-level workflows using the existing manifest ownership resolver."""
    return {
        path
        for path in template_paths
        if PurePosixPath(path).parent == PurePosixPath(".github/workflows")
        and PurePosixPath(path).suffix in {".yml", ".yaml"}
        and (relation := selected_relation_for_path(path, mappings)) is not None
        and relation.is_retained_by(included_modules)
    }


def render_workflow_contract(
    template_root: Path,
    mappings: tuple[ManifestMapping, ...],
    included_modules: Collection[str],
    *,
    template_paths: Collection[str] | None = None,
    expected_workflows: set[str] | None = None,
) -> str:
    """Prune reviewed workflow controls using only manifest and inline-marker ownership.

    Validate the complete source before deriving fingerprints of pruned shell bodies.
    This is essential: unchecked regeneration would bless altered validation commands.
    The resulting runtime contract has no dependency on template-sync support.
    """
    try:
        import jsonschema
    except ImportError as error:
        raise MaterializationError(
            "Workflow contract rendering requires jsonschema. Install it in the "
            "Python environment running the materializer."
        ) from error
    validator_path = trusted_tool_path(".github/scripts/validate_workflow_security.py")
    spec = importlib.util.spec_from_file_location("workflow_policy_renderer", validator_path)
    if spec is None or spec.loader is None:
        raise MaterializationError("Cannot load workflow security validator")
    validator = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(validator)
    trusted_tool_path(validator.SCHEMA)
    if template_paths is None:
        template_paths, _ = iter_safe_repository_files(template_root)
    if expected_workflows is None:
        expected_workflows = retained_workflow_paths(template_paths, mappings, included_modules)
    if not expected_workflows:
        raise MaterializationError("Cannot render a workflow contract without retained workflows")
    try:
        validator.validate_repository(template_root, schema_root=TRUSTED_TOOL_ROOT)
        contract = validator.load_contract(template_root, schema_root=TRUSTED_TOOL_ROOT)
        rendered: dict[str, Any] = {}
        for path in contract["workflows"]:
            relation = selected_relation_for_path(path, mappings)
            if relation is None:
                raise MaterializationError(f"Unmapped workflow contract path: {path}")
            if relation.is_retained_by(included_modules):
                text = remove_inline_blocks_for_modules(
                    validator.read_text(template_root, path),
                    included_modules,
                    relative_path=path,
                )
                rendered[path] = validator.validate_workflow(text)
        if set(rendered) != expected_workflows:
            raise MaterializationError(
                "Retained workflow inventory differs from the reviewed contract; "
                f"missing: {', '.join(sorted(expected_workflows - set(rendered))) or 'none'}; "
                f"unexpected: {', '.join(sorted(set(rendered) - expected_workflows)) or 'none'}. "
                "Review the source contract and manifest together."
            )
        expected_examples: set[str] = set()
        for path in template_paths:
            if PurePosixPath(path).suffix not in {".md", ".mdc"}:
                continue
            relation = selected_relation_for_path(path, mappings)
            if relation is None or not relation.is_retained_by(included_modules):
                continue
            text = remove_inline_blocks_for_modules(
                validator.read_text(template_root, path),
                included_modules,
                relative_path=path,
            )
            if validator.check_examples(text):
                expected_examples.add(path)

        examples: list[str] = []
        for path in contract["examples"]:
            relation = selected_relation_for_path(path, mappings)
            if relation is None:
                raise MaterializationError(f"Unmapped workflow example: {path}")
            if relation.is_retained_by(included_modules):
                examples.append(path)
        if set(examples) != expected_examples:
            raise MaterializationError(
                "Retained workflow example inventory differs from the reviewed contract; "
                f"missing: {', '.join(sorted(expected_examples - set(examples))) or 'none'}; "
                f"unexpected: {', '.join(sorted(set(examples) - expected_examples)) or 'none'}. "
                "Review the source contract and manifest together."
            )
        return format_marker_yaml({"version": 1, "workflows": rendered, "examples": examples})
    except (
        ValueError,
        OSError,
        KeyError,
        TypeError,
        jsonschema.ValidationError,
        jsonschema.SchemaError,
    ) as error:
        raise MaterializationError(
            f"Workflow contract rendering failed: {format_cli_error(error)}; "
            "if the source format is newer, "
            "review and update the installed materializer tool bundle."
        ) from error


def placeholder_requested(args: argparse.Namespace) -> bool:
    """Return whether any placeholder replacement input was supplied."""
    return any(getattr(args, destination) is not None for destination in PLACEHOLDER_DESTS)


HELPER_FAILURE_OUTPUT_LINE_LIMIT = 20


def summarize_helper_failure(
    *,
    returncode: int,
    stdout: str,
    stderr: str,
    line_limit: int = HELPER_FAILURE_OUTPUT_LINE_LIMIT,
) -> str:
    """Return a bounded, path-safe placeholder-helper failure summary.

    The helper reports findings and ``PlaceholderError`` messages using
    repository-relative paths and only operates on the temporary staging tree,
    so a bounded tail of its captured ``stdout``/``stderr`` is safe to surface
    and lets the operator diagnose the failure without rerunning. Each stream is
    truncated to its most recent ``line_limit`` lines to keep the message
    minimal.
    """
    message_parts = [f"Placeholder helper failed with exit code {returncode}."]
    for label, stream in (("stdout", stdout), ("stderr", stderr)):
        lines = stream.splitlines()
        if not lines:
            continue
        shown = lines[-line_limit:]
        if len(lines) > len(shown):
            message_parts.append(
                f"Helper {label} (showing last {len(shown)} of {len(lines)} lines):"
            )
        else:
            message_parts.append(f"Helper {label}:")
        message_parts.extend(shown)
    return "\n".join(message_parts)


def run_placeholder_helper(
    *,
    args: argparse.Namespace,
    template_root: Path,
    staging_root: Path,
    decisions: Decisions,
    summary: Summary,
) -> None:
    """Run the existing approved placeholder helper against the staging tree."""
    if not placeholder_requested(args):
        summary.placeholder_notes.append("skipped: no placeholder inputs supplied")
        return
    helper_path = trusted_tool_path(PLACEHOLDER_HELPER_PATH)

    command = [
        sys.executable,
        str(helper_path),
        "replace",
        "--repo-root",
        str(staging_root),
        "--scan-mode",
        "retained-hard",
        "--manifest",
        str(template_root / DEFAULT_MANIFEST_PATH),
    ]
    for module_name in sorted(decisions.included_modules):
        command.extend(["--retained-module", module_name])
    if args.host_provider is not None:
        command.extend(["--host-provider", args.host_provider])
    if args.repository is not None:
        command.extend(["--repository", args.repository])
    optional_pairs = (
        ("--github-host", args.github_host),
        ("--azure-devops-organization", args.azure_devops_organization),
        ("--azure-devops-organization-url", args.azure_devops_organization_url),
        ("--azure-devops-project", args.azure_devops_project),
        ("--azure-devops-project-url", args.azure_devops_project_url),
        ("--azure-devops-repository", args.azure_devops_repository),
        ("--azure-devops-repository-url", args.azure_devops_repository_url),
        ("--azure-devops-clone-url", args.azure_devops_clone_url),
        ("--azure-devops-default-branch", args.azure_devops_default_branch),
        ("--azure-boards-policy", args.azure_boards_policy),
        ("--azure-repos-pr-template-policy", args.azure_repos_pr_template_policy),
        (
            "--azure-branch-policy-reviewer-guidance",
            args.azure_branch_policy_reviewer_guidance,
        ),
        ("--azure-security-intake-policy", args.azure_security_intake_policy),
        ("--azure-security-product-enablement", args.azure_security_product_enablement),
        ("--azure-dependency-update-policy", args.azure_dependency_update_policy),
        ("--codeowners-owner", args.codeowners_owner),
        ("--conduct-contact", args.conduct_contact),
        ("--conduct-contact-sentence", args.conduct_contact_sentence),
        ("--security-contact", args.security_contact),
        ("--security-contact-section", args.security_contact_section),
        ("--security-reporting-mode", args.security_reporting_mode),
        ("--issue-label-policy", args.issue_label_policy),
        ("--discussions-policy", args.discussions_policy),
        (
            "--collaboration-policy-follow-up-status",
            args.collaboration_policy_follow_up_status,
        ),
        ("--vscode-title", args.vscode_title),
        ("--package-name", args.package_name),
        ("--package-description", args.package_description),
        ("--package-author", args.package_author),
        ("--package-version", args.package_version),
    )
    for flag, value in optional_pairs:
        if value is not None:
            command.extend([flag, value])
    if args.issue_labels is not None:
        for label in args.issue_labels:
            command.extend(["--issue-label", label])
    if args.package_keywords is not None:
        for keyword in args.package_keywords:
            command.extend(["--package-keyword", keyword])

    with tempfile.TemporaryDirectory(prefix="template-adoption-byte-only-") as byte_directory:
        byte_root = Path(byte_directory)
        scan_input_path = byte_root / "placeholder-scan-inputs.json"
        scan_input = {
            "local_overrides": decisions.raw_marker_fields.get("local_overrides", []),
            "placeholder_waivers": decisions.raw_marker_fields.get("placeholder_waivers", []),
        }
        scan_input_path.write_text(json.dumps(scan_input, indent=2) + "\n", encoding="utf-8")
        command.extend(["--local-overrides-file", str(scan_input_path)])
        command.extend(["--placeholder-waivers-file", str(scan_input_path)])
        moved_byte_only_paths: list[tuple[Path, Path]] = []
        for relative_path in sorted(summary.byte_only):
            source = staging_root / relative_path
            if not source.exists():
                continue
            destination = byte_root / relative_path
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(source), str(destination))
            moved_byte_only_paths.append((source, destination))

        try:
            result = subprocess.run(
                command,
                cwd=TRUSTED_TOOL_ROOT,
                check=False,
                capture_output=True,
                text=True,
            )
        except OSError as error:
            raise MaterializationError(
                f"Unable to run placeholder helper: {os_error_summary(error)}"
            ) from error
        finally:
            for source, destination in moved_byte_only_paths:
                source.parent.mkdir(parents=True, exist_ok=True)
                if destination.exists():
                    shutil.move(str(destination), str(source))

    for line in result.stdout.splitlines():
        if line.startswith("  - ") and ":" in line:
            path = line.removeprefix("  - ").split(":", 1)[0].strip()
            if path:
                summary.placeholder_related.add(path)
        elif line:
            summary.placeholder_notes.append(line)
    for line in result.stderr.splitlines():
        if line:
            summary.placeholder_notes.append(f"stderr: {line}")
    if result.returncode != 0:
        raise MaterializationError(
            summarize_helper_failure(
                returncode=result.returncode,
                stdout=result.stdout,
                stderr=result.stderr,
            )
        )
    if not summary.placeholder_related:
        summary.placeholder_notes.append("no approved placeholders were replaced")


def apply_license_preservation(
    *,
    staging_root: Path,
    target_root: Path,
    staged_paths: Iterable[str],
    preservation: LicensePreservation,
    summary: Summary,
) -> tuple[str, ...]:
    """Write preserved downstream license text and suppress staged template licenses."""
    target_license_path = resolve_safe_repository_target_path(
        target_root,
        LICENSE_TARGET_PATH,
        field_name="license target path",
    )
    if target_license_path.exists():
        raise MaterializationError(
            "Cannot preserve downstream license text to root LICENSE because root "
            "LICENSE already exists. Resolve the license conflict before rerunning."
        )
    write_staged_bytes(target_license_path, LICENSE_TARGET_PATH, preservation.bytes_content)
    summary.created.add(LICENSE_TARGET_PATH)
    summary.residual_cleanup_paths.add(preservation.source_path)
    summary.license_notes.append(
        f"preserved downstream license text from {preservation.source_path} to LICENSE"
    )
    summary.license_notes.append(
        "suppressed template LICENSE and added local_overrides SKIP for LICENSE "
        "to the computed marker decision"
    )

    suppressed_paths = {LICENSE_TARGET_PATH, preservation.source_path}
    remaining_paths: list[str] = []
    for relative_path in staged_paths:
        if relative_path not in suppressed_paths:
            remaining_paths.append(relative_path)
            continue
        staged_path = resolve_safe_repository_target_path(
            staging_root,
            relative_path,
            field_name="staged path",
        )
        if staged_path.exists():
            staged_path.unlink()
    return tuple(sorted(remaining_paths))


def most_specific_local_override(
    relative_path: str,
    local_overrides: tuple[LocalOverride, ...],
) -> LocalOverride | None:
    """Return the most specific local override matching a repository path."""
    matches = [override for override in local_overrides if override.matches(relative_path)]
    if not matches:
        return None
    # Scan from the end so the last-listed override wins when specificity keys
    # tie, preserving the behavior of the earlier ``sorted(...)[-1]`` selection.
    return max(
        reversed(matches),
        key=lambda override: (len(override.path), not override.is_directory),
    )


def protected_decision_for_path(
    relative_path: str,
    protected_decisions: tuple[ProtectedFileDecision, ...],
) -> ProtectedFileDecision | None:
    """Return the exact protected-file decision for a path, if present."""
    for decision in protected_decisions:
        if decision.path == relative_path:
            return decision
    return None


def deferred_candidate_for_path(
    relative_path: str,
    deferred_candidates: tuple[DeferredProtectedCandidate, ...],
) -> DeferredProtectedCandidate | None:
    """Return the exact deferred protected candidate for a path, if present."""
    for candidate in deferred_candidates:
        if candidate.path == relative_path:
            return candidate
    return None


def conflict_for_non_protected_path(
    relative_path: str,
    *,
    recorded: bool,
    decision: str | None,
) -> Conflict:
    """Return an actionable non-protected conflict description."""
    if recorded:
        assert decision is not None
        return Conflict(
            path=relative_path,
            classification=f"recorded local_overrides {decision}",
            resolution=(
                "change the local_overrides entry to TAKE or SKIP to resolve, "
                "or keep it as a recorded unresolved decision"
            ),
            recorded=True,
        )
    return Conflict(
        path=relative_path,
        classification="unrecorded non-protected overwrite conflict",
        resolution=(
            "add template_sync.local_overrides with default_decision TAKE or SKIP, "
            "or record MERGE, DEFER, PROTECTED-REVIEW, or REMOVE-LOCAL"
        ),
        recorded=False,
    )


def conflict_for_protected_path(
    relative_path: str,
    *,
    recorded: bool,
    decision: str | None,
) -> Conflict:
    """Return an actionable protected-file conflict description."""
    if recorded:
        assert decision is not None
        return Conflict(
            path=relative_path,
            classification=f"recorded protected decision {decision}",
            resolution=(
                "change protected_file_decisions to TAKE or SKIP to resolve, "
                "or keep the protected path recorded as unresolved"
            ),
            recorded=True,
        )
    return Conflict(
        path=relative_path,
        classification="unrecorded protected-file decision required",
        resolution=(
            "add template_sync.protected_file_decisions or "
            "template_sync.deferred_protected_candidates for this concrete path"
        ),
        recorded=False,
    )


def write_staged_bytes(target_path: Path, relative_path: str, bytes_content: bytes) -> None:
    """Write materialized bytes to the target path."""
    write_repository_file_bytes(target_path, relative_path, bytes_content)


def ensure_regular_target(target_path: Path, relative_path: str) -> None:
    """Raise an actionable error when an existing target path is not a regular file.

    A directory or symlink where a regular file is expected would otherwise surface
    later as a path-stripped ``IsADirectoryError`` from a read or write. Failing here
    with the repository-relative ``relative_path`` lets the adopter resolve the
    conflict safely.
    """
    try:
        ensure_regular_repository_file_target(target_path, relative_path)
    except TemplateSyncMaterializationError as error:
        raise MaterializationError(str(error)) from error


def reconcile_staged_files(
    *,
    staging_root: Path,
    target_root: Path,
    staged_paths: Iterable[str],
    decisions: Decisions,
    summary: Summary,
) -> None:
    """Reconcile staged candidates into the target without overwriting conflicts."""
    for relative_path in sorted(staged_paths):
        staged_path = resolve_safe_repository_target_path(
            staging_root,
            relative_path,
            field_name="staged path",
        )
        target_path = resolve_safe_repository_target_path(
            target_root,
            relative_path,
            field_name="target path",
        )
        staged_bytes = staged_path.read_bytes()
        ensure_regular_target(target_path, relative_path)
        target_exists = target_path.exists()
        target_bytes = target_path.read_bytes() if target_exists else None
        is_identical = target_bytes == staged_bytes if target_exists else False
        is_protected = is_protected_instruction_path(relative_path)

        if is_protected:
            summary.protected.add(relative_path)
            decision = protected_decision_for_path(
                relative_path,
                decisions.marker_data.protected_decisions,
            )
            deferred_candidate = deferred_candidate_for_path(
                relative_path,
                decisions.marker_data.deferred_candidates,
            )
            if deferred_candidate is not None and decision is None:
                summary.deferred.add(relative_path)
                summary.conflicts.append(
                    conflict_for_protected_path(
                        relative_path,
                        recorded=True,
                        decision="deferred_protected_candidates",
                    )
                )
                continue
            if decision is None:
                summary.conflicts.append(
                    conflict_for_protected_path(
                        relative_path,
                        recorded=False,
                        decision=None,
                    )
                )
                continue
            if decision.decision == TEMPLATE_SKIP_DECISION:
                summary.skipped.add(relative_path)
                continue
            if decision.decision == TEMPLATE_TAKE_DECISION:
                if is_identical:
                    summary.unchanged.add(relative_path)
                elif target_exists:
                    write_staged_bytes(target_path, relative_path, staged_bytes)
                    summary.updated.add(relative_path)
                else:
                    write_staged_bytes(target_path, relative_path, staged_bytes)
                    summary.created.add(relative_path)
                continue
            if decision.decision == REMOVAL_DECISION:
                summary.recorded_removals.add(relative_path)
            else:
                summary.deferred.add(relative_path)
            summary.conflicts.append(
                conflict_for_protected_path(
                    relative_path,
                    recorded=True,
                    decision=decision.decision,
                )
            )
            continue

        local_override = most_specific_local_override(
            relative_path,
            decisions.marker_data.local_overrides,
        )
        if local_override is not None:
            summary.locally_overridden.add(relative_path)
            local_override_decision = local_override.default_decision
            if local_override_decision == TEMPLATE_SKIP_DECISION:
                summary.skipped.add(relative_path)
                continue
            if local_override_decision == TEMPLATE_TAKE_DECISION:
                if is_identical:
                    summary.unchanged.add(relative_path)
                elif target_exists:
                    write_staged_bytes(target_path, relative_path, staged_bytes)
                    summary.updated.add(relative_path)
                else:
                    write_staged_bytes(target_path, relative_path, staged_bytes)
                    summary.created.add(relative_path)
                continue
            if local_override_decision == REMOVAL_DECISION:
                summary.recorded_removals.add(relative_path)
            else:
                summary.deferred.add(relative_path)
            summary.conflicts.append(
                conflict_for_non_protected_path(
                    relative_path,
                    recorded=True,
                    decision=local_override_decision,
                )
            )
            continue

        if not target_exists:
            write_staged_bytes(target_path, relative_path, staged_bytes)
            summary.created.add(relative_path)
        elif is_identical:
            summary.unchanged.add(relative_path)
        else:
            summary.conflicts.append(
                conflict_for_non_protected_path(
                    relative_path,
                    recorded=False,
                    decision=None,
                )
            )


def reconcile_marker(
    *,
    target_root: Path,
    decisions: Decisions,
    summary: Summary,
    computed_marker_text: str,
) -> None:
    """Apply marker advancement rules independently from file reconciliation."""
    has_support = "template-sync-support" in decisions.included_modules
    if not has_support:
        summary.marker_status = "preview-only"
        summary.marker_reason = "template-sync-support is not included"
        summary.computed_marker_preview = computed_marker_text
        return
    if summary.unrecorded_conflicts:
        summary.marker_status = "preview-only"
        summary.marker_reason = "unrecorded conflicts remain"
        summary.computed_marker_preview = computed_marker_text
        return

    marker_path = marker_target_path(target_root)
    ensure_regular_target(marker_path, DEFAULT_MARKER_PATH)
    marker_bytes = computed_marker_text.encode("utf-8")
    existing_bytes = marker_path.read_bytes() if marker_path.exists() else None
    if existing_bytes == marker_bytes:
        summary.marker_status = "unchanged"
        summary.marker_reason = "existing marker already equals computed marker"
        return
    write_repository_relative_file_bytes(
        target_root,
        DEFAULT_MARKER_PATH,
        marker_bytes,
        field_name="marker path",
    )
    summary.marker_status = "updated"
    summary.marker_reason = "computed marker written"


def sorted_items(items: Iterable[str]) -> list[str]:
    """Return deterministic display items."""
    return sorted(set(items))


def print_section(title: str, items: Iterable[str]) -> None:
    """Print a deterministic bullet section."""
    values = sorted_items(items)
    print(f"{title}:")
    if not values:
        print("  - (none)")
        return
    for value in values:
        print(f"  - {value}")


def print_conflicts(conflicts: Iterable[Conflict]) -> None:
    """Print actionable conflict details."""
    conflict_items = sorted(conflicts, key=lambda conflict: conflict.path)
    print("Conflicted paths:")
    if not conflict_items:
        print("  - (none)")
        return
    for conflict in conflict_items:
        state = "recorded" if conflict.recorded else "unrecorded"
        print(f"  - {conflict.path}: {state}; {conflict.classification}")
        print(f"    authorize: {conflict.resolution}")


def print_source_summary(source: SourceSummary | None) -> None:
    """Print template source identity and cleanup status."""
    print("Source:")
    if source is None:
        print("  - (not recorded)")
        return
    print(f"  - target root: {source.target_root}")
    print(f"  - template root: {source.template_root}")
    print(f"  - source mode: {source.source_mode}")
    if source.source_mode == "template-ref":
        print(f"  - source ref: {source.source_value}")
    elif source.source_mode == "template-revision":
        print(f"  - source revision: {source.source_value}")
    else:
        print(f"  - source value: {source.source_value}")
    if source.source_mode == "template-root":
        if source.source_worktree_root is not None:
            print(f"  - source worktree root: {source.source_worktree_root}")
        if source.observed_source_sha is not None:
            if source.resolved_source_sha is not None:
                print(
                    "  - source commit: "
                    f"observed {source.observed_source_sha}; accepted for reviewed-state "
                    "assertion"
                )
            elif source.source_not_stampable_reason is not None:
                print(
                    "  - source commit: "
                    f"observed {source.observed_source_sha}; not stampable "
                    f"({source.source_not_stampable_reason})"
                )
            else:
                print(
                    "  - source commit: "
                    f"observed {source.observed_source_sha}; not accepted as "
                    "reviewed state"
                )
        else:
            reason = source.source_not_stampable_reason or "not discovered"
            print(f"  - source commit: not stampable ({reason})")
    else:
        print(f"  - resolved source SHA: {source.resolved_source_sha or '(not resolved)'}")
    if source.accepted_reviewed_commit is None:
        print(
            "  - accepted reviewed commit for computed marker: "
            "(none; complete review, then rerun with "
            "--stamp-resolved-source-as-reviewed or supply "
            "--last-reviewed-template-commit FULL_SHA)"
        )
    else:
        reviewed_source = source.reviewed_commit_source or "computed marker input"
        print(
            "  - accepted reviewed commit for computed marker: "
            f"{source.accepted_reviewed_commit} ({reviewed_source})"
        )
    if source.source_cross_check_note is not None:
        print(f"  - source cross-check note: {source.source_cross_check_note}")
    print(f"  - source repository: {source.source_repository or '(not used)'}")
    print(f"  - temporary checkout path: {source.temporary_checkout_path or '(not used)'}")
    print(f"  - cleanup status: {source.cleanup_status}")
    if source.cleanup_failure is not None:
        print(f"  - cleanup failure: {source.cleanup_failure}")
    if source.manual_cleanup_command is not None:
        print(f"  - manual cleanup command: {source.manual_cleanup_command}")


def print_summary(summary: Summary) -> None:
    """Emit the deterministic human-readable materialization summary."""
    print("Materialization summary")
    print_source_summary(summary.source)
    print(f"Default adoption mode: {summary.default_adoption_mode}")
    print_section("Retained modules", summary.retained_modules)
    print_section("Excluded modules", summary.excluded_modules)
    print_section("Created paths", summary.created)
    print_section("Updated paths", summary.updated)
    print_section("Unchanged paths", summary.unchanged)
    print_section("Skipped paths", summary.skipped)
    print_conflicts(summary.conflicts)
    print_section("Protected paths", summary.protected)
    print_section("Locally overridden paths", summary.locally_overridden)
    print_section("Deferred paths", summary.deferred)
    print_section("Recorded but not applied removals", summary.recorded_removals)
    print_section("Unmapped template paths", summary.unmapped)
    print_section("Excluded template-managed paths", summary.excluded_paths)
    print_section("Byte-only paths", summary.byte_only)
    print_section("Placeholder-related paths", summary.placeholder_related)
    print_section("Placeholder notes", summary.placeholder_notes)
    print_section("License preservation notes", summary.license_notes)
    print_section("Residual manual-cleanup paths", summary.residual_cleanup_paths)
    print_section("Pruning preview", summary.pruning_preview)
    print("Marker:")
    print(f"  - {summary.marker_status}: {summary.marker_reason}")
    if summary.recorded_conflicts:
        print("Recorded unresolved decisions remain:")
        for conflict in sorted(summary.recorded_conflicts, key=lambda item: item.path):
            print(f"  - {conflict.path}: {conflict.classification}")
    if summary.computed_marker_preview is not None:
        print("Computed marker preview:")
        for line in summary.computed_marker_preview.rstrip("\n").splitlines():
            print(f"  {line}")


def materialize(args: argparse.Namespace) -> Summary:
    """Run first-adoption materialization and return the operation summary."""
    target_root = resolve_existing_root(
        args.target_root,
        default=None,
        name="--target-root",
    )
    source_checkout: SourceCheckout | None = None
    summary: Summary | None = None
    primary_error: Exception | None = None

    try:
        source_checkout = resolve_template_source(args, target_root)
        template_root = source_checkout.template_root
        _manifest, module_order, mappings = load_validated_manifest_context(template_root)
        decisions = load_decisions(
            args,
            template_root=template_root,
            target_root=target_root,
            module_order=module_order,
        )
        apply_marker_placeholder_values(args, decisions)
        license_preservation = resolve_license_preservation(args, target_root=target_root)
        if license_preservation is not None:
            decisions = append_license_preservation_override(decisions, license_preservation)
        decisions = apply_reviewed_source_assertion(args, decisions, source_checkout)
        validate_reviewed_commit_matches_source(decisions, source_checkout)

        marker_document = computed_marker_document(
            decisions=decisions,
            module_order=module_order,
        )
        validate_computed_marker(marker_document, template_root=template_root)
        computed_marker_text = marker_yaml(marker_document)

        retained_modules = ordered_modules(module_order, decisions.included_modules)
        excluded_modules = [
            module for module in module_order if module not in decisions.included_modules
        ]
        summary = Summary(
            retained_modules=retained_modules,
            excluded_modules=excluded_modules,
            default_adoption_mode=args.default_adoption_mode,
            source=source_checkout.summary,
        )

        with tempfile.TemporaryDirectory(prefix="template-adoption-") as temporary_directory:
            staging_root = Path(temporary_directory)
            staged_paths = write_staged_candidate(
                template_root=template_root,
                staging_root=staging_root,
                mappings=mappings,
                included_modules=decisions.included_modules,
                summary=summary,
            )
            run_placeholder_helper(
                args=args,
                template_root=template_root,
                staging_root=staging_root,
                decisions=decisions,
                summary=summary,
            )
            render_instruction_profile(
                staging_root=staging_root,
                target_root=target_root,
                marker_document=marker_document,
            )
            if license_preservation is not None:
                staged_paths = apply_license_preservation(
                    staging_root=staging_root,
                    target_root=target_root,
                    staged_paths=staged_paths,
                    preservation=license_preservation,
                    summary=summary,
                )
            reconcile_staged_files(
                staging_root=staging_root,
                target_root=target_root,
                staged_paths=staged_paths,
                decisions=decisions,
                summary=summary,
            )

        reconcile_marker(
            target_root=target_root,
            decisions=decisions,
            summary=summary,
            computed_marker_text=computed_marker_text,
        )
    except Exception as error:  # noqa: BLE001
        # Catch every failure so the temporary source checkout below is always
        # cleaned up before the original error is re-raised, with any cleanup
        # failure attached to it.
        primary_error = error

    cleanup_failure = (
        cleanup_source_checkout(source_checkout) if source_checkout is not None else None
    )
    if primary_error is not None:
        if cleanup_failure is not None:
            raise MaterializationError(
                f"{format_cli_error(primary_error)}\n\n"
                + format_cleanup_failure_diagnostic(
                    cleanup_failure,
                    materialization_succeeded=False,
                )
            ) from primary_error
        raise primary_error
    assert summary is not None
    return summary


def format_cli_error(error: Exception) -> str:
    """Return a user-facing error message that never leaks filesystem paths.

    ``OSError`` and its subclasses (including ``shutil.Error``) render their
    default string form with the offending absolute local path, so summarize
    them through :func:`os_error_summary`. Domain errors already carry
    intentional, path-safe messages and are returned unchanged.
    """
    if isinstance(error, OSError):
        return os_error_summary(error)
    return f"{error}"


def main(argv: Sequence[str] | None = None) -> int:
    """Run the materialization CLI."""
    try:
        args = parse_args(argv)
        summary = materialize(args)
    except (
        MaterializationError,
        TemplateSyncMaterializationError,
        OSError,
    ) as error:
        print(f"ERROR: {format_cli_error(error)}", file=sys.stderr)
        return EXIT_RUNTIME_FAILURE

    print_summary(summary)
    if summary.source is not None and summary.source.cleanup_status == "failed":
        print(
            format_cleanup_failure_diagnostic(
                CleanupFailure(
                    detail=summary.source.cleanup_failure or "cleanup failed",
                    residual_worktree_path=Path(
                        summary.source.temporary_checkout_path or "(unknown)"
                    ),
                    source_repository=Path(summary.source.source_repository or "(unknown)"),
                    manual_cleanup_command=summary.source.manual_cleanup_command
                    or "git worktree remove --force <temporary-checkout-path>",
                ),
                materialization_succeeded=True,
            ),
            file=sys.stderr,
        )
        return EXIT_CLEANUP_FAILURE
    if summary.unrecorded_conflicts and not args.allow_conflicts:
        return EXIT_DECISIONS_REQUIRED
    return EXIT_SUCCESS


if __name__ == "__main__":
    sys.exit(main())
