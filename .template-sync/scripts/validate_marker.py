"""Validate a downstream template sync marker against the on-disk file set."""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import NoReturn

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from template_sync_materialization_helpers import (  # noqa: E402
    DEFAULT_MANIFEST_PATH,
    DEFAULT_MANIFEST_SCHEMA_PATH,
    DEFAULT_MARKER_PATH,
    DEFAULT_MARKER_SCHEMA_PATH,
    DEFAULT_REMOVE_LOCAL_AUTHORIZATION_TOKENS,
    REMOVAL_DECISION,
    DeferredProtectedCandidate,
    LocalOverride,
    LocalPathOwnership,
    ManifestMapping,
    MarkerPathOverlap,
    PathRelation,
    ProtectedFileDecision,
    ProtectedGuideContractWaiver,
    TemplateSyncMaterializationError,
    format_overlap_block,
    git_present_paths,
    is_locally_overridden,
    is_locally_owned_path,
    iter_safe_repository_files,
    load_json_mapping,
    load_yaml_mapping,
    local_path_ownership_summary,
    manifest_covers_directory,
    parse_manifest_compatibility_groups,
    parse_manifest_mappings,
    parse_marker_decision_data,
    path_has_symlink_component,
    repository_relative_path,
    resolve_repo_path,
    resolve_repo_root,
    selected_relation_for_path,
    unresolved_concrete_manifest_patterns,
    validate_module_compatibility,
    validate_protected_file_decisions,
    validate_schema,
)

# Public alias kept for callers such as validate_downstream_adoption and
# validate_instruction_contracts, which catch ``validate_marker.MarkerValidationError``.
MarkerValidationError = TemplateSyncMaterializationError

LOCAL_PATH_SUGGESTION_LIMIT = 10
LOCAL_PATH_SUGGESTION_COVERED_LIMIT = 3
UNRECORDED_LOCAL_PATH_LIMIT = 20


@dataclass(frozen=True)
class LocalPathOwnershipSuggestion:
    """A bounded, copy-ready local ownership record suggestion."""

    path: str
    covered_paths: tuple[str, ...]


@dataclass(frozen=True)
class MarkerValidationReport:
    """Validation result details to print for the operator."""

    included_modules: tuple[str, ...]
    manifest_mapping_count: int
    unsafe_managed_paths: tuple[str, ...]
    missing_expected_files: tuple[tuple[str, PathRelation], ...]
    leftover_files: tuple[tuple[str, PathRelation], ...]
    local_overrides: tuple[LocalOverride, ...]
    local_path_ownership: tuple[LocalPathOwnership, ...]
    unsafe_local_paths: tuple[str, ...]
    unrecorded_local_paths: tuple[str, ...]
    local_path_ownership_suggestions: tuple[LocalPathOwnershipSuggestion, ...]
    omitted_local_path_suggestion_count: int
    deferred_candidates: tuple[DeferredProtectedCandidate, ...]
    protected_decisions: tuple[ProtectedFileDecision, ...]
    protected_guide_contract_waivers: tuple[ProtectedGuideContractWaiver, ...]
    marker_path_overlaps: tuple[MarkerPathOverlap, ...]
    strict_local_path_ownership: bool

    @property
    def has_failures(self) -> bool:
        """Return whether validation failures were found.

        Retained-template inconsistencies (unsafe template-managed paths,
        missing expected files, and excluded leftovers) always count as
        failures. When ``strict_local_path_ownership`` is enabled, local path
        ownership findings (unsafe local symlinks and unrecorded local paths)
        are treated as failures too.
        """
        if self.unsafe_managed_paths or self.missing_expected_files or self.leftover_files:
            return True
        return self.strict_local_path_ownership and bool(
            self.unsafe_local_paths or self.unrecorded_local_paths
        )

    @property
    def has_local_path_ownership_findings(self) -> bool:
        """Return whether local path ownership warnings were found."""
        return bool(self.unsafe_local_paths or self.unrecorded_local_paths)


def parse_args(argv: list[str]) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Validate .template-sync/marker.yml against .template-sync/manifest.yml "
            "and the current repository files."
        ),
        epilog=(
            "Protected-file decision overlap checks fail when the same path has "
            "different protected_file_decisions.decision and local_overrides[].default_decision "
            "values, or when the same path appears in both protected_file_decisions and "
            "deferred_protected_candidates."
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
        default=DEFAULT_MARKER_PATH,
        help=f"Marker path relative to the repository root. Default: {DEFAULT_MARKER_PATH}",
    )
    parser.add_argument(
        "--manifest",
        default=DEFAULT_MANIFEST_PATH,
        help=f"Manifest path relative to the repository root. Default: {DEFAULT_MANIFEST_PATH}",
    )
    parser.add_argument(
        "--marker-schema",
        default=DEFAULT_MARKER_SCHEMA_PATH,
        help=(
            "Marker JSON Schema path relative to the repository root. "
            f"Default: {DEFAULT_MARKER_SCHEMA_PATH}"
        ),
    )
    parser.add_argument(
        "--manifest-schema",
        default=DEFAULT_MANIFEST_SCHEMA_PATH,
        help=(
            "Manifest JSON Schema path relative to the repository root. "
            f"Default: {DEFAULT_MANIFEST_SCHEMA_PATH}"
        ),
    )
    parser.add_argument(
        "--require-marker",
        action="store_true",
        help="Fail when the marker file is absent instead of treating the run as a no-op.",
    )
    parser.add_argument(
        "--strict-remove-local-phrasing",
        action="store_true",
        help=(
            "Fail REMOVE-LOCAL protected-file decisions whose authorization_basis "
            "does not contain a configured removal token. Off by default."
        ),
    )
    parser.add_argument(
        "--strict-local-path-ownership",
        action="store_true",
        help=(
            "Fail on Git-visible local path ownership findings: paths that are "
            "neither template-managed nor covered by template_sync.local_path_ownership, "
            "and local paths that are symlinks or otherwise resolve unsafely. "
            "Off by default."
        ),
    )
    parser.add_argument(
        "--remove-local-authorization-token",
        action="append",
        default=None,
        metavar="TOKEN",
        help=(
            "Case-insensitive substring token accepted by --strict-remove-local-phrasing. "
            "May be repeated. Defaults to: remov, delet."
        ),
    )
    parser.add_argument(
        "--remove-local-authorization-tokens",
        default=None,
        metavar="TOKENS",
        help=(
            "Comma-separated case-insensitive substring tokens accepted by "
            "--strict-remove-local-phrasing. Overrides are combined with repeated "
            "--remove-local-authorization-token values."
        ),
    )
    return parser.parse_args(argv)


def remove_local_authorization_tokens(
    repeated_tokens: list[str] | None,
    comma_separated_tokens: str | None,
) -> tuple[str, ...]:
    """Return normalized REMOVE-LOCAL strict-phrasing tokens."""
    raw_tokens: list[str] = []
    if comma_separated_tokens:
        raw_tokens.extend(comma_separated_tokens.split(","))
    if repeated_tokens:
        raw_tokens.extend(repeated_tokens)
    tokens = tuple(token.strip().lower() for token in raw_tokens if token.strip())
    return tokens or DEFAULT_REMOVE_LOCAL_AUTHORIZATION_TOKENS


def parse_marker(
    marker: dict[str, object],
) -> tuple[
    set[str],
    tuple[LocalOverride, ...],
    tuple[LocalPathOwnership, ...],
    tuple[DeferredProtectedCandidate, ...],
    tuple[ProtectedFileDecision, ...],
]:
    """Extract included modules, local path records, deferred candidates, and decisions."""
    marker_data = parse_marker_decision_data(marker)

    return (
        set(marker_data.included_modules),
        marker_data.local_overrides,
        marker_data.local_path_ownership,
        marker_data.deferred_candidates,
        marker_data.protected_decisions,
    )


def validate_local_path_ownership_path(
    repo_root: Path,
    local_path_ownership: LocalPathOwnership,
) -> None:
    """Validate local ownership path containment without requiring the leaf to exist."""
    if path_has_symlink_component(repo_root, local_path_ownership.path):
        raise MarkerValidationError(
            "template_sync.local_path_ownership[].path must not traverse a symlink: "
            f"{local_path_ownership.display_path}"
        )

    target_path = repo_root / local_path_ownership.path
    trusted_root = repo_root.resolve()
    try:
        target_path.resolve(strict=False).relative_to(trusted_root)
    except (OSError, ValueError) as error:
        raise MarkerValidationError(
            "template_sync.local_path_ownership[].path escapes the repository root: "
            f"{local_path_ownership.display_path}"
        ) from error

    if local_path_ownership.is_directory and target_path.exists() and not target_path.is_dir():
        raise MarkerValidationError(
            "template_sync.local_path_ownership[].path uses directory-prefix notation "
            f"but the path is not a directory: {local_path_ownership.display_path}"
        )


def broad_manifest_patterns_for_local_ownership(
    local_path_ownership: LocalPathOwnership,
    mappings: tuple[ManifestMapping, ...],
) -> tuple[str, ...]:
    """Return broad manifest patterns that overlap a local ownership record."""
    patterns: list[str] = []
    for mapping in mappings:
        if mapping.is_concrete:
            continue
        pattern = mapping.pattern
        if selected_relation_for_path(local_path_ownership.path, (mapping,)) is not None:
            patterns.append(pattern)
            continue
        if local_path_ownership.is_directory and pattern.startswith(
            f"{local_path_ownership.path}/"
        ):
            patterns.append(pattern)
    return tuple(sorted(set(patterns)))


def concrete_manifest_collisions_for_local_ownership(
    local_path_ownership: LocalPathOwnership,
    mappings: tuple[ManifestMapping, ...],
) -> tuple[str, ...]:
    """Return concrete manifest paths that exactly collide with local ownership."""
    return tuple(
        sorted(
            mapping.pattern
            for mapping in mappings
            if mapping.is_concrete and mapping.pattern == local_path_ownership.path
        )
    )


def validate_local_path_ownership_semantics(
    repo_root: Path,
    local_path_ownership: tuple[LocalPathOwnership, ...],
    mappings: tuple[ManifestMapping, ...],
) -> None:
    """Validate local ownership records against repository and manifest context."""
    errors: list[str] = []
    for record in local_path_ownership:
        try:
            validate_local_path_ownership_path(repo_root, record)
        except MarkerValidationError as error:
            errors.append(str(error))
            continue

        exact_collisions = concrete_manifest_collisions_for_local_ownership(record, mappings)
        broad_patterns = broad_manifest_patterns_for_local_ownership(record, mappings)
        if exact_collisions:
            errors.append(
                f"{record.display_path} exactly collides with concrete manifest-owned "
                "path(s): " + ", ".join(exact_collisions)
            )
            if record.overlap_exception_reason is not None:
                errors.append(
                    f"{record.display_path} overlap_exception_reason cannot override an "
                    "exact upstream-owned file collision."
                )
            continue
        if broad_patterns and record.overlap_exception_reason is None:
            errors.append(
                f"{record.display_path} overlaps broad manifest-owned pattern(s): "
                + ", ".join(broad_patterns)
                + "; add overlap_exception_reason to confirm the local ownership exception."
            )
        if record.overlap_exception_reason is not None and not broad_patterns:
            errors.append(
                f"{record.display_path} overlap_exception_reason is only permitted when "
                "the local ownership record overlaps a broad manifest-owned pattern."
            )
    if errors:
        raise MarkerValidationError(
            "Invalid local_path_ownership record(s):\n"
            + "\n".join(f"  - {error}" for error in errors)
        )


def suggested_local_ownership_path(relative_path: str) -> str:
    """Return a concise ownership path suggestion for an unrecorded file."""
    if "/" not in relative_path:
        return relative_path
    return relative_path.split("/", maxsplit=1)[0] + "/"


def build_local_path_ownership_suggestions(
    unrecorded_paths: tuple[str, ...],
) -> tuple[tuple[LocalPathOwnershipSuggestion, ...], int]:
    """Return bounded copy-ready ownership suggestions for unrecorded local paths."""
    grouped_paths: dict[str, list[str]] = {}
    for relative_path in sorted(unrecorded_paths):
        grouped_paths.setdefault(suggested_local_ownership_path(relative_path), []).append(
            relative_path
        )

    suggestions = tuple(
        LocalPathOwnershipSuggestion(path=path, covered_paths=tuple(grouped_paths[path]))
        for path in sorted(grouped_paths)[:LOCAL_PATH_SUGGESTION_LIMIT]
    )
    omitted_count = max(0, len(grouped_paths) - LOCAL_PATH_SUGGESTION_LIMIT)
    return suggestions, omitted_count


def unrecorded_local_path_findings(
    repo_root: Path,
    mappings: tuple[ManifestMapping, ...],
    local_path_ownership: tuple[LocalPathOwnership, ...],
    present_paths: set[str],
    skipped_symlinks: tuple[str, ...],
) -> tuple[tuple[str, ...], tuple[str, ...]]:
    """Return unsafe and unrecorded Git-visible local paths."""
    unsafe_paths: set[str] = set()
    for skipped_path in skipped_symlinks:
        is_directory_symlink = skipped_path.endswith("/")
        normalized_path = skipped_path.rstrip("/")
        is_git_visible = normalized_path in present_paths or any(
            path.startswith(f"{normalized_path}/") for path in present_paths
        )
        if not is_git_visible:
            continue
        if selected_relation_for_path(normalized_path, mappings) is not None:
            continue
        # A symlinked directory is stored by Git as a single entry, so manifest
        # globs under it (for example ``templates/python/**``) never match the
        # directory path itself. Treat it as template-managed so it surfaces as a
        # fatal managed-path finding instead of a non-fatal local warning.
        if is_directory_symlink and manifest_covers_directory(normalized_path, mappings):
            continue
        unsafe_paths.add(normalized_path)

    repo_root_resolved = repo_root.resolve()
    unrecorded_paths: list[str] = []
    for relative_path in sorted(present_paths):
        if selected_relation_for_path(relative_path, mappings) is not None:
            continue
        if is_locally_owned_path(relative_path, local_path_ownership):
            continue
        # Derive the unrecorded set from the full Git-visible list (the documented
        # ``git ls-files --cached --others --exclude-standard`` contract) and apply
        # per-path safety checks directly. Filtering by the safe-walk list would
        # hide files under walk-pruned directories such as build/, dist/, and
        # node_modules/ that a downstream may intentionally track. The safe walk
        # also prunes those directories, so a Git-visible symlink (or repo-escaping
        # path) tracked under one is absent from skipped_symlinks; classify such a
        # path as unsafe here so it is not dropped from both outputs. A directory
        # symlink over a manifest-managed tree is reported as a fatal managed-path
        # finding instead, so it is left out of the local set.
        if path_has_symlink_component(repo_root, relative_path):
            if not manifest_covers_directory(relative_path, mappings):
                unsafe_paths.add(relative_path)
            continue
        candidate = repo_root / relative_path
        try:
            candidate.resolve().relative_to(repo_root_resolved)
        except (OSError, ValueError):
            unsafe_paths.add(relative_path)
            continue
        if not candidate.is_file():
            continue
        unrecorded_paths.append(relative_path)

    return tuple(sorted(unsafe_paths)), tuple(unrecorded_paths)


def validate_marker_state(
    repo_root: Path,
    marker_path: Path,
    manifest_path: Path,
    marker_schema_path: Path,
    manifest_schema_path: Path,
    *,
    strict_remove_local_phrasing: bool = False,
    strict_local_path_ownership: bool = False,
    remove_local_tokens: tuple[str, ...] = DEFAULT_REMOVE_LOCAL_AUTHORIZATION_TOKENS,
) -> MarkerValidationReport:
    """Validate marker decisions against manifest mappings and on-disk files."""
    marker = load_yaml_mapping(marker_path, repo_root)
    manifest = load_yaml_mapping(manifest_path, repo_root)
    marker_schema = load_json_mapping(marker_schema_path, repo_root)
    manifest_schema = load_json_mapping(manifest_schema_path, repo_root)

    validate_schema(marker, marker_schema, marker_path, repo_root)
    validate_schema(manifest, manifest_schema, manifest_path, repo_root)

    manifest_modules, mappings = parse_manifest_mappings(manifest)
    marker_data = parse_marker_decision_data(marker)
    included_modules = set(marker_data.included_modules)
    local_overrides = marker_data.local_overrides
    local_path_ownership = marker_data.local_path_ownership
    deferred_candidates = marker_data.deferred_candidates
    protected_decisions = marker_data.protected_decisions
    overlaps = validate_protected_file_decisions(
        protected_decisions,
        local_overrides,
        deferred_candidates,
        strict_remove_local_phrasing=strict_remove_local_phrasing,
        remove_local_tokens=remove_local_tokens,
    )
    unknown_included_modules = included_modules - manifest_modules
    if unknown_included_modules:
        raise MarkerValidationError(
            "Marker includes module(s) that are not defined by the manifest: "
            + ", ".join(sorted(unknown_included_modules))
        )
    compatibility_groups = parse_manifest_compatibility_groups(
        manifest,
        module_names=manifest_modules,
    )
    compatibility_errors = validate_module_compatibility(
        included_modules,
        compatibility_groups,
    )
    if compatibility_errors:
        raise MarkerValidationError(
            "Unsupported module compatibility selection(s):\n"
            + "\n".join(f"  - {error}" for error in compatibility_errors)
        )

    validate_local_path_ownership_semantics(repo_root, local_path_ownership, mappings)

    present_paths = set(git_present_paths(repo_root))
    repository_files, skipped_symlinks = iter_safe_repository_files(repo_root)
    unsafe_local_paths, unrecorded_local_paths = unrecorded_local_path_findings(
        repo_root=repo_root,
        mappings=mappings,
        local_path_ownership=local_path_ownership,
        present_paths=present_paths,
        skipped_symlinks=skipped_symlinks,
    )
    local_path_ownership_suggestions, omitted_local_path_suggestion_count = (
        build_local_path_ownership_suggestions(unrecorded_local_paths)
    )
    unsafe_managed_paths: list[str] = []
    for relative_path in skipped_symlinks:
        is_directory_symlink = relative_path.endswith("/")
        normalized_relative_path = relative_path.rstrip("/")
        if normalized_relative_path not in present_paths:
            continue
        if is_locally_overridden(normalized_relative_path, local_overrides):
            continue
        relation = selected_relation_for_path(normalized_relative_path, mappings)
        # A symlinked directory over a manifest-managed tree is not matched by
        # ``selected_relation_for_path`` (glob patterns like ``dir/**`` never match
        # the directory itself), so recognize it as managed when the manifest has
        # patterns under it. This keeps directory symlinks a fatal finding.
        is_managed = relation is not None or (
            is_directory_symlink and manifest_covers_directory(normalized_relative_path, mappings)
        )
        if is_managed:
            unsafe_managed_paths.append(normalized_relative_path)

    leftover_files: list[tuple[str, PathRelation]] = []
    for relative_path in repository_files:
        if relative_path not in present_paths:
            continue
        if is_locally_overridden(relative_path, local_overrides):
            continue
        relation = selected_relation_for_path(relative_path, mappings)
        if relation is None:
            continue
        if not relation.is_retained_by(included_modules):
            leftover_files.append((relative_path, relation))

    missing_expected_files = unresolved_concrete_manifest_patterns(
        mappings=mappings,
        present_paths=present_paths,
        included_modules=included_modules,
        local_overrides=local_overrides,
    )

    return MarkerValidationReport(
        included_modules=tuple(sorted(included_modules)),
        manifest_mapping_count=len(mappings),
        unsafe_managed_paths=tuple(sorted(unsafe_managed_paths)),
        missing_expected_files=tuple(missing_expected_files),
        leftover_files=tuple(leftover_files),
        local_overrides=local_overrides,
        local_path_ownership=local_path_ownership,
        unsafe_local_paths=unsafe_local_paths,
        unrecorded_local_paths=unrecorded_local_paths,
        local_path_ownership_suggestions=local_path_ownership_suggestions,
        omitted_local_path_suggestion_count=omitted_local_path_suggestion_count,
        deferred_candidates=deferred_candidates,
        protected_decisions=protected_decisions,
        protected_guide_contract_waivers=marker_data.protected_guide_contract_waivers,
        marker_path_overlaps=overlaps,
        strict_local_path_ownership=strict_local_path_ownership,
    )


def print_report(report: MarkerValidationReport) -> None:
    """Print a human-readable validation report."""
    if report.has_failures:
        print("Marker-aware template sync validation failed.")
    else:
        print("Marker-aware template sync validation passed.")
    print(f"Included modules: {', '.join(report.included_modules)}")
    print(f"Manifest mappings checked: {report.manifest_mapping_count}")

    if report.unsafe_managed_paths:
        print("\nTemplate-managed paths that are symlinks or resolve unsafely:")
        for relative_path in report.unsafe_managed_paths:
            print(f"  - {relative_path}")

    if report.leftover_files:
        print("\nFiles present on disk but not retained by included modules:")
        for relative_path, relation in report.leftover_files:
            print(f"  - {relative_path} ({relation.description})")

    if report.missing_expected_files:
        print("\nConcrete mapped files expected for included modules but missing:")
        for relative_path, relation in report.missing_expected_files:
            print(f"  - {relative_path} ({relation.description})")

    if report.local_overrides:
        print("\nLocal overrides skipped:")
        for local_override in report.local_overrides:
            suffix = "/" if local_override.is_directory else ""
            print(
                f"  - {local_override.path}{suffix} "
                f"({local_override.default_decision}): {local_override.reason}"
            )

    if report.local_path_ownership:
        print("\nLocal path ownership records:")
        for record in report.local_path_ownership:
            print(f"  - {local_path_ownership_summary(record)}")

    if report.unsafe_local_paths:
        print("\nGit-visible local paths that are symlinks or resolve unsafely:")
        for relative_path in report.unsafe_local_paths:
            print(f"  - {relative_path}")

    if report.unrecorded_local_paths:
        print(
            "\nGit-visible paths not mapped by the manifest or "
            "template_sync.local_path_ownership:"
        )
        sampled_paths = report.unrecorded_local_paths[:UNRECORDED_LOCAL_PATH_LIMIT]
        for relative_path in sampled_paths:
            print(f"  - {relative_path}")
        remaining_paths = len(report.unrecorded_local_paths) - len(sampled_paths)
        if remaining_paths > 0:
            print(f"  - ... {remaining_paths} more path(s)")

        print("\nSuggested template_sync.local_path_ownership records:")
        if report.omitted_local_path_suggestion_count > 0:
            print(
                "  # Showing "
                f"{len(report.local_path_ownership_suggestions)} of "
                f"{len(report.local_path_ownership_suggestions) + report.omitted_local_path_suggestion_count} "
                "suggested record(s)."
            )
        for suggestion in report.local_path_ownership_suggestions:
            print(f"  - path: {suggestion.path}")
            if suggestion.path.endswith("/"):
                print("    reason: Downstream project owns this path family.")
            else:
                print("    reason: Downstream project owns this path.")
            sampled_covered_paths = suggestion.covered_paths[:LOCAL_PATH_SUGGESTION_COVERED_LIMIT]
            print("    # covers: " + ", ".join(sampled_covered_paths))
            remaining_covered_paths = len(suggestion.covered_paths) - len(sampled_covered_paths)
            if remaining_covered_paths > 0:
                print(f"    # plus {remaining_covered_paths} more path(s)")

    if report.has_local_path_ownership_findings:
        if report.strict_local_path_ownership:
            print("\nLocal path ownership findings are fatal because strict mode is enabled.")
        else:
            print(
                "\nLocal path ownership findings are warnings by default; rerun with "
                "--strict-local-path-ownership to make them fatal."
            )

    if report.deferred_candidates:
        print("\nDeferred protected candidates:")
        for candidate in report.deferred_candidates:
            print(f"  - {candidate.path} at {candidate.source_commit}: {candidate.reason}")

    if report.protected_decisions:
        print("\nProtected file decisions:")
        for protected_decision in report.protected_decisions:
            print(f"  - {protected_decision.path}: {protected_decision.decision}")
            if protected_decision.adoption_mode is not None:
                print(f"    adoption_mode: {protected_decision.adoption_mode}")
            if protected_decision.authorization_basis is not None:
                print(f"    authorization_basis: {protected_decision.authorization_basis}")
            if protected_decision.authorized_scope is not None:
                print(f"    authorized_scope: {protected_decision.authorized_scope}")
            if protected_decision.tailored_authorization_basis is not None:
                print(
                    "    tailored_authorization_basis: "
                    f"{protected_decision.tailored_authorization_basis}"
                )
            if protected_decision.reason is not None:
                print(f"    reason: {protected_decision.reason}")

    if report.protected_guide_contract_waivers:
        print("\nProtected guide contract waivers:")
        for waiver in report.protected_guide_contract_waivers:
            print(f"  - {waiver.path}: {waiver.contract_key}")
            if waiver.target_path is not None:
                print(f"    target_path: {waiver.target_path}")
            if waiver.target_module is not None:
                print(f"    target_module: {waiver.target_module}")
            if waiver.linked_local_override_path is not None:
                print(f"    linked_local_override_path: {waiver.linked_local_override_path}")
            print(f"    reason: {waiver.reason}")
            print(f"    authorization_basis: {waiver.authorization_basis}")

    remove_local_decisions = tuple(
        protected_decision
        for protected_decision in report.protected_decisions
        if protected_decision.decision == REMOVAL_DECISION
    )
    if remove_local_decisions:
        print("\nREMOVE-LOCAL authorizations:")
        for protected_decision in remove_local_decisions:
            print(f"  - {protected_decision.path}")
            print(f"    authorization_basis: {protected_decision.authorization_basis}")
            print(f"    authorized_scope: {protected_decision.authorized_scope}")
            print(f"    reason: {protected_decision.reason}")

    if report.marker_path_overlaps:
        print("\nProtected decision overlaps:")
        for overlap in report.marker_path_overlaps:
            print(format_overlap_block(overlap))

    if not report.has_failures:
        print("\nNo retained-template inconsistencies found.")


def fail(message: str) -> NoReturn:
    """Print an error and exit non-zero."""
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(1)


def main(argv: list[str] | None = None) -> int:
    """Run marker-aware downstream sync validation."""
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        repo_root = resolve_repo_root(args.repo_root)
        marker_path = resolve_repo_path(repo_root, args.marker)
        manifest_path = resolve_repo_path(repo_root, args.manifest)
        marker_schema_path = resolve_repo_path(repo_root, args.marker_schema)
        manifest_schema_path = resolve_repo_path(repo_root, args.manifest_schema)

        if not marker_path.exists():
            marker_relative_path = repository_relative_path(marker_path, repo_root)
            if args.require_marker:
                raise MarkerValidationError(
                    f"Marker is required but was not found at {marker_relative_path}."
                )
            print(f"No marker found at {marker_relative_path}; nothing to validate.")
            return 0

        report = validate_marker_state(
            repo_root=repo_root,
            marker_path=marker_path,
            manifest_path=manifest_path,
            marker_schema_path=marker_schema_path,
            manifest_schema_path=manifest_schema_path,
            strict_remove_local_phrasing=args.strict_remove_local_phrasing,
            strict_local_path_ownership=args.strict_local_path_ownership,
            remove_local_tokens=remove_local_authorization_tokens(
                args.remove_local_authorization_token,
                args.remove_local_authorization_tokens,
            ),
        )
    except MarkerValidationError as error:
        fail(str(error))

    print_report(report)
    return 1 if report.has_failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
