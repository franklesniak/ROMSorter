"""Report cleanup surfaces for modules excluded from a template adoption."""

from __future__ import annotations

import argparse
import fnmatch
import re
import sys
from collections import Counter, defaultdict
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import NoReturn, cast
from urllib.parse import unquote, urlsplit

import yaml  # type: ignore[import-untyped]

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

_SHARED_SCRIPTS = Path(__file__).resolve().parents[2] / ".github" / "scripts"
if str(_SHARED_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SHARED_SCRIPTS))

import instruction_contract_core  # noqa: E402
from template_sync_materialization_helpers import (  # noqa: E402
    DEFAULT_MANIFEST_PATH,
    DEFAULT_MANIFEST_SCHEMA_PATH,
    DEFAULT_MARKER_PATH,
    DEFAULT_MARKER_SCHEMA_PATH,
    INLINE_BLOCK_ANY_MODULES,
    DeferredProtectedCandidate,
    InlineBlockError,
    LocalOverride,
    ManifestMapping,
    PathRelation,
    ProtectedFileDecision,
    TemplateSyncMaterializationError,
    collect_live_inline_block_spans,
    fence_context_for_path,
    git_present_paths,
    inline_block_families_to_prune,
    inline_block_module_requirement,
    is_locally_overridden,
    is_protected_instruction_path,
    is_protected_prose_path,
    iter_safe_repository_files,
    live_inline_marker_lines,
    load_json_mapping,
    load_validated_marker_decision_data,
    load_yaml_mapping,
    manifest_pattern_matches_path,
    markdown_lines,
    os_error_summary,
    parse_manifest_mappings,
    repository_relative_path,
    resolve_repo_path,
    resolve_repo_root,
    selected_relation_for_path,
    validate_schema,
)

REFERENCE_LINK_FILE_SUFFIXES = frozenset({".md", ".mdc", ".yml", ".yaml"})
PRE_COMMIT_HOOK_FIELD_RE = re.compile(r"^\s+(?:-\s+)?(?P<field>id|alias):\s*(?P<value>[^#\n]+)")
WORKFLOW_RUN_RE = re.compile(r"^\s*run:\s*(?P<command>.+?)\s*$")
# Capture the contact-link ``url:`` value as a quoted scalar (kept intact, so a
# ``#`` fragment inside quotes is preserved) or an unquoted scalar whose ``#`` is
# only treated as a YAML comment when it follows whitespace. The caller strips
# any surrounding quotes from the captured value.
CONTACT_LINK_URL_RE = re.compile(
    r"""^\s*url:\s*(?P<target>(?P<quote>['"])[^'"\n]*(?P=quote)|[^\s#\n]\S*)\s*(?:#.*)?$"""
)
UPSTREAM_BLOB_PREFIX = "/franklesniak/copilot-repo-template/blob/HEAD/"
DEPENDABOT_ECOSYSTEM_MODULES: dict[str, tuple[str, tuple[str, ...]]] = {
    "npm": ("markdown", ("package.json", "package-lock.json")),
    "pip": (
        "python",
        (
            "requirements-pre-commit.txt",
            "pyproject.toml",
            "requirements.txt",
            "setup.py",
            "setup.cfg",
        ),
    ),
    # Directory surfaces end with "/" so the prefix branch of
    # dependency_file_is_retained_or_present() treats them as directories.
    "github-actions": ("github-actions", (".github/workflows/",)),
    "pre-commit": ("baseline", (".pre-commit-config.yaml",)),
}
GENERAL_VALIDATION_REFERENCES: dict[str, dict[str, tuple[str, ...]]] = {
    "json": {
        ".pre-commit-config.yaml": ("check-json",),
        ".github/workflows/data-ci.yml": ("pre-commit run check-json --all-files",),
    },
    "yaml": {
        ".pre-commit-config.yaml": ("check-yaml",),
        ".github/workflows/data-ci.yml": ("pre-commit run check-yaml --all-files",),
    },
}
PROTECTED_DOCUMENT_TOOLING_REFERENCES: dict[str, tuple[str, ...]] = {
    "python": (
        # ``pytest tests/ -m`` covers the split slow/non-slow invocations and
        # ``python -m pyright`` the type-check command introduced with Pyright.
        "pytest tests/ -m",
        "python -m pyright",
        "mypy src tests",
        "Black (Python)",
        "Ruff (Python)",
    ),
    "terraform": (
        "terraform fmt -check -recursive",
        "terraform test",
        "tflint --recursive",
        "terraform_version",
        "tflint_version",
        "terraform-fmt",
        "terraform-validate",
        "terraform-tflint",
        "hashicorp/setup-terraform",
        "terraform-linters/setup-tflint",
        "TFLint",
        "Terraform Test",
    ),
    "yaml": (
        "pre-commit run yamllint --all-files",
        "yamllint",
    ),
    "schema": (
        "pytest tests/test_schema_examples.py -v",
        "validate-example-config",
        "worked-example schema",
        "schema example fixtures",
    ),
}
COLLABORATION_TEMPLATE_TOOLING_REFERENCES: Mapping[str, Mapping[str, tuple[str, ...]]] = {
    "python": {
        ".github/pull_request_template.md": (
            "Python-Specific",
            "Minimum Python version",
            "`pytest`",
            "`mypy`",
        ),
    },
    "powershell": {
        ".github/pull_request_template.md": (
            "PowerShell-Specific",
            "PSScriptAnalyzer",
            "Invoke-Pester",
        ),
    },
    "schema": {
        ".github/pull_request_template.md": (
            "Schema-Specific",
            "schemas/examples/",
            "pytest tests/test_schema_examples.py -v",
            "check-jsonschema",
        ),
    },
    "github-actions": {
        ".github/pull_request_template.md": (
            "GitHub Actions-Specific",
            "actionlint",
        ),
    },
}
UPSTREAM_TEMPLATE_URL_RE = re.compile(
    re.escape("https://github.com/franklesniak/copilot-repo-template/") + r"(?:blob|tree)/HEAD/\S+"
)
UPSTREAM_TEMPLATE_MARKDOWN_LINK_RE = re.compile(
    r"\[[^\]\n]+\]\("
    + re.escape("https://github.com/franklesniak/copilot-repo-template/")
    + r"(?:blob|tree)/HEAD/[^)\s]+"
    + r"\)"
)

CATEGORY_REQUIRED = "required_cleanup"
CATEGORY_OPTIONAL_REFERENCE = "optional_reference_only_cleanup"
CATEGORY_PROTECTED = "protected_file_authorization_needed"
CATEGORY_LOCAL_OVERRIDE = "local_override_recorded"
CATEGORY_DEFERRED = "deferred_decision_recorded"
CATEGORY_DOCUMENTED_REFERENCE = "documented_reference_only_retention"
CATEGORY_FALSE_POSITIVE = "likely_false_positive_documented_reference"


class ReportError(Exception):
    """Raised when excluded-module reporting cannot continue."""


@dataclass(frozen=True)
class Finding:
    """One deterministic report finding."""

    rule_id: str
    category: str
    module: str
    path: str
    detail: str
    line_number: int | None = None

    @property
    def location(self) -> str:
        """Return the display location for this finding."""
        if self.line_number is None:
            return self.path
        return f"{self.path}:{self.line_number}"

    def sort_key(self) -> tuple[str, str, str, int, str]:
        """Return a stable sort key."""
        line_number = -1 if self.line_number is None else self.line_number
        return (self.module, self.rule_id, self.path, line_number, self.detail)

    def render(self) -> str:
        """Return a compact single-line representation."""
        return (
            f"{self.rule_id} | {self.category} | {self.module} | "
            f"{self.location} | {self.detail}"
        )


@dataclass(frozen=True)
class InlineBlock:
    """One complete template-sync inline block in a retained file."""

    path: str
    marker_name: str
    line_number: int
    content_lines: tuple[str, ...]


@dataclass(frozen=True)
class ReportState:
    """Parsed manifest, module, marker, and file state for reporting."""

    source: str
    manifest_modules: frozenset[str]
    mappings: tuple[ManifestMapping, ...]
    included_modules: frozenset[str]
    excluded_modules: frozenset[str]
    local_overrides: tuple[LocalOverride, ...]
    deferred_candidates: tuple[DeferredProtectedCandidate, ...]
    protected_decisions: tuple[ProtectedFileDecision, ...]
    present_files: frozenset[str]
    safe_files: tuple[str, ...]


@dataclass(frozen=True)
class ExcludedModuleReport:
    """The complete excluded-module cleanup report."""

    state: ReportState
    scopes: tuple[str, ...]
    findings: tuple[Finding, ...]


def parse_args(argv: list[str]) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Report cleanup surfaces and retained references for modules excluded "
            "from a partial template adoption."
        ),
    )
    parser.add_argument(
        "--repo-root",
        default=None,
        help=(
            "Repository root to inspect. Defaults to the parent of the .template-sync "
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
        "--included-module",
        action="append",
        default=None,
        metavar="MODULE",
        help=(
            "Explicit retained module for pre-marker cleanup planning. May be repeated. "
            "Do not pass this when .template-sync/marker.yml is present."
        ),
    )
    return parser.parse_args(argv)


def load_validated_manifest(
    repo_root: Path,
    manifest_path: Path,
    manifest_schema_path: Path,
) -> tuple[frozenset[str], tuple[ManifestMapping, ...]]:
    """Load the schema-valid template-sync manifest."""
    manifest = load_yaml_mapping(manifest_path, repo_root)
    manifest_schema = load_json_mapping(manifest_schema_path, repo_root)
    validate_schema(manifest, manifest_schema, manifest_path, repo_root)
    return parse_manifest_mappings(manifest)


def included_modules_from_args(
    *,
    explicit_included_modules: list[str] | None,
    marker_path: Path,
    marker_schema_path: Path,
    repo_root: Path,
    manifest_modules: frozenset[str],
) -> tuple[
    str,
    frozenset[str],
    tuple[LocalOverride, ...],
    tuple[DeferredProtectedCandidate, ...],
    tuple[ProtectedFileDecision, ...],
]:
    """Return included modules from exactly one deterministic source."""
    has_explicit_modules = bool(explicit_included_modules)
    marker_exists = marker_path.is_file()
    if has_explicit_modules and marker_exists:
        raise ReportError(
            "Ambiguous included-module state: marker-derived modules and explicit "
            "--included-module values were both supplied."
        )
    if has_explicit_modules:
        included_modules = frozenset(explicit_included_modules or ())
        unknown_modules = included_modules - manifest_modules
        if unknown_modules:
            raise ReportError(
                "Explicit included module(s) are not defined by the manifest: "
                + ", ".join(sorted(unknown_modules))
            )
        return (
            "explicit --included-module input",
            included_modules,
            (),
            (),
            (),
        )
    if not marker_exists:
        relative_marker_path = repository_relative_path(marker_path, repo_root)
        raise ReportError(
            f"No marker found at {relative_marker_path}; supply repeated "
            "--included-module values for pre-marker planning."
        )

    marker_data = load_validated_marker_decision_data(repo_root, marker_path, marker_schema_path)
    unknown_modules = marker_data.included_modules - manifest_modules
    if unknown_modules:
        raise ReportError(
            "Marker includes module(s) that are not defined by the manifest: "
            + ", ".join(sorted(unknown_modules))
        )
    relative_marker_path = repository_relative_path(marker_path, repo_root)
    return (
        f"marker ({relative_marker_path})",
        marker_data.included_modules,
        marker_data.local_overrides,
        marker_data.deferred_candidates,
        marker_data.protected_decisions,
    )


def build_state(
    *,
    repo_root: Path,
    marker_path: Path,
    marker_schema_path: Path,
    manifest_path: Path,
    manifest_schema_path: Path,
    explicit_included_modules: list[str] | None,
) -> ReportState:
    """Load manifest, module selection, marker data, and safe file state."""
    manifest_modules, mappings = load_validated_manifest(
        repo_root,
        manifest_path,
        manifest_schema_path,
    )
    (
        source,
        included_modules,
        local_overrides,
        deferred_candidates,
        protected_decisions,
    ) = included_modules_from_args(
        explicit_included_modules=explicit_included_modules,
        marker_path=marker_path,
        marker_schema_path=marker_schema_path,
        repo_root=repo_root,
        manifest_modules=manifest_modules,
    )
    present_files = frozenset(git_present_paths(repo_root))
    safe_files, _skipped_symlinks = iter_safe_repository_files(repo_root)
    safe_present_files = tuple(path for path in safe_files if path in present_files)
    return ReportState(
        source=source,
        manifest_modules=manifest_modules,
        mappings=mappings,
        included_modules=included_modules,
        excluded_modules=frozenset(manifest_modules - included_modules),
        local_overrides=local_overrides,
        deferred_candidates=deferred_candidates,
        protected_decisions=protected_decisions,
        present_files=present_files,
        safe_files=safe_present_files,
    )


def relation_modules(relation: PathRelation) -> frozenset[str]:
    """Return all modules named by a path relation."""
    return relation.requires_all | relation.requires_any


def missing_modules_for_relation(
    relation: PathRelation,
    included_modules: frozenset[str],
) -> tuple[str, ...]:
    """Return modules whose absence makes a relation excluded."""
    missing = set(relation.requires_all - included_modules)
    if relation.requires_any and not relation.requires_any.intersection(included_modules):
        missing.update(relation.requires_any)
    return tuple(sorted(missing))


def finding_category_for_path(relative_path: str, default_category: str) -> str:
    """Return the cleanup category for a path-sensitive finding."""
    if is_protected_instruction_path(relative_path):
        return CATEGORY_PROTECTED
    return default_category


def selected_modules_for_path(
    relative_path: str,
    mappings: tuple[ManifestMapping, ...],
) -> tuple[str, ...]:
    """Return selected manifest modules for a path."""
    relation = selected_relation_for_path(relative_path, mappings)
    if relation is None:
        return ()
    return tuple(sorted(relation_modules(relation)))


def mapping_modules(mapping: ManifestMapping) -> tuple[str, ...]:
    """Return modules directly named by one manifest mapping."""
    return tuple(sorted(mapping.requires_all | mapping.requires_any))


def pattern_is_workflow(pattern: str) -> bool:
    """Return whether a manifest pattern names GitHub Actions workflow files."""
    return pattern.startswith(".github/workflows/") and pattern.endswith((".yml", ".yaml"))


def load_yaml_file(repo_root: Path, relative_path: str) -> dict[str, object] | None:
    """Load a YAML mapping if a retained file exists and parses as one."""
    path = resolve_repo_path(repo_root, relative_path)
    if not path.is_file():
        return None
    try:
        parsed: object = yaml.safe_load(path.read_text(encoding="utf-8-sig"))
    except (OSError, yaml.YAMLError):
        return None
    return cast(dict[str, object], parsed) if isinstance(parsed, dict) else None


def workflow_jobs(repo_root: Path, relative_path: str) -> tuple[str, ...]:
    """Return workflow job identifiers for a workflow file, when readable."""
    document = load_yaml_file(repo_root, relative_path)
    if document is None:
        return ()
    jobs = document.get("jobs")
    if not isinstance(jobs, dict):
        return ()
    jobs = cast(dict[object, object], jobs)
    return tuple(sorted(str(job_name) for job_name in jobs if isinstance(job_name, str)))


def build_scope_lines(repo_root: Path, state: ReportState) -> tuple[str, ...]:
    """Return deterministic per-module cleanup scope summaries."""
    lines: list[str] = []
    mappings_by_module: dict[str, list[str]] = defaultdict(list)
    retained_cross_module_paths: dict[str, set[str]] = defaultdict(set)
    inline_families_by_module: dict[str, set[str]] = defaultdict(set)
    workflow_details_by_module: dict[str, set[str]] = defaultdict(set)
    protected_paths_by_module: dict[str, set[str]] = defaultdict(set)

    for mapping in state.mappings:
        for module in mapping_modules(mapping):
            if module in state.excluded_modules:
                mappings_by_module[module].append(mapping.pattern)
                if pattern_is_workflow(mapping.pattern):
                    jobs = workflow_jobs(repo_root, mapping.pattern)
                    if jobs:
                        workflow_details_by_module[module].add(
                            f"{mapping.pattern} jobs={', '.join(jobs)}"
                        )
                    else:
                        workflow_details_by_module[module].add(mapping.pattern)
                if is_protected_instruction_path(mapping.pattern):
                    protected_paths_by_module[module].add(mapping.pattern)

    for relative_path in state.safe_files:
        if is_locally_overridden(relative_path, state.local_overrides):
            continue
        relation = selected_relation_for_path(relative_path, state.mappings)
        if relation is None or not relation.is_retained_by(state.included_modules):
            continue
        try:
            text = resolve_repo_path(repo_root, relative_path).read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError):
            continue
        try:
            marker_families = inline_block_families_to_prune(
                text,
                state.included_modules,
                relative_path=relative_path,
            )
        except InlineBlockError:
            continue
        for name in marker_families:
            required_modules = inline_block_module_requirement(name)
            if required_modules is None:
                continue
            if name in INLINE_BLOCK_ANY_MODULES and not required_modules.isdisjoint(
                state.included_modules
            ):
                # OR-retention block is materialized via another member module,
                # so it is not a retained cross-module surface for the excluded
                # members.
                continue
            for module in sorted(required_modules & state.excluded_modules):
                retained_cross_module_paths[module].add(relative_path)
                inline_families_by_module[module].add(name)
                if is_protected_instruction_path(relative_path):
                    protected_paths_by_module[module].add(relative_path)

    for module in sorted(state.excluded_modules):
        module_lines = [
            "manifest-owned paths: " + join_or_none(sorted(mappings_by_module[module])),
            "retained cross-module paths: "
            + join_or_none(sorted(retained_cross_module_paths[module])),
            "inline block families: " + join_or_none(sorted(inline_families_by_module[module])),
            "workflow files/jobs: " + join_or_none(sorted(workflow_details_by_module[module])),
            "protected instruction paths: "
            + join_or_none(sorted(protected_paths_by_module[module])),
        ]
        lines.append(f"{module} | " + " | ".join(module_lines))
    return tuple(lines)


def join_or_none(values: Iterable[str]) -> str:
    """Join values for compact report output."""
    rendered = tuple(values)
    return "; ".join(rendered) if rendered else "None."


def manifest_owned_path_findings(state: ReportState) -> tuple[Finding, ...]:
    """Return findings for excluded-module-owned files still present."""
    findings: list[Finding] = []
    for relative_path in state.safe_files:
        if is_locally_overridden(relative_path, state.local_overrides):
            continue
        relation = selected_relation_for_path(relative_path, state.mappings)
        if relation is None or relation.is_retained_by(state.included_modules):
            continue
        category = finding_category_for_path(relative_path, CATEGORY_REQUIRED)
        for module in missing_modules_for_relation(relation, state.included_modules):
            findings.append(
                Finding(
                    rule_id="manifest-owned-path",
                    category=category,
                    module=module,
                    path=relative_path,
                    detail=f"Present file is not retained by included modules ({relation.description}).",
                )
            )
    return tuple(findings)


def collect_inline_blocks(repo_root: Path, relative_path: str) -> tuple[InlineBlock, ...]:
    """Return complete template-sync inline blocks in a text file."""
    path = resolve_repo_path(repo_root, relative_path)
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError):
        return ()

    return tuple(
        InlineBlock(
            path=relative_path,
            marker_name=block.marker_name,
            line_number=block.line_number,
            content_lines=block.content_lines,
        )
        for block in collect_live_inline_block_spans(
            text,
            relative_path=relative_path,
        )
    )


def inline_block_findings(repo_root: Path, state: ReportState) -> tuple[Finding, ...]:
    """Return stale inline-block and derived command/hook findings."""
    findings: list[Finding] = []
    for relative_path in state.safe_files:
        if is_locally_overridden(relative_path, state.local_overrides):
            continue
        relation = selected_relation_for_path(relative_path, state.mappings)
        if relation is None or not relation.is_retained_by(state.included_modules):
            continue
        try:
            blocks = collect_inline_blocks(repo_root, relative_path)
        except InlineBlockError as error:
            findings.append(
                Finding(
                    rule_id="inline-block.invalid",
                    category=finding_category_for_path(relative_path, CATEGORY_REQUIRED),
                    module="unknown",
                    path=relative_path,
                    line_number=error.line_number,
                    # Use the bare message: Finding.render() prints the location
                    # separately via ``location``, so str(error) (which embeds a
                    # "{path}:{line}: " prefix) would duplicate it in the report.
                    detail=error.message,
                )
            )
            continue
        for block in blocks:
            required_modules = inline_block_module_requirement(block.marker_name)
            if required_modules is None:
                findings.append(
                    Finding(
                        rule_id="inline-block.unknown",
                        category=finding_category_for_path(relative_path, CATEGORY_REQUIRED),
                        module="unknown",
                        path=relative_path,
                        line_number=block.line_number,
                        detail=f"Unknown template-sync inline marker family: {block.marker_name}.",
                    )
                )
                continue
            # OR-retention: the block is materialized whenever *any* named
            # module is included, so it is only stale when *every* named
            # module is excluded. Skip it while at least one is retained to
            # avoid flagging a correctly retained block for cleanup.
            is_retained_by_any_module = (
                block.marker_name in INLINE_BLOCK_ANY_MODULES
                and not required_modules.isdisjoint(state.included_modules)
            )
            if is_retained_by_any_module:
                continue
            missing_modules = tuple(sorted(required_modules - state.included_modules))
            if not missing_modules:
                continue
            for module in missing_modules:
                reference_only = block.marker_name.endswith("reference-only")
                if reference_only:
                    default_category = CATEGORY_OPTIONAL_REFERENCE
                    rule_id = "inline-block.reference-only"
                else:
                    default_category = CATEGORY_REQUIRED
                    rule_id = "inline-block.stale"
                findings.append(
                    Finding(
                        rule_id=rule_id,
                        category=finding_category_for_path(relative_path, default_category),
                        module=module,
                        path=relative_path,
                        line_number=block.line_number,
                        detail=(
                            f"Inline block {block.marker_name} remains but requires "
                            f"excluded module(s): {', '.join(missing_modules)}."
                        ),
                    )
                )
                findings.extend(derived_inline_block_findings(state, block, module))
                if reference_only and relation.notes:
                    findings.append(
                        Finding(
                            rule_id="manifest.reference-only-note",
                            category=CATEGORY_DOCUMENTED_REFERENCE,
                            module=module,
                            path=relative_path,
                            line_number=block.line_number,
                            detail="Manifest notes document reference-only inline-block handling.",
                        )
                    )
    return tuple(findings)


def derived_inline_block_findings(
    state: ReportState,
    block: InlineBlock,
    module: str,
) -> tuple[Finding, ...]:
    """Return hook and workflow command findings derived from stale inline blocks."""
    findings: list[Finding] = []
    category = finding_category_for_path(block.path, CATEGORY_REQUIRED)
    if block.path == ".pre-commit-config.yaml":
        hooks = extract_pre_commit_hook_names(block.content_lines)
        if hooks:
            findings.append(
                Finding(
                    rule_id="pre-commit-hook-group.stale",
                    category=category,
                    module=module,
                    path=block.path,
                    line_number=block.line_number,
                    detail=f"Hooks: {', '.join(hooks)}.",
                )
            )
    if block.path.startswith(".github/workflows/"):
        commands = extract_workflow_run_commands(block.content_lines)
        for command in commands:
            findings.append(
                Finding(
                    rule_id="workflow-validation-command.stale",
                    category=category,
                    module=module,
                    path=block.path,
                    line_number=block.line_number,
                    detail=f"Command: {command}.",
                )
            )
    return tuple(findings)


def extract_pre_commit_hook_names(lines: tuple[str, ...]) -> tuple[str, ...]:
    """Return hook ids and aliases mentioned inside a pre-commit inline block."""
    hooks: list[str] = []
    for line in lines:
        match = PRE_COMMIT_HOOK_FIELD_RE.match(line)
        if match is None:
            continue
        value = match.group("value").strip().strip("'\"")
        if value and value not in hooks:
            hooks.append(value)
    return tuple(hooks)


def extract_workflow_run_commands(lines: tuple[str, ...]) -> tuple[str, ...]:
    """Return run commands mentioned inside a workflow inline block."""
    commands: list[str] = []
    for line in lines:
        match = WORKFLOW_RUN_RE.match(line)
        if match is None:
            continue
        command = match.group("command").strip().strip("'\"")
        if command:
            commands.append(command)
    return tuple(commands)


def link_targets_outside_fences(path: Path) -> tuple[tuple[int, str], ...]:
    """Return Markdown-style link targets outside fenced code blocks."""
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return ()

    return instruction_contract_core.markdown_link_targets_from_text(
        text, fence_context=fence_context_for_path(path.name)
    )


def normalize_markdown_target(target: str) -> str:
    """Strip Markdown angle brackets from a link target."""
    if target.startswith("<") and target.endswith(">"):
        return target[1:-1]
    return target


resolve_relative_markdown_target = instruction_contract_core.resolve_relative_markdown_target


def resolve_upstream_blob_target(target: str) -> str | None:
    """Return the repo path from an upstream template blob URL, if present."""
    try:
        parsed = instruction_contract_core.split_markdown_destination(target)
    except ValueError:
        return None
    if parsed.scheme not in {"http", "https"} or parsed.netloc.lower() != "github.com":
        return None
    if not parsed.path.startswith(UPSTREAM_BLOB_PREFIX):
        return None
    return unquote(parsed.path.removeprefix(UPSTREAM_BLOB_PREFIX))


def resolve_github_blob_target(target: str) -> str | None:
    """Return the repo path from a GitHub blob/HEAD URL, if present."""
    try:
        parsed = urlsplit(target)
    except ValueError:
        return None
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        return None
    path_parts = [part for part in parsed.path.split("/") if part]
    if len(path_parts) < 5 or path_parts[2:4] != ["blob", "HEAD"]:
        return None
    return unquote("/".join(path_parts[4:]))


def reference_link_findings(repo_root: Path, state: ReportState) -> tuple[Finding, ...]:
    """Return stale and documented-reference findings for retained links."""
    findings: list[Finding] = []
    for relative_path in state.safe_files:
        if Path(relative_path).suffix not in REFERENCE_LINK_FILE_SUFFIXES:
            continue
        if is_locally_overridden(relative_path, state.local_overrides):
            continue
        relation = selected_relation_for_path(relative_path, state.mappings)
        if relation is None or not relation.is_retained_by(state.included_modules):
            continue
        path = resolve_repo_path(repo_root, relative_path)
        for line_number, target in link_targets_outside_fences(path):
            target_path = resolve_relative_markdown_target(relative_path, target)
            if target_path is not None:
                target_relation = selected_relation_for_path(target_path, state.mappings)
                if target_relation is None or target_relation.is_retained_by(
                    state.included_modules
                ):
                    continue
                if is_locally_overridden(target_path, state.local_overrides):
                    continue
                for module in missing_modules_for_relation(
                    target_relation,
                    state.included_modules,
                ):
                    findings.append(
                        Finding(
                            rule_id="markdown-link.excluded-target",
                            category=finding_category_for_path(relative_path, CATEGORY_REQUIRED),
                            module=module,
                            path=relative_path,
                            line_number=line_number,
                            detail=(
                                f"{target} resolves to excluded path {target_path} "
                                f"({target_relation.description})."
                            ),
                        )
                    )
                continue

            upstream_path = resolve_upstream_blob_target(target)
            if upstream_path is None:
                continue
            upstream_relation = selected_relation_for_path(upstream_path, state.mappings)
            if upstream_relation is None or upstream_relation.is_retained_by(
                state.included_modules
            ):
                continue
            for module in missing_modules_for_relation(upstream_relation, state.included_modules):
                findings.append(
                    Finding(
                        rule_id="markdown-link.upstream-reference",
                        category=CATEGORY_FALSE_POSITIVE,
                        module=module,
                        path=relative_path,
                        line_number=line_number,
                        detail=(
                            f"{target} points at excluded upstream template path "
                            f"{upstream_path}; this is documentation, not a local file."
                        ),
                    )
                )
    return tuple(findings)


def line_without_upstream_template_urls(line: str) -> str:
    """Remove intentional upstream template URLs before local-reference scanning."""
    without_links = UPSTREAM_TEMPLATE_MARKDOWN_LINK_RE.sub("", line)
    return UPSTREAM_TEMPLATE_URL_RE.sub("", without_links)


def lines_outside_inline_blocks(
    text: str,
    *,
    relative_path: str,
) -> tuple[tuple[int, str], ...]:
    """Return text lines outside template-sync inline blocks."""
    lines: list[tuple[int, str]] = []
    active_block_name: str | None = None

    try:
        live_lines = live_inline_marker_lines(text, relative_path=relative_path)
    except InlineBlockError:
        return tuple(enumerate(markdown_lines(text), 1))

    for line_number, line, marker in live_lines:
        line = line.rstrip("\r\n")
        if marker is not None:
            if marker.kind == "begin" and active_block_name is None:
                active_block_name = marker.name
                continue
            if marker.kind == "end" and marker.name == active_block_name:
                active_block_name = None
                continue

        if active_block_name is not None:
            continue
        lines.append((line_number, line))

    return tuple(lines)


def directory_is_excluded_module_owned(
    directory: str,
    module: str,
    state: ReportState,
) -> bool:
    """Return whether every present path below ``directory`` is excluded-owned."""
    child_paths = tuple(path for path in state.present_files if path.startswith(directory))
    if not child_paths:
        return False

    module_has_child = False
    for child_path in child_paths:
        if is_locally_overridden(child_path, state.local_overrides):
            return False
        child_relation = selected_relation_for_path(child_path, state.mappings)
        if child_relation is None or child_relation.is_retained_by(state.included_modules):
            return False
        missing_modules = missing_modules_for_relation(child_relation, state.included_modules)
        module_has_child = module_has_child or module in missing_modules

    return module_has_child


def parent_directory_tokens(relative_path: str) -> tuple[str, ...]:
    """Return parent directory tokens for a repository-relative path."""
    parts = relative_path.split("/")[:-1]
    return tuple(f"{'/'.join(parts[:index])}/" for index in range(1, len(parts) + 1))


def excluded_path_reference_tokens(state: ReportState) -> tuple[tuple[str, str, str], ...]:
    """Return module-owned local path tokens that should not survive exclusion."""
    token_rows: dict[tuple[str, str], str] = {}

    for relative_path in sorted(state.present_files):
        if is_locally_overridden(relative_path, state.local_overrides):
            continue
        relation = selected_relation_for_path(relative_path, state.mappings)
        if relation is None or relation.is_retained_by(state.included_modules):
            continue
        for module in missing_modules_for_relation(relation, state.included_modules):
            token_rows.setdefault((module, relative_path), relative_path)
            if relative_path.endswith(".schema.json"):
                token_rows.setdefault(
                    (module, relative_path.removesuffix(".schema.json")),
                    relative_path,
                )
            for directory in parent_directory_tokens(relative_path):
                if directory_is_excluded_module_owned(directory, module, state):
                    token_rows.setdefault((module, directory), relative_path)
                    token_rows.setdefault((module, directory.rstrip("/")), relative_path)

    return tuple(
        sorted(
            (
                (module, token, resolved_path)
                for (module, token), resolved_path in token_rows.items()
            ),
            key=lambda row: (-len(row[1]), row[0], row[1], row[2]),
        )
    )


def protected_document_prose_reference_findings_for_text(
    relative_path: str,
    text: str,
    state: ReportState,
    path_tokens: tuple[tuple[str, str, str], ...] | None = None,
) -> tuple[Finding, ...]:
    """Return stale excluded-module prose references in protected instruction text.

    ``path_tokens`` lets a caller compute ``excluded_path_reference_tokens(state)``
    once and reuse it across every scanned protected file; that helper sorts and
    walks ``state.present_files``, so recomputing it per file would scale poorly as
    the repository grows. The tokens are computed on demand when not supplied.
    """
    findings: list[Finding] = []
    if path_tokens is None:
        path_tokens = excluded_path_reference_tokens(state)

    for line_number, line in lines_outside_inline_blocks(
        text,
        relative_path=relative_path,
    ):
        searchable_line = line_without_upstream_template_urls(line)
        for module, token, resolved_path in path_tokens:
            if token not in searchable_line:
                continue
            findings.append(
                Finding(
                    rule_id="protected-document.prose-reference",
                    category=finding_category_for_path(relative_path, CATEGORY_REQUIRED),
                    module=module,
                    path=relative_path,
                    line_number=line_number,
                    detail=f"{token} references excluded path {resolved_path}.",
                )
            )
        for module, tokens in PROTECTED_DOCUMENT_TOOLING_REFERENCES.items():
            if module not in state.excluded_modules:
                continue
            for token in tokens:
                if token not in searchable_line:
                    continue
                findings.append(
                    Finding(
                        rule_id="protected-document.prose-reference",
                        category=finding_category_for_path(relative_path, CATEGORY_REQUIRED),
                        module=module,
                        path=relative_path,
                        line_number=line_number,
                        detail=f"{token} documents excluded {module} tooling.",
                    )
                )

    return tuple(findings)


def protected_document_prose_reference_findings(
    repo_root: Path,
    state: ReportState,
) -> tuple[Finding, ...]:
    """Return stale prose-reference findings for retained protected documents."""
    findings: list[Finding] = []
    path_tokens = excluded_path_reference_tokens(state)
    for relative_path in state.safe_files:
        if not is_protected_prose_path(relative_path):
            continue
        if is_locally_overridden(relative_path, state.local_overrides):
            continue
        relation = selected_relation_for_path(relative_path, state.mappings)
        if relation is None or not relation.is_retained_by(state.included_modules):
            continue
        try:
            text = resolve_repo_path(repo_root, relative_path).read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError):
            continue
        findings.extend(
            protected_document_prose_reference_findings_for_text(
                relative_path, text, state, path_tokens
            )
        )
    return tuple(findings)


def active_contact_link_urls(
    text: str,
    *,
    relative_path: str,
) -> tuple[tuple[int, str], ...]:
    """Return active issue-template contact-link URL values outside inline blocks."""
    targets: list[tuple[int, str]] = []
    for line_number, line in lines_outside_inline_blocks(
        text,
        relative_path=relative_path,
    ):
        if line.lstrip().startswith("#"):
            continue
        match = CONTACT_LINK_URL_RE.match(line)
        if match is None:
            continue
        target = match.group("target").strip().strip("'\"")
        targets.append((line_number, target))
    return tuple(targets)


def collaboration_template_reference_findings(
    repo_root: Path,
    state: ReportState,
) -> tuple[Finding, ...]:
    """Return stale excluded-module references in issue and PR collaboration files."""
    findings: list[Finding] = []
    collaboration_paths = {
        path
        for paths_and_tokens in COLLABORATION_TEMPLATE_TOOLING_REFERENCES.values()
        for path in paths_and_tokens
    } | {".github/ISSUE_TEMPLATE/config.yml"}

    for relative_path in sorted(collaboration_paths):
        if relative_path not in state.present_files or not retained_path(relative_path, state):
            continue
        if is_locally_overridden(relative_path, state.local_overrides):
            continue
        try:
            text = resolve_repo_path(repo_root, relative_path).read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError):
            continue

        for module, paths_and_tokens in COLLABORATION_TEMPLATE_TOOLING_REFERENCES.items():
            if module not in state.excluded_modules:
                continue
            for token in paths_and_tokens.get(relative_path, ()):
                for line_number, line in lines_outside_inline_blocks(
                    text,
                    relative_path=relative_path,
                ):
                    searchable_line = line_without_upstream_template_urls(line)
                    if token not in searchable_line:
                        continue
                    findings.append(
                        Finding(
                            rule_id="collaboration-template.prose-reference",
                            category=finding_category_for_path(
                                relative_path,
                                CATEGORY_REQUIRED,
                            ),
                            module=module,
                            path=relative_path,
                            line_number=line_number,
                            detail=f"{token} documents excluded {module} tooling.",
                        )
                    )

        if relative_path != ".github/ISSUE_TEMPLATE/config.yml":
            continue
        for line_number, target in active_contact_link_urls(
            text,
            relative_path=relative_path,
        ):
            target_path = resolve_github_blob_target(target)
            if target_path is None:
                continue
            target_relation = selected_relation_for_path(target_path, state.mappings)
            if target_relation is None or target_relation.is_retained_by(state.included_modules):
                continue
            for module in missing_modules_for_relation(target_relation, state.included_modules):
                findings.append(
                    Finding(
                        rule_id="contact-link.excluded-target",
                        category=finding_category_for_path(relative_path, CATEGORY_REQUIRED),
                        module=module,
                        path=relative_path,
                        line_number=line_number,
                        detail=(
                            f"{target} points at excluded path {target_path} "
                            f"({target_relation.description})."
                        ),
                    )
                )

    return tuple(findings)


def dependency_file_is_retained_or_present(
    dependency_path: str,
    state: ReportState,
) -> bool:
    """Return whether a Dependabot-scanned path is retained or intentionally kept.

    A path qualifies when the selected modules retain it or when it is present
    but covered by a marker local override. An overridden file is intentionally
    kept downstream, so it can still justify its Dependabot ecosystem entry and
    must not make that ecosystem look stale.
    """
    if dependency_path.endswith("/"):
        prefix = dependency_path.lstrip("/")
        return any(
            path.startswith(prefix)
            and (retained_path(path, state) or is_locally_overridden(path, state.local_overrides))
            for path in state.safe_files
        )
    normalized = dependency_path.lstrip("/")
    return normalized in state.present_files and (
        retained_path(normalized, state) or is_locally_overridden(normalized, state.local_overrides)
    )


def retained_path(relative_path: str, state: ReportState) -> bool:
    """Return whether a path is retained by the selected modules."""
    relation = selected_relation_for_path(relative_path, state.mappings)
    return relation is None or relation.is_retained_by(state.included_modules)


def dependabot_findings(repo_root: Path, state: ReportState) -> tuple[Finding, ...]:
    """Return findings for stale Dependabot ecosystems."""
    relative_path = ".github/dependabot.yml"
    if relative_path not in state.present_files or not retained_path(relative_path, state):
        return ()
    document = load_yaml_file(repo_root, relative_path)
    if document is None:
        return ()
    updates = document.get("updates")
    if not isinstance(updates, list):
        return ()
    updates = cast(list[object], updates)

    findings: list[Finding] = []
    for update in updates:
        if not isinstance(update, dict):
            continue
        update = cast(dict[object, object], update)
        ecosystem = update.get("package-ecosystem")
        if not isinstance(ecosystem, str):
            continue
        ecosystem_data = DEPENDABOT_ECOSYSTEM_MODULES.get(ecosystem)
        if ecosystem_data is None:
            continue
        module, scanned_paths = ecosystem_data
        if module not in state.excluded_modules:
            continue
        if any(dependency_file_is_retained_or_present(path, state) for path in scanned_paths):
            continue
        findings.append(
            Finding(
                rule_id="dependabot-ecosystem.stale",
                category=CATEGORY_REQUIRED,
                module=module,
                path=relative_path,
                detail=(
                    f"{ecosystem} scans {', '.join(scanned_paths)}, but no retained "
                    "scanned file or workflow surface is present."
                ),
            )
        )
    return tuple(findings)


def general_validation_reference_findings(
    repo_root: Path, state: ReportState
) -> tuple[Finding, ...]:
    """Return documented false-positive findings for general validation references."""
    findings: list[Finding] = []
    for module, paths_and_tokens in GENERAL_VALIDATION_REFERENCES.items():
        if module not in state.excluded_modules:
            continue
        for relative_path, tokens in paths_and_tokens.items():
            if relative_path not in state.present_files or not retained_path(relative_path, state):
                continue
            try:
                text = resolve_repo_path(repo_root, relative_path).read_text(encoding="utf-8")
            except OSError as error:
                summary = os_error_summary(error)
                raise ReportError(f"Unable to read {relative_path}: {summary}") from error
            for token in tokens:
                if token not in text:
                    continue
                findings.append(
                    Finding(
                        rule_id="validation-command.documented-reference",
                        category=CATEGORY_FALSE_POSITIVE,
                        module=module,
                        path=relative_path,
                        detail=(
                            f"{token} remains in a retained general validation surface; "
                            "review only if no retained files use that generic validator."
                        ),
                    )
                )
    return tuple(findings)


def marker_decision_findings(state: ReportState) -> tuple[Finding, ...]:
    """Return findings for local overrides and deferred protected decisions."""
    findings: list[Finding] = []
    for local_override in state.local_overrides:
        path = f"{local_override.path}/" if local_override.is_directory else local_override.path
        for module in modules_for_marker_path(local_override.path, state) or ("unknown",):
            findings.append(
                Finding(
                    rule_id="marker.local-override",
                    category=CATEGORY_LOCAL_OVERRIDE,
                    module=module,
                    path=path,
                    detail=(f"{local_override.default_decision}; reason: {local_override.reason}."),
                )
            )
    for candidate in state.deferred_candidates:
        for module in modules_for_marker_path(candidate.path, state) or ("unknown",):
            findings.append(
                Finding(
                    rule_id="marker.deferred-protected-decision",
                    category=CATEGORY_DEFERRED,
                    module=module,
                    path=candidate.path,
                    detail=f"{candidate.source_commit}: {candidate.reason}.",
                )
            )
    for decision in state.protected_decisions:
        for module in modules_for_marker_path(decision.path, state) or ("unknown",):
            details = [decision.decision]
            if decision.reason:
                details.append(f"reason: {decision.reason}")
            findings.append(
                Finding(
                    rule_id="marker.protected-file-decision",
                    category=CATEGORY_PROTECTED,
                    module=module,
                    path=decision.path,
                    detail="; ".join(details) + ".",
                )
            )
    return tuple(findings)


def modules_for_marker_path(raw_path: str, state: ReportState) -> tuple[str, ...]:
    """Return excluded modules related to a marker-scoped path."""
    normalized_path = raw_path.strip("/")
    related_modules: set[str] = set()
    relation = selected_relation_for_path(normalized_path, state.mappings)
    if relation is not None:
        related_modules.update(relation_modules(relation) & state.excluded_modules)
    for mapping in state.mappings:
        if marker_path_matches_mapping(normalized_path, mapping):
            related_modules.update(
                (mapping.requires_all | mapping.requires_any) & state.excluded_modules
            )
    return tuple(sorted(related_modules))


def marker_path_matches_mapping(path: str, mapping: ManifestMapping) -> bool:
    """Return whether a marker path or directory overlaps a manifest mapping."""
    if manifest_pattern_matches_path(mapping.pattern, path):
        return True
    prefix = f"{path}/"
    return mapping.pattern.startswith(prefix) or fnmatch.fnmatchcase(
        f"{prefix}placeholder", mapping.pattern
    )


def build_report(repo_root: Path, state: ReportState) -> ExcludedModuleReport:
    """Build the complete excluded-module cleanup report."""
    findings = [
        *manifest_owned_path_findings(state),
        *inline_block_findings(repo_root, state),
        *reference_link_findings(repo_root, state),
        *protected_document_prose_reference_findings(repo_root, state),
        *collaboration_template_reference_findings(repo_root, state),
        *dependabot_findings(repo_root, state),
        *general_validation_reference_findings(repo_root, state),
        *marker_decision_findings(state),
    ]
    unique_findings = tuple(
        sorted(
            {finding.render(): finding for finding in findings}.values(),
            key=lambda finding: finding.sort_key(),
        )
    )
    return ExcludedModuleReport(
        state=state,
        scopes=build_scope_lines(repo_root, state),
        findings=unique_findings,
    )


def print_named_section(name: str, items: Iterable[str], *, empty_text: str = "None.") -> None:
    """Print a named report section."""
    print(f"\n{name}:")
    rendered_items = tuple(items)
    if not rendered_items:
        print(f"  {empty_text}")
        return
    for item in rendered_items:
        print(f"  - {item}")


def print_report(report: ExcludedModuleReport) -> None:
    """Print a human-readable excluded-module cleanup report."""
    print("Excluded-module cleanup report.")
    print(f"State source: {report.state.source}")
    print("Exit model: findings are informational; runtime or input failures exit nonzero.")
    print_named_section("Retained modules", sorted(report.state.included_modules))
    print_named_section("Excluded modules", sorted(report.state.excluded_modules))
    print_named_section("Excluded module scopes", report.scopes)

    category_counts = Counter(finding.category for finding in report.findings)
    category_lines = [f"{category}: {count}" for category, count in sorted(category_counts.items())]
    print_named_section("Finding categories", category_lines)
    print_named_section(
        "Findings",
        (finding.render() for finding in report.findings),
        empty_text="No excluded-module cleanup findings.",
    )


def fail(message: str) -> NoReturn:
    """Print an error and exit non-zero."""
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(1)


def main(argv: list[str] | None = None) -> int:
    """Run excluded-module cleanup reporting."""
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        repo_root = resolve_repo_root(args.repo_root)
        marker_path = resolve_repo_path(repo_root, args.marker)
        marker_schema_path = resolve_repo_path(repo_root, args.marker_schema)
        manifest_path = resolve_repo_path(repo_root, args.manifest)
        manifest_schema_path = resolve_repo_path(repo_root, args.manifest_schema)
        state = build_state(
            repo_root=repo_root,
            marker_path=marker_path,
            marker_schema_path=marker_schema_path,
            manifest_path=manifest_path,
            manifest_schema_path=manifest_schema_path,
            explicit_included_modules=args.included_module,
        )
        report = build_report(repo_root, state)
    except (ReportError, TemplateSyncMaterializationError) as error:
        fail(str(error))

    print_report(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
