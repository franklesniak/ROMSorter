"""Generate a marker-aware template sync candidate table."""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from collections.abc import Callable
from dataclasses import dataclass, replace
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, NoReturn, cast
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode, urlsplit, urlunsplit
from urllib.request import Request, urlopen

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from first_adoption_state import (  # noqa: E402
    DEFAULT_ADOPTION_JOURNAL_PATH,
    FirstAdoptionState,
    format_first_adoption_state,
    inspect_first_adoption_state,
)
from template_sync_materialization_helpers import (  # noqa: E402
    AGENT_INSTRUCTION_EXACT_PATHS,
    DEFAULT_MANIFEST_PATH,
    DEFAULT_MANIFEST_SCHEMA_PATH,
    DEFAULT_MARKER_PATH,
    DEFAULT_MARKER_SCHEMA_PATH,
    REMOVAL_DECISION,
    DeferredProtectedCandidate,
    LocalOverride,
    LocalPathOwnership,
    ManifestMapping,
    PathRelation,
    ProtectedFileDecision,
    TemplateSyncMaterializationError,
    deferred_candidate_summary,
    directory_prefix_relation,
    has_wildcard,
    is_protected_instruction_path,
    is_protected_manifest_pattern,
    load_json_mapping,
    load_yaml_mapping,
    local_override_summary,
    manifest_pattern_matches_path,
    parse_manifest_mappings,
    parse_marker_decision_data,
    path_has_symlink_component,
    protected_decision_summary,
    repository_path_exists,
    repository_relative_path,
    resolve_repo_path,
    resolve_repo_root,
    selected_relation_for_path,
    validate_protected_file_decisions,
    validate_schema,
)

DEFAULT_RANGE_HEAD_REF = "template/main"
DEFAULT_TODO_PATH = "_TODO-repo-init.md"
TEMPLATE_UPDATE_PROCEDURE_PATH = "TEMPLATE_UPDATE_PROCEDURE.md"
GITHUB_METADATA_MANUAL_SETTINGS = (
    "private vulnerability reporting",
    "repository rulesets or classic branch protection",
    "required status checks",
    "bypass roles",
)
DEFAULT_GITHUB_API_BASE = "https://api.github.com"
GITHUB_REST_ACCEPT_HEADER = "application/vnd.github+json"
GITHUB_REST_API_VERSION = "2022-11-28"
GITHUB_REST_TIMEOUT_SECONDS = 5.0
GITHUB_REST_USER_AGENT = "copilot-repo-template-preflight/1.0"
GITHUB_REST_MAX_PAGES = 20
GITHUB_METADATA_SETTING_ORDER = (
    "repository",
    "visibility",
    "default_branch",
    "discussions",
    "labels",
    "rulesets",
    "branch_protection",
    "private_vulnerability_reporting",
    "security_and_analysis",
)
GITHUB_METADATA_SETTING_LABELS = {
    "repository": "Repository owner/name",
    "visibility": "Visibility",
    "default_branch": "Default branch",
    "discussions": "Discussions",
    "labels": "Labels",
    "rulesets": "Repository rulesets",
    "branch_protection": "Default-branch protection",
    "private_vulnerability_reporting": "Private vulnerability reporting",
    "security_and_analysis": "Security and analysis",
}
GITHUB_OBSERVATION_BASES = (
    "gh",
    "rest",
    "owner-confirmed",
    "auth-required",
    "permission-required",
    "rate-limited",
    "not-found-or-not-visible",
    "unsupported-endpoint",
    "dependency-unobserved",
    "transport-error",
    "malformed-response",
    "manual-review",
)
GITHUB_OBSERVATION_PAGINATION = (
    "not-applicable",
    "complete",
    "partial",
    "unavailable",
)
GITHUB_SECURITY_AND_ANALYSIS_KEYS = frozenset(
    {
        "advanced_security",
        "code_security",
        "dependabot_security_updates",
        "secret_scanning",
        "secret_scanning_ai_detection",
        "secret_scanning_non_provider_patterns",
        "secret_scanning_push_protection",
        "secret_scanning_validity_checks",
    }
)
DISCOVERY_SKIP_DIRS = frozenset(
    {
        ".git",
        ".mypy_cache",
        ".pytest_cache",
        ".ruff_cache",
        ".terraform",
        ".venv",
        "__pycache__",
        "build",
        "dist",
        "node_modules",
        "venv",
    }
)
ADOPTION_MODE_MODULES = frozenset(
    {
        "agent-instructions",
        "agent-copilot",
        "agent-codex",
        "agent-claude",
        "agent-cursor",
        "agent-gemini",
        "agent-hermes",
        "instruction-enforcement",
        "azure-devops-collaboration",
        "azure-devops-platform",
        "azure-pipelines",
        "baseline",
        "github-actions",
        "github-platform",
        "github-templates",
        "template-onboarding",
        "template-sync-support",
    }
)
INLINE_TRIM_NOTE_KEYWORDS = (
    "inline block",
    "inline blocks",
    "delimited",
    "strip",
)
TODO_KNOWN_SECTIONS = (
    "Discoverable Repository State",
    "Manual GitHub Settings",
    "Maintainer Policy Decisions",
    "Protected-File Adoption Decisions",
    "Unresolved Settings",
    "Resolution Evidence",
)
TODO_DECISION_SECTIONS = frozenset(
    {
        "Manual GitHub Settings",
        "Maintainer Policy Decisions",
        "Protected-File Adoption Decisions",
        "Unresolved Settings",
    }
)
TODO_LINK_LIMIT = 3
VALIDATION_COMMANDS_BY_MODULE: dict[str, tuple[str, ...]] = {
    "agent-instructions": ("manual protected-file authorization review",),
    "agent-copilot": ("manual Copilot instructions and retained setup-workflow review",),
    "agent-codex": ("manual Codex entry-point and required protocol review",),
    "agent-claude": ("manual Claude entry-point, imports, and required protocol review",),
    "agent-cursor": ("manual Cursor rule applicability and required protocol review",),
    "agent-gemini": ("manual Gemini entry-point and required protocol review",),
    "agent-hermes": ("manual Hermes entry-point and required protocol review",),
    "instruction-enforcement": (
        "python .github/scripts/validate_instruction_profile.py",
        "review explicit profile mode, selected modules, and local exception declarations",
    ),
    "azure-devops-collaboration": ("manual Azure DevOps collaboration template and policy review",),
    "azure-devops-platform": (
        "manual Azure DevOps Services project, security product, dependency scanning, and dependency update policy review",
    ),
    "azure-pipelines": ("manual Azure Pipelines validation or pipeline run review",),
    "baseline": (
        "pre-commit run --all-files",
        "placeholder and repository-identity review",
    ),
    "github-actions": (
        "pre-commit run actionlint --all-files",
        "python .github/scripts/validate_workflow_security.py",
    ),
    "github-platform": (
        "pre-commit run validate-dependabot-config --all-files",
        "pre-commit run validate-dependabot-config-valid-examples --all-files",
        "pytest tests/test_dependabot_schema.py -v",
    ),
    "github-templates": ("manual GitHub template rendering review",),
    "git-lfs": ("pytest tests/test_gitattributes.py -v",),
    "json": ("pre-commit run check-json --all-files",),
    "markdown": (
        "npm run lint:md",
        "npm run lint:md:links",
        "npm run lint:md:nested",
    ),
    "powershell": ("Invoke-Pester -Path tests/ -Output Detailed",),
    "python": (
        'pytest tests/ -m "not slow" -v --cov --cov-report=term-missing',
        "pytest tests/ -m slow -v --no-cov",
        "python -m pyright --project pyrightconfig.json",
        "pre-commit run check-toml --all-files",
    ),
    "schema": (
        "pre-commit run validate-example-config-valid-examples --all-files",
        "pre-commit run validate-example-config-schema --all-files",
        "pytest tests/test_schema_examples.py -v",
    ),
    "template-onboarding": (
        "npm run lint:md",
        "npm run lint:md:links",
        "npm run lint:md:nested",
    ),
    "template-sync-support": (
        "python .template-sync/scripts/validate_marker.py --require-marker",
        "python .template-sync/scripts/validate_instruction_contracts.py --mode downstream --require-marker",
        "pre-commit run validate-template-sync-marker-valid-examples --all-files",
        "pre-commit run validate-template-sync-instruction-contracts-valid-examples --all-files",
        "pre-commit run validate-template-sync-manifest-schema --all-files",
        "pre-commit run validate-template-sync-marker-schema --all-files",
        "pre-commit run validate-template-sync-instruction-contracts-schema --all-files",
        "pre-commit run validate-template-sync-manifest --all-files",
        "pre-commit run validate-template-sync-marker --all-files",
        "pre-commit run validate-template-sync-instruction-contracts --all-files",
        "pre-commit run validate-instruction-contracts-upstream --all-files",
        "pre-commit run validate-instruction-contracts-downstream --all-files",
        "pytest tests/test_schema_examples.py -v",
    ),
    "terraform": (
        "terraform fmt -check -recursive -diff",
        "tflint --init",
        'tflint --recursive --config "$(pwd)/.tflint.hcl"',
        "terraform test -verbose",
        "pytest tests/test_terraform_hooks.py -v",
    ),
    "yaml": (
        "pre-commit run check-yaml --all-files",
        "pre-commit run yamllint --all-files",
    ),
}


class CandidateGenerationError(Exception):
    """Raised when a sync candidate table cannot be generated."""


@dataclass(frozen=True)
class MarkerData:
    """Marker values needed to evaluate changed upstream paths."""

    last_reviewed_template_commit: str | None
    included_modules: frozenset[str]
    local_overrides: tuple[LocalOverride, ...]
    local_path_ownership: tuple[LocalPathOwnership, ...]
    deferred_candidates: tuple[DeferredProtectedCandidate, ...]
    protected_decisions: tuple[ProtectedFileDecision, ...]


@dataclass(frozen=True)
class DiffEntry:
    """One path-level change from ``git diff --name-status``."""

    status: str
    path: str
    old_path: str | None = None

    @property
    def display_path(self) -> str:
        """Return the path text shown in the candidate table."""
        if self.old_path is None:
            return self.path
        return f"{self.old_path} -> {self.path}"

    @property
    def change_kind(self) -> str:
        """Return a human-readable change kind."""
        status_code = self.status[0]
        if status_code == "A":
            return "Added"
        if status_code == "C":
            return f"Copied ({self.status})"
        if status_code == "D":
            return "Deleted"
        if status_code == "M":
            return "Modified"
        if status_code == "R":
            return f"Renamed ({self.status})"
        if status_code == "T":
            return "Type changed"
        return self.status


@dataclass(frozen=True)
class CandidateRow:
    """A rendered decision-aid row for one changed path."""

    path: str
    change: str
    module_relation: str
    retained_status: str
    local_override_status: str
    deferred_status: str
    protected_decision_status: str
    protected_status: str
    notes: tuple[str, ...]


@dataclass(frozen=True)
class TodoItem:
    """One first-adoption checklist item that can be linked from the ledger."""

    line_number: int
    text: str
    is_complete: bool


@dataclass(frozen=True)
class SummaryTodoItem:
    """One machine-interpretable unchecked checklist item for the concise summary."""

    line_number: int
    section: str
    text: str


@dataclass(frozen=True)
class SummaryTodoState:
    """Machine-interpretability state for the first-adoption checklist summary."""

    exists: bool
    is_interpretable: bool
    unchecked_items: tuple[SummaryTodoItem, ...]


@dataclass(frozen=True)
class LedgerRow:
    """A rendered adoption-ledger row for one manifest or manual setup item."""

    path: str
    manifest_modules: str
    decision: str
    reason: str
    protected_file: str
    requires_maintainer_decision: str
    adoption_mode: str
    todo_link: str
    validation_commands: str


@dataclass(frozen=True)
class RemoteInfo:
    """One configured Git remote discovered from local repository state."""

    name: str
    url: str
    purpose: str


@dataclass(frozen=True)
class RestRequest:
    """One read-only REST request issued by the metadata preflight probe."""

    method: str
    url: str
    headers: tuple[tuple[str, str], ...]
    timeout: float


@dataclass(frozen=True)
class RestResult:
    """A small HTTP response envelope for injected REST clients."""

    status_code: int | None
    body: str
    error: str | None = None
    headers: tuple[tuple[str, str], ...] = ()


@dataclass(frozen=True)
class RestJsonRead:
    """Parsed REST JSON plus sanitized response evidence."""

    payload: object | None
    basis: str
    diagnostics: tuple[tuple[str, str], ...]
    headers: tuple[tuple[str, str], ...] = ()


JsonValue = bool | int | float | str | None | list["JsonValue"] | dict[str, "JsonValue"]


@dataclass(frozen=True)
class GitHubObservation:
    """One typed GitHub setting observation with common provenance metadata."""

    setting: str
    value: JsonValue
    observed: bool
    basis: str
    source: str
    observed_at: str
    pagination: str = "not-applicable"
    partial: bool = False
    partial_fields: tuple[str, ...] = ()
    diagnostics: tuple[tuple[str, str], ...] = ()


@dataclass(frozen=True)
class GitHubMetadata:
    """Optional GitHub metadata represented as typed setting observations."""

    requested: bool
    observations: tuple[GitHubObservation, ...] = ()

    def observation(self, setting: str) -> GitHubObservation | None:
        """Return one observation by setting key."""
        for observation in self.observations:
            if observation.setting == setting:
                return observation
        return None

    @property
    def available(self) -> bool:
        """Return whether at least one GitHub setting was directly observed."""
        return any(observation.observed for observation in self.observations)

    @property
    def source(self) -> str:
        """Return the primary observation source for compatibility callers."""
        if not self.requested:
            return "not requested"
        for source in ("gh", "rest"):
            if any(
                observation.observed and observation.basis == source
                for observation in self.observations
            ):
                return source
        # Nothing was directly observed via ``gh`` or ``rest``. Preserve the legacy
        # source vocabulary ("not requested" / "manual" / "gh" / "rest") instead of
        # leaking a per-observation basis token (e.g. "auth-required") that callers
        # would misread as a discovery source. The specific basis remains visible
        # per setting in the ledger and in ``error``.
        return "manual"

    @property
    def repository(self) -> str | None:
        """Return the observed repository owner/name, when available."""
        return github_observation_string_value(self.observation("repository"))

    @property
    def visibility(self) -> str | None:
        """Return the observed repository visibility, when available."""
        return github_observation_string_value(self.observation("visibility"))

    @property
    def default_branch(self) -> str | None:
        """Return the observed default branch, when available."""
        return github_observation_string_value(self.observation("default_branch"))

    @property
    def discussions_enabled(self) -> str | None:
        """Return human-readable Discussions state for compatibility callers."""
        value = github_observation_boolean_value(self.observation("discussions"))
        if value is None:
            return None
        return "enabled" if value else "disabled"

    @property
    def labels(self) -> tuple[str, ...]:
        """Return observed label names, when available."""
        value = github_observation_list_value(self.observation("labels"))
        if value is None:
            return ()
        labels = [label for label in value if isinstance(label, str)]
        return tuple(labels)

    @property
    def error(self) -> str | None:
        """Return combined observation diagnostics for compatibility callers."""
        diagnostics: list[str] = []
        for observation in self.observations:
            detail = github_observation_diagnostic_text(observation)
            if detail:
                diagnostics.append(detail)
        return combine_metadata_errors(*diagnostics)

    @property
    def repository_source(self) -> str:
        """Return legacy source text for the repository observation."""
        return github_legacy_observation_source(self.observation("repository"))

    @property
    def visibility_source(self) -> str:
        """Return legacy source text for the visibility observation."""
        return github_legacy_observation_source(self.observation("visibility"))

    @property
    def default_branch_source(self) -> str:
        """Return legacy source text for the default branch observation."""
        return github_legacy_observation_source(self.observation("default_branch"))

    @property
    def discussions_source(self) -> str:
        """Return legacy source text for the Discussions observation."""
        return github_legacy_observation_source(self.observation("discussions"))

    @property
    def labels_source(self) -> str:
        """Return legacy source text for the labels observation."""
        return github_legacy_observation_source(self.observation("labels"))


@dataclass(frozen=True)
class RepositoryDiscovery:
    """Read-only local repository state for the first-adoption preflight report."""

    owner_name: str | None
    remotes: tuple[RemoteInfo, ...]
    current_branch: str | None
    likely_default_branch: str | None
    working_tree_entries: tuple[str, ...]
    workflows: tuple[str, ...]
    issue_templates: tuple[str, ...]
    pr_templates: tuple[str, ...]
    security_files: tuple[str, ...]
    agent_instruction_files: tuple[str, ...]
    tooling_stacks: tuple[str, ...]
    marker_present: bool
    adoption_note_present: bool
    github_metadata: GitHubMetadata


CommandRunner = Callable[[Path, list[str]], subprocess.CompletedProcess[str]]
RestClient = Callable[[RestRequest], RestResult]
Clock = Callable[[], datetime]


JsonObject = dict[str, object]


def json_object(value: object) -> JsonObject | None:
    """Return ``value`` as a JSON object after runtime validation."""
    return cast(JsonObject, value) if isinstance(value, dict) else None


def json_array(value: object) -> list[object] | None:
    """Return ``value`` as a JSON array after runtime validation."""
    return cast(list[object], value) if isinstance(value, list) else None


def json_string_field(mapping: JsonObject, field_name: str) -> str | None:
    """Return one optional string field from a JSON object."""
    value = mapping.get(field_name)
    return value if isinstance(value, str) else None


def utc_now() -> datetime:
    """Return the current UTC time for observation timestamps."""
    return datetime.now(UTC)


def observed_at_timestamp(clock: Clock) -> str:
    """Return a full RFC 3339 UTC timestamp with a ``Z`` suffix."""
    value = clock()
    if value.tzinfo is None:
        value = value.replace(tzinfo=UTC)
    return value.astimezone(UTC).isoformat(timespec="seconds").replace("+00:00", "Z")


def bounded_diagnostic_text(value: str, *, limit: int = 160) -> str:
    """Return a single-line, redacted diagnostic bounded for ledger evidence."""
    text = redact_url_userinfo(value).replace("\r", " ").replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)
    if len(text) <= limit:
        return text
    return f"{text[: limit - 3]}..."


def sanitize_accepted_permissions(value: str) -> str:
    """Return an allowlisted permissions hint from ``X-Accepted-GitHub-Permissions``."""
    sanitized = re.sub(r"[^A-Za-z0-9_.,;:=/ -]", "", value)
    return bounded_diagnostic_text(sanitized, limit=200)


def observation_diagnostic_pairs(
    *pairs: tuple[str, str | None],
) -> tuple[tuple[str, str], ...]:
    """Return sanitized non-empty observation diagnostic pairs."""
    diagnostics: list[tuple[str, str]] = []
    for key, value in pairs:
        if value is None:
            continue
        sanitized = bounded_diagnostic_text(str(value))
        if sanitized:
            diagnostics.append((key, sanitized))
    return tuple(diagnostics)


def validate_observation_envelope(basis: str, pagination: str) -> None:
    """Validate closed-token observation envelope fields."""
    if basis not in GITHUB_OBSERVATION_BASES:
        raise ValueError(f"Unknown GitHub observation basis: {basis!r}")
    if pagination not in GITHUB_OBSERVATION_PAGINATION:
        raise ValueError(f"Unknown GitHub observation pagination state: {pagination!r}")


def github_observation(
    *,
    setting: str,
    value: JsonValue,
    observed: bool,
    basis: str,
    source: str,
    observed_at: str,
    pagination: str = "not-applicable",
    partial: bool = False,
    partial_fields: tuple[str, ...] = (),
    diagnostics: tuple[tuple[str, str], ...] = (),
) -> GitHubObservation:
    """Build one validated GitHub observation."""
    validate_observation_envelope(basis, pagination)
    return GitHubObservation(
        setting=setting,
        value=value,
        observed=observed,
        basis=basis,
        source=source,
        observed_at=observed_at,
        pagination=pagination,
        partial=partial,
        partial_fields=partial_fields,
        diagnostics=diagnostics,
    )


def unobserved_github_observation(
    *,
    setting: str,
    basis: str,
    source: str,
    observed_at: str,
    pagination: str = "not-applicable",
    diagnostics: tuple[tuple[str, str], ...] = (),
) -> GitHubObservation:
    """Build one unavailable setting observation with a null value."""
    return github_observation(
        setting=setting,
        value=None,
        observed=False,
        basis=basis,
        source=source,
        observed_at=observed_at,
        pagination=pagination,
        diagnostics=diagnostics,
    )


def unobserved_github_observations(
    *,
    basis: str,
    source: str,
    observed_at: str,
    diagnostics: tuple[tuple[str, str], ...] = (),
) -> tuple[GitHubObservation, ...]:
    """Build unavailable observations for every supported GitHub setting."""
    return tuple(
        unobserved_github_observation(
            setting=setting,
            basis=basis,
            source=source,
            observed_at=observed_at,
            pagination=("unavailable" if setting in {"labels", "rulesets"} else "not-applicable"),
            diagnostics=diagnostics,
        )
        for setting in GITHUB_METADATA_SETTING_ORDER
    )


def github_observation_string_value(observation: GitHubObservation | None) -> str | None:
    """Return an observed string value from one observation."""
    if observation is None or not observation.observed or not isinstance(observation.value, str):
        return None
    return observation.value


def github_observation_boolean_value(observation: GitHubObservation | None) -> bool | None:
    """Return an observed boolean value from one observation."""
    if observation is None or not observation.observed or not isinstance(observation.value, bool):
        return None
    return observation.value


def github_observation_list_value(
    observation: GitHubObservation | None,
) -> list[JsonValue] | None:
    """Return an observed list value from one observation."""
    if observation is None or not observation.observed or not isinstance(observation.value, list):
        return None
    return observation.value


def github_legacy_observation_source(observation: GitHubObservation | None) -> str:
    """Return old source vocabulary for compatibility properties."""
    if observation is None or not observation.observed:
        return "manual"
    return observation.basis


def github_observation_diagnostic_text(observation: GitHubObservation) -> str | None:
    """Return compact diagnostics for a metadata observation."""
    details: list[str] = []
    if not observation.observed:
        details.append(
            f"{GITHUB_METADATA_SETTING_LABELS.get(observation.setting, observation.setting)} "
            f"unobserved ({observation.basis})"
        )
    elif observation.partial:
        details.append(
            f"{GITHUB_METADATA_SETTING_LABELS.get(observation.setting, observation.setting)} "
            "partially observed"
        )
    details.extend(f"{key}: {value}" for key, value in observation.diagnostics)
    if observation.partial_fields:
        fields = ", ".join(f"`{field}`" for field in observation.partial_fields)
        details.append(f"partial fields: {fields}")
    if not details:
        return None
    return "; ".join(details)


def github_primary_diagnostic(metadata: GitHubMetadata) -> str | None:
    """Return the first sanitized metadata diagnostic message."""
    for observation in metadata.observations:
        for key, value in observation.diagnostics:
            if key == "message":
                return value
    return metadata.error


def github_label_names(value: object) -> tuple[str, ...] | None:
    """Return sorted label names from a GitHub label JSON array."""
    label_records = json_array(value)
    if label_records is None:
        return None

    names: list[str] = []
    for raw_label in label_records:
        label = json_object(raw_label)
        if label is None:
            continue
        name = json_string_field(label, "name")
        if name is not None:
            names.append(name)
    return tuple(sorted(names))


def parse_args(argv: list[str]) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Generate a Markdown template sync candidate table from the marker, "
            "manifest, and upstream git range."
        )
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
        "--range-base",
        default=None,
        help=(
            "Base commit or ref for the upstream comparison. Defaults to "
            "template_sync.last_reviewed_template_commit from the marker."
        ),
    )
    parser.add_argument(
        "--range-head",
        default=None,
        help=(
            "Head commit or ref for the upstream comparison. Defaults to the local "
            "template/main ref when it is present. The helper does not fetch."
        ),
    )
    parser.add_argument(
        "--write-candidates",
        default=None,
        metavar="PATH",
        help=(
            "Write the rendered candidate table to PATH while still printing the full "
            "report to stdout. PATH must stay inside the repository root."
        ),
    )
    ledger_mode_group = parser.add_mutually_exclusive_group()
    ledger_mode_group.add_argument(
        "--ledger",
        action="store_true",
        help=(
            "Include a generated adoption ledger in stdout after the candidate table. "
            "The ledger is a review artifact only."
        ),
    )
    ledger_mode_group.add_argument(
        "--ledger-only",
        action="store_true",
        help=(
            "Emit only the generated adoption ledger and skip git range inspection. "
            "Use this for first-adoption or full-reconciliation review when a delta "
            "candidate table is not available."
        ),
    )
    ledger_mode_group.add_argument(
        "--summary",
        action="store_true",
        help=(
            "Emit a concise deterministic template adoption summary for PR "
            "descriptions and skip git range inspection."
        ),
    )
    ledger_mode_group.add_argument(
        "--preflight",
        "--questionnaire",
        action="store_true",
        dest="preflight",
        help=(
            "Emit a read-only first-adoption preflight report, maintainer "
            "questionnaire, _TODO-repo-init.md starter, and reused adoption ledger. "
            "This mode skips git range inspection and never writes files."
        ),
    )
    parser.add_argument(
        "--write-ledger",
        default=None,
        metavar="PATH",
        help=(
            "Write the generated adoption ledger snapshot to PATH. PATH must stay "
            "inside the repository root."
        ),
    )
    parser.add_argument(
        "--todo-file",
        default=DEFAULT_TODO_PATH,
        help=(
            "First-adoption checklist path relative to the repository root. "
            f"Default: {DEFAULT_TODO_PATH}"
        ),
    )
    parser.add_argument(
        "--adoption-mode",
        choices=("minimal-preservation", "tailored"),
        default="minimal-preservation",
        help=(
            "Default adoption mode to show for protected and template-derived files "
            "in the generated ledger. Default: minimal-preservation"
        ),
    )
    parser.add_argument(
        "--include-github-metadata",
        action="store_true",
        help=(
            "Opt in to read-only GitHub metadata lookup through the gh CLI, falling "
            "back to unauthenticated public REST reads when gh is unavailable or "
            "unusable. Without this flag, GitHub-only settings are labeled manual "
            "review required."
        ),
    )
    parser.add_argument(
        "--github-api-base",
        default=None,
        metavar="URL",
        help=(
            "GitHub REST API base URL for --include-github-metadata, such as "
            "https://github.example.com/api/v3 for GHES. Defaults to "
            f"{DEFAULT_GITHUB_API_BASE}."
        ),
    )
    parser.add_argument(
        "--full-state",
        action="store_true",
        help=(
            "With --preflight, list every raw first-adoption state entry instead of "
            "bounded deterministic samples."
        ),
    )
    return parser.parse_args(argv)


def run_git(
    repo_root: Path, args: list[str], *, check: bool = True
) -> subprocess.CompletedProcess[str]:
    """Run a read-only Git command in ``repo_root``."""
    result = subprocess.run(
        ["git", "-C", str(repo_root), *args],
        check=False,
        capture_output=True,
        text=True,
    )
    if check and result.returncode != 0:
        detail = result.stderr.strip() or result.stdout.strip() or f"exit {result.returncode}"
        raise CandidateGenerationError(f"git {' '.join(args)} failed: {detail}")
    return result


def run_command(repo_root: Path, command: list[str]) -> subprocess.CompletedProcess[str]:
    """Run a read-only discovery command in ``repo_root``."""
    return subprocess.run(
        command,
        cwd=repo_root,
        check=False,
        capture_output=True,
        text=True,
    )


def run_rest_request(request: RestRequest) -> RestResult:
    """Run one bounded read-only REST request through the standard library.

    The request method is restricted to ``GET`` so this helper cannot mutate
    remote state; any other method raises ``ValueError`` before a network call
    is made, enforcing the read-only contract at the request boundary.
    """
    if request.method != "GET":
        raise ValueError(
            f"run_rest_request only issues read-only GET requests, not {request.method!r}."
        )
    headers = dict(request.headers)
    urllib_request = Request(
        request.url,
        headers=headers,
        method=request.method,
    )
    try:
        with urlopen(urllib_request, timeout=request.timeout) as response:
            body = response.read().decode("utf-8", errors="replace")
            return RestResult(
                status_code=response.status,
                body=body,
                headers=tuple(response.headers.items()),
            )
    except HTTPError as error:
        body = error.read().decode("utf-8", errors="replace")
        return RestResult(
            status_code=error.code,
            body=body,
            headers=tuple(error.headers.items()) if error.headers is not None else (),
        )
    except URLError as error:
        return RestResult(
            status_code=None,
            body="",
            error=urllib_error_summary(error),
        )
    except TimeoutError as error:
        return RestResult(
            status_code=None,
            body="",
            error=f"{type(error).__name__}: timed out",
        )
    except OSError as error:
        return RestResult(
            status_code=None,
            body="",
            error=os_error_summary(error),
        )


def command_detail(result: subprocess.CompletedProcess[str]) -> str:
    """Return a short diagnostic string from a command result."""
    for stream_text in (result.stderr, result.stdout):
        for line in stream_text.splitlines():
            if line.strip():
                return line.strip()
    return f"exit {result.returncode}"


def os_error_summary(error: OSError) -> str:
    """Return an OSError summary that avoids exposing implicit filesystem paths."""
    return f"{type(error).__name__}: {error.strerror or 'I/O error'}"


def urllib_error_summary(error: URLError) -> str:
    """Return a short urllib diagnostic without relying on raw OSError strings."""
    reason = error.reason
    if isinstance(reason, OSError):
        reason_text = f"{type(reason).__name__}: {reason.strerror or 'I/O error'}"
    else:
        reason_text = str(reason)
    return f"{type(error).__name__}: {reason_text}"


def git_ref_exists(repo_root: Path, raw_ref: str) -> bool:
    """Return whether ``raw_ref`` resolves to a local commit."""
    result = run_git(
        repo_root,
        ["rev-parse", "--verify", "--end-of-options", f"{raw_ref}^{{commit}}"],
        check=False,
    )
    return result.returncode == 0


def resolve_commit(repo_root: Path, raw_ref: str, label: str) -> str:
    """Resolve a commit-ish ref to a full commit SHA."""
    result = run_git(
        repo_root,
        ["rev-parse", "--verify", "--end-of-options", f"{raw_ref}^{{commit}}"],
        check=False,
    )
    if result.returncode != 0:
        detail = result.stderr.strip() or result.stdout.strip() or "not found"
        raise CandidateGenerationError(
            f"Unable to resolve {label} {raw_ref!r} to a commit: {detail}"
        )
    return result.stdout.strip().splitlines()[0]


def resolve_range_head_ref(repo_root: Path, raw_range_head: str | None) -> tuple[str, str]:
    """Resolve the requested range head or the local default ref."""
    if raw_range_head is not None:
        return raw_range_head, resolve_commit(repo_root, raw_range_head, "range head")

    if git_ref_exists(repo_root, DEFAULT_RANGE_HEAD_REF):
        return (
            DEFAULT_RANGE_HEAD_REF,
            resolve_commit(repo_root, DEFAULT_RANGE_HEAD_REF, "range head"),
        )

    raise CandidateGenerationError(
        "No --range-head was supplied and local template/main was not found. "
        "Fetch the template remote first or pass --range-head explicitly."
    )


def resolve_range_base_ref(
    repo_root: Path,
    raw_range_base: str | None,
    marker_data: MarkerData,
) -> tuple[str, str, str]:
    """Resolve the comparison base from CLI input or marker state."""
    if raw_range_base is not None:
        return (
            raw_range_base,
            resolve_commit(repo_root, raw_range_base, "range base"),
            "--range-base",
        )

    marker_base = marker_data.last_reviewed_template_commit
    if marker_base:
        return (
            marker_base,
            resolve_commit(
                repo_root,
                marker_base,
                "marker template_sync.last_reviewed_template_commit",
            ),
            "marker template_sync.last_reviewed_template_commit",
        )

    raise CandidateGenerationError(
        "No range base was provided and .template-sync/marker.yml does not set "
        "template_sync.last_reviewed_template_commit. Pass --range-base for a "
        "first-sync delta review; this helper will not guess a baseline."
    )


def verify_reachable_range(repo_root: Path, range_base_sha: str, range_head_sha: str) -> None:
    """Reject a range whose base is not an ancestor of its head."""
    result = run_git(
        repo_root,
        ["merge-base", "--is-ancestor", range_base_sha, range_head_sha],
        check=False,
    )
    if result.returncode == 0:
        return
    if result.returncode == 1:
        raise CandidateGenerationError(
            f"Range base {range_base_sha} is not an ancestor of range head {range_head_sha}."
        )
    detail = result.stderr.strip() or result.stdout.strip() or f"exit {result.returncode}"
    raise CandidateGenerationError(f"Unable to check range reachability: {detail}")


def parse_name_status_line(line: str) -> DiffEntry:
    """Parse one ``git diff --name-status`` line."""
    parts = line.split("\t")
    if len(parts) < 2:
        raise CandidateGenerationError(f"Unexpected git diff --name-status line: {line!r}")

    status = parts[0]
    status_code = status[0]
    if status_code in {"R", "C"}:
        if len(parts) != 3:
            raise CandidateGenerationError(f"Unexpected rename/copy diff line: {line!r}")
        return DiffEntry(status=status, old_path=parts[1], path=parts[2])

    return DiffEntry(status=status, path=parts[1])


def changed_paths(
    repo_root: Path, range_base_sha: str, range_head_sha: str
) -> tuple[DiffEntry, ...]:
    """Return upstream path changes for ``range_base_sha..range_head_sha``."""
    result = run_git(
        repo_root,
        [
            "diff",
            "--name-status",
            "-M",
            f"{range_base_sha}..{range_head_sha}",
            "--",
        ],
    )
    entries = [parse_name_status_line(line) for line in result.stdout.splitlines() if line.strip()]
    return tuple(entries)


def modeled_diff_command(range_base_sha: str, range_head_sha: str) -> str:
    """Return the copyable diff command modeled by candidate generation."""
    return f"git diff --name-status -M {range_base_sha}..{range_head_sha} --"


def stale_procedure_warning(repo_root: Path, range_head_sha: str) -> str | None:
    """Return a warning when the local procedure differs from range-head upstream."""
    upstream_result = run_git(
        repo_root,
        ["show", f"{range_head_sha}:{TEMPLATE_UPDATE_PROCEDURE_PATH}"],
        check=False,
    )
    if upstream_result.returncode != 0:
        return None

    local_path = repo_root / TEMPLATE_UPDATE_PROCEDURE_PATH
    show_command = f"git show {range_head_sha}:{TEMPLATE_UPDATE_PROCEDURE_PATH}"
    try:
        local_text = local_path.read_text(encoding="utf-8")
    except OSError as error:
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        return (
            f"WARNING: Unable to read local `{TEMPLATE_UPDATE_PROCEDURE_PATH}` "
            f"for comparison with range head `{range_head_sha}`: {error_summary}. "
            f"Review the current upstream procedure with `{show_command}` before "
            "following local procedure text."
        )

    if local_text == upstream_result.stdout:
        return None

    return (
        f"WARNING: Local `{TEMPLATE_UPDATE_PROCEDURE_PATH}` may be stale; it differs "
        f"from the upstream copy at range head `{range_head_sha}`. Review the current "
        f"upstream procedure with `{show_command}` before following local procedure text."
    )


def parse_marker(marker: dict[str, Any]) -> MarkerData:
    """Extract marker fields needed for candidate generation."""
    marker_data = parse_marker_decision_data(
        marker,
        validate_protected_decision_integrity=True,
    )

    return MarkerData(
        last_reviewed_template_commit=marker_data.last_reviewed_template_commit,
        included_modules=marker_data.included_modules,
        local_overrides=marker_data.local_overrides,
        local_path_ownership=marker_data.local_path_ownership,
        deferred_candidates=marker_data.deferred_candidates,
        protected_decisions=marker_data.protected_decisions,
    )


def parse_manifest(manifest: dict[str, Any]) -> tuple[frozenset[str], tuple[ManifestMapping, ...]]:
    """Extract manifest modules and mapping rows without making owner decisions."""
    return parse_manifest_mappings(manifest, reject_unknown_modules=False)


def describe_relation(relation: PathRelation | None) -> str:
    """Return a human-readable module relation description, or ``unmapped``."""
    return relation.description if relation is not None else "unmapped"


def relation_modules(relation: PathRelation | None) -> frozenset[str]:
    """Return all manifest modules for a relation."""
    if relation is None:
        return frozenset()
    return relation.requires_all | relation.requires_any


def relations_match(left: PathRelation | None, right: PathRelation | None) -> bool:
    """Return whether two relations resolve to the same module mapping."""
    if left is None or right is None:
        return left is None and right is None
    return left.requires_all == right.requires_all and left.requires_any == right.requires_any


def matching_local_overrides_for_pattern(
    pattern: str,
    local_overrides: tuple[LocalOverride, ...],
) -> tuple[LocalOverride, ...]:
    """Return marker local overrides that apply to a manifest pattern."""
    matches: list[LocalOverride] = []
    for local_override in local_overrides:
        if not has_wildcard(pattern) and local_override.matches(pattern):
            matches.append(local_override)
            continue
        if local_override.is_directory and pattern.startswith(f"{local_override.path}/"):
            matches.append(local_override)
            continue
        if local_override.path == pattern:
            matches.append(local_override)
    return tuple(matches)


def matching_deferred_candidates_for_pattern(
    pattern: str,
    deferred_candidates: tuple[DeferredProtectedCandidate, ...],
) -> tuple[DeferredProtectedCandidate, ...]:
    """Return deferred protected candidates that apply to a manifest pattern."""
    return tuple(
        candidate
        for candidate in deferred_candidates
        if manifest_pattern_matches_path(pattern, candidate.path)
    )


def matching_protected_decisions_for_pattern(
    pattern: str,
    protected_decisions: tuple[ProtectedFileDecision, ...],
) -> tuple[ProtectedFileDecision, ...]:
    """Return protected-file decisions that apply to a manifest pattern."""
    return tuple(
        protected_decision
        for protected_decision in protected_decisions
        if manifest_pattern_matches_path(pattern, protected_decision.path)
    )


def matching_local_overrides(
    entry: DiffEntry,
    local_overrides: tuple[LocalOverride, ...],
) -> tuple[LocalOverride, ...]:
    """Return marker local overrides that match the changed path."""
    paths = [entry.path]
    if entry.old_path is not None:
        paths.append(entry.old_path)
    matches: list[LocalOverride] = []
    for local_override in local_overrides:
        if local_override in matches:
            continue
        if any(local_override.matches(path) for path in paths):
            matches.append(local_override)
    return tuple(matches)


def matching_deferred_candidates(
    entry: DiffEntry,
    deferred_candidates: tuple[DeferredProtectedCandidate, ...],
) -> tuple[DeferredProtectedCandidate, ...]:
    """Return marker deferred protected candidates that match the changed path."""
    paths = {entry.path}
    if entry.old_path is not None:
        paths.add(entry.old_path)
    return tuple(candidate for candidate in deferred_candidates if candidate.path in paths)


def matching_protected_decisions(
    entry: DiffEntry,
    protected_decisions: tuple[ProtectedFileDecision, ...],
) -> tuple[ProtectedFileDecision, ...]:
    """Return marker protected-file decisions that match the changed path."""
    paths = {entry.path}
    if entry.old_path is not None:
        paths.add(entry.old_path)
    return tuple(decision for decision in protected_decisions if decision.path in paths)


def protected_decision_status(
    protected_decisions: tuple[ProtectedFileDecision, ...],
) -> str:
    """Return candidate-table status text for matching protected decisions."""
    if not protected_decisions:
        return "None"
    return "; ".join(protected_decision_summary(decision) for decision in protected_decisions)


def build_candidate_row(
    entry: DiffEntry,
    marker_data: MarkerData,
    mappings: tuple[ManifestMapping, ...],
) -> CandidateRow:
    """Build one Markdown candidate row from marker and manifest context."""
    relation = selected_relation_for_path(entry.path, mappings)
    notes: list[str] = []

    if relation is None:
        module_relation = "UNMAPPED"
        retained_status = "Unmapped"
        notes.append("Unmapped path; owner must assign or confirm a module before final review.")
    else:
        module_relation = relation.description
        retained_status = (
            "Retained" if relation.is_retained_by(marker_data.included_modules) else "Excluded"
        )
        if retained_status == "Excluded":
            notes.append("Excluded by marker included_modules.")
        if relation.unknown_modules:
            notes.append("Unknown module(s): " + ", ".join(sorted(relation.unknown_modules)) + ".")
        if relation.is_cross_module:
            notes.append("Cross-module manifest relation matched.")
        for manifest_note in relation.notes:
            notes.append(f"Manifest note: {manifest_note}")

    local_overrides = matching_local_overrides(entry, marker_data.local_overrides)
    if local_overrides:
        local_override_status = "; ".join(
            f"{override.default_decision}: {override.reason}" for override in local_overrides
        )
        notes.append("Local override present; use it as a default, not an automatic decision.")
    else:
        local_override_status = "None"

    deferred_candidates = matching_deferred_candidates(entry, marker_data.deferred_candidates)
    if deferred_candidates:
        deferred_status = "; ".join(
            f"{candidate.source_commit}: {candidate.reason}" for candidate in deferred_candidates
        )
    else:
        deferred_status = "None"

    protected_decisions = matching_protected_decisions(entry, marker_data.protected_decisions)
    protected_decision_text = protected_decision_status(protected_decisions)
    if protected_decisions:
        notes.append("Protected file decision record present in marker.")

    protected_paths = [entry.path]
    if entry.old_path is not None:
        protected_paths.append(entry.old_path)
        is_copy = entry.status.startswith("C")
        change_verb = "Copied" if is_copy else "Renamed"
        change_noun = "Copy" if is_copy else "Rename"
        notes.append(f"{change_verb} from {entry.old_path}.")
        old_relation = selected_relation_for_path(entry.old_path, mappings)
        if not relations_match(old_relation, relation):
            notes.append(
                f"{change_noun} crosses module mapping boundaries: old path resolves to "
                f"{describe_relation(old_relation)}; new path resolves to "
                f"{describe_relation(relation)}. Review both mappings before deciding."
            )
    protected_status = (
        "Yes" if any(is_protected_instruction_path(path) for path in protected_paths) else "No"
    )
    if protected_status == "Yes":
        notes.append(
            "Protected instruction/governance file; explicit owner authorization is required."
        )

    if entry.status.startswith("D"):
        notes.append("Upstream deletion; owner must decide whether to remove the local file.")

    return CandidateRow(
        path=entry.display_path,
        change=entry.change_kind,
        module_relation=module_relation,
        retained_status=retained_status,
        local_override_status=local_override_status,
        deferred_status=deferred_status,
        protected_decision_status=protected_decision_text,
        protected_status=protected_status,
        notes=tuple(notes),
    )


def format_manifest_modules(requires_all: frozenset[str], requires_any: frozenset[str]) -> str:
    """Return a compact module relation summary for the adoption ledger."""
    parts: list[str] = []
    if requires_all:
        parts.append("all: " + ", ".join(sorted(requires_all)))
    if requires_any:
        parts.append("any: " + ", ".join(sorted(requires_any)))
    return "; ".join(parts) if parts else "none"


def relation_module_names(mapping: ManifestMapping) -> frozenset[str]:
    """Return all manifest modules referenced by ``mapping``."""
    return mapping.requires_all | mapping.requires_any


def validation_commands_for_modules(modules: frozenset[str]) -> str:
    """Return validation commands affected by a set of manifest modules."""
    commands: list[str] = []
    for module in sorted(modules):
        for command in VALIDATION_COMMANDS_BY_MODULE.get(module, ()):
            if command not in commands:
                commands.append(command)
    return "<br>".join(commands) if commands else "manual review"


def local_override_reason(local_overrides: tuple[LocalOverride, ...]) -> str:
    """Return the ledger reason text for matching local overrides."""
    return "; ".join(
        f"Marker local override defaults to `{override.default_decision}`: {override.reason}"
        for override in local_overrides
    )


def protected_decision_reason(
    protected_decisions: tuple[ProtectedFileDecision, ...],
    local_overrides: tuple[LocalOverride, ...] = (),
) -> str:
    """Return ledger reason text for matching protected-file decisions."""
    reason_parts: list[str] = []
    for protected_decision in protected_decisions:
        if protected_decision.decision in {"TAKE", "MERGE"}:
            if protected_decision.adoption_mode == "tailored":
                reason_parts.append("Authorized tailored protected-file rewrite.")
            else:
                reason_parts.append("Authorized minimal-preservation protected-file edit.")
            reason_parts.append(f"authorization_basis: {protected_decision.authorization_basis}")
            reason_parts.append(f"authorized_scope: {protected_decision.authorized_scope}")
            if protected_decision.tailored_authorization_basis is not None:
                reason_parts.append(
                    "tailored_authorization_basis: "
                    f"{protected_decision.tailored_authorization_basis}"
                )
        elif protected_decision.decision == REMOVAL_DECISION:
            reason_parts.append(
                "Authorized protected-file removal; see REMOVE-LOCAL authorizations section."
            )
        else:
            reason = protected_decision.reason or "No reason recorded."
            reason_parts.append(f"Protected decision `{protected_decision.decision}`: {reason}")

    for local_override in local_overrides:
        reason_parts.append(f"Overlapping local override: {local_override_summary(local_override)}")

    return " ".join(reason_parts)


def protected_decision_requires_maintainer_decision(
    protected_decisions: tuple[ProtectedFileDecision, ...],
) -> bool:
    """Return whether matching protected decisions still require maintainer action."""
    return any(
        decision.decision in {"DEFER", "PROTECTED-REVIEW"} for decision in protected_decisions
    )


def protected_decision_adoption_mode(
    protected_decisions: tuple[ProtectedFileDecision, ...],
) -> str:
    """Return the adoption-mode ledger value for protected decisions."""
    modes: list[str] = []
    for protected_decision in protected_decisions:
        if protected_decision.adoption_mode == "tailored":
            modes.append("tailored (authorized protected-file decision)")
        elif protected_decision.adoption_mode == "minimal-preservation":
            modes.append("minimal-preservation (protected-file decision)")
        elif protected_decision.decision == REMOVAL_DECISION:
            modes.append("not applicable (authorized removal)")
        else:
            modes.append("not applicable")
    return "; ".join(dict.fromkeys(modes))


def has_trim_note(notes: tuple[str, ...]) -> bool:
    """Return whether manifest notes indicate module-scoped trim work."""
    return any(keyword in note.lower() for note in notes for keyword in INLINE_TRIM_NOTE_KEYWORDS)


def notes_mention_modules(notes: tuple[str, ...], modules: frozenset[str]) -> bool:
    """Return whether manifest notes mention any module in ``modules``."""
    normalized_notes = " ".join(notes).lower()
    return any(
        module in normalized_notes or module.replace("-", " ") in normalized_notes
        for module in modules
    )


def retention_gap_reason(mapping: ManifestMapping, included_modules: frozenset[str]) -> str:
    """Return why a manifest mapping is not retained by included modules."""
    missing_all = mapping.requires_all - included_modules
    if missing_all:
        return "Required module(s) not included: " + ", ".join(sorted(missing_all)) + "."
    if mapping.requires_any and not mapping.requires_any.intersection(included_modules):
        return (
            "None of the required alternative module(s) are included: "
            + ", ".join(sorted(mapping.requires_any))
            + "."
        )
    return "Excluded by marker included_modules."


def adoption_mode_for_modules(
    modules: frozenset[str],
    local_overrides: tuple[LocalOverride, ...],
    is_protected: bool,
    default_adoption_mode: str,
) -> str:
    """Return the adoption mode displayed for a ledger row."""
    for local_override in local_overrides:
        lowered_reason = local_override.reason.lower()
        if "tailored" in lowered_reason:
            return "tailored (marker local override)"
        if "minimal-preservation" in lowered_reason:
            return "minimal-preservation (marker local override)"

    if is_protected or modules.intersection(ADOPTION_MODE_MODULES):
        return default_adoption_mode
    return "not applicable"


def todo_link_target(
    todo_path: Path,
    repo_root: Path,
    ledger_output_dir: Path | None,
) -> str:
    """Return the link target string for the checklist file.

    When ``ledger_output_dir`` is provided, the target is computed relative to
    that directory so the link resolves correctly when the rendered ledger is
    committed under it. Otherwise the target is repo-root-relative, which is
    informative when reading the ledger in a terminal but is not a clickable
    link to the checklist file when the ledger is rendered from a PR or issue
    body (GitHub resolves relative Markdown link targets in PR/issue bodies
    relative to the PR/issue URL, not the repository root). Callers that need
    clickable rendered links should pass ``ledger_output_dir`` and commit the
    saved ledger so it is rendered under that directory on GitHub.
    """
    if ledger_output_dir is None:
        return repository_relative_path(todo_path, repo_root)
    return os.path.relpath(str(todo_path), str(ledger_output_dir)).replace(os.sep, "/")


def todo_item_link(
    todo_path: Path,
    repo_root: Path,
    item: TodoItem,
    ledger_output_dir: Path | None = None,
) -> str:
    """Return a Markdown link to one checklist line."""
    todo_relative = todo_link_target(todo_path, repo_root, ledger_output_dir)
    return f"[line {item.line_number}]({todo_relative}#L{item.line_number})"


def matching_todo_links(
    *,
    path: str,
    modules: frozenset[str],
    is_protected: bool,
    todo_items: tuple[TodoItem, ...],
    todo_path: Path,
    repo_root: Path,
    ledger_output_dir: Path | None = None,
) -> str:
    """Return links to first-adoption checklist rows relevant to a ledger row."""
    if not todo_items:
        return "None"

    normalized_path = path.rstrip("/")
    path_parts = [normalized_path.lower()]
    if "/" in normalized_path:
        path_parts.append(normalized_path.rsplit("/", maxsplit=1)[-1].lower())
    needles = set(path_parts)
    needles.update(module.lower() for module in modules)
    if is_protected:
        needles.update({"protected", "adoption mode"})

    links: list[str] = []
    for item in todo_items:
        text = item.text.lower()
        if any(needle and needle in text for needle in needles):
            links.append(todo_item_link(todo_path, repo_root, item, ledger_output_dir))

    if not links:
        return "None"
    if len(links) > TODO_LINK_LIMIT:
        visible_links = links[:TODO_LINK_LIMIT]
        visible_links.append(f"{len(links) - TODO_LINK_LIMIT} more")
        return ", ".join(visible_links)
    return ", ".join(links)


def load_todo_items(todo_path: Path, repo_root: Path) -> tuple[TodoItem, ...]:
    """Load checkbox items from the first-adoption checklist when it exists."""
    if not todo_path.exists():
        return ()
    try:
        lines = todo_path.read_text(encoding="utf-8").splitlines()
    except OSError as error:
        todo_relative = repository_relative_path(todo_path, repo_root)
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        raise CandidateGenerationError(
            f"Unable to read adoption checklist {todo_relative}: {error_summary}"
        ) from error

    items: list[TodoItem] = []
    for line_number, line in enumerate(lines, start=1):
        match = re.match(r"^\s*[-*]\s+\[(?P<status>[ xX])\]\s+(?P<text>.+?)\s*$", line)
        if match is None:
            continue
        items.append(
            TodoItem(
                line_number=line_number,
                text=match.group("text"),
                is_complete=match.group("status").lower() == "x",
            )
        )
    return tuple(items)


def load_summary_todo_state(todo_path: Path, repo_root: Path) -> SummaryTodoState:
    """Load unchecked checklist items only from the documented TODO shape."""
    if not todo_path.exists():
        return SummaryTodoState(exists=False, is_interpretable=True, unchecked_items=())
    try:
        lines = todo_path.read_text(encoding="utf-8").splitlines()
    except OSError as error:
        todo_relative = repository_relative_path(todo_path, repo_root)
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        raise CandidateGenerationError(
            f"Unable to read adoption checklist {todo_relative}: {error_summary}"
        ) from error

    first_content_line = next(
        (line.strip() for line in lines if line.strip()),
        None,
    )
    has_documented_title = first_content_line == "# Repository Initialization Checklist"
    current_section: str | None = None
    found_known_section = False
    unchecked_items: list[SummaryTodoItem] = []
    known_sections = frozenset(TODO_KNOWN_SECTIONS)

    for line_number, line in enumerate(lines, start=1):
        heading_match = re.match(r"^##\s+(?P<section>.+?)\s*$", line)
        if heading_match is not None:
            section = heading_match.group("section")
            current_section = section if section in known_sections else None
            found_known_section = found_known_section or current_section is not None
            continue

        if current_section is None:
            continue

        task_match = re.match(
            r"^\s*[-*]\s+\[(?P<status>[ xX])\]\s+(?P<text>.+?)\s*$",
            line,
        )
        if task_match is None or task_match.group("status").lower() == "x":
            continue

        unchecked_items.append(
            SummaryTodoItem(
                line_number=line_number,
                section=current_section,
                text=task_match.group("text"),
            )
        )

    is_interpretable = has_documented_title and found_known_section
    return SummaryTodoState(
        exists=True,
        is_interpretable=is_interpretable,
        unchecked_items=tuple(unchecked_items) if is_interpretable else (),
    )


def empty_marker_data() -> MarkerData:
    """Return marker data for repositories that have not created a marker yet."""
    return MarkerData(
        last_reviewed_template_commit=None,
        included_modules=frozenset(),
        local_overrides=(),
        local_path_ownership=(),
        deferred_candidates=(),
        protected_decisions=(),
    )


def load_preflight_inputs(
    repo_root: Path,
    marker_path: Path,
    manifest_path: Path,
    marker_schema_path: Path,
    manifest_schema_path: Path,
) -> tuple[MarkerData, frozenset[str], tuple[ManifestMapping, ...]]:
    """Load manifest state and optional marker state for preflight reporting."""
    manifest = load_yaml_mapping(manifest_path, repo_root)
    manifest_schema = load_json_mapping(manifest_schema_path, repo_root)
    validate_schema(manifest, manifest_schema, manifest_path, repo_root)
    manifest_modules, mappings = parse_manifest(manifest)

    if not marker_path.exists():
        return empty_marker_data(), manifest_modules, mappings

    marker = load_yaml_mapping(marker_path, repo_root)
    marker_schema = load_json_mapping(marker_schema_path, repo_root)
    validate_schema(marker, marker_schema, marker_path, repo_root)
    return parse_marker(marker), manifest_modules, mappings


def iter_repo_files(repo_root: Path) -> tuple[Path, ...]:
    """Return regular repository files without following symlinks."""
    trusted_root = repo_root.resolve()
    discovered: list[Path] = []
    for dirpath, dirnames, filenames in os.walk(repo_root, followlinks=False):
        current_dir = Path(dirpath)
        kept_dirs: list[str] = []
        for dirname in dirnames:
            child_dir = current_dir / dirname
            if dirname in DISCOVERY_SKIP_DIRS or child_dir.is_symlink():
                continue
            try:
                child_dir.resolve().relative_to(trusted_root)
            except (OSError, ValueError):
                continue
            kept_dirs.append(dirname)
        dirnames[:] = kept_dirs

        for filename in filenames:
            child_file = current_dir / filename
            if child_file.is_symlink() or not child_file.is_file():
                continue
            try:
                child_file.resolve().relative_to(trusted_root)
            except (OSError, ValueError):
                continue
            discovered.append(child_file)
    return tuple(sorted(discovered))


def relative_file_set(repo_root: Path) -> frozenset[str]:
    """Return repository-relative file paths for stack discovery."""
    return frozenset(path.relative_to(repo_root).as_posix() for path in iter_repo_files(repo_root))


def list_existing_paths(repo_root: Path, candidates: tuple[str, ...]) -> tuple[str, ...]:
    """Return candidate paths that exist as non-symlink repository paths."""
    return tuple(path for path in candidates if repository_path_exists(repo_root, path))


def list_directory_files(
    repo_root: Path,
    relative_dir: str,
    suffixes: tuple[str, ...],
) -> tuple[str, ...]:
    """Return immediate files in a repository directory that match ``suffixes``."""
    if path_has_symlink_component(repo_root, relative_dir):
        return ()

    trusted_root = repo_root.resolve()
    directory = repo_root / relative_dir
    try:
        directory.resolve().relative_to(trusted_root)
    except (OSError, ValueError):
        return ()
    if not directory.is_dir():
        return ()

    paths: list[str] = []
    for child in sorted(directory.iterdir()):
        if child.is_symlink() or not child.is_file():
            continue
        if suffixes and child.suffix.lower() not in suffixes:
            continue
        try:
            child.resolve().relative_to(trusted_root)
        except (OSError, ValueError):
            continue
        paths.append(child.relative_to(repo_root).as_posix())
    return tuple(paths)


def parse_git_remotes(repo_root: Path) -> tuple[RemoteInfo, ...]:
    """Return configured Git remotes from local config."""
    result = run_git(repo_root, ["remote", "-v"], check=False)
    if result.returncode != 0:
        return ()

    remotes: list[RemoteInfo] = []
    for line in result.stdout.splitlines():
        parts = line.split()
        if len(parts) < 3:
            continue
        purpose = parts[2].strip("()")
        remotes.append(RemoteInfo(name=parts[0], url=parts[1], purpose=purpose))
    return tuple(remotes)


def normalize_dns_host(host: str) -> str:
    """Normalize a DNS host for exact comparison."""
    return host.lower().removesuffix(".")


def is_github_host(host: str | None, extra_allowed_hosts: tuple[str, ...] = ()) -> bool:
    """Return whether ``host`` is a known GitHub or explicit GHES host."""
    if not host:
        return False
    normalized_host = normalize_dns_host(host)
    normalized_extra_hosts = tuple(
        normalize_dns_host(extra_host) for extra_host in extra_allowed_hosts
    )
    return (
        normalized_host == "github.com"
        or normalized_host.endswith(".github.com")
        or normalized_host in normalized_extra_hosts
    )


def infer_repository_name(
    remotes: tuple[RemoteInfo, ...],
    *,
    extra_allowed_hosts: tuple[str, ...] = (),
) -> str | None:
    """Infer ``owner/name`` from a configured GitHub remote URL.

    The host is matched against the parsed URL authority (or the SCP-style host),
    never an arbitrary substring, so look-alike hosts such as ``notgithub.com`` or
    ``github.com.example.com`` are not misclassified as GitHub remotes. GHES hosts
    are trusted only when an explicit API base contributes an allowed host.
    """
    for remote in remotes:
        if "://" in remote.url:
            try:
                parsed = urlsplit(remote.url)
            except ValueError:
                continue
            host, path = parsed.hostname, parsed.path
        else:
            scp_match = re.match(r"^(?:[^@/]+@)?(?P<host>[^/:]+):(?P<path>.+)$", remote.url)
            if scp_match is None:
                continue
            host, path = scp_match.group("host"), scp_match.group("path")
        if not is_github_host(host, extra_allowed_hosts):
            continue
        segments = [segment for segment in path.split("/") if segment]
        if len(segments) < 2:
            continue
        owner, repo = segments[0], segments[1].removesuffix(".git")
        if owner and repo:
            return f"{owner}/{repo}"
    return None


def normalize_github_api_base(raw_api_base: str | None) -> str:
    """Validate and normalize the GitHub REST API base URL."""
    if raw_api_base is None:
        return DEFAULT_GITHUB_API_BASE

    try:
        parsed = urlsplit(raw_api_base)
    except ValueError as error:
        raise CandidateGenerationError("GitHub API base URL is malformed.") from error

    if parsed.scheme not in {"http", "https"} or parsed.hostname is None:
        raise CandidateGenerationError("GitHub API base URL must be an absolute HTTP(S) URL.")
    if parsed.username is not None or parsed.password is not None:
        raise CandidateGenerationError("GitHub API base URL must not contain credentials.")
    if parsed.query or parsed.fragment:
        raise CandidateGenerationError("GitHub API base URL must not include a query or fragment.")

    normalized_path = parsed.path.rstrip("/")
    return urlunsplit((parsed.scheme, parsed.netloc, normalized_path, "", ""))


def github_api_base_host(raw_api_base: str | None) -> str | None:
    """Return the explicit API base host when one was supplied."""
    if raw_api_base is None:
        return None
    api_base = normalize_github_api_base(raw_api_base)
    return urlsplit(api_base).hostname


def rest_headers(*, include_version_header: bool) -> tuple[tuple[str, str], ...]:
    """Return safe headers for GitHub REST reads."""
    headers = [
        ("Accept", GITHUB_REST_ACCEPT_HEADER),
        ("User-Agent", GITHUB_REST_USER_AGENT),
    ]
    if include_version_header:
        headers.append(("X-GitHub-Api-Version", GITHUB_REST_API_VERSION))
    return tuple(headers)


def github_rest_url(api_base: str, endpoint_path: str, query: dict[str, str] | None = None) -> str:
    """Build a GitHub REST URL from a normalized base and endpoint path."""
    url = f"{api_base.rstrip('/')}/{endpoint_path.lstrip('/')}"
    if query:
        url = f"{url}?{urlencode(query)}"
    return url


def github_rest_request(
    api_base: str,
    endpoint_path: str,
    *,
    include_version_header: bool,
    query: dict[str, str] | None = None,
) -> RestRequest:
    """Build one read-only GitHub REST request."""
    return RestRequest(
        method="GET",
        url=github_rest_url(api_base, endpoint_path, query),
        headers=rest_headers(include_version_header=include_version_header),
        timeout=GITHUB_REST_TIMEOUT_SECONDS,
    )


def github_rest_repo_path(owner_name: str) -> str | None:
    """Return the encoded repository REST endpoint path."""
    parts = owner_name.split("/")
    if len(parts) != 2 or not all(parts):
        return None
    owner, repository = (quote(part, safe="") for part in parts)
    return f"/repos/{owner}/{repository}"


def github_rest_endpoint_path(owner_name: str, suffix: str) -> str | None:
    """Return an encoded repository REST endpoint path with a suffix."""
    repo_path = github_rest_repo_path(owner_name)
    if repo_path is None:
        return None
    return f"{repo_path}/{suffix.lstrip('/')}"


def github_rest_branch_protection_path(owner_name: str, branch: str) -> str | None:
    """Return the encoded branch-protection REST endpoint path."""
    repo_path = github_rest_repo_path(owner_name)
    if repo_path is None or not branch:
        return None
    return f"{repo_path}/branches/{quote(branch, safe='')}/protection"


def rest_header(result: RestResult | RestJsonRead, name: str) -> str | None:
    """Return one response header value by case-insensitive name."""
    expected = name.casefold()
    for key, value in result.headers:
        if key.casefold() == expected:
            return value
    return None


def rest_json_message(result: RestResult) -> str | None:
    """Return a bounded message from a REST response body."""
    if not result.body.strip():
        return None
    try:
        payload: object = json.loads(result.body)
    except json.JSONDecodeError:
        return bounded_diagnostic_text(result.body.strip().splitlines()[0])

    payload_object = json_object(payload)
    if payload_object is None:
        return None
    message = json_string_field(payload_object, "message")
    if message is None:
        return None
    return bounded_diagnostic_text(message)


def rest_response_detail(result: RestResult) -> str:
    """Return a compact REST response diagnostic."""
    if result.error is not None:
        return result.error
    if result.status_code is None:
        return "transport error"

    detail = rest_json_message(result) or "non-2xx response"
    if len(detail) > 160:
        detail = f"{detail[:157]}..."
    return f"HTTP {result.status_code}: {detail}"


def rest_result_is_success(result: RestResult) -> bool:
    """Return whether a REST result has a successful HTTP status."""
    return result.status_code is not None and 200 <= result.status_code < 300


def rest_rate_limit_evidence(result: RestResult) -> bool:
    """Return whether a REST response has primary or secondary rate-limit evidence."""
    remaining = rest_header(result, "x-ratelimit-remaining")
    retry_after = rest_header(result, "retry-after")
    message = (rest_json_message(result) or "").casefold()
    return (
        remaining == "0"
        or retry_after is not None
        or "rate limit" in message
        or "secondary rate" in message
    )


def rest_failure_basis(result: RestResult) -> str:
    """Classify a REST failure into the observation basis vocabulary."""
    status_code = result.status_code
    message = (rest_json_message(result) or "").casefold()
    if status_code is None:
        return "transport-error"
    if status_code in {403, 429} and rest_rate_limit_evidence(result):
        return "rate-limited"
    if status_code == 401:
        return "auth-required"
    if status_code == 403:
        return "permission-required"
    if status_code == 404:
        if "unsupported" in message or "not supported" in message:
            return "unsupported-endpoint"
        return "not-found-or-not-visible"
    if status_code in {410, 415} or "unsupported" in message or "not supported" in message:
        return "unsupported-endpoint"
    return "manual-review"


def rest_failure_diagnostics(label: str, result: RestResult) -> tuple[tuple[str, str], ...]:
    """Return sanitized REST failure diagnostics."""
    diagnostics: list[tuple[str, str]] = [
        ("message", f"{label} unavailable: {rest_response_detail(result)}")
    ]
    if result.status_code is not None:
        diagnostics.append(("http_status", str(result.status_code)))
    message = rest_json_message(result)
    if message is not None:
        diagnostics.append(("response_message", message))
    for header_name, diagnostic_key in (
        ("x-ratelimit-remaining", "rate_limit_remaining"),
        ("x-ratelimit-reset", "rate_limit_reset"),
        ("retry-after", "retry_after"),
    ):
        value = rest_header(result, header_name)
        if value is not None:
            diagnostics.append((diagnostic_key, bounded_diagnostic_text(value, limit=80)))
    accepted_permissions = rest_header(result, "x-accepted-github-permissions")
    if accepted_permissions is not None:
        diagnostics.append(
            ("accepted_permissions", sanitize_accepted_permissions(accepted_permissions))
        )
    if result.error is not None:
        diagnostics.append(("transport", bounded_diagnostic_text(result.error)))
    return tuple(diagnostics)


def is_version_header_rejection(result: RestResult) -> bool:
    """Return whether a REST failure looks like API-version header incompatibility."""
    if result.status_code not in {400, 415}:
        return False
    body = result.body.lower()
    return (
        "x-github-api-version" in body
        or ("api version" in body and "header" in body)
        or "unsupported api version" in body
    )


def read_github_rest_json(
    *,
    api_base: str,
    endpoint_path: str,
    label: str,
    rest_client: RestClient,
    query: dict[str, str] | None = None,
) -> RestJsonRead:
    """Read a GitHub REST endpoint with one GHES version-header retry."""
    request_with_version = github_rest_request(
        api_base,
        endpoint_path,
        include_version_header=True,
        query=query,
    )
    result = rest_client(request_with_version)
    diagnostics: list[tuple[str, str]] = []

    if is_version_header_rejection(result):
        request_without_version = github_rest_request(
            api_base,
            endpoint_path,
            include_version_header=False,
            query=query,
        )
        retry_result = rest_client(request_without_version)
        warning = (
            f"{label} rejected the GitHub API version header; retried the same "
            "read-only endpoint without it."
        )
        diagnostics.append(("message", warning))
        if not rest_result_is_success(retry_result):
            diagnostics.extend(
                rest_failure_diagnostics(
                    label,
                    RestResult(
                        status_code=retry_result.status_code,
                        body=retry_result.body,
                        error=retry_result.error,
                        headers=retry_result.headers,
                    ),
                )
            )
            diagnostics.append(("message", f"Retry failed: {rest_response_detail(retry_result)}"))
            return RestJsonRead(
                payload=None,
                basis=rest_failure_basis(retry_result),
                diagnostics=tuple(diagnostics),
                headers=retry_result.headers,
            )
        result = retry_result

    if not rest_result_is_success(result):
        return RestJsonRead(
            payload=None,
            basis=rest_failure_basis(result),
            diagnostics=rest_failure_diagnostics(label, result),
            headers=result.headers,
        )

    try:
        return RestJsonRead(
            payload=json.loads(result.body),
            basis="rest",
            diagnostics=tuple(diagnostics),
            headers=result.headers,
        )
    except json.JSONDecodeError as error:
        return RestJsonRead(
            payload=None,
            basis="malformed-response",
            diagnostics=observation_diagnostic_pairs(
                ("message", f"{label} returned malformed JSON: {error.msg}")
            ),
            headers=result.headers,
        )


def link_header_has_next(headers: tuple[tuple[str, str], ...]) -> bool:
    """Return whether a REST response includes Link evidence for another page."""
    for key, value in headers:
        if key.casefold() != "link":
            continue
        # Link parameters and relation types are case-insensitive (RFC 8288), and
        # servers/clients may emit `REL="next"` or `rel="Next"`. Casefold each
        # fragment before matching so pagination is not stopped early.
        parts = [part.strip().casefold() for part in value.split(",")]
        if any('rel="next"' in part or "rel=next" in part for part in parts):
            return True
    return False


@dataclass(frozen=True)
class RestCollectionRead:
    """A paginated REST collection read with completeness metadata."""

    items: tuple[object, ...]
    observed: bool
    basis: str
    pagination: str
    diagnostics: tuple[tuple[str, str], ...]


def read_github_rest_collection(
    *,
    api_base: str,
    endpoint_path: str,
    label: str,
    rest_client: RestClient,
) -> RestCollectionRead:
    """Read a GitHub REST collection, following Link pagination when possible."""
    items: list[object] = []
    diagnostics: list[tuple[str, str]] = []
    page = 1
    while True:
        page_read = read_github_rest_json(
            api_base=api_base,
            endpoint_path=endpoint_path,
            label=label,
            query={"per_page": "100", "page": str(page)},
            rest_client=rest_client,
        )
        diagnostics.extend(page_read.diagnostics)
        if page_read.payload is None:
            if items:
                diagnostics.append(
                    ("pagination", f"page {page} unavailable; earlier pages retained")
                )
                return RestCollectionRead(
                    items=tuple(items),
                    observed=True,
                    basis="rest",
                    pagination="partial",
                    diagnostics=tuple(diagnostics),
                )
            return RestCollectionRead(
                items=(),
                observed=False,
                basis=page_read.basis,
                pagination="unavailable",
                diagnostics=tuple(diagnostics),
            )

        page_items = json_array(page_read.payload)
        if page_items is None:
            diagnostics.append(("message", f"{label} returned a non-list JSON payload."))
            if items:
                return RestCollectionRead(
                    items=tuple(items),
                    observed=True,
                    basis="rest",
                    pagination="partial",
                    diagnostics=tuple(diagnostics),
                )
            return RestCollectionRead(
                items=(),
                observed=False,
                basis="malformed-response",
                pagination="unavailable",
                diagnostics=tuple(diagnostics),
            )

        items.extend(page_items)
        has_next = link_header_has_next(page_read.headers)
        if not has_next:
            return RestCollectionRead(
                items=tuple(items),
                observed=True,
                basis="rest",
                pagination="complete",
                diagnostics=tuple(diagnostics),
            )

        diagnostics.append(
            (
                "pagination",
                f'{label} Link header included rel="next" after page {page}.',
            )
        )
        page += 1
        if page > GITHUB_REST_MAX_PAGES:
            diagnostics.append(
                (
                    "pagination",
                    f'stopped after {GITHUB_REST_MAX_PAGES} pages while rel="next" remained',
                )
            )
            return RestCollectionRead(
                items=tuple(items),
                observed=True,
                basis="rest",
                pagination="partial",
                diagnostics=tuple(diagnostics),
            )


def combine_metadata_errors(*errors: str | None) -> str | None:
    """Combine optional metadata diagnostics into one report line."""
    present_errors = [error for error in errors if error]
    if not present_errors:
        return None
    return "; ".join(present_errors)


def github_source_label(kind: str, endpoint_path: str) -> str:
    """Return a sanitized source label for a GitHub command or endpoint family."""
    return f"{kind} {endpoint_path}"


def github_boolean_observation(
    *,
    setting: str,
    value: object,
    source: str,
    observed_at: str,
    basis: str,
    diagnostics: tuple[tuple[str, str], ...] = (),
) -> GitHubObservation:
    """Return an observed, unobserved, or malformed boolean setting observation."""
    if isinstance(value, bool):
        return github_observation(
            setting=setting,
            value=value,
            observed=True,
            basis=basis,
            source=source,
            observed_at=observed_at,
            diagnostics=diagnostics,
        )
    if value is None:
        # A missing field in an otherwise-successful payload is unobserved, not
        # malformed (e.g. Discussions on older API versions, or an endpoint object
        # that omits the expected boolean). Record it as manual-review with whatever
        # caller diagnostics already explain the absence, consistent with the other
        # missing-field observations.
        return unobserved_github_observation(
            setting=setting,
            basis="manual-review",
            source=source,
            observed_at=observed_at,
            diagnostics=diagnostics,
        )
    return unobserved_github_observation(
        setting=setting,
        basis="malformed-response",
        source=source,
        observed_at=observed_at,
        diagnostics=diagnostics
        + observation_diagnostic_pairs(("message", f"{setting} was not a boolean.")),
    )


def security_and_analysis_summary(value: object) -> dict[str, JsonValue] | None:
    """Return an allowlisted security-and-analysis summary from repository metadata."""
    payload = json_object(value)
    if payload is None:
        return None

    summary: dict[str, JsonValue] = {}
    for key in sorted(GITHUB_SECURITY_AND_ANALYSIS_KEYS):
        raw_setting = json_object(payload.get(key))
        if raw_setting is None:
            continue
        status = json_string_field(raw_setting, "status")
        if status is not None:
            summary[key] = {"status": status}
    return summary


def branch_protection_summary(value: object) -> dict[str, JsonValue] | None:
    """Return a bounded branch-protection summary without raw URLs."""
    payload = json_object(value)
    if payload is None:
        return None

    enforce_admins = json_object(payload.get("enforce_admins"))
    enforce_admins_enabled = enforce_admins.get("enabled") if enforce_admins is not None else None
    return {
        "enabled": True,
        "required_status_checks": json_object(payload.get("required_status_checks")) is not None,
        "required_pull_request_reviews": (
            json_object(payload.get("required_pull_request_reviews")) is not None
        ),
        "enforce_admins": (
            enforce_admins_enabled if isinstance(enforce_admins_enabled, bool) else None
        ),
        "restrictions": json_object(payload.get("restrictions")) is not None,
    }


def ruleset_summaries(
    values: tuple[object, ...],
) -> tuple[list[JsonValue], tuple[str, ...], tuple[tuple[str, str], ...]]:
    """Return bounded ruleset summaries plus partial-field diagnostics.

    Always returns a concrete list (possibly empty): malformed items are skipped
    and recorded as diagnostics rather than collapsing the whole result to ``None``.
    """
    summaries: list[JsonValue] = []
    partial_fields: set[str] = set()
    diagnostics: list[tuple[str, str]] = []

    for index, raw_ruleset in enumerate(values):
        ruleset = json_object(raw_ruleset)
        if ruleset is None:
            diagnostics.append(("message", f"ruleset item {index + 1} was not an object"))
            continue
        summary: dict[str, JsonValue] = {}
        for field_name in ("id", "name", "target", "enforcement"):
            field_value = ruleset.get(field_name)
            if isinstance(field_value, (str, int, float, bool)) or field_value is None:
                summary[field_name] = field_value
        bypass_actors = ruleset.get("bypass_actors")
        if isinstance(bypass_actors, list):
            summary["bypass_actors_count"] = len(bypass_actors)
        else:
            partial_fields.add("rulesets[].bypass_actors")
        summaries.append(summary)

    return summaries, tuple(sorted(partial_fields)), tuple(diagnostics)


@dataclass(frozen=True)
class GitHubCliJsonRead:
    """Parsed ``gh api`` JSON plus sanitized diagnostics."""

    payload: object | None
    basis: str
    diagnostics: tuple[tuple[str, str], ...]


def github_cli_failure_basis(detail: str) -> str:
    """Classify a ``gh`` failure into the observation basis vocabulary."""
    normalized = detail.casefold()
    if "auth" in normalized or "login" in normalized or "credential" in normalized:
        return "auth-required"
    if "rate limit" in normalized or "secondary rate" in normalized:
        return "rate-limited"
    if "not found" in normalized or "404" in normalized:
        return "not-found-or-not-visible"
    if "permission" in normalized or "forbidden" in normalized or "403" in normalized:
        return "permission-required"
    if "unsupported" in normalized or "not supported" in normalized:
        return "unsupported-endpoint"
    return "manual-review"


def read_github_cli_json(
    *,
    repo_root: Path,
    endpoint_path: str,
    label: str,
    command_runner: CommandRunner,
    paginate: bool = False,
    hostname: str | None = None,
) -> GitHubCliJsonRead:
    """Read one GitHub API JSON payload through the GitHub CLI.

    ``hostname`` targets a non-default (e.g. GHES) GitHub host via ``gh api
    --hostname``; when ``None`` the GitHub CLI uses its configured default host.
    """
    command = ["gh", "api"]
    if hostname is not None:
        command.extend(["--hostname", hostname])
    if paginate:
        command.extend(["--paginate", "--slurp"])
    command.append(endpoint_path)
    try:
        result = command_runner(repo_root, command)
    except OSError as error:
        return GitHubCliJsonRead(
            payload=None,
            basis="transport-error",
            diagnostics=observation_diagnostic_pairs(
                ("message", f"gh unavailable: {os_error_summary(error)}")
            ),
        )

    if result.returncode != 0:
        detail = command_detail(result)
        return GitHubCliJsonRead(
            payload=None,
            basis=github_cli_failure_basis(detail),
            diagnostics=observation_diagnostic_pairs(("message", f"{label} unavailable: {detail}")),
        )

    try:
        payload: object = json.loads(result.stdout)
    except json.JSONDecodeError as error:
        return GitHubCliJsonRead(
            payload=None,
            basis="malformed-response",
            diagnostics=observation_diagnostic_pairs(
                ("message", f"{label} returned malformed JSON: {error.msg}")
            ),
        )
    return GitHubCliJsonRead(payload=payload, basis="gh", diagnostics=())


def gh_paginated_collection_items(payload: object) -> tuple[object, ...] | None:
    """Return collection items from ``gh api --paginate --slurp`` output."""
    pages = json_array(payload)
    if pages is None:
        return None
    if all(isinstance(page, list) for page in pages):
        items: list[object] = []
        for page in pages:
            items.extend(cast(list[object], page))
        return tuple(items)
    return tuple(pages)


def collection_observation_from_items(
    *,
    setting: str,
    source: str,
    observed_at: str,
    basis: str,
    items: tuple[object, ...],
    pagination: str,
    diagnostics: tuple[tuple[str, str], ...] = (),
) -> GitHubObservation:
    """Build labels or rulesets observations from collection records."""
    if setting == "labels":
        label_names = github_label_names(list(items))
        if label_names is None:
            return unobserved_github_observation(
                setting=setting,
                basis="malformed-response",
                source=source,
                observed_at=observed_at,
                pagination="unavailable",
                diagnostics=diagnostics
                + observation_diagnostic_pairs(
                    ("message", "Label metadata returned a non-list JSON payload.")
                ),
            )
        return github_observation(
            setting=setting,
            value=list(label_names),
            observed=True,
            basis=basis,
            source=source,
            observed_at=observed_at,
            pagination=pagination,
            diagnostics=diagnostics,
        )

    if setting == "rulesets":
        # ruleset_summaries always returns a concrete list (malformed items are
        # skipped and surfaced as diagnostics), so there is no None case to handle.
        summaries, partial_fields, summary_diagnostics = ruleset_summaries(items)
        return github_observation(
            setting=setting,
            value=summaries,
            observed=True,
            basis=basis,
            source=source,
            observed_at=observed_at,
            pagination=pagination,
            partial=bool(partial_fields or pagination == "partial"),
            partial_fields=partial_fields,
            diagnostics=diagnostics + summary_diagnostics,
        )

    raise ValueError(f"Unsupported collection setting: {setting}")


def repository_observations_from_payload(
    *,
    payload: JsonObject,
    basis: str,
    source: str,
    observed_at: str,
) -> tuple[GitHubObservation, ...]:
    """Return core repository observations from a repository metadata payload."""
    observations: list[GitHubObservation] = []

    repository = json_string_field(payload, "full_name") or json_string_field(
        payload, "nameWithOwner"
    )
    if repository is not None:
        observations.append(
            github_observation(
                setting="repository",
                value=repository,
                observed=True,
                basis=basis,
                source=source,
                observed_at=observed_at,
            )
        )
    else:
        observations.append(
            unobserved_github_observation(
                setting="repository",
                basis="manual-review",
                source=source,
                observed_at=observed_at,
                diagnostics=observation_diagnostic_pairs(
                    ("message", "repository full_name was not present in the payload")
                ),
            )
        )

    visibility = json_string_field(payload, "visibility")
    if visibility is None:
        private = payload.get("private")
        if isinstance(private, bool):
            visibility = "private" if private else "public"
    observations.append(
        github_observation(
            setting="visibility",
            value=visibility,
            observed=visibility is not None,
            basis=basis if visibility is not None else "manual-review",
            source=source,
            observed_at=observed_at,
            diagnostics=(
                ()
                if visibility is not None
                else observation_diagnostic_pairs(
                    ("message", "repository visibility was not present in the payload")
                )
            ),
        )
        if visibility is not None
        else unobserved_github_observation(
            setting="visibility",
            basis="manual-review",
            source=source,
            observed_at=observed_at,
            diagnostics=observation_diagnostic_pairs(
                ("message", "repository visibility was not present in the payload")
            ),
        )
    )

    default_branch = json_string_field(payload, "default_branch")
    default_branch_ref = json_object(payload.get("defaultBranchRef"))
    if default_branch is None and default_branch_ref is not None:
        default_branch = json_string_field(default_branch_ref, "name")
    observations.append(
        github_observation(
            setting="default_branch",
            value=default_branch,
            observed=True,
            basis=basis,
            source=source,
            observed_at=observed_at,
        )
        if default_branch is not None
        else unobserved_github_observation(
            setting="default_branch",
            basis="manual-review",
            source=source,
            observed_at=observed_at,
            diagnostics=observation_diagnostic_pairs(
                ("message", "default branch was not present in the payload")
            ),
        )
    )

    discussions = payload.get("has_discussions")
    if not isinstance(discussions, bool):
        discussions = payload.get("hasDiscussionsEnabled")
    observations.append(
        github_boolean_observation(
            setting="discussions",
            value=discussions,
            basis=basis,
            source=source,
            observed_at=observed_at,
            diagnostics=(
                ()
                if isinstance(discussions, bool)
                else observation_diagnostic_pairs(
                    ("message", "Discussions state was not present in the payload")
                )
            ),
        )
    )

    security_summary = security_and_analysis_summary(payload.get("security_and_analysis"))
    if security_summary is None:
        observations.append(
            unobserved_github_observation(
                setting="security_and_analysis",
                # A missing field in an otherwise-successful repository payload is not
                # proof of a permission failure: the block is also absent on older API
                # versions, on GitHub Enterprise Server, and for callers without admin
                # access. Classify it as manual-review (consistent with the other
                # missing-field observations above) and keep the permission hint in the
                # diagnostic instead of asserting it in the basis.
                basis="manual-review",
                source=source,
                observed_at=observed_at,
                diagnostics=observation_diagnostic_pairs(
                    (
                        "message",
                        (
                            "security_and_analysis was not present in the payload; it may require "
                            "repository admin access or may be unavailable on this API version or "
                            "GitHub Enterprise Server"
                        ),
                    )
                ),
            )
        )
    else:
        observations.append(
            github_observation(
                setting="security_and_analysis",
                value=security_summary,
                observed=True,
                basis=basis,
                source=source,
                observed_at=observed_at,
                partial=False,
            )
        )

    return tuple(observations)


def private_vulnerability_reporting_observation(
    *,
    payload: object | None,
    basis: str,
    source: str,
    observed_at: str,
    diagnostics: tuple[tuple[str, str], ...],
) -> GitHubObservation:
    """Return a private vulnerability reporting observation from an endpoint payload."""
    # ``payload is None`` is how the REST/gh reads represent an unavailable endpoint
    # (404, auth/permission failure, transport error, ...). Treat that as unobserved
    # while preserving the read's own basis/diagnostics, matching
    # ``branch_protection_observation``. Only a present-but-non-object payload is a
    # genuine malformed response.
    if payload is None:
        return unobserved_github_observation(
            setting="private_vulnerability_reporting",
            basis=basis,
            source=source,
            observed_at=observed_at,
            diagnostics=diagnostics,
        )
    payload_object = json_object(payload)
    if payload_object is None:
        return unobserved_github_observation(
            setting="private_vulnerability_reporting",
            basis="malformed-response",
            source=source,
            observed_at=observed_at,
            diagnostics=diagnostics
            + observation_diagnostic_pairs(
                ("message", "private vulnerability reporting payload was not an object")
            ),
        )
    return github_boolean_observation(
        setting="private_vulnerability_reporting",
        value=payload_object.get("enabled"),
        basis=basis,
        source=source,
        observed_at=observed_at,
        diagnostics=diagnostics,
    )


def branch_protection_observation(
    *,
    payload: object | None,
    basis: str,
    source: str,
    observed_at: str,
    diagnostics: tuple[tuple[str, str], ...],
) -> GitHubObservation:
    """Return a branch-protection observation from an endpoint payload."""
    if payload is None:
        return unobserved_github_observation(
            setting="branch_protection",
            basis=basis,
            source=source,
            observed_at=observed_at,
            diagnostics=diagnostics,
        )
    summary = branch_protection_summary(payload)
    if summary is None:
        return unobserved_github_observation(
            setting="branch_protection",
            basis="malformed-response",
            source=source,
            observed_at=observed_at,
            diagnostics=diagnostics
            + observation_diagnostic_pairs(
                ("message", "branch protection payload was not an object")
            ),
        )
    return github_observation(
        setting="branch_protection",
        value=summary,
        observed=True,
        basis=basis,
        source=source,
        observed_at=observed_at,
        diagnostics=diagnostics,
    )


def ordered_github_metadata(
    *,
    requested: bool,
    observations: tuple[GitHubObservation, ...],
) -> GitHubMetadata:
    """Return metadata with observations ordered by known setting keys."""
    by_setting = {observation.setting: observation for observation in observations}
    ordered = tuple(
        by_setting[setting] for setting in GITHUB_METADATA_SETTING_ORDER if setting in by_setting
    )
    extras = tuple(
        observation
        for observation in observations
        if observation.setting not in GITHUB_METADATA_SETTING_ORDER
    )
    return GitHubMetadata(requested=requested, observations=ordered + extras)


def discover_github_metadata_with_gh(
    repo_root: Path,
    owner_name: str,
    *,
    clock: Clock,
    command_runner: CommandRunner = run_command,
    api_base_host: str | None = None,
) -> GitHubMetadata:
    """Discover GitHub metadata through explicit read-only ``gh api`` calls.

    ``api_base_host`` targets a non-default (e.g. GHES) host so authenticated ``gh``
    reads work against that instance; when ``None`` the CLI uses its default host.
    """
    observed_at = observed_at_timestamp(clock)
    repo_path = github_rest_repo_path(owner_name)
    if repo_path is None:
        return ordered_github_metadata(
            requested=True,
            observations=unobserved_github_observations(
                basis="manual-review",
                source="gh api",
                observed_at=observed_at,
                diagnostics=observation_diagnostic_pairs(
                    ("message", "Repository owner/name is not in owner/repo form.")
                ),
            ),
        )

    repo_source = github_source_label("gh api", repo_path)
    repo_read = read_github_cli_json(
        repo_root=repo_root,
        endpoint_path=repo_path,
        label="gh repo metadata",
        command_runner=command_runner,
        hostname=api_base_host,
    )
    repo_payload = json_object(repo_read.payload)
    if repo_payload is None:
        diagnostics = repo_read.diagnostics or observation_diagnostic_pairs(
            ("message", "gh repo metadata returned a non-object JSON payload.")
        )
        return ordered_github_metadata(
            requested=True,
            observations=unobserved_github_observations(
                basis=repo_read.basis,
                source=repo_source,
                observed_at=observed_at,
                diagnostics=diagnostics,
            ),
        )

    observations = list(
        repository_observations_from_payload(
            payload=repo_payload,
            basis="gh",
            source=repo_source,
            observed_at=observed_at,
        )
    )

    labels_path = github_rest_endpoint_path(owner_name, "labels?per_page=100")
    if labels_path is not None:
        labels_source = github_source_label("gh api --paginate", labels_path)
        labels_read = read_github_cli_json(
            repo_root=repo_root,
            endpoint_path=labels_path,
            label="gh labels metadata",
            command_runner=command_runner,
            paginate=True,
            hostname=api_base_host,
        )
        label_items = gh_paginated_collection_items(labels_read.payload)
        if label_items is None:
            observations.append(
                unobserved_github_observation(
                    setting="labels",
                    basis=labels_read.basis,
                    source=labels_source,
                    observed_at=observed_at,
                    pagination="unavailable",
                    diagnostics=labels_read.diagnostics
                    or observation_diagnostic_pairs(
                        ("message", "gh labels metadata returned a non-list JSON payload.")
                    ),
                )
            )
        else:
            observations.append(
                collection_observation_from_items(
                    setting="labels",
                    source=labels_source,
                    observed_at=observed_at,
                    basis="gh",
                    items=label_items,
                    pagination="complete",
                    diagnostics=labels_read.diagnostics,
                )
            )

    rulesets_path = github_rest_endpoint_path(owner_name, "rulesets?per_page=100")
    if rulesets_path is not None:
        rulesets_source = github_source_label("gh api --paginate", rulesets_path)
        rulesets_read = read_github_cli_json(
            repo_root=repo_root,
            endpoint_path=rulesets_path,
            label="gh rulesets metadata",
            command_runner=command_runner,
            paginate=True,
            hostname=api_base_host,
        )
        ruleset_items = gh_paginated_collection_items(rulesets_read.payload)
        if ruleset_items is None:
            observations.append(
                unobserved_github_observation(
                    setting="rulesets",
                    basis=rulesets_read.basis,
                    source=rulesets_source,
                    observed_at=observed_at,
                    pagination="unavailable",
                    diagnostics=rulesets_read.diagnostics
                    or observation_diagnostic_pairs(
                        ("message", "gh rulesets metadata returned a non-list JSON payload.")
                    ),
                )
            )
        else:
            observations.append(
                collection_observation_from_items(
                    setting="rulesets",
                    source=rulesets_source,
                    observed_at=observed_at,
                    basis="gh",
                    items=ruleset_items,
                    pagination="complete",
                    diagnostics=rulesets_read.diagnostics,
                )
            )

    default_branch = github_observation_string_value(
        next(
            (
                observation
                for observation in observations
                if observation.setting == "default_branch"
            ),
            None,
        )
    )
    if default_branch is None:
        observations.append(
            unobserved_github_observation(
                setting="branch_protection",
                basis="dependency-unobserved",
                source="gh api branch protection",
                observed_at=observed_at,
                diagnostics=observation_diagnostic_pairs(
                    (
                        "message",
                        "default branch was not observed, so branch protection was not queried",
                    )
                ),
            )
        )
    else:
        branch_path = github_rest_branch_protection_path(owner_name, default_branch)
        assert branch_path is not None
        branch_source = github_source_label("gh api", branch_path)
        branch_read = read_github_cli_json(
            repo_root=repo_root,
            endpoint_path=branch_path,
            label="gh branch protection metadata",
            command_runner=command_runner,
            hostname=api_base_host,
        )
        observations.append(
            branch_protection_observation(
                payload=branch_read.payload,
                basis=branch_read.basis,
                source=branch_source,
                observed_at=observed_at,
                diagnostics=branch_read.diagnostics,
            )
        )

    private_vulnerability_path = github_rest_endpoint_path(
        owner_name, "private-vulnerability-reporting"
    )
    if private_vulnerability_path is not None:
        private_vulnerability_source = github_source_label("gh api", private_vulnerability_path)
        private_vulnerability_read = read_github_cli_json(
            repo_root=repo_root,
            endpoint_path=private_vulnerability_path,
            label="gh private vulnerability reporting metadata",
            command_runner=command_runner,
            hostname=api_base_host,
        )
        observations.append(
            private_vulnerability_reporting_observation(
                payload=private_vulnerability_read.payload,
                basis=private_vulnerability_read.basis,
                source=private_vulnerability_source,
                observed_at=observed_at,
                diagnostics=private_vulnerability_read.diagnostics,
            )
        )

    return ordered_github_metadata(requested=True, observations=tuple(observations))


def discover_github_metadata_with_rest(
    owner_name: str,
    *,
    github_api_base: str,
    rest_client: RestClient,
    gh_error: str | None,
    clock: Clock,
) -> GitHubMetadata:
    """Discover public-safe GitHub metadata through unauthenticated REST reads."""
    observed_at = observed_at_timestamp(clock)
    repo_path = github_rest_repo_path(owner_name)
    if repo_path is None:
        return ordered_github_metadata(
            requested=True,
            observations=unobserved_github_observations(
                basis="manual-review",
                source="REST repository metadata",
                observed_at=observed_at,
                diagnostics=observation_diagnostic_pairs(
                    ("message", "Repository owner/name is not in owner/repo form.")
                ),
            ),
        )

    diagnostics: list[str] = []
    if gh_error is not None:
        # gh diagnostics are already self-describing: the OSError/transport case emits
        # "gh unavailable: <error>", and command failures emit "gh <label> unavailable:
        # <detail>" (e.g. an auth/permission/rate-limit condition). Append verbatim
        # instead of blanket-prefixing "gh unavailable:", which mislabels those
        # non-transport failures as gh being missing or broken.
        diagnostics.append(gh_error)

    repo_read = read_github_rest_json(
        api_base=github_api_base,
        endpoint_path=repo_path,
        label="REST repo metadata",
        rest_client=rest_client,
    )
    if repo_read.payload is None:
        return ordered_github_metadata(
            requested=True,
            observations=unobserved_github_observations(
                basis=repo_read.basis,
                source=github_source_label("REST", repo_path),
                observed_at=observed_at,
                diagnostics=observation_diagnostic_pairs(
                    ("message", combine_metadata_errors(*diagnostics))
                )
                + repo_read.diagnostics,
            ),
        )
    repo_payload_object = json_object(repo_read.payload)
    if repo_payload_object is None:
        return ordered_github_metadata(
            requested=True,
            observations=unobserved_github_observations(
                basis="malformed-response",
                source=github_source_label("REST", repo_path),
                observed_at=observed_at,
                diagnostics=observation_diagnostic_pairs(
                    ("message", combine_metadata_errors(*diagnostics)),
                    ("message", "REST repo metadata returned a non-object JSON payload."),
                ),
            ),
        )

    observation_diagnostics = (
        observation_diagnostic_pairs(("message", combine_metadata_errors(*diagnostics)))
        + repo_read.diagnostics
    )
    observations = list(
        repository_observations_from_payload(
            payload=repo_payload_object,
            basis="rest",
            source=github_source_label("REST", repo_path),
            observed_at=observed_at,
        )
    )
    if observation_diagnostics:
        observations[0] = replace(
            observations[0],
            diagnostics=observations[0].diagnostics + observation_diagnostics,
        )

    labels_path = github_rest_endpoint_path(owner_name, "labels")
    if labels_path is not None:
        labels_read = read_github_rest_collection(
            api_base=github_api_base,
            endpoint_path=labels_path,
            label="REST labels metadata",
            rest_client=rest_client,
        )
        if labels_read.observed:
            observations.append(
                collection_observation_from_items(
                    setting="labels",
                    source=github_source_label("REST", labels_path),
                    observed_at=observed_at,
                    basis="rest",
                    items=labels_read.items,
                    pagination=labels_read.pagination,
                    diagnostics=labels_read.diagnostics,
                )
            )
        else:
            observations.append(
                unobserved_github_observation(
                    setting="labels",
                    basis=labels_read.basis,
                    source=github_source_label("REST", labels_path),
                    observed_at=observed_at,
                    pagination=labels_read.pagination,
                    diagnostics=labels_read.diagnostics,
                )
            )

    rulesets_path = github_rest_endpoint_path(owner_name, "rulesets")
    if rulesets_path is not None:
        rulesets_read = read_github_rest_collection(
            api_base=github_api_base,
            endpoint_path=rulesets_path,
            label="REST rulesets metadata",
            rest_client=rest_client,
        )
        if rulesets_read.observed:
            observations.append(
                collection_observation_from_items(
                    setting="rulesets",
                    source=github_source_label("REST", rulesets_path),
                    observed_at=observed_at,
                    basis="rest",
                    items=rulesets_read.items,
                    pagination=rulesets_read.pagination,
                    diagnostics=rulesets_read.diagnostics,
                )
            )
        else:
            observations.append(
                unobserved_github_observation(
                    setting="rulesets",
                    basis=rulesets_read.basis,
                    source=github_source_label("REST", rulesets_path),
                    observed_at=observed_at,
                    pagination=rulesets_read.pagination,
                    diagnostics=rulesets_read.diagnostics,
                )
            )

    default_branch = github_observation_string_value(
        next(
            (
                observation
                for observation in observations
                if observation.setting == "default_branch"
            ),
            None,
        )
    )
    if default_branch is None:
        observations.append(
            unobserved_github_observation(
                setting="branch_protection",
                basis="dependency-unobserved",
                source="REST branch protection",
                observed_at=observed_at,
                diagnostics=observation_diagnostic_pairs(
                    (
                        "message",
                        "default branch was not observed, so branch protection was not queried",
                    )
                ),
            )
        )
    else:
        branch_path = github_rest_branch_protection_path(owner_name, default_branch)
        assert branch_path is not None
        branch_read = read_github_rest_json(
            api_base=github_api_base,
            endpoint_path=branch_path,
            label="REST branch protection metadata",
            rest_client=rest_client,
        )
        observations.append(
            branch_protection_observation(
                payload=branch_read.payload,
                basis=branch_read.basis,
                source=github_source_label("REST", branch_path),
                observed_at=observed_at,
                diagnostics=branch_read.diagnostics,
            )
        )

    private_vulnerability_path = github_rest_endpoint_path(
        owner_name, "private-vulnerability-reporting"
    )
    if private_vulnerability_path is not None:
        private_vulnerability_read = read_github_rest_json(
            api_base=github_api_base,
            endpoint_path=private_vulnerability_path,
            label="REST private vulnerability reporting metadata",
            rest_client=rest_client,
        )
        observations.append(
            private_vulnerability_reporting_observation(
                payload=private_vulnerability_read.payload,
                basis=private_vulnerability_read.basis,
                source=github_source_label("REST", private_vulnerability_path),
                observed_at=observed_at,
                diagnostics=private_vulnerability_read.diagnostics,
            )
        )

    return ordered_github_metadata(requested=True, observations=tuple(observations))


def discover_github_metadata(
    repo_root: Path,
    owner_name: str | None,
    *,
    include_metadata: bool,
    command_runner: CommandRunner = run_command,
    rest_client: RestClient = run_rest_request,
    github_api_base: str | None = None,
    clock: Clock = utc_now,
) -> GitHubMetadata:
    """Optionally discover GitHub metadata through ``gh`` or public REST reads."""
    if not include_metadata:
        return GitHubMetadata(
            requested=False,
        )

    observed_at = observed_at_timestamp(clock)
    if owner_name is None:
        return ordered_github_metadata(
            requested=True,
            observations=unobserved_github_observations(
                basis="manual-review",
                source="local remotes",
                observed_at=observed_at,
                diagnostics=observation_diagnostic_pairs(
                    ("message", "Repository owner/name could not be inferred from local remotes.")
                ),
            ),
        )

    normalized_api_base = normalize_github_api_base(github_api_base)
    # Always attempt authenticated ``gh`` first, then fall back to public REST. For a
    # non-default (GHES) API base, target that host via ``gh api --hostname`` instead
    # of skipping ``gh`` entirely; skipping it drops observable authenticated GHES
    # metadata down to the unauthenticated REST path and breaks the documented
    # gh-to-REST fallback. If ``gh`` is unavailable for the host, the REST fallback
    # below still runs, so this is never worse than the REST-only path.
    api_base_host = (
        None
        if normalized_api_base == DEFAULT_GITHUB_API_BASE
        else github_api_base_host(github_api_base)
    )

    gh_metadata = discover_github_metadata_with_gh(
        repo_root,
        owner_name,
        clock=clock,
        command_runner=command_runner,
        api_base_host=api_base_host,
    )
    if gh_metadata.available:
        return gh_metadata

    return discover_github_metadata_with_rest(
        owner_name,
        github_api_base=normalized_api_base,
        rest_client=rest_client,
        gh_error=github_primary_diagnostic(gh_metadata),
        clock=clock,
    )


def current_branch(repo_root: Path) -> str | None:
    """Return the current branch name, or detached HEAD when applicable."""
    result = run_git(repo_root, ["branch", "--show-current"], check=False)
    branch = result.stdout.strip()
    if result.returncode == 0 and branch:
        return branch

    detached = run_git(repo_root, ["rev-parse", "--short", "HEAD"], check=False)
    if detached.returncode == 0 and detached.stdout.strip():
        return f"detached at {detached.stdout.strip()}"
    return None


def likely_default_branch(repo_root: Path, github_metadata: GitHubMetadata) -> str | None:
    """Return the likely default branch from local remote refs or metadata."""
    result = run_git(
        repo_root,
        ["symbolic-ref", "--quiet", "--short", "refs/remotes/origin/HEAD"],
        check=False,
    )
    remote_head = result.stdout.strip()
    if result.returncode == 0 and remote_head:
        return remote_head.split("/", maxsplit=1)[-1]
    return github_metadata.default_branch


def working_tree_entries(repo_root: Path) -> tuple[str, ...]:
    """Return porcelain working-tree entries."""
    result = run_git(repo_root, ["status", "--short"], check=False)
    if result.returncode != 0:
        return (f"status unavailable: {command_detail(result)}",)
    return tuple(line for line in result.stdout.splitlines() if line.strip())


def discover_workflows(repo_root: Path) -> tuple[str, ...]:
    """Return existing GitHub Actions workflow files."""
    return list_directory_files(repo_root, ".github/workflows", (".yml", ".yaml"))


def discover_issue_templates(repo_root: Path) -> tuple[str, ...]:
    """Return existing issue-template files."""
    direct_templates = list_existing_paths(repo_root, (".github/ISSUE_TEMPLATE.md",))
    directory_templates = list_directory_files(
        repo_root,
        ".github/ISSUE_TEMPLATE",
        (".md", ".yml", ".yaml"),
    )
    return tuple(sorted((*direct_templates, *directory_templates)))


def discover_pr_templates(repo_root: Path) -> tuple[str, ...]:
    """Return existing pull-request template files."""
    direct_templates = list_existing_paths(repo_root, (".github/pull_request_template.md",))
    directory_templates = list_directory_files(
        repo_root,
        ".github/PULL_REQUEST_TEMPLATE",
        (".md",),
    )
    return tuple(sorted((*direct_templates, *directory_templates)))


def discover_security_files(repo_root: Path) -> tuple[str, ...]:
    """Return existing governance and community-health files relevant to preflight."""
    return list_existing_paths(
        repo_root,
        (
            "SECURITY.md",
            ".github/SECURITY.md",
            "CODE_OF_CONDUCT.md",
            ".github/CODE_OF_CONDUCT.md",
            "CODEOWNERS",
            ".github/CODEOWNERS",
            "docs/CODEOWNERS",
            ".github/dependabot.yml",
        ),
    )


def discover_agent_instruction_files(repo_root: Path) -> tuple[str, ...]:
    """Return existing agent-instruction and modular instruction files."""
    direct_files = list_existing_paths(repo_root, tuple(sorted(AGENT_INSTRUCTION_EXACT_PATHS)))
    modular_files = list_directory_files(repo_root, ".github/instructions", (".md",))
    cursor_rules = list_directory_files(repo_root, ".cursor/rules", (".mdc",))
    return tuple(sorted((*direct_files, *modular_files, *cursor_rules)))


def discover_tooling_stacks(repo_root: Path, workflows: tuple[str, ...]) -> tuple[str, ...]:
    """Infer language and tooling stacks from local files without owner decisions."""
    files = relative_file_set(repo_root)
    stacks: list[str] = []

    def add_stack(name: str, condition: bool) -> None:
        if condition:
            stacks.append(name)

    add_stack("GitHub Actions", bool(workflows))
    add_stack("Markdown", any(path.endswith((".md", ".mdc")) for path in files))
    add_stack("Node/npm", "package.json" in files or "package-lock.json" in files)
    add_stack(
        "Python",
        any(path.endswith(".py") for path in files)
        or any(path in files for path in ("pyproject.toml", "requirements.txt", "setup.py")),
    )
    add_stack("PowerShell", any(path.endswith((".ps1", ".psd1")) for path in files))
    add_stack(
        "Terraform",
        any(
            path.endswith((".tf", ".tfvars", ".tftest.hcl", ".tf.json", ".tftpl", ".tfbackend"))
            for path in files
        )
        or ".tflint.hcl" in files,
    )
    add_stack("JSON", any(path.endswith((".json", ".jsonc")) for path in files))
    add_stack("YAML", any(path.endswith((".yml", ".yaml")) for path in files))
    add_stack(
        "JSON Schema",
        any(path.startswith("schemas/") and path.endswith(".schema.json") for path in files),
    )
    add_stack(
        "Template sync support",
        any(path.startswith(".template-sync/") for path in files),
    )
    return tuple(stacks)


def discover_repository_state(
    *,
    repo_root: Path,
    marker_path: Path,
    todo_path: Path,
    include_github_metadata: bool,
    github_api_base: str | None = None,
    command_runner: CommandRunner = run_command,
    rest_client: RestClient = run_rest_request,
    clock: Clock = utc_now,
) -> RepositoryDiscovery:
    """Discover read-only repository state for the preflight report."""
    remotes = parse_git_remotes(repo_root)
    explicit_api_host = github_api_base_host(github_api_base)
    owner_name = infer_repository_name(
        remotes,
        extra_allowed_hosts=((explicit_api_host,) if explicit_api_host else ()),
    )
    github_metadata = discover_github_metadata(
        repo_root,
        owner_name,
        include_metadata=include_github_metadata,
        command_runner=command_runner,
        rest_client=rest_client,
        github_api_base=github_api_base,
        clock=clock,
    )
    workflows = discover_workflows(repo_root)
    return RepositoryDiscovery(
        owner_name=github_metadata.repository or owner_name,
        remotes=remotes,
        current_branch=current_branch(repo_root),
        likely_default_branch=likely_default_branch(repo_root, github_metadata),
        working_tree_entries=working_tree_entries(repo_root),
        workflows=workflows,
        issue_templates=discover_issue_templates(repo_root),
        pr_templates=discover_pr_templates(repo_root),
        security_files=discover_security_files(repo_root),
        agent_instruction_files=discover_agent_instruction_files(repo_root),
        tooling_stacks=discover_tooling_stacks(repo_root, workflows),
        marker_present=marker_path.exists(),
        adoption_note_present=todo_path.exists(),
        github_metadata=github_metadata,
    )


def build_manifest_ledger_row(
    *,
    mapping: ManifestMapping,
    marker_data: MarkerData,
    manifest_modules: frozenset[str],
    todo_items: tuple[TodoItem, ...],
    todo_path: Path,
    repo_root: Path,
    default_adoption_mode: str,
    ledger_output_dir: Path | None = None,
) -> tuple[LedgerRow, tuple[LocalOverride, ...], tuple[ProtectedFileDecision, ...]]:
    """Build one adoption-ledger row from a manifest mapping."""
    modules = relation_module_names(mapping)
    local_overrides = matching_local_overrides_for_pattern(
        mapping.pattern,
        marker_data.local_overrides,
    )
    deferred_candidates = matching_deferred_candidates_for_pattern(
        mapping.pattern,
        marker_data.deferred_candidates,
    )
    protected_decisions = matching_protected_decisions_for_pattern(
        mapping.pattern,
        marker_data.protected_decisions,
    )
    is_protected = is_protected_manifest_pattern(mapping.pattern)
    retained = not mapping.unknown_modules and PathRelation(
        patterns=(mapping.pattern,),
        requires_all=mapping.requires_all,
        requires_any=mapping.requires_any,
        notes=(mapping.notes,) if mapping.notes else (),
        unknown_modules=mapping.unknown_modules,
    ).is_retained_by(marker_data.included_modules)
    should_trim = (
        retained
        and mapping.notes is not None
        and has_trim_note((mapping.notes,))
        and notes_mention_modules(
            (mapping.notes,),
            manifest_modules - marker_data.included_modules,
        )
    )

    decision = "retain"
    requires_decision = False
    reason_parts: list[str] = []

    if mapping.unknown_modules:
        decision = "needs maintainer decision"
        requires_decision = True
        reason_parts.append(
            "Manifest references unknown module(s): "
            + ", ".join(sorted(mapping.unknown_modules))
            + "."
        )
    elif protected_decisions:
        decision = "protected decision: " + ", ".join(
            protected_decision.decision for protected_decision in protected_decisions
        )
        requires_decision = protected_decision_requires_maintainer_decision(protected_decisions)
        reason_parts.append(protected_decision_reason(protected_decisions, local_overrides))
    elif local_overrides:
        decision = "local override"
        reason_parts.append(local_override_reason(local_overrides))
        if any(
            override.default_decision in {"DEFER", "PROTECTED-REVIEW"}
            for override in local_overrides
        ):
            requires_decision = True
    elif deferred_candidates:
        decision = "needs maintainer decision"
        requires_decision = True
        reason_parts.append(
            "Deferred protected candidate: "
            + "; ".join(
                f"{candidate.source_commit}: {candidate.reason}"
                for candidate in deferred_candidates
            )
        )
    elif is_protected:
        decision = "needs maintainer decision"
        requires_decision = True
        if retained:
            reason_parts.append("Protected retained path requires explicit owner authorization.")
        else:
            reason_parts.append(
                "Protected path is not retained by included modules; confirm deferral or skip."
            )
    elif not retained:
        decision = "skip"
        reason_parts.append(retention_gap_reason(mapping, marker_data.included_modules))
    elif should_trim:
        decision = "trim"
        reason_parts.append(
            "Retained path has manifest notes for module-scoped inline blocks; "
            "remove blocks for unadopted modules and keep the retained file."
        )
    else:
        reason_parts.append(
            "Retained because marker included_modules satisfy the manifest relation."
        )

    if is_protected and decision != "needs maintainer decision" and not protected_decisions:
        requires_decision = True
        reason_parts.append(
            "Protected file; explicit owner authorization is required before edits."
        )
    if mapping.notes:
        reason_parts.append(f"Manifest note: {mapping.notes}")

    return (
        LedgerRow(
            path=mapping.pattern,
            manifest_modules=format_manifest_modules(mapping.requires_all, mapping.requires_any),
            decision=decision,
            reason=" ".join(reason_parts),
            protected_file="Yes" if is_protected else "No",
            requires_maintainer_decision="Yes" if requires_decision else "No",
            adoption_mode=(
                adoption_mode_for_modules(
                    modules,
                    local_overrides,
                    is_protected,
                    default_adoption_mode,
                )
                if not protected_decisions
                else protected_decision_adoption_mode(protected_decisions)
            ),
            todo_link=matching_todo_links(
                path=mapping.pattern,
                modules=modules,
                is_protected=is_protected,
                todo_items=todo_items,
                todo_path=todo_path,
                repo_root=repo_root,
                ledger_output_dir=ledger_output_dir,
            ),
            validation_commands=validation_commands_for_modules(modules),
        ),
        local_overrides,
        protected_decisions,
    )


def build_unmatched_local_override_row(
    *,
    local_override: LocalOverride,
    mappings: tuple[ManifestMapping, ...],
    todo_items: tuple[TodoItem, ...],
    todo_path: Path,
    repo_root: Path,
    default_adoption_mode: str,
    ledger_output_dir: Path | None = None,
) -> LedgerRow:
    """Build a ledger row for a marker local override not consumed by any manifest mapping row."""
    display_path = local_override.path + ("/" if local_override.is_directory else "")
    relation = selected_relation_for_path(local_override.path, mappings)
    modules = relation_modules(relation)
    is_protected = is_protected_instruction_path(local_override.path)
    requires_decision = (
        is_protected
        or relation is None
        or local_override.default_decision in {"DEFER", "PROTECTED-REVIEW"}
    )
    reason = (
        f"Marker local override defaults to `{local_override.default_decision}`: "
        f"{local_override.reason}"
    )
    if relation is None:
        reason += " No manifest mapping currently matches this override path."

    return LedgerRow(
        path=display_path,
        manifest_modules=(
            format_manifest_modules(relation.requires_all, relation.requires_any)
            if relation is not None
            else "unmapped"
        ),
        decision="local override",
        reason=reason,
        protected_file="Yes" if is_protected else "No",
        requires_maintainer_decision="Yes" if requires_decision else "No",
        adoption_mode=adoption_mode_for_modules(
            modules,
            (local_override,),
            is_protected,
            default_adoption_mode,
        ),
        todo_link=matching_todo_links(
            path=display_path,
            modules=modules,
            is_protected=is_protected,
            todo_items=todo_items,
            todo_path=todo_path,
            repo_root=repo_root,
            ledger_output_dir=ledger_output_dir,
        ),
        validation_commands=validation_commands_for_modules(modules),
    )


def build_local_path_ownership_row(
    *,
    local_path_ownership: LocalPathOwnership,
    mappings: tuple[ManifestMapping, ...],
    todo_items: tuple[TodoItem, ...],
    todo_path: Path,
    repo_root: Path,
    ledger_output_dir: Path | None = None,
) -> LedgerRow:
    """Build a ledger row for one downstream local path ownership record."""
    relation = selected_relation_for_path(local_path_ownership.path, mappings)
    if relation is None and local_path_ownership.is_directory:
        # A directory record (for example ``schemas/``) is never matched by a glob
        # such as ``schemas/**``, so fall back to the manifest mappings under the
        # directory prefix. This keeps the ledger's manifest_modules and proximity
        # note consistent with the broad-overlap checks in marker validation.
        relation = directory_prefix_relation(local_path_ownership.path, mappings)
    modules = relation_modules(relation)
    is_protected = is_protected_instruction_path(local_path_ownership.path)
    reason = f"Marker local path ownership: {local_path_ownership.reason}"
    if local_path_ownership.overlap_exception_reason is not None:
        reason += (
            f" Broad manifest overlap exception: {local_path_ownership.overlap_exception_reason}"
        )
    if relation is not None:
        reason += " Manifest proximity is documented by this ownership record."

    return LedgerRow(
        path=local_path_ownership.display_path,
        manifest_modules=(
            format_manifest_modules(relation.requires_all, relation.requires_any)
            if relation is not None
            else "local"
        ),
        decision="local ownership",
        reason=reason,
        protected_file="Yes" if is_protected else "No",
        requires_maintainer_decision="Yes" if is_protected else "No",
        adoption_mode="downstream-owned",
        todo_link=matching_todo_links(
            path=local_path_ownership.display_path,
            modules=modules,
            is_protected=is_protected,
            todo_items=todo_items,
            todo_path=todo_path,
            repo_root=repo_root,
            ledger_output_dir=ledger_output_dir,
        ),
        validation_commands=validation_commands_for_modules(modules),
    )


def build_unmatched_protected_decision_row(
    *,
    protected_decision: ProtectedFileDecision,
    mappings: tuple[ManifestMapping, ...],
    todo_items: tuple[TodoItem, ...],
    todo_path: Path,
    repo_root: Path,
    ledger_output_dir: Path | None = None,
) -> LedgerRow:
    """Build a ledger row for a protected decision not consumed by a manifest mapping."""
    relation = selected_relation_for_path(protected_decision.path, mappings)
    modules = relation_modules(relation)
    is_protected = is_protected_instruction_path(protected_decision.path)
    return LedgerRow(
        path=protected_decision.path,
        manifest_modules=(
            format_manifest_modules(relation.requires_all, relation.requires_any)
            if relation is not None
            else "unmapped"
        ),
        decision=f"protected decision: {protected_decision.decision}",
        reason=protected_decision_reason((protected_decision,)),
        protected_file="Yes" if is_protected else "No",
        requires_maintainer_decision=(
            "Yes"
            if protected_decision_requires_maintainer_decision((protected_decision,))
            else "No"
        ),
        adoption_mode=protected_decision_adoption_mode((protected_decision,)),
        todo_link=matching_todo_links(
            path=protected_decision.path,
            modules=modules,
            is_protected=is_protected,
            todo_items=todo_items,
            todo_path=todo_path,
            repo_root=repo_root,
            ledger_output_dir=ledger_output_dir,
        ),
        validation_commands=validation_commands_for_modules(modules),
    )


def build_manual_todo_ledger_row(
    *,
    todo_item: TodoItem,
    todo_path: Path,
    repo_root: Path,
    ledger_output_dir: Path | None = None,
) -> LedgerRow:
    """Build a ledger row for one first-adoption checklist item."""
    status = "complete" if todo_item.is_complete else "open"
    return LedgerRow(
        path=repository_relative_path(todo_path, repo_root),
        manifest_modules="manual setup",
        decision="manual TODO",
        reason=f"{status}: {todo_item.text}",
        protected_file="No",
        requires_maintainer_decision="No" if todo_item.is_complete else "Yes",
        adoption_mode="recorded in checklist",
        todo_link=todo_item_link(todo_path, repo_root, todo_item, ledger_output_dir),
        validation_commands="manual first-adoption review",
    )


def build_adoption_ledger_rows(
    *,
    marker_data: MarkerData,
    manifest_modules: frozenset[str],
    mappings: tuple[ManifestMapping, ...],
    todo_items: tuple[TodoItem, ...],
    todo_path: Path,
    repo_root: Path,
    default_adoption_mode: str,
    ledger_output_dir: Path | None = None,
) -> tuple[LedgerRow, ...]:
    """Build the generated adoption ledger rows."""
    rows: list[LedgerRow] = []
    matched_overrides: set[LocalOverride] = set()
    matched_protected_decisions: set[ProtectedFileDecision] = set()
    for mapping in mappings:
        row, local_overrides, protected_decisions = build_manifest_ledger_row(
            mapping=mapping,
            marker_data=marker_data,
            manifest_modules=manifest_modules,
            todo_items=todo_items,
            todo_path=todo_path,
            repo_root=repo_root,
            default_adoption_mode=default_adoption_mode,
            ledger_output_dir=ledger_output_dir,
        )
        rows.append(row)
        matched_overrides.update(local_overrides)
        matched_protected_decisions.update(protected_decisions)

    for local_override in sorted(
        set(marker_data.local_overrides) - matched_overrides,
        key=lambda override: override.path,
    ):
        rows.append(
            build_unmatched_local_override_row(
                local_override=local_override,
                mappings=mappings,
                todo_items=todo_items,
                todo_path=todo_path,
                repo_root=repo_root,
                default_adoption_mode=default_adoption_mode,
                ledger_output_dir=ledger_output_dir,
            )
        )

    for local_path_ownership in sorted(
        marker_data.local_path_ownership,
        key=lambda record: record.display_path,
    ):
        rows.append(
            build_local_path_ownership_row(
                local_path_ownership=local_path_ownership,
                mappings=mappings,
                todo_items=todo_items,
                todo_path=todo_path,
                repo_root=repo_root,
                ledger_output_dir=ledger_output_dir,
            )
        )

    for protected_decision in sorted(
        set(marker_data.protected_decisions) - matched_protected_decisions,
        key=lambda decision: decision.path,
    ):
        rows.append(
            build_unmatched_protected_decision_row(
                protected_decision=protected_decision,
                mappings=mappings,
                todo_items=todo_items,
                todo_path=todo_path,
                repo_root=repo_root,
                ledger_output_dir=ledger_output_dir,
            )
        )

    for todo_item in todo_items:
        rows.append(
            build_manual_todo_ledger_row(
                todo_item=todo_item,
                todo_path=todo_path,
                repo_root=repo_root,
                ledger_output_dir=ledger_output_dir,
            )
        )

    return tuple(rows)


def markdown_cell(value: str) -> str:
    """Escape a value for inclusion in a Markdown table cell."""
    return value.replace("\n", "<br>").replace("|", r"\|")


def format_table(rows: tuple[CandidateRow, ...]) -> str:
    """Render candidate rows as a Markdown table."""
    header = (
        "| Path | Change | Matched module relation | Retained status | Local override | "
        "Deferred protected candidate | Protected file decision | "
        "Protected instruction/governance file | Notes |"
    )
    divider = "| --- | --- | --- | --- | --- | --- | --- | --- | --- |"
    rendered_rows = [header, divider]
    for row in rows:
        rendered_rows.append(
            "| "
            + " | ".join(
                markdown_cell(cell)
                for cell in (
                    row.path,
                    row.change,
                    row.module_relation,
                    row.retained_status,
                    row.local_override_status,
                    row.deferred_status,
                    row.protected_decision_status,
                    row.protected_status,
                    "<br>".join(row.notes) if row.notes else "None",
                )
            )
            + " |"
        )
    return "\n".join(rendered_rows)


def format_ledger_table(rows: tuple[LedgerRow, ...]) -> str:
    """Render adoption-ledger rows as a Markdown table."""
    header = (
        "| Path | Manifest module(s) | Decision | Reason | Protected file | "
        "Requires maintainer decision | Adoption mode | `_TODO-repo-init.md` link | "
        "Validation command affected |"
    )
    divider = "| --- | --- | --- | --- | --- | --- | --- | --- | --- |"
    rendered_rows = [header, divider]
    for row in rows:
        rendered_rows.append(
            "| "
            + " | ".join(
                markdown_cell(cell)
                for cell in (
                    row.path,
                    row.manifest_modules,
                    row.decision,
                    row.reason,
                    row.protected_file,
                    row.requires_maintainer_decision,
                    row.adoption_mode,
                    row.todo_link,
                    row.validation_commands,
                )
            )
            + " |"
        )
    return "\n".join(rendered_rows)


def format_protected_decision_records(marker_data: MarkerData) -> str | None:
    """Render protected-file decision records and overlapping marker context."""
    if not marker_data.protected_decisions:
        return None

    overlaps_by_path = {
        overlap.path: overlap
        for overlap in validate_protected_file_decisions(
            marker_data.protected_decisions,
            marker_data.local_overrides,
            marker_data.deferred_candidates,
        )
    }
    lines = [
        "## Protected File Decision Records",
        "",
        "These entries come from `template_sync.protected_file_decisions`.",
        "",
    ]
    for protected_decision in sorted(
        marker_data.protected_decisions,
        key=lambda decision: decision.path,
    ):
        lines.append(f"- `{protected_decision.path}`")
        lines.append(
            f"  - protected_file_decisions: {protected_decision_summary(protected_decision)}"
        )
        overlap = overlaps_by_path.get(protected_decision.path)
        if overlap is not None:
            for local_override in overlap.local_overrides:
                lines.append(f"  - local_overrides: {local_override_summary(local_override)}")
            for candidate in overlap.deferred_candidates:
                lines.append(
                    f"  - deferred_protected_candidates: {deferred_candidate_summary(candidate)}"
                )
    return "\n".join(lines)


def format_remove_local_authorizations(marker_data: MarkerData) -> str | None:
    """Render REMOVE-LOCAL authorization details in a distinct section."""
    remove_local_decisions = tuple(
        protected_decision
        for protected_decision in marker_data.protected_decisions
        if protected_decision.decision == REMOVAL_DECISION
    )
    if not remove_local_decisions:
        return None

    lines = [
        "## REMOVE-LOCAL Authorizations",
        "",
        (
            "Reviewers must verify that each authorization basis names the removed "
            "file and explicitly authorizes removal."
        ),
        "",
    ]
    for protected_decision in sorted(
        remove_local_decisions,
        key=lambda decision: decision.path,
    ):
        lines.append(f"- `{protected_decision.path}`")
        lines.append(f"  - authorization_basis: {protected_decision.authorization_basis}")
        lines.append(f"  - authorized_scope: {protected_decision.authorized_scope}")
        lines.append(f"  - reason: {protected_decision.reason}")
    return "\n".join(lines)


def format_adoption_ledger(
    *,
    marker_path: Path,
    manifest_path: Path,
    todo_path: Path,
    repo_root: Path,
    marker_data: MarkerData,
    rows: tuple[LedgerRow, ...],
    default_adoption_mode: str,
) -> str:
    """Render the generated adoption ledger as a Markdown snapshot."""
    marker_relative = repository_relative_path(marker_path, repo_root)
    manifest_relative = repository_relative_path(manifest_path, repo_root)
    todo_relative = repository_relative_path(todo_path, repo_root)
    todo_status = "found" if todo_path.exists() else "not found"
    included_modules = (
        ", ".join(f"`{module}`" for module in sorted(marker_data.included_modules)) or "none"
    )

    lines = [
        "# Template Adoption Ledger",
        "",
        (
            "Generated snapshot; review artifact only. "
            f"`{manifest_relative}` and `{marker_relative}` remain the authoritative "
            "machine-readable state. Regenerate this ledger before first-adoption, "
            "full-reconciliation, or protected-file review."
        ),
        "",
        f"- Marker: `{marker_relative}`",
        f"- Manifest: `{manifest_relative}`",
        f"- First-adoption checklist: `{todo_relative}` ({todo_status})",
        f"- Included modules: {included_modules}",
        f"- Default adoption mode: `{default_adoption_mode}`",
        "",
        format_ledger_table(rows) if rows else "No manifest mappings or manual TODO rows found.",
    ]
    protected_decision_records = format_protected_decision_records(marker_data)
    if protected_decision_records is not None:
        lines.extend(["", protected_decision_records])
    remove_local_authorizations = format_remove_local_authorizations(marker_data)
    if remove_local_authorizations is not None:
        lines.extend(["", remove_local_authorizations])
    return "\n".join(lines)


def format_summary_module_list(modules: frozenset[str]) -> str:
    """Render a deterministic bullet list of module names."""
    if not modules:
        return "- None"
    return "\n".join(f"- `{module}`" for module in sorted(modules))


def protected_decision_authorization_status(protected_decision: ProtectedFileDecision) -> str:
    """Return concise authorization status for one protected-file decision."""
    if protected_decision.decision in {"TAKE", "MERGE", REMOVAL_DECISION}:
        missing_fields: list[str] = []
        if protected_decision.authorization_basis is None:
            missing_fields.append("authorization_basis")
        if protected_decision.authorized_scope is None:
            missing_fields.append("authorized_scope")
        if (
            protected_decision.adoption_mode == "tailored"
            and protected_decision.tailored_authorization_basis is None
        ):
            missing_fields.append("tailored_authorization_basis")
        if missing_fields:
            return "missing " + ", ".join(missing_fields)
        return "authorized"

    if protected_decision.decision in {"DEFER", "PROTECTED-REVIEW"}:
        return "deferred"

    return "not applicable"


def format_summary_protected_decisions(marker_data: MarkerData) -> str:
    """Render protected-file decisions for the concise adoption summary."""
    if not marker_data.protected_decisions:
        return "None recorded."

    header = (
        "| Path | Decision | Adoption mode | Authorization status | Authorization basis | "
        "Authorized scope | Tailored authorization basis | Reason |"
    )
    divider = "| --- | --- | --- | --- | --- | --- | --- | --- |"
    lines = [header, divider]
    for protected_decision in sorted(
        marker_data.protected_decisions,
        key=lambda decision: decision.path,
    ):
        lines.append(
            "| "
            + " | ".join(
                markdown_cell(cell)
                for cell in (
                    protected_decision.path,
                    protected_decision.decision,
                    protected_decision.adoption_mode or "not recorded",
                    protected_decision_authorization_status(protected_decision),
                    protected_decision.authorization_basis or "not recorded",
                    protected_decision.authorized_scope or "not recorded",
                    protected_decision.tailored_authorization_basis or "not recorded",
                    protected_decision.reason or "not recorded",
                )
            )
            + " |"
        )
    return "\n".join(lines)


def format_summary_local_overrides(marker_data: MarkerData) -> str:
    """Render marker local overrides for the concise adoption summary."""
    if not marker_data.local_overrides:
        return "None recorded."

    lines = [
        "| Path | Default decision | Reason |",
        "| --- | --- | --- |",
    ]
    for local_override in sorted(
        marker_data.local_overrides,
        key=lambda override: override.path,
    ):
        display_path = local_override.path + ("/" if local_override.is_directory else "")
        lines.append(
            "| "
            + " | ".join(
                markdown_cell(cell)
                for cell in (
                    display_path,
                    local_override.default_decision,
                    local_override.reason,
                )
            )
            + " |"
        )
    return "\n".join(lines)


def format_summary_local_path_ownership(marker_data: MarkerData) -> str:
    """Render marker local path ownership records for the concise adoption summary."""
    if not marker_data.local_path_ownership:
        return "None recorded."

    lines = [
        "| Path | Reason | Overlap exception reason |",
        "| --- | --- | --- |",
    ]
    for record in sorted(
        marker_data.local_path_ownership,
        key=lambda item: item.display_path,
    ):
        lines.append(
            "| "
            + " | ".join(
                markdown_cell(cell)
                for cell in (
                    record.display_path,
                    record.reason,
                    record.overlap_exception_reason or "not recorded",
                )
            )
            + " |"
        )
    return "\n".join(lines)


def format_summary_manual_todo_items(
    *,
    todo_state: SummaryTodoState,
    todo_path: Path,
    repo_root: Path,
) -> str:
    """Render machine-interpretable unchecked TODO items for the summary."""
    todo_relative = repository_relative_path(todo_path, repo_root)
    if not todo_state.exists:
        return f"`{todo_relative}` not found."
    if not todo_state.is_interpretable:
        return (
            f"`{todo_relative}` is present but not machine-interpretable; "
            "unchecked items were not parsed."
        )
    if not todo_state.unchecked_items:
        return "No unchecked items in machine-interpretable checklist sections."

    lines: list[str] = []
    for section in TODO_KNOWN_SECTIONS:
        for item in todo_state.unchecked_items:
            if item.section == section:
                lines.append(
                    f"- {section}: {item.text} " + f"(`{todo_relative}` line {item.line_number})"
                )
    return "\n".join(lines)


def append_unique_summary_item(
    items: list[str],
    seen_items: set[str],
    item: str,
) -> None:
    """Append one summary item while preserving deterministic first occurrence."""
    if item in seen_items:
        return
    seen_items.add(item)
    items.append(item)


def deferred_candidate_rendered_in_rows(
    candidate: DeferredProtectedCandidate,
    rows: tuple[LedgerRow, ...],
) -> bool:
    """Return True when a deferred-candidate ledger row already covers this candidate.

    A deferred protected candidate is folded into a manifest ledger row only when that
    row takes the deferred-candidate branch, which embeds ``Deferred protected
    candidate:`` in the row reason and matches the candidate through the row's manifest
    pattern. Matching on the candidate path via the manifest pattern, rather than
    substring-matching the reason text, avoids suppressing distinct candidates that
    merely share a templated reason such as ``Awaiting protected-file authorization.``.
    """
    return any(
        row.requires_maintainer_decision == "Yes"
        and "Deferred protected candidate:" in row.reason
        and manifest_pattern_matches_path(row.path, candidate.path)
        for row in rows
    )


def format_summary_unresolved_decisions(
    *,
    rows: tuple[LedgerRow, ...],
    marker_data: MarkerData,
    todo_state: SummaryTodoState,
    todo_path: Path,
    repo_root: Path,
) -> str:
    """Render unresolved maintainer decisions from available machine-readable state."""
    todo_relative = repository_relative_path(todo_path, repo_root)
    unresolved: list[str] = []
    seen_unresolved: set[str] = set()

    for row in rows:
        if row.requires_maintainer_decision == "Yes" and row.decision != "manual TODO":
            append_unique_summary_item(
                unresolved,
                seen_unresolved,
                f"`{row.path}`: {row.decision}; {row.reason}",
            )

    for candidate in sorted(
        marker_data.deferred_candidates,
        key=lambda item: (item.path, item.source_commit),
    ):
        if not deferred_candidate_rendered_in_rows(candidate, rows):
            append_unique_summary_item(
                unresolved,
                seen_unresolved,
                (
                    f"`{candidate.path}`: deferred protected candidate from "
                    f"`{candidate.source_commit}`; {candidate.reason}"
                ),
            )

    if todo_state.exists and not todo_state.is_interpretable:
        append_unique_summary_item(
            unresolved,
            seen_unresolved,
            (
                f"`{todo_relative}` is present but not machine-interpretable; "
                "manual review is required."
            ),
        )
    elif todo_state.is_interpretable:
        for item in todo_state.unchecked_items:
            if item.section in TODO_DECISION_SECTIONS:
                append_unique_summary_item(
                    unresolved,
                    seen_unresolved,
                    (
                        f"`{todo_relative}` line {item.line_number}: "
                        + f"{item.section}: {item.text}"
                    ),
                )

    if unresolved:
        return "\n".join(
            [
                (
                    "- Unresolved maintainer decisions remain in available "
                    "machine-readable state: Yes."
                ),
                *(f"- {item}" for item in unresolved),
            ]
        )

    return "\n".join(
        [
            "- Unresolved maintainer decisions remain in available machine-readable state: No.",
            (
                "- No unresolved decisions were found in marker data, protected-file "
                "decisions, deferred protected candidates, local overrides, or "
                "machine-interpretable TODO state."
            ),
        ]
    )


def summary_validation_commands(marker_data: MarkerData) -> tuple[str, ...]:
    """Return validation commands for retained modules in summary order."""
    commands = [
        "python .template-sync/scripts/run_first_adoption_checks.py",
        "pre-commit run --all-files",
    ]
    for command in validation_commands_for_modules(marker_data.included_modules).split("<br>"):
        if command and command != "manual review" and command not in commands:
            commands.append(command)
    return tuple(commands)


def format_summary_validation_commands(marker_data: MarkerData) -> str:
    """Render retained-module validation commands for the concise summary."""
    return "\n".join(f"- `{command}`" for command in summary_validation_commands(marker_data))


def format_todo_summary_status(todo_state: SummaryTodoState) -> str:
    """Return a short machine-interpretability status for the TODO file."""
    if not todo_state.exists:
        return "not found"
    if not todo_state.is_interpretable:
        return "found; not machine-interpretable"
    return "found; machine-interpretable"


def format_adoption_summary(
    *,
    marker_path: Path,
    manifest_path: Path,
    todo_path: Path,
    repo_root: Path,
    marker_data: MarkerData,
    manifest_modules: frozenset[str],
    rows: tuple[LedgerRow, ...],
    todo_state: SummaryTodoState,
    default_adoption_mode: str,
) -> str:
    """Render a concise deterministic template adoption summary."""
    marker_relative = repository_relative_path(marker_path, repo_root)
    manifest_relative = repository_relative_path(manifest_path, repo_root)
    todo_relative = repository_relative_path(todo_path, repo_root)
    marker_status = "found" if marker_path.exists() else "not found"
    excluded_modules = manifest_modules - marker_data.included_modules

    return "\n".join(
        [
            "# Template Adoption Summary",
            "",
            (
                "Concise deterministic review artifact for adoption PR descriptions. "
                "`--ledger-only` remains the detailed path-level review artifact; "
                f"`{manifest_relative}` and `{marker_relative}` remain authoritative."
            ),
            "",
            f"- Marker: `{marker_relative}` ({marker_status})",
            f"- Manifest: `{manifest_relative}`",
            f"- First-adoption checklist: `{todo_relative}` ({format_todo_summary_status(todo_state)})",
            f"- Default adoption mode: `{default_adoption_mode}`",
            "",
            "## Included Modules",
            "",
            format_summary_module_list(marker_data.included_modules),
            "",
            "## Excluded Modules",
            "",
            format_summary_module_list(excluded_modules),
            "",
            "## Protected-File Decisions",
            "",
            format_summary_protected_decisions(marker_data),
            "",
            "## Local Overrides",
            "",
            format_summary_local_overrides(marker_data),
            "",
            "## Local Path Ownership",
            "",
            format_summary_local_path_ownership(marker_data),
            "",
            "## Unresolved Maintainer Decisions",
            "",
            format_summary_unresolved_decisions(
                rows=rows,
                marker_data=marker_data,
                todo_state=todo_state,
                todo_path=todo_path,
                repo_root=repo_root,
            ),
            "",
            "## Manual TODO Items",
            "",
            format_summary_manual_todo_items(
                todo_state=todo_state,
                todo_path=todo_path,
                repo_root=repo_root,
            ),
            "",
            "## Validation Commands",
            "",
            format_summary_validation_commands(marker_data),
        ]
    )


def format_limited_path_list(paths: tuple[str, ...], empty_text: str, *, limit: int = 8) -> str:
    """Return a compact Markdown list of repository paths."""
    if not paths:
        return empty_text
    visible = [f"`{path}`" for path in paths[:limit]]
    if len(paths) > limit:
        visible.append(f"{len(paths) - limit} more")
    return ", ".join(visible)


def redact_url_userinfo(value: str) -> str:
    """Redact URL user-info before displaying a URL-like value.

    Handles both ``scheme://userinfo@host`` and scheme-relative ``//userinfo@host``
    forms. SCP-style remotes such as ``git@github.com:org/repo`` have no URL
    authority and are returned unchanged.
    """
    try:
        parsed = urlsplit(value)
    except ValueError:
        # Unparseable input (for example, a malformed IPv6 authority): redact
        # defensively instead of risking a leak by returning the raw value.
        if "@" not in value:
            return value
        return f"***@{value.rsplit('@', maxsplit=1)[1]}"
    if not parsed.netloc or "@" not in parsed.netloc:
        return value
    safe_netloc = f"***@{parsed.netloc.rsplit('@', maxsplit=1)[1]}"
    return urlunsplit((parsed.scheme, safe_netloc, parsed.path, parsed.query, parsed.fragment))


def format_remotes(remotes: tuple[RemoteInfo, ...]) -> str:
    """Return a compact remote summary."""
    if not remotes:
        return "none found"
    return "; ".join(
        f"`{remote.name}` {remote.purpose}: `{redact_url_userinfo(remote.url)}`"
        for remote in remotes
    )


def format_working_tree(entries: tuple[str, ...]) -> str:
    """Return a compact working-tree summary."""
    if not entries:
        return "clean"
    visible = entries[:8]
    suffix = f"; {len(entries) - len(visible)} more" if len(entries) > len(visible) else ""
    return f"{len(entries)} changed or untracked path(s): " + "; ".join(visible) + suffix


def markdown_code_span(value: str) -> str:
    """Return ``value`` as a Markdown code span that is safe inside a GFM table cell.

    The fence is chosen one backtick longer than the longest backtick run in
    ``value`` so embedded backticks cannot terminate the span, and a space pads the
    span when ``value`` starts or ends with a backtick (per CommonMark). Pipes are
    escaped because a GFM table row is split on unescaped pipes before inline
    parsing, even inside code spans, and CR/LF are flattened to spaces so a value
    cannot break out of its row. Backslashes are left untouched: they are literal
    inside a code span, so doubling them would corrupt the displayed text.
    """
    text = value.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    text = text.replace("|", "\\|")
    longest_run = 0
    current_run = 0
    for char in text:
        if char == "`":
            current_run += 1
            longest_run = max(longest_run, current_run)
        else:
            current_run = 0
    fence = "`" * (longest_run + 1)
    pad = " " if text.startswith("`") or text.endswith("`") else ""
    return f"{fence}{pad}{text}{pad}{fence}"


def format_code_span_list(items: tuple[str, ...], empty_text: str, *, limit: int = 8) -> str:
    """Return a compact, backtick-safe Markdown list of code spans for table cells."""
    if not items:
        return empty_text
    visible = [markdown_code_span(item) for item in items[:limit]]
    if len(items) > limit:
        visible.append(f"{len(items) - limit} more")
    return ", ".join(visible)


def format_json_value(value: JsonValue, *, limit: int = 180) -> str:
    """Return a bounded JSON value rendering for table output."""
    rendered = json.dumps(value, sort_keys=True)
    if len(rendered) > limit:
        rendered = f"{rendered[: limit - 3]}..."
    return markdown_code_span(rendered)


def format_github_observation_value(observation: GitHubObservation) -> str:
    """Return a human-readable typed observation value."""
    if not observation.observed:
        return f"{markdown_code_span('null')} (not observed)"
    value = observation.value
    if isinstance(value, bool):
        state = "enabled" if value else "disabled"
        return f"{markdown_code_span(str(value).lower())} ({state}, observed)"
    if isinstance(value, str):
        return markdown_code_span(value)
    if isinstance(value, list):
        if not value:
            return "observed empty list"
        if all(isinstance(item, str) for item in value):
            return format_code_span_list(tuple(cast(list[str], value)), "observed empty list")
        return format_json_value(value)
    if isinstance(value, dict):
        if not value:
            return "observed empty object"
        return format_json_value(value)
    if value is None:
        return markdown_code_span("null")
    return format_json_value(value)


def format_github_observation_state(observation: GitHubObservation) -> str:
    """Return the observation state and completeness."""
    if not observation.observed:
        return f"unobserved ({markdown_code_span(observation.basis)})"
    completeness = "partial" if observation.partial else "complete"
    return f"observed {completeness} ({markdown_code_span(observation.basis)})"


def format_github_observation_evidence(observation: GitHubObservation) -> str:
    """Return compact, sanitized observation evidence."""
    details = [
        f"source: {markdown_code_span(observation.source)}",
        f"observed_at: {markdown_code_span(observation.observed_at)}",
        f"pagination: {markdown_code_span(observation.pagination)}",
    ]
    if observation.partial_fields:
        details.append(
            "partial fields: "
            + ", ".join(markdown_code_span(field) for field in observation.partial_fields)
        )
    details.extend(f"{key}: {markdown_code_span(value)}" for key, value in observation.diagnostics)
    return "; ".join(details)


def format_github_metadata(metadata: GitHubMetadata) -> str:
    """Return a Markdown observation ledger for optional GitHub metadata."""
    manual_settings = ", ".join(GITHUB_METADATA_MANUAL_SETTINGS)
    if not metadata.requested:
        return "\n".join(
            [
                (
                    "- GitHub metadata: not requested; run with "
                    "`--include-github-metadata` to opt in to read-only `gh` or "
                    "public REST lookups."
                ),
                f"- GitHub-only settings requiring manual review: {manual_settings}.",
            ]
        )

    lines = [
        f"- GitHub metadata source: `{metadata.source}`",
        f"- GitHub-only settings requiring manual review: {manual_settings}.",
        "",
        "| Setting | Value | Observation | Evidence |",
        "| --- | --- | --- | --- |",
    ]
    for observation in metadata.observations:
        label = GITHUB_METADATA_SETTING_LABELS.get(observation.setting, observation.setting)
        # The value/state/evidence renderers already emit table-cell-safe content via
        # markdown_code_span; the label is a controlled constant. So no separate
        # whole-cell escaping pass is needed (and a backslash-doubling one would
        # corrupt the code spans built above).
        lines.append(
            "| "
            + " | ".join(
                (
                    label,
                    format_github_observation_value(observation),
                    format_github_observation_state(observation),
                    format_github_observation_evidence(observation),
                )
            )
            + " |"
        )
    if metadata.error is not None:
        lines.append(f"- Partial metadata warning: {metadata.error}")
    return "\n".join(lines)


def format_marker_summary(marker_data: MarkerData, marker_present: bool) -> str:
    """Return a compact marker-state summary."""
    if not marker_present:
        return "not found; selected modules and protected-file decisions are unrecorded"
    included_modules = (
        ", ".join(f"`{module}`" for module in sorted(marker_data.included_modules)) or "none"
    )
    return (
        "found; included modules: "
        f"{included_modules}; local overrides: {len(marker_data.local_overrides)}; "
        f"local path ownership records: {len(marker_data.local_path_ownership)}; "
        f"deferred protected candidates: {len(marker_data.deferred_candidates)}; "
        f"protected-file decisions: {len(marker_data.protected_decisions)}"
    )


def format_module_selection_aid(
    manifest_modules: frozenset[str],
    marker_data: MarkerData,
) -> str:
    """Render a module-selection decision aid without recomputing path decisions."""
    lines = [
        "| Module | Current marker selection | First-adoption action |",
        "| --- | --- | --- |",
    ]
    for module in sorted(manifest_modules):
        selected = "selected" if module in marker_data.included_modules else "not selected"
        action = (
            "confirm retained module scope"
            if module in marker_data.included_modules
            else "decide whether this module is adopted"
        )
        lines.append(f"| {module} | {selected} | {action} |")
    return "\n".join(lines)


def format_protected_authorization_questions(rows: tuple[LedgerRow, ...]) -> str:
    """Render protected-file authorization questions from reused ledger rows."""
    protected_rows = tuple(row for row in rows if row.protected_file == "Yes")
    if not protected_rows:
        return "No protected instruction/governance paths were flagged by the adoption ledger."

    lines: list[str] = []
    for row in protected_rows:
        if row.requires_maintainer_decision == "Yes":
            lines.append(
                "- [ ] Does the maintainer authorize the selected action for "
                f"`{row.path}`? Ledger decision: `{row.decision}`; adoption mode: "
                f"`{row.adoption_mode}`; reason: {row.reason}"
            )
        else:
            lines.append(
                "- [ ] Verify the recorded protected-file authorization for "
                f"`{row.path}` remains valid. Ledger decision: `{row.decision}`; "
                f"adoption mode: `{row.adoption_mode}`."
            )
    return "\n".join(lines)


def format_structural_alignment_candidates(discovery: RepositoryDiscovery) -> str:
    """Render structural-alignment candidates discovered during preflight."""
    candidates: list[str] = []
    if not discovery.adoption_note_present:
        candidates.append(
            "Create or update `_TODO-repo-init.md` before finalizing files that depend on "
            "manual GitHub settings or maintainer policy."
        )
    if not discovery.marker_present:
        candidates.append(
            "Create `.template-sync/marker.yml` if future template sync support is retained; "
            "record selected modules, local overrides, and protected-file decisions there."
        )
    if discovery.workflows:
        candidates.append(
            "Confirm retained workflow files under `.github/workflows/` still point at the "
            "downstream repository's actual validation commands and paths."
        )
    if discovery.agent_instruction_files:
        candidates.append(
            "Classify protected instruction files before editing them and record any "
            "authorized protected-file decisions."
        )
    if "Terraform" in discovery.tooling_stacks:
        candidates.append(
            "Confirm Terraform module/test roots and `.tflint.hcl` paths match retained "
            "Terraform workflow and pre-commit hooks."
        )
    if "Python" in discovery.tooling_stacks:
        candidates.append(
            "Confirm Python source, test, and packaging paths match retained Python CI and "
            "pre-commit hooks, or document downstream path remaps."
        )
    if not candidates:
        return "No structural-alignment candidates were detected from local files."
    return "\n".join(f"- [ ] {candidate}" for candidate in candidates)


def validation_plan_suggestions(rows: tuple[LedgerRow, ...]) -> tuple[str, ...]:
    """Return validation suggestions from ledger rows and first-adoption helpers."""
    commands = [
        "python .template-sync/scripts/run_first_adoption_checks.py",
        "pre-commit run --all-files",
    ]
    for row in rows:
        for command in row.validation_commands.split("<br>"):
            if command and command != "manual review" and command not in commands:
                commands.append(command)
    return tuple(commands)


def format_validation_plan(rows: tuple[LedgerRow, ...]) -> str:
    """Render validation-plan suggestions for the preflight report."""
    return "\n".join(f"- [ ] `{command}`" for command in validation_plan_suggestions(rows))


def format_maintainer_questionnaire(discovery: RepositoryDiscovery) -> str:
    """Render maintainer questions for policy values that must not be guessed."""
    default_branch = discovery.likely_default_branch or "[recorded default branch]"
    return "\n".join(
        [
            "- [ ] Which manifest modules should this repository retain?",
            "- [ ] Which contact path should `CODE_OF_CONDUCT.md` publish?",
            "- [ ] Which vulnerability reporting channel should `SECURITY.md` publish?",
            "- [ ] Which CODEOWNERS owner or team should be used?",
            "- [ ] Should template issue labels be used, remapped, or removed?",
            "- [ ] Should GitHub Discussions be enabled or left disabled?",
            (
                "- [ ] Should private vulnerability reporting be enabled, left disabled, or marked "
                "not available?"
            ),
            f"- [ ] Should the default branch remain `{default_branch}`, be renamed, or be deferred?",
            (
                "- [ ] Should the default branch use a ruleset, classic branch protection, or no "
                "new protection?"
            ),
            "- [ ] Is there a GHES host override, or should `github.com` remain the expected host?",
            "- [ ] Are protected instruction-file edits authorized, deferred, or out of scope?",
            "- [ ] Are any protected instruction-file removals authorized?",
            "- [ ] Are any template-derived files explicitly approved for `tailored` adoption mode?",
            (
                "- [ ] Which structural findings are required for adoption and which become "
                "post-adoption issues?"
            ),
        ]
    )


def format_todo_starter(
    discovery: RepositoryDiscovery,
    marker_data: MarkerData,
) -> str:
    """Render starter content for ``_TODO-repo-init.md``."""
    owner_name = discovery.owner_name or "OWNER/REPO"
    default_branch = discovery.likely_default_branch or "[recorded default branch]"
    visibility = discovery.github_metadata.visibility or "[public, private, or internal]"
    marker_status = "found" if discovery.marker_present else "not found"
    adoption_note_status = "found" if discovery.adoption_note_present else "not found"
    included_modules = (
        ", ".join(sorted(marker_data.included_modules))
        if marker_data.included_modules
        else "[selected modules]"
    )

    return "\n".join(
        [
            "# Repository Initialization Checklist",
            "",
            (
                "This file records first-adoption decisions for this downstream repository. "
                "It is downstream-owned state, not an upstream template file."
            ),
            "",
            "## Discoverable Repository State",
            "",
            f"- [ ] Repository owner/name recorded: `{owner_name}`",
            f"- [ ] Repository visibility recorded: `{visibility}`",
            f"- [ ] Repository default branch recorded: `{default_branch}`",
            (
                "- [ ] Existing governance, security, CODEOWNERS, issue-template, PR-template, "
                "and Dependabot files reviewed before replacement."
            ),
            (
                "- [ ] Existing `.template-sync/marker.yml`, `_TODO-repo-init.md`, or "
                "equivalent adoption note checked before asking repeated questions "
                f"(marker: {marker_status}; checklist: {adoption_note_status})."
            ),
            f"- [ ] Included module candidates reviewed: `{included_modules}`",
            "",
            "## Manual GitHub Settings",
            "",
            (
                "- [ ] Private vulnerability reporting decision: "
                "`[enable, leave disabled, or not available]`"
            ),
            "- [ ] GitHub Discussions decision: `[enable or leave disabled]`",
            "- [ ] Labels decision, including `triage`: `[create, skip, or already present]`",
            (
                "- [ ] Default branch decision: "
                "`[keep recorded default branch, rename, or defer]`"
            ),
            (
                "- [ ] Repository ruleset or branch-protection decision: "
                "`[create ruleset, use classic branch protection, leave unchanged, or defer]`"
            ),
            "",
            "## Maintainer Policy Decisions",
            "",
            (
                "- [ ] Code of Conduct reporting contact method: "
                "`[explicit contact, profile contact, approved alternate, or deferred]`"
            ),
            (
                "- [ ] Security vulnerability reporting channel: "
                "`[private vulnerability reporting, monitored email, both, or deferred]`"
            ),
            "- [ ] CODEOWNERS owner/team identity: `[@user or @org/team]`",
            (
                "- [ ] Label-dependent issue-template behavior: "
                "`[use labels, remap labels, remove labels, or deferred]`"
            ),
            (
                "- [ ] Adoption mode for protected and template-derived files: "
                "`minimal-preservation` by default; list any `tailored` opt-ins."
            ),
            "- [ ] GHES host override: `[none or github.company.com]`",
            "",
            "## Protected-File Adoption Decisions",
            "",
            (
                "- [ ] Protected instruction files identified before editing: "
                "`.github/copilot-instructions.md`, `.github/instructions/*.instructions.md`, "
                "`.cursor/rules/*.mdc`, `.hermes.md`, `AGENTS.md`, `CLAUDE.md`, and "
                "`GEMINI.md`."
            ),
            "- [ ] Protected-file edits authorized by maintainer: `[none, path-scoped, or deferred]`",
            (
                "- [ ] Protected-file removals authorized by maintainer: "
                "`[none, path-scoped, or deferred]`"
            ),
            (
                "- [ ] `.template-sync/marker.yml` protected-file decisions updated if "
                "template sync support is retained."
            ),
            "",
            "## Unresolved Settings",
            "",
            (
                "- [ ] Items that must be completed later: Question: `[concrete owner "
                "question]`; state: `[not yet asked, asked and deferred, or unavailable "
                "through current safe tooling / manual review required]`; dependency: "
                "`[file, workflow, or GitHub setting]`; dependent-file status: "
                "`[not finalized, blocked, placeholder removed, local default withheld, "
                "or other concrete status]`"
            ),
            "",
            "## Resolution Evidence",
            "",
            "- [ ] No generated or adopted file ships unresolved placeholders or misleading defaults.",
            (
                "- [ ] Evidence captured for every resolved GitHub setting or policy decision: "
                "`[commit, PR, screenshot/link, connector result, API response note, or "
                "maintainer note]`"
            ),
            "- [ ] Deferred items copied into the adoption PR description, sync summary, or issue.",
        ]
    )


def format_issue_draft_skeletons() -> str:
    """Render copyable post-adoption issue draft skeletons."""
    return "\n".join(
        [
            "```markdown",
            "## Structural Follow-Up",
            "",
            "### Context",
            "",
            (
                "Template adoption is complete. This issue handles one deferred structural "
                "follow-up in this repository only."
            ),
            "",
            "### Scope",
            "",
            "- [Specific paths, commands, or workflow roots to change]",
            "- Protected instruction files in scope: [yes/no].",
            "",
            "### Non-Goals",
            "",
            "- Do not revisit template adoption decisions.",
            "- Do not change unrelated structure or formatting.",
            "",
            "### Acceptance Criteria",
            "",
            "- [Observable result]",
            "- Any touched file with `Last Updated` or `**Version:**` metadata is synchronized.",
            "",
            "### Validation",
            "",
            "- [Command or manual check]",
            "",
            "## Policy Follow-Up",
            "",
            "### Context",
            "",
            (
                "A first-adoption policy decision was deferred and must be resolved before the "
                "dependent file is final."
            ),
            "",
            "### Scope",
            "",
            "- Decision: [concrete maintainer question]",
            "- Dependent files/settings: [paths or GitHub setting]",
            "",
            "### Acceptance Criteria",
            "",
            (
                "- The decision is recorded in `_TODO-repo-init.md`, `.template-sync/marker.yml`, "
                "or an equivalent committed adoption note."
            ),
            "- Dependent files no longer contain unresolved placeholders or misleading defaults.",
            "",
            "### Validation",
            "",
            "- [Command or manual check]",
            "```",
        ]
    )


def format_preflight_report(
    *,
    marker_path: Path,
    manifest_path: Path,
    todo_path: Path,
    repo_root: Path,
    marker_data: MarkerData,
    manifest_modules: frozenset[str],
    discovery: RepositoryDiscovery,
    first_adoption_state: FirstAdoptionState,
    full_state: bool,
    ledger_rows: tuple[LedgerRow, ...],
    ledger_document: str,
) -> str:
    """Render the read-only first-adoption preflight report."""
    marker_relative = repository_relative_path(marker_path, repo_root)
    manifest_relative = repository_relative_path(manifest_path, repo_root)
    todo_relative = repository_relative_path(todo_path, repo_root)

    lines = [
        "# First-Adoption Preflight Report",
        "",
        (
            "Read-only planning artifact. This mode inspects local repository state, "
            "optionally reads GitHub metadata only when explicitly requested, and does "
            "not write files or modify repository settings."
        ),
        "",
        "## Repository Discovery",
        "",
        f"- Repository owner/name: `{discovery.owner_name or 'unknown'}`",
        f"- Remotes: {format_remotes(discovery.remotes)}",
        f"- Current branch: `{discovery.current_branch or 'unknown'}`",
        f"- Likely default branch: `{discovery.likely_default_branch or 'manual review required'}`",
        f"- Working tree: {format_working_tree(discovery.working_tree_entries)}",
        f"- Workflows: {format_limited_path_list(discovery.workflows, 'none found')}",
        (
            "- Issue templates: "
            f"{format_limited_path_list(discovery.issue_templates, 'none found')}"
        ),
        f"- PR templates: {format_limited_path_list(discovery.pr_templates, 'none found')}",
        (
            "- Security/conduct/CODEOWNERS files: "
            f"{format_limited_path_list(discovery.security_files, 'none found')}"
        ),
        (
            "- Agent instruction files: "
            f"{format_limited_path_list(discovery.agent_instruction_files, 'none found')}"
        ),
        (
            "- Language/tooling stacks inferred from local files: "
            + (", ".join(discovery.tooling_stacks) if discovery.tooling_stacks else "none found")
        ),
        f"- Marker: `{marker_relative}` ({format_marker_summary(marker_data, discovery.marker_present)})",
        (
            f"- First-adoption checklist: `{todo_relative}` "
            f"({'found' if discovery.adoption_note_present else 'not found'})"
        ),
        "",
        format_github_metadata(discovery.github_metadata),
        "",
        "## Raw First-Adoption State",
        "",
        format_first_adoption_state(first_adoption_state, full_state=full_state),
        "",
        "## Maintainer Questionnaire",
        "",
        format_maintainer_questionnaire(discovery),
        "",
        "## `_TODO-repo-init.md` Starter",
        "",
        "```markdown",
        format_todo_starter(discovery, marker_data),
        "```",
        "",
        "## Module-Selection Decision Aid",
        "",
        (
            "Use this table to choose marker `included_modules`. The adoption ledger "
            "below remains the reused path-level decision view."
        ),
        "",
        format_module_selection_aid(manifest_modules, marker_data),
        "",
        "## Protected-File Authorization Questions",
        "",
        format_protected_authorization_questions(ledger_rows),
        "",
        "## Required Structural Alignment Candidates",
        "",
        format_structural_alignment_candidates(discovery),
        "",
        "## Post-Adoption Issue Draft Skeletons",
        "",
        format_issue_draft_skeletons(),
        "",
        "## Validation Plan Suggestions",
        "",
        format_validation_plan(ledger_rows),
        "",
        "## Reused Adoption Ledger",
        "",
        (
            f"The following section is rendered by the existing adoption-ledger helpers from "
            f"`{manifest_relative}` and `{marker_relative}`."
        ),
        "",
        ledger_document,
    ]
    return "\n".join(lines)


def write_candidate_table(repo_root: Path, output_path: Path, candidate_table: str) -> None:
    """Write the rendered candidate table to a repository-contained path."""
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(candidate_table + "\n", encoding="utf-8")
    except OSError as error:
        output_relative = repository_relative_path(output_path, repo_root)
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        raise CandidateGenerationError(
            f"Unable to write candidate table to {output_relative}: {error_summary}"
        ) from error


def write_adoption_ledger(repo_root: Path, output_path: Path, ledger_document: str) -> None:
    """Write the rendered adoption ledger to a repository-contained path."""
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(ledger_document + "\n", encoding="utf-8")
    except OSError as error:
        output_relative = repository_relative_path(output_path, repo_root)
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        raise CandidateGenerationError(
            f"Unable to write adoption ledger to {output_relative}: {error_summary}"
        ) from error


def load_and_validate_inputs(
    repo_root: Path,
    marker_path: Path,
    manifest_path: Path,
    marker_schema_path: Path,
    manifest_schema_path: Path,
) -> tuple[MarkerData, frozenset[str], tuple[ManifestMapping, ...]]:
    """Load and schema-validate marker and manifest inputs."""
    marker = load_yaml_mapping(marker_path, repo_root)
    manifest = load_yaml_mapping(manifest_path, repo_root)
    marker_schema = load_json_mapping(marker_schema_path, repo_root)
    manifest_schema = load_json_mapping(manifest_schema_path, repo_root)

    validate_schema(marker, marker_schema, marker_path, repo_root)
    validate_schema(manifest, manifest_schema, manifest_path, repo_root)

    marker_data = parse_marker(marker)
    manifest_modules, mappings = parse_manifest(manifest)
    return marker_data, manifest_modules, mappings


def print_report(
    *,
    marker_path: Path,
    manifest_path: Path,
    repo_root: Path,
    range_base_ref: str,
    range_base_sha: str,
    range_base_source: str,
    range_head_ref: str,
    range_head_sha: str,
    manifest_modules: frozenset[str],
    marker_data: MarkerData,
    rows: tuple[CandidateRow, ...],
    candidate_table: str,
    procedure_warning: str | None,
    write_candidates_path: Path | None,
    write_ledger_path: Path | None,
    ledger_document: str | None,
) -> None:
    """Print the Markdown candidate table report."""
    marker_relative = repository_relative_path(marker_path, repo_root)
    manifest_relative = repository_relative_path(manifest_path, repo_root)
    unknown_marker_modules = marker_data.included_modules - manifest_modules

    print("# Template Sync Candidate Table")
    print()
    print(f"- Marker: `{marker_relative}`")
    print(f"- Manifest: `{manifest_relative}`")
    print(f"- Range base: `{range_base_sha}` from `{range_base_ref}` ({range_base_source})")
    print(f"- Range head: `{range_head_sha}` from `{range_head_ref}`")
    print(f"- Modeled diff command: `{modeled_diff_command(range_base_sha, range_head_sha)}`")
    if procedure_warning is not None:
        print(f"- {procedure_warning}")
    if write_candidates_path is not None:
        print(
            "- Saved candidate table: "
            f"`{repository_relative_path(write_candidates_path, repo_root)}`"
        )
    if write_ledger_path is not None:
        print(
            "- Saved adoption ledger: "
            f"`{repository_relative_path(write_ledger_path, repo_root)}`"
        )
    print(
        "- Included modules: "
        + (", ".join(f"`{module}`" for module in sorted(marker_data.included_modules)) or "none")
    )
    if unknown_marker_modules:
        print(
            "- Unknown marker modules: "
            + ", ".join(f"`{module}`" for module in sorted(unknown_marker_modules))
        )
    print()
    print(
        "This table is a decision aid only. The manual review process in "
        "`TEMPLATE_UPDATE_PROCEDURE.md` remains authoritative."
    )
    print()

    if not rows:
        print("No changed paths found in the reviewed range.")
        if ledger_document is not None:
            print()
            print(ledger_document)
        return

    print(candidate_table)
    if ledger_document is not None:
        print()
        print(ledger_document)


def print_ledger_only_report(
    *,
    ledger_document: str,
    write_ledger_path: Path | None,
    repo_root: Path,
) -> None:
    """Print the adoption-ledger-only report."""
    if write_ledger_path is not None:
        print(f"Saved adoption ledger: `{repository_relative_path(write_ledger_path, repo_root)}`")
        print()
    print(ledger_document)


def fail(message: str) -> NoReturn:
    """Print an error and exit non-zero."""
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(1)


def main(argv: list[str] | None = None) -> int:
    """Generate the sync candidate table."""
    args = parse_args(sys.argv[1:] if argv is None else argv)
    ledger_document: str | None = None
    try:
        if args.ledger_only and args.write_candidates is not None:
            raise CandidateGenerationError("--write-candidates cannot be used with --ledger-only.")
        if args.summary and args.write_candidates is not None:
            raise CandidateGenerationError("--write-candidates cannot be used with --summary.")
        if args.summary and args.write_ledger is not None:
            raise CandidateGenerationError("--write-ledger cannot be used with --summary.")
        if args.preflight and args.write_candidates is not None:
            raise CandidateGenerationError("--write-candidates cannot be used with --preflight.")
        if args.preflight and args.write_ledger is not None:
            raise CandidateGenerationError("--write-ledger cannot be used with --preflight.")
        if args.include_github_metadata and not args.preflight:
            raise CandidateGenerationError(
                "--include-github-metadata can only be used with --preflight."
            )
        if args.github_api_base is not None and not args.preflight:
            raise CandidateGenerationError("--github-api-base can only be used with --preflight.")
        if args.github_api_base is not None and not args.include_github_metadata:
            raise CandidateGenerationError(
                "--github-api-base can only be used with --include-github-metadata."
            )
        if args.full_state and not args.preflight:
            raise CandidateGenerationError("--full-state can only be used with --preflight.")

        repo_root = resolve_repo_root(args.repo_root)
        marker_path = resolve_repo_path(repo_root, args.marker)
        manifest_path = resolve_repo_path(repo_root, args.manifest)
        marker_schema_path = resolve_repo_path(repo_root, args.marker_schema)
        manifest_schema_path = resolve_repo_path(repo_root, args.manifest_schema)
        todo_path = resolve_repo_path(repo_root, args.todo_file)
        journal_path = resolve_repo_path(repo_root, DEFAULT_ADOPTION_JOURNAL_PATH)
        write_candidates_path = (
            resolve_repo_path(repo_root, args.write_candidates)
            if args.write_candidates is not None
            else None
        )
        write_ledger_path = (
            resolve_repo_path(repo_root, args.write_ledger)
            if args.write_ledger is not None
            else None
        )

        if args.preflight:
            marker_data, manifest_modules, mappings = load_preflight_inputs(
                repo_root=repo_root,
                marker_path=marker_path,
                manifest_path=manifest_path,
                marker_schema_path=marker_schema_path,
                manifest_schema_path=manifest_schema_path,
            )
            todo_items = load_todo_items(todo_path, repo_root)
            ledger_rows = build_adoption_ledger_rows(
                marker_data=marker_data,
                manifest_modules=manifest_modules,
                mappings=mappings,
                todo_items=todo_items,
                todo_path=todo_path,
                repo_root=repo_root,
                default_adoption_mode=args.adoption_mode,
            )
            ledger_document = format_adoption_ledger(
                marker_path=marker_path,
                manifest_path=manifest_path,
                todo_path=todo_path,
                repo_root=repo_root,
                marker_data=marker_data,
                rows=ledger_rows,
                default_adoption_mode=args.adoption_mode,
            )
            discovery = discover_repository_state(
                repo_root=repo_root,
                marker_path=marker_path,
                todo_path=todo_path,
                include_github_metadata=args.include_github_metadata,
                github_api_base=args.github_api_base,
            )
            first_adoption_state = inspect_first_adoption_state(
                repo_root=repo_root,
                marker_path=marker_path,
                todo_path=todo_path,
                journal_path=journal_path,
            )
            print(
                format_preflight_report(
                    marker_path=marker_path,
                    manifest_path=manifest_path,
                    todo_path=todo_path,
                    repo_root=repo_root,
                    marker_data=marker_data,
                    manifest_modules=manifest_modules,
                    discovery=discovery,
                    first_adoption_state=first_adoption_state,
                    full_state=args.full_state,
                    ledger_rows=ledger_rows,
                    ledger_document=ledger_document,
                )
            )
            return 0

        if args.summary:
            if marker_path.exists():
                marker_data, manifest_modules, mappings = load_and_validate_inputs(
                    repo_root=repo_root,
                    marker_path=marker_path,
                    manifest_path=manifest_path,
                    marker_schema_path=marker_schema_path,
                    manifest_schema_path=manifest_schema_path,
                )
            else:
                marker_data, manifest_modules, mappings = load_preflight_inputs(
                    repo_root=repo_root,
                    marker_path=marker_path,
                    manifest_path=manifest_path,
                    marker_schema_path=marker_schema_path,
                    manifest_schema_path=manifest_schema_path,
                )
            todo_state = load_summary_todo_state(todo_path, repo_root)
            summary_rows = build_adoption_ledger_rows(
                marker_data=marker_data,
                manifest_modules=manifest_modules,
                mappings=mappings,
                todo_items=(),
                todo_path=todo_path,
                repo_root=repo_root,
                default_adoption_mode=args.adoption_mode,
            )
            print(
                format_adoption_summary(
                    marker_path=marker_path,
                    manifest_path=manifest_path,
                    todo_path=todo_path,
                    repo_root=repo_root,
                    marker_data=marker_data,
                    manifest_modules=manifest_modules,
                    rows=summary_rows,
                    todo_state=todo_state,
                    default_adoption_mode=args.adoption_mode,
                )
            )
            return 0

        if args.ledger_only and not marker_path.exists():
            marker_data, manifest_modules, mappings = load_preflight_inputs(
                repo_root=repo_root,
                marker_path=marker_path,
                manifest_path=manifest_path,
                marker_schema_path=marker_schema_path,
                manifest_schema_path=manifest_schema_path,
            )
        else:
            marker_data, manifest_modules, mappings = load_and_validate_inputs(
                repo_root=repo_root,
                marker_path=marker_path,
                manifest_path=manifest_path,
                marker_schema_path=marker_schema_path,
                manifest_schema_path=manifest_schema_path,
            )
        ledger_document = None
        if args.ledger or args.ledger_only or write_ledger_path is not None:
            todo_items = load_todo_items(todo_path, repo_root)
            ledger_output_dir = write_ledger_path.parent if write_ledger_path is not None else None
            ledger_rows = build_adoption_ledger_rows(
                marker_data=marker_data,
                manifest_modules=manifest_modules,
                mappings=mappings,
                todo_items=todo_items,
                todo_path=todo_path,
                repo_root=repo_root,
                default_adoption_mode=args.adoption_mode,
                ledger_output_dir=ledger_output_dir,
            )
            ledger_document = format_adoption_ledger(
                marker_path=marker_path,
                manifest_path=manifest_path,
                todo_path=todo_path,
                repo_root=repo_root,
                marker_data=marker_data,
                rows=ledger_rows,
                default_adoption_mode=args.adoption_mode,
            )

        if args.ledger_only:
            if ledger_document is None:
                raise CandidateGenerationError("Unable to generate adoption ledger.")
            if write_ledger_path is not None:
                write_adoption_ledger(repo_root, write_ledger_path, ledger_document)
            print_ledger_only_report(
                ledger_document=ledger_document,
                write_ledger_path=write_ledger_path,
                repo_root=repo_root,
            )
            return 0

        range_base_ref, range_base_sha, range_base_source = resolve_range_base_ref(
            repo_root,
            args.range_base,
            marker_data,
        )
        range_head_ref, range_head_sha = resolve_range_head_ref(repo_root, args.range_head)
        verify_reachable_range(repo_root, range_base_sha, range_head_sha)

        entries = changed_paths(repo_root, range_base_sha, range_head_sha)
        rows = tuple(build_candidate_row(entry, marker_data, mappings) for entry in entries)
        candidate_table = format_table(rows)
        procedure_warning = stale_procedure_warning(repo_root, range_head_sha)
        if write_candidates_path is not None:
            write_candidate_table(repo_root, write_candidates_path, candidate_table)
        if write_ledger_path is not None:
            if ledger_document is None:
                raise CandidateGenerationError("Unable to generate adoption ledger.")
            write_adoption_ledger(repo_root, write_ledger_path, ledger_document)
    except (CandidateGenerationError, TemplateSyncMaterializationError) as error:
        fail(str(error))

    print_report(
        marker_path=marker_path,
        manifest_path=manifest_path,
        repo_root=repo_root,
        range_base_ref=range_base_ref,
        range_base_sha=range_base_sha,
        range_base_source=range_base_source,
        range_head_ref=range_head_ref,
        range_head_sha=range_head_sha,
        manifest_modules=manifest_modules,
        marker_data=marker_data,
        rows=rows,
        candidate_table=candidate_table,
        procedure_warning=procedure_warning,
        write_candidates_path=write_candidates_path,
        write_ledger_path=write_ledger_path,
        ledger_document=ledger_document if args.ledger else None,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
