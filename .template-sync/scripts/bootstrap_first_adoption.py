"""Bootstrap durable first-adoption state before template pruning."""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, NoReturn, TextIO, TypedDict, cast

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import first_adoption_state  # noqa: E402
import generate_sync_candidates as sync_candidates  # noqa: E402
import initialize_adoption_journal  # noqa: E402
import run_first_adoption_checks as first_adoption_checks  # noqa: E402
from template_sync_materialization_helpers import (  # noqa: E402
    DEFAULT_MANIFEST_PATH,
    DEFAULT_MANIFEST_SCHEMA_PATH,
    DEFAULT_MARKER_PATH,
    DEFAULT_MARKER_SCHEMA_PATH,
    TemplateSyncMaterializationError,
    format_marker_yaml,
    is_string_list,
    os_error_summary,
    repository_relative_path,
    resolve_safe_repository_target_path,
    validate_marker_yaml_text,
    write_repository_relative_file_bytes,
)

BOOTSTRAP_STATE_SCHEMA_VERSION = 1
BOOTSTRAP_STATE_BEGIN = "<!-- first-adoption-bootstrap-state: begin v1"
BOOTSTRAP_STATE_END = "first-adoption-bootstrap-state: end -->"
DEFAULT_SOURCE_REPO = "https://github.com/franklesniak/copilot-repo-template.git"
DEFAULT_ADOPTION_MODE = "minimal-preservation"
RESOLVED_STATUS = "resolved"
DEFERRED_STATUS = "asked_and_deferred"
UNRESOLVED_STATUSES = frozenset(
    {
        "not_yet_asked",
        DEFERRED_STATUS,
        "unavailable_manual_review_required",
    }
)
STATUS_DISPLAY = {
    "not_yet_asked": "not yet asked",
    DEFERRED_STATUS: "asked and deferred",
    "unavailable_manual_review_required": (
        "unavailable through current safe tooling / manual review required"
    ),
    RESOLVED_STATUS: "resolved",
}


class OwnerDecisionTemplate(TypedDict):
    """Static template describing one owner decision recorded before pruning."""

    key: str
    category: str
    question: str
    dependent_files: tuple[str, ...]


OWNER_DECISION_TEMPLATES: tuple[OwnerDecisionTemplate, ...] = (
    {
        "key": "manual.github.private_vulnerability_reporting",
        "category": "manual_github_setting",
        "question": (
            "Should private vulnerability reporting be enabled, left disabled, "
            "or marked not available?"
        ),
        "dependent_files": ("SECURITY.md",),
    },
    {
        "key": "manual.github.discussions",
        "category": "manual_github_setting",
        "question": "Should GitHub Discussions be enabled or left disabled?",
        "dependent_files": (".github/ISSUE_TEMPLATE/**",),
    },
    {
        "key": "manual.github.labels",
        "category": "manual_github_setting",
        "question": "Should template issue labels be used, remapped, or removed?",
        "dependent_files": (".github/ISSUE_TEMPLATE/**",),
    },
    {
        "key": "manual.github.default_branch",
        "category": "manual_github_setting",
        "question": "Should the recorded default branch be kept, renamed, or deferred?",
        "dependent_files": (".github/workflows/**", "README.md"),
    },
    {
        "key": "manual.github.branch_protection",
        "category": "manual_github_setting",
        "question": (
            "Should the default branch use a ruleset, classic branch protection, "
            "or no new protection?"
        ),
        "dependent_files": (".github/workflows/**",),
    },
    {
        "key": "policy.conduct_contact",
        "category": "maintainer_policy",
        "question": "Which contact path should CODE_OF_CONDUCT.md publish?",
        "dependent_files": ("CODE_OF_CONDUCT.md",),
    },
    {
        "key": "policy.security_reporting",
        "category": "maintainer_policy",
        "question": "Which vulnerability reporting channel should SECURITY.md publish?",
        "dependent_files": ("SECURITY.md",),
    },
    {
        "key": "policy.codeowners",
        "category": "maintainer_policy",
        "question": "Which CODEOWNERS owner or team should be used?",
        "dependent_files": (".github/CODEOWNERS",),
    },
    {
        "key": "policy.adoption_mode",
        "category": "maintainer_policy",
        "question": (
            "Are protected and template-derived files staying in minimal-preservation "
            "mode, or are path-scoped tailored opt-ins approved?"
        ),
        "dependent_files": ("template-derived governance and workflow files",),
    },
    {
        "key": "policy.ghes_host",
        "category": "maintainer_policy",
        "question": "Is there a GHES host override, or should github.com remain expected?",
        "dependent_files": ("repository URLs and GitHub API calls",),
    },
    {
        "key": "template_sync.included_modules",
        "category": "template_sync_state",
        "question": "Which manifest modules should this repository retain?",
        "dependent_files": (DEFAULT_MARKER_PATH,),
    },
    {
        "key": "structural.required_alignment",
        "category": "structural_alignment",
        "question": (
            "Which structural findings are required for adoption and which become "
            "post-adoption issues?"
        ),
        "dependent_files": ("selected module paths and validation commands",),
    },
)
PROTECTED_QUESTION_PATH_PATTERN = re.compile(r"`(?P<path>[^`]+)`")
QUESTION_KEY_PHRASES = (
    ("Which manifest modules", "template_sync.included_modules"),
    ("Which contact path", "policy.conduct_contact"),
    ("Which vulnerability reporting channel", "policy.security_reporting"),
    ("Which CODEOWNERS owner", "policy.codeowners"),
    ("Should template issue labels", "manual.github.labels"),
    ("Should GitHub Discussions", "manual.github.discussions"),
    ("Should private vulnerability reporting", "manual.github.private_vulnerability_reporting"),
    ("Should the default branch remain", "manual.github.default_branch"),
    ("Should the default branch use", "manual.github.branch_protection"),
    ("Is there a GHES host override", "policy.ghes_host"),
    ("Are any template-derived files explicitly approved", "policy.adoption_mode"),
    ("Which structural findings", "structural.required_alignment"),
)


class FirstAdoptionBootstrapError(RuntimeError):
    """Raised when the bootstrap command cannot safely complete."""


@dataclass(frozen=True)
class ExistingBootstrapState:
    """Parsed state-block details from an existing checklist."""

    state: dict[str, Any] | None
    start: int | None
    end: int | None


@dataclass(frozen=True)
class BootstrapInputs:
    """Repository facts and helper-rendered state used by the bootstrap report."""

    marker_path: Path
    manifest_path: Path
    marker_schema_path: Path
    manifest_schema_path: Path
    todo_path: Path
    journal_path: Path
    repo_root: Path
    marker_data: sync_candidates.MarkerData
    manifest_modules: frozenset[str]
    ledger_rows: tuple[sync_candidates.LedgerRow, ...]
    ledger_document: str
    discovery: sync_candidates.RepositoryDiscovery
    raw_state: first_adoption_state.FirstAdoptionState


@dataclass(frozen=True)
class BootstrapWriteResult:
    """Filesystem effects from an explicit bootstrap write run."""

    todo_status: str
    journal_status: str
    marker_status: str


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Compose first-adoption preflight, journal, raw-state, marker, and "
            "validation-plan helpers into one safe bootstrap report."
        )
    )
    parser.add_argument(
        "--repo-root",
        default=None,
        help=(
            "Repository root. Defaults to the parent of the .template-sync directory "
            "that contains this script."
        ),
    )
    parser.add_argument(
        "--write",
        action="store_true",
        help=(
            "Create missing bootstrap-owned files. Existing checklist, journal, and "
            "marker content is preserved unless a narrower update flag is also used."
        ),
    )
    parser.add_argument(
        "--update-existing-todo-state",
        action="store_true",
        help=(
            "With --write, insert or refresh only the bootstrap-owned structured state "
            "block in an existing _TODO-repo-init.md while preserving Markdown notes "
            "outside the block."
        ),
    )
    parser.add_argument(
        "--write-draft-marker",
        action="store_true",
        help=(
            "With --write, create a missing draft .template-sync/marker.yml after "
            "schema and protected-decision integrity validation. Existing marker "
            "content is never overwritten."
        ),
    )
    parser.add_argument(
        "--draft-marker",
        choices=("auto", "always", "never"),
        default="auto",
        help=(
            "Control whether the report prints a draft marker. auto prints it when an "
            "existing marker retains template-sync support, or when no marker exists yet "
            "but a manifest is present. Default: auto."
        ),
    )
    parser.add_argument(
        "--source-repo",
        default=DEFAULT_SOURCE_REPO,
        help=f"Source repository URL for a new draft marker. Default: {DEFAULT_SOURCE_REPO}",
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
        "--todo-file",
        default=sync_candidates.DEFAULT_TODO_PATH,
        help=(
            "First-adoption checklist path relative to the repository root. "
            f"Default: {sync_candidates.DEFAULT_TODO_PATH}"
        ),
    )
    parser.add_argument(
        "--journal-path",
        default=first_adoption_state.DEFAULT_ADOPTION_JOURNAL_PATH,
        help=(
            "Adoption difficulties journal path relative to the repository root. "
            f"Default: {first_adoption_state.DEFAULT_ADOPTION_JOURNAL_PATH}"
        ),
    )
    parser.add_argument(
        "--adoption-mode",
        choices=("minimal-preservation", "tailored"),
        default=DEFAULT_ADOPTION_MODE,
        help="Default adoption mode to show in the generated ledger.",
    )
    parser.add_argument(
        "--include-github-metadata",
        action="store_true",
        help=(
            "Opt in to the same read-only GitHub metadata discovery used by "
            "generate_sync_candidates.py --preflight."
        ),
    )
    parser.add_argument(
        "--github-api-base",
        default=None,
        metavar="URL",
        help="GitHub REST API base URL for --include-github-metadata.",
    )
    parser.add_argument(
        "--full-state",
        action="store_true",
        help="List every raw first-adoption state entry instead of bounded samples.",
    )
    parser.add_argument(
        "--max-command-length",
        type=int,
        default=first_adoption_checks.DEFAULT_MAX_COMMAND_LENGTH,
        help=(
            "Maximum formatted command length to use when surfacing the runnable "
            "first-adoption validation plan."
        ),
    )
    return parser.parse_args(argv)


def state_block_text(state: dict[str, Any]) -> str:
    """Return the hidden, machine-readable bootstrap state block."""
    state_text = json.dumps(state, indent=2, sort_keys=True)
    return f"{BOOTSTRAP_STATE_BEGIN}\n{state_text}\n{BOOTSTRAP_STATE_END}"


def find_existing_state(text: str) -> ExistingBootstrapState:
    """Return the parsed bootstrap-owned state block from ``text`` when present."""
    begin_index = text.find(BOOTSTRAP_STATE_BEGIN)
    if begin_index == -1:
        return ExistingBootstrapState(state=None, start=None, end=None)

    json_start = text.find("\n", begin_index)
    if json_start == -1:
        raise FirstAdoptionBootstrapError("Bootstrap state block begin marker is incomplete.")
    json_start += 1
    # Bound the end-marker search to the current block (up to the next begin marker, or
    # end of text) and take the last end marker within it, so an end-marker string copied
    # into a JSON value cannot truncate parsing early.
    next_begin = text.find(BOOTSTRAP_STATE_BEGIN, json_start)
    block_limit = next_begin if next_begin != -1 else len(text)
    end_index = text.rfind(BOOTSTRAP_STATE_END, json_start, block_limit)
    if end_index == -1:
        raise FirstAdoptionBootstrapError("Bootstrap state block end marker is missing.")

    raw_json = text[json_start:end_index].strip()
    try:
        parsed = json.loads(raw_json)
    except json.JSONDecodeError as error:
        raise FirstAdoptionBootstrapError(
            f"Bootstrap state block is not valid JSON: {error}"
        ) from error
    if not isinstance(parsed, dict):
        raise FirstAdoptionBootstrapError("Bootstrap state block must contain a JSON object.")
    parsed = cast(dict[str, Any], parsed)
    if parsed.get("schema_version") != BOOTSTRAP_STATE_SCHEMA_VERSION:
        raise FirstAdoptionBootstrapError(
            f"Unsupported bootstrap state schema_version: {parsed.get('schema_version')!r}"
        )
    block_end = end_index + len(BOOTSTRAP_STATE_END)
    return ExistingBootstrapState(state=parsed, start=begin_index, end=block_end)


def normalize_decision(raw_decision: dict[str, Any]) -> dict[str, Any]:
    """Return one state decision with stable required keys."""
    key = raw_decision.get("key")
    category = raw_decision.get("category")
    question = raw_decision.get("question")
    if not isinstance(key, str) or not key:
        raise FirstAdoptionBootstrapError("Every bootstrap decision must define a string key.")
    if not isinstance(category, str) or not category:
        raise FirstAdoptionBootstrapError(
            f"Bootstrap decision {key} must define a string category."
        )
    if not isinstance(question, str) or not question:
        raise FirstAdoptionBootstrapError(
            f"Bootstrap decision {key} must define a string question."
        )

    status = raw_decision.get("status", "not_yet_asked")
    if status not in UNRESOLVED_STATUSES and status != RESOLVED_STATUS:
        raise FirstAdoptionBootstrapError(
            f"Bootstrap decision {key} has unsupported status: {status!r}"
        )

    raw_dependent_files = raw_decision.get("dependent_files", [])
    if not is_string_list(raw_dependent_files):
        raise FirstAdoptionBootstrapError(
            f"Bootstrap decision {key} dependent_files must be a string list."
        )

    return {
        "key": key,
        "category": category,
        "question": question,
        "status": status,
        "answer": raw_decision.get("answer", ""),
        "evidence": raw_decision.get("evidence", ""),
        "owner": raw_decision.get("owner", ""),
        "updated_utc": raw_decision.get("updated_utc", ""),
        "dependent_files": raw_dependent_files,
        "dependent_file_status": raw_decision.get("dependent_file_status", "not_finalized"),
        "notes": raw_decision.get("notes", ""),
    }


def decision_by_key(state: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Return normalized decisions keyed by stable decision key."""
    raw_decisions = state.get("decisions", [])
    if not isinstance(raw_decisions, list):
        raise FirstAdoptionBootstrapError("Bootstrap state decisions must be a list.")

    decisions: dict[str, dict[str, Any]] = {}
    for raw_decision in cast(list[object], raw_decisions):
        if not isinstance(raw_decision, dict):
            raise FirstAdoptionBootstrapError("Every bootstrap decision must be a mapping.")
        decision = normalize_decision(cast(dict[str, Any], raw_decision))
        key = decision["key"]
        if key in decisions:
            raise FirstAdoptionBootstrapError(f"Duplicate bootstrap decision key: {key}")
        decisions[key] = decision
    return decisions


def protected_decision_key(path: str) -> str:
    """Return the stable state key for a protected-file authorization question."""
    return f"protected_file_authorization:{path}"


def protected_marker_resolved_paths(
    marker_data: sync_candidates.MarkerData,
) -> frozenset[str]:
    """Return protected paths with complete marker authorization state."""
    resolved_paths: set[str] = set()
    for protected_decision in marker_data.protected_decisions:
        if (
            sync_candidates.protected_decision_authorization_status(protected_decision)
            == "authorized"
        ):
            resolved_paths.add(protected_decision.path)
    return frozenset(resolved_paths)


def default_decision(
    *,
    key: str,
    category: str,
    question: str,
    dependent_files: tuple[str, ...],
    dependent_file_status: str = "not_finalized",
) -> dict[str, Any]:
    """Return one unresolved decision record with durable answer/evidence fields."""
    return {
        "key": key,
        "category": category,
        "question": question,
        "status": "not_yet_asked",
        "answer": "",
        "evidence": "",
        "owner": "",
        "updated_utc": "",
        "dependent_files": list(dependent_files),
        "dependent_file_status": dependent_file_status,
        "notes": "",
    }


def default_bootstrap_state(
    *,
    ledger_rows: tuple[sync_candidates.LedgerRow, ...],
    marker_data: sync_candidates.MarkerData,
) -> dict[str, Any]:
    """Return the bootstrap-owned durable state contract for a first run."""
    decisions: list[dict[str, Any]] = []
    for template in OWNER_DECISION_TEMPLATES:
        decision = default_decision(
            key=template["key"],
            category=template["category"],
            question=template["question"],
            dependent_files=tuple(template["dependent_files"]),
        )
        if template["key"] == "template_sync.included_modules" and marker_data.included_modules:
            decision["status"] = RESOLVED_STATUS
            decision["answer"] = "Recorded modules: " + ", ".join(
                sorted(marker_data.included_modules)
            )
            decision["evidence"] = DEFAULT_MARKER_PATH
            decision["dependent_file_status"] = "recorded"
        decisions.append(decision)

    marker_resolved_paths = protected_marker_resolved_paths(marker_data)
    for row in ledger_rows:
        if row.protected_file != "Yes":
            continue
        status = RESOLVED_STATUS if row.path in marker_resolved_paths else "not_yet_asked"
        decision = default_decision(
            key=protected_decision_key(row.path),
            category="protected_file_authorization",
            question=(
                "Does the maintainer authorize the selected action for "
                f"{row.path}? Ledger decision: {row.decision}; adoption mode: "
                f"{row.adoption_mode}; reason: {row.reason}"
            ),
            dependent_files=(row.path,),
        )
        decision["status"] = status
        if status == RESOLVED_STATUS:
            decision["answer"] = "Recorded in template_sync.protected_file_decisions."
            decision["evidence"] = DEFAULT_MARKER_PATH
            decision["dependent_file_status"] = "authorized"
        decisions.append(decision)

    return {
        "schema_version": BOOTSTRAP_STATE_SCHEMA_VERSION,
        "state_owner": "bootstrap_first_adoption.py",
        "contract": {
            "description": (
                "Tools parse only this JSON block between exact bootstrap delimiters. "
                "Markdown checkboxes and prose outside the block are not structured state."
            ),
            "resolved_status": RESOLVED_STATUS,
            "unresolved_statuses": sorted(UNRESOLVED_STATUSES),
            "required_decision_fields": [
                "key",
                "category",
                "question",
                "status",
                "answer",
                "evidence",
                "owner",
                "updated_utc",
                "dependent_files",
                "dependent_file_status",
                "notes",
            ],
        },
        "decisions": decisions,
    }


def merge_bootstrap_state(
    default_state: dict[str, Any],
    existing_state: dict[str, Any] | None,
) -> dict[str, Any]:
    """Merge existing answers into the current default state contract."""
    if existing_state is None:
        return default_state

    default_decisions = decision_by_key(default_state)
    existing_decisions = decision_by_key(existing_state)
    merged_decisions: list[dict[str, Any]] = []
    for key, default_decision_record in default_decisions.items():
        existing_decision = existing_decisions.pop(key, None)
        if existing_decision is None:
            merged_decisions.append(default_decision_record)
            continue
        merged = dict(default_decision_record)
        for field_name in (
            "status",
            "answer",
            "evidence",
            "owner",
            "updated_utc",
            "dependent_file_status",
            "notes",
        ):
            merged[field_name] = existing_decision.get(field_name, merged[field_name])
        merged["dependent_files"] = existing_decision.get(
            "dependent_files",
            merged["dependent_files"],
        )
        merged_decisions.append(normalize_decision(merged))

    for leftover_key in sorted(existing_decisions):
        merged_decisions.append(existing_decisions[leftover_key])

    merged_state = dict(default_state)
    merged_state["decisions"] = merged_decisions
    return merged_state


def replace_or_append_state_block(text: str, state: dict[str, Any]) -> str:
    """Insert or replace only the bootstrap-owned hidden state block."""
    existing = find_existing_state(text)
    block = state_block_text(state)
    if existing.start is None or existing.end is None:
        # Trim only trailing newlines so rendering-significant trailing spaces in
        # free-form notes survive; plain rstrip() would also strip those spaces.
        body = text.rstrip("\n")
        return f"{body}\n\n{bootstrap_contract_section(state)}\n"
    return text[: existing.start] + block + text[existing.end :]


def bootstrap_contract_section(state: dict[str, Any]) -> str:
    """Return the visible TODO contract note plus hidden structured state."""
    return "\n".join(
        [
            "## Bootstrap Durable State Contract",
            "",
            (
                "The hidden JSON block below is the bootstrap-owned durable state. "
                "Later tooling parses only the exact "
                "`first-adoption-bootstrap-state` delimiters and the stable `key`, "
                "`status`, `answer`, and `evidence` fields. Free-form Markdown notes "
                "outside the block are preserved and are not used to suppress questions."
            ),
            "",
            (
                "Use `status: resolved` only after the maintainer answer and evidence "
                "are recorded. Use `not_yet_asked`, `asked_and_deferred`, or "
                "`unavailable_manual_review_required` for unresolved or blocked "
                "decisions, and keep dependent files unfinalized until their state "
                "allows finalization."
            ),
            "",
            state_block_text(state),
        ]
    )


def format_bootstrap_todo(
    discovery: sync_candidates.RepositoryDiscovery,
    marker_data: sync_candidates.MarkerData,
    state: dict[str, Any],
) -> str:
    """Return the proposed ``_TODO-repo-init.md`` with structured bootstrap state."""
    return "\n\n".join(
        [
            sync_candidates.format_todo_starter(discovery, marker_data),
            bootstrap_contract_section(state),
        ]
    )


def read_existing_todo_state(todo_path: Path) -> tuple[str | None, ExistingBootstrapState]:
    """Return existing checklist text and parsed bootstrap state details."""
    # Reject symlinks before the existence check: Path.exists() follows symlinks,
    # so a broken symlink would otherwise read as "missing" instead of being flagged
    # (write mode refuses symlinks at this path), matching write_todo_if_requested.
    if todo_path.is_symlink():
        raise FirstAdoptionBootstrapError(
            f"First-adoption checklist `{todo_path.name}` is a symlink; "
            "replace it with a regular file or remove it before bootstrapping."
        )
    if not todo_path.exists():
        return None, ExistingBootstrapState(state=None, start=None, end=None)
    if not todo_path.is_file():
        raise FirstAdoptionBootstrapError(
            f"First-adoption checklist exists but is not a regular file: {todo_path.name}"
        )
    try:
        text = todo_path.read_text(encoding="utf-8")
    except OSError as error:
        raise FirstAdoptionBootstrapError(
            f"Unable to read {todo_path.name}: {os_error_summary(error)}"
        ) from error
    return text, find_existing_state(text)


def load_bootstrap_inputs(
    *,
    repo_root: Path,
    marker_path: Path,
    manifest_path: Path,
    marker_schema_path: Path,
    manifest_schema_path: Path,
    todo_path: Path,
    journal_path: Path,
    adoption_mode: str,
    include_github_metadata: bool,
    github_api_base: str | None,
) -> BootstrapInputs:
    """Load existing helper state for the bootstrap report."""
    marker_data, manifest_modules, mappings = sync_candidates.load_preflight_inputs(
        repo_root=repo_root,
        marker_path=marker_path,
        manifest_path=manifest_path,
        marker_schema_path=marker_schema_path,
        manifest_schema_path=manifest_schema_path,
    )
    todo_items = sync_candidates.load_todo_items(todo_path, repo_root)
    ledger_rows = sync_candidates.build_adoption_ledger_rows(
        marker_data=marker_data,
        manifest_modules=manifest_modules,
        mappings=mappings,
        todo_items=todo_items,
        todo_path=todo_path,
        repo_root=repo_root,
        default_adoption_mode=adoption_mode,
    )
    ledger_document = sync_candidates.format_adoption_ledger(
        marker_path=marker_path,
        manifest_path=manifest_path,
        todo_path=todo_path,
        repo_root=repo_root,
        marker_data=marker_data,
        rows=ledger_rows,
        default_adoption_mode=adoption_mode,
    )
    discovery = sync_candidates.discover_repository_state(
        repo_root=repo_root,
        marker_path=marker_path,
        todo_path=todo_path,
        include_github_metadata=include_github_metadata,
        github_api_base=github_api_base,
    )
    raw_state = first_adoption_state.inspect_first_adoption_state(
        repo_root=repo_root,
        marker_path=marker_path,
        todo_path=todo_path,
        journal_path=journal_path,
    )
    return BootstrapInputs(
        marker_path=marker_path,
        manifest_path=manifest_path,
        marker_schema_path=marker_schema_path,
        manifest_schema_path=manifest_schema_path,
        todo_path=todo_path,
        journal_path=journal_path,
        repo_root=repo_root,
        marker_data=marker_data,
        manifest_modules=manifest_modules,
        ledger_rows=ledger_rows,
        ledger_document=ledger_document,
        discovery=discovery,
        raw_state=raw_state,
    )


def draft_marker_document(
    *,
    source_repo: str,
    manifest_modules: frozenset[str],
    existing_marker_data: sync_candidates.MarkerData,
) -> dict[str, Any]:
    """Return a schema-valid draft marker document for pre-pruning state."""
    included_modules = (
        existing_marker_data.included_modules
        if existing_marker_data.included_modules
        else manifest_modules
    )
    return {
        "template_sync": {
            "source_repo": source_repo,
            "included_modules": sorted(included_modules),
        }
    }


def validate_marker_document(
    marker_document: dict[str, Any],
    *,
    marker_path: Path,
    marker_schema_path: Path,
    repo_root: Path,
) -> None:
    """Validate the formatted marker schema and protected-decision integrity."""
    validate_marker_yaml_text(
        format_marker_yaml(marker_document),
        repo_root=repo_root,
        marker_path=repository_relative_path(marker_path, repo_root),
        marker_schema_path=repository_relative_path(marker_schema_path, repo_root),
    )


def should_print_draft_marker(
    *,
    mode: str,
    marker_path: Path,
    marker_data: sync_candidates.MarkerData,
    manifest_path: Path,
) -> bool:
    """Return whether the report should include a draft marker preview."""
    if mode == "always":
        return True
    if mode == "never":
        return False
    if not marker_path.exists():
        return manifest_path.exists()
    return "template-sync-support" in marker_data.included_modules


def write_draft_marker_if_requested(
    *,
    marker_path: Path,
    marker_text: str,
    repo_root: Path,
    write: bool,
    write_draft_marker: bool,
) -> str:
    """Create a missing draft marker only when the explicit marker flag is used."""
    marker_relative = repository_relative_path(marker_path, repo_root)
    if not write:
        return "not written; default mode is read-only"
    if not write_draft_marker:
        return "not written; --write-draft-marker was not selected"
    if marker_path.exists() or marker_path.is_symlink():
        if marker_path.is_file() and not marker_path.is_symlink():
            try:
                existing_text = marker_path.read_text(encoding="utf-8")
            except OSError as error:
                raise FirstAdoptionBootstrapError(
                    f"Unable to read existing marker {marker_relative}: {os_error_summary(error)}"
                ) from error
            if existing_text == marker_text:
                return "unchanged; existing marker already matches the validated draft"
        raise FirstAdoptionBootstrapError(
            f"Refusing to overwrite existing marker `{marker_relative}`. "
            "Edit or validate the existing marker instead."
        )
    try:
        write_repository_relative_file_bytes(
            repo_root,
            marker_relative,
            marker_text.encode("utf-8"),
            field_name="marker path",
        )
    except (OSError, TemplateSyncMaterializationError) as error:
        if isinstance(error, OSError):
            error_summary = os_error_summary(error)
        else:
            error_summary = str(error)
        raise FirstAdoptionBootstrapError(
            f"Unable to write draft marker {marker_relative}: {error_summary}"
        ) from error
    return "created from validated draft marker"


def write_todo_if_requested(
    *,
    todo_path: Path,
    todo_text: str,
    merged_state: dict[str, Any],
    repo_root: Path,
    write: bool,
    update_existing_todo_state: bool,
) -> str:
    """Create or update the checklist only through documented safe paths."""
    todo_relative = repository_relative_path(todo_path, repo_root)
    if not write:
        return "not written; default mode is read-only"
    if todo_path.is_symlink():
        raise FirstAdoptionBootstrapError(
            f"Refusing to write checklist `{todo_relative}` because it is a symlink. "
            "Replace it with a regular file or remove it before writing."
        )
    if not todo_path.exists():
        try:
            todo_path.parent.mkdir(parents=True, exist_ok=True)
            todo_path.write_text(todo_text + "\n", encoding="utf-8")
        except OSError as error:
            raise FirstAdoptionBootstrapError(
                f"Unable to write checklist {todo_relative}: {os_error_summary(error)}"
            ) from error
        return "created"
    if not todo_path.is_file():
        raise FirstAdoptionBootstrapError(
            f"Refusing to update checklist `{todo_relative}` because it is not a regular file."
        )
    if not update_existing_todo_state:
        return "unchanged; existing checklist content was preserved"
    try:
        existing_text = todo_path.read_text(encoding="utf-8")
        updated_text = replace_or_append_state_block(existing_text, merged_state)
        if updated_text == existing_text:
            return "unchanged; structured state block was already current"
        todo_path.write_text(updated_text, encoding="utf-8")
    except OSError as error:
        raise FirstAdoptionBootstrapError(
            f"Unable to update checklist {todo_relative}: {os_error_summary(error)}"
        ) from error
    return "updated bootstrap-owned structured state block only"


def write_journal_if_requested(
    *,
    repo_root: Path,
    journal_path: Path,
    write: bool,
) -> str:
    """Create the adoption difficulties journal only when requested."""
    if not write:
        return "not written; default mode is read-only"
    template_path = resolve_safe_repository_target_path(
        repo_root,
        initialize_adoption_journal.DEFAULT_JOURNAL_TEMPLATE_PATH,
        field_name="journal scaffold path",
    )
    result = initialize_adoption_journal.create_adoption_journal(
        repo_root=repo_root,
        journal_path=journal_path,
        template_path=template_path,
    )
    return "created" if result.created else "unchanged; existing journal content was preserved"


def perform_writes(
    *,
    inputs: BootstrapInputs,
    todo_text: str,
    merged_state: dict[str, Any],
    marker_text: str | None,
    write: bool,
    update_existing_todo_state: bool,
    write_draft_marker: bool,
) -> BootstrapWriteResult:
    """Apply explicit write-mode changes and return their statuses."""
    todo_status = write_todo_if_requested(
        todo_path=inputs.todo_path,
        todo_text=todo_text,
        merged_state=merged_state,
        repo_root=inputs.repo_root,
        write=write,
        update_existing_todo_state=update_existing_todo_state,
    )
    journal_status = write_journal_if_requested(
        repo_root=inputs.repo_root,
        journal_path=inputs.journal_path,
        write=write,
    )
    if marker_text is None:
        marker_status = "not applicable; draft marker preview was not requested"
    else:
        marker_status = write_draft_marker_if_requested(
            marker_path=inputs.marker_path,
            marker_text=marker_text,
            repo_root=inputs.repo_root,
            write=write,
            write_draft_marker=write_draft_marker,
        )
    return BootstrapWriteResult(
        todo_status=todo_status,
        journal_status=journal_status,
        marker_status=marker_status,
    )


def active_and_resolved_decisions(
    state: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Split decisions into active and resolved records."""
    active: list[dict[str, Any]] = []
    resolved: list[dict[str, Any]] = []
    for decision in decision_by_key(state).values():
        if decision["status"] == RESOLVED_STATUS:
            resolved.append(decision)
        else:
            active.append(decision)
    return (
        sorted(active, key=lambda item: item["key"]),
        sorted(resolved, key=lambda item: item["key"]),
    )


def format_decision_list(decisions: list[dict[str, Any]]) -> str:
    """Render decision records as Markdown bullets."""
    if not decisions:
        return "- None"

    lines: list[str] = []
    for decision in decisions:
        dependent_files = ", ".join(f"`{path}`" for path in decision["dependent_files"])
        status = STATUS_DISPLAY.get(decision["status"], decision["status"])
        lines.append(
            f"- `{decision['key']}` ({decision['category']}): {decision['question']} "
            f"Status: `{status}`; dependent files/settings: {dependent_files or '`none`'}; "
            f"dependent-file status: `{decision['dependent_file_status']}`; "
            f"answer field: `{decision['answer'] or '[record answer]'}`; "
            f"evidence field: `{decision['evidence'] or '[record evidence]'}`."
        )
    return "\n".join(lines)


def line_path(line: str) -> str | None:
    """Extract the first path-like code span from a protected question line."""
    match = PROTECTED_QUESTION_PATH_PATTERN.search(line)
    if match is None:
        return None
    return match.group("path")


def format_filtered_protected_questions(
    *,
    ledger_rows: tuple[sync_candidates.LedgerRow, ...],
    state: dict[str, Any],
) -> str:
    """Render preflight protected questions, suppressing only resolved state keys."""
    resolved_keys = {
        decision["key"]
        for decision in decision_by_key(state).values()
        if decision["status"] == RESOLVED_STATUS
    }
    rendered = sync_candidates.format_protected_authorization_questions(ledger_rows)
    if rendered.startswith("No protected"):
        return rendered

    kept_lines: list[str] = []
    suppressed_paths: list[str] = []
    for line in rendered.splitlines():
        path = line_path(line)
        if path is not None and protected_decision_key(path) in resolved_keys:
            suppressed_paths.append(path)
            continue
        kept_lines.append(line)

    lines: list[str] = []
    if kept_lines:
        lines.extend(kept_lines)
    else:
        lines.append("No unresolved protected-file authorization questions remain.")
    if suppressed_paths:
        lines.extend(
            [
                "",
                "Resolved protected-file questions suppressed from active prompts:",
            ]
        )
        lines.extend(f"- `{path}`" for path in sorted(suppressed_paths))
    return "\n".join(lines)


def format_filtered_maintainer_questionnaire(
    *,
    discovery: sync_candidates.RepositoryDiscovery,
    state: dict[str, Any],
) -> str:
    """Render preflight maintainer questions with resolved state keys suppressed."""
    resolved_keys = {
        decision["key"]
        for decision in decision_by_key(state).values()
        if decision["status"] == RESOLVED_STATUS
    }
    kept_lines: list[str] = []
    suppressed_keys: list[str] = []
    for line in sync_candidates.format_maintainer_questionnaire(discovery).splitlines():
        matching_key = next(
            (key for phrase, key in QUESTION_KEY_PHRASES if phrase.casefold() in line.casefold()),
            None,
        )
        if matching_key is not None and matching_key in resolved_keys:
            suppressed_keys.append(matching_key)
            continue
        kept_lines.append(line)

    lines: list[str] = []
    if kept_lines:
        lines.extend(kept_lines)
    else:
        lines.append("No unresolved maintainer questionnaire items remain.")
    if suppressed_keys:
        lines.extend(["", "Resolved maintainer questions suppressed from active prompts:"])
        lines.extend(f"- `{key}`" for key in sorted(set(suppressed_keys)))
    return "\n".join(lines)


def format_validation_plan(
    *,
    repo_root: Path,
    max_command_length: int,
) -> str:
    """Render the runnable validation plan from run_first_adoption_checks.py."""
    collection = first_adoption_checks.collect_present_regular_files(repo_root)
    plan = first_adoption_checks.build_check_plan(
        repo_root=repo_root,
        files=collection.files,
        max_command_length=max_command_length,
    )

    lines = [
        (
            "Source: `.template-sync/scripts/run_first_adoption_checks.py` "
            "`collect_present_regular_files` + `build_check_plan`."
        ),
        "",
        f"- Git-visible regular files: {len(collection.files)}",
    ]
    if collection.skipped_non_regular_paths:
        lines.append(
            "- Skipped non-regular Git-visible paths: "
            + ", ".join(f"`{path}`" for path in collection.skipped_non_regular_paths)
        )
    if plan.notes:
        lines.append("- Notes:")
        lines.extend(f"  - {note}" for note in plan.notes)
    lines.append(f"- Planned validation commands: {len(plan.commands)}")
    if not plan.commands:
        lines.append("  - No first-adoption checks are available to run.")
    else:
        for index, planned_command in enumerate(plan.commands, start=1):
            lines.append(
                "  - "
                f"{index}. [{planned_command.group_label}] "
                f"`{first_adoption_checks.format_command(planned_command.command)}`"
            )
    lines.extend(
        [
            "",
            "Run before the first adoption commit:",
            "",
            "```bash",
            "python .template-sync/scripts/run_first_adoption_checks.py --check",
            "```",
        ]
    )
    return "\n".join(lines)


def existing_durable_path_status(path: Path) -> str:
    """Describe a durable-state path, flagging symlinks the writer would refuse.

    ``Path.is_file()`` follows symlinks, so a symlinked target would read as
    ``found`` and a broken symlink as ``missing`` even though write mode rejects
    symlinks. Report symlinks distinctly so the summary matches write behavior.
    """
    if path.is_symlink():
        return "present as a symlink (write mode will refuse it)"
    return "found" if path.is_file() else "missing"


def format_existing_state_summary(
    *,
    todo_text: str | None,
    existing_state: ExistingBootstrapState,
    inputs: BootstrapInputs,
) -> str:
    """Render detected durable-state files and parser behavior."""
    todo_status = "found" if todo_text is not None else "missing"
    state_status = "found" if existing_state.state is not None else "missing"
    journal_status = existing_durable_path_status(inputs.journal_path)
    marker_status = existing_durable_path_status(inputs.marker_path)
    return "\n".join(
        [
            f"- `{repository_relative_path(inputs.todo_path, inputs.repo_root)}`: {todo_status}",
            f"- Bootstrap structured state block: {state_status}",
            (
                "- Loose Markdown checklist state: ignored for resumed-run suppression; "
                "only the structured state block or another documented structured source counts."
            ),
            f"- `{repository_relative_path(inputs.journal_path, inputs.repo_root)}`: {journal_status}",
            f"- `{repository_relative_path(inputs.marker_path, inputs.repo_root)}`: {marker_status}",
        ]
    )


def format_intended_operations(
    *,
    inputs: BootstrapInputs,
    write_result: BootstrapWriteResult,
) -> str:
    """Render write/print operations and preservation constraints."""
    return "\n".join(
        [
            "| Artifact | Intended operation | Result | Preservation rule |",
            "| --- | --- | --- | --- |",
            (
                f"| `{repository_relative_path(inputs.todo_path, inputs.repo_root)}` | "
                "Print proposed checklist; write creates it when absent or updates only the "
                "bootstrap state block with `--update-existing-todo-state`. | "
                f"{write_result.todo_status} | Preserve Markdown notes outside the delimited "
                "bootstrap state block. |"
            ),
            (
                f"| `{repository_relative_path(inputs.journal_path, inputs.repo_root)}` | "
                "Print journal scaffold; write creates it when absent. | "
                f"{write_result.journal_status} | Never overwrite existing journal content. |"
            ),
            (
                f"| `{repository_relative_path(inputs.marker_path, inputs.repo_root)}` | "
                "Print draft marker when template-sync support is retained or requested; "
                "write requires `--write-draft-marker`. | "
                f"{write_result.marker_status} | Never overwrite an existing marker. |"
            ),
        ]
    )


def format_protocol_constraints() -> str:
    """Render protocol-preservation constraints for the bootstrap report."""
    return "\n".join(
        [
            "- Do not edit protected instruction files without explicit path-scoped maintainer authorization in the current task.",
            "- Preserve retained agent platform protocol sections when protected entry points are later authorized for editing.",
            "- Do not treat `minimal-preservation` as protected-file authorization; it only limits an already authorized edit.",
            "- Do not finalize policy-dependent files while their structured state is `not_yet_asked` or `unavailable_manual_review_required`.",
            "- `asked_and_deferred` requires dependent-file status and follow-up evidence before dependent files can avoid misleading defaults.",
            "- Do not infer resolved owner decisions from Markdown checkboxes, prose wording, or copied preflight output.",
        ]
    )


def format_report(
    *,
    inputs: BootstrapInputs,
    todo_text: str,
    journal_text: str,
    marker_text: str | None,
    existing_todo_text: str | None,
    existing_state: ExistingBootstrapState,
    merged_state: dict[str, Any],
    write_result: BootstrapWriteResult,
    full_state: bool,
    max_command_length: int,
) -> str:
    """Render the full first-adoption bootstrap report."""
    active_decisions, resolved_decisions = active_and_resolved_decisions(merged_state)
    lines = [
        "# First-Adoption Bootstrap Report",
        "",
        (
            "Default mode is read-only. This report composes existing preflight, "
            "journal, raw-state, marker-validation, and first-adoption validation-plan "
            "helpers so owner decisions survive context transitions before pruning."
        ),
        "",
        "## Existing Durable State",
        "",
        format_existing_state_summary(
            todo_text=existing_todo_text,
            existing_state=existing_state,
            inputs=inputs,
        ),
        "",
        "## Intended Operations",
        "",
        format_intended_operations(inputs=inputs, write_result=write_result),
        "",
        "## Protocol-Preservation Constraints",
        "",
        format_protocol_constraints(),
        "",
        "## Active Owner Questions And Blocked State",
        "",
        format_decision_list(active_decisions),
        "",
        "## Resolved Structured Decisions",
        "",
        format_decision_list(resolved_decisions),
        "",
        "## Maintainer Questionnaire",
        "",
        format_filtered_maintainer_questionnaire(
            discovery=inputs.discovery,
            state=merged_state,
        ),
        "",
        "## Protected-File Authorization Questions",
        "",
        format_filtered_protected_questions(
            ledger_rows=inputs.ledger_rows,
            state=merged_state,
        ),
        "",
        "## Raw First-Adoption State",
        "",
        first_adoption_state.format_first_adoption_state(
            inputs.raw_state,
            full_state=full_state,
        ),
        "",
        "## Runnable Validation Plan",
        "",
        format_validation_plan(
            repo_root=inputs.repo_root,
            max_command_length=max_command_length,
        ),
        "",
        "## Proposed `_TODO-repo-init.md`",
        "",
        "```markdown",
        todo_text,
        "```",
        "",
        "## Proposed `_ADOPTION-DIFFICULTIES.md`",
        "",
        "```markdown",
        journal_text.rstrip(),
        "```",
    ]
    if marker_text is not None:
        lines.extend(
            [
                "",
                "## Draft `.template-sync/marker.yml`",
                "",
                (
                    "The draft marker is validated for YAML serialization, JSON Schema, "
                    "and protected-decision integrity before printing. Treat the included "
                    "module list as a pre-pruning draft and edit it after maintainer "
                    "module decisions are recorded."
                ),
                "",
                "```yaml",
                marker_text.rstrip(),
                "```",
            ]
        )
    lines.extend(
        [
            "",
            "## Reused Adoption Ledger",
            "",
            inputs.ledger_document,
        ]
    )
    return "\n".join(lines)


def fail(message: str) -> NoReturn:
    """Print an error and exit non-zero."""
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(1)


def main(argv: list[str] | None = None, *, stdout: TextIO = sys.stdout) -> int:
    """Run the first-adoption bootstrap command."""
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        if args.github_api_base is not None and not args.include_github_metadata:
            raise FirstAdoptionBootstrapError(
                "--github-api-base can only be used with --include-github-metadata."
            )
        if args.update_existing_todo_state and not args.write:
            raise FirstAdoptionBootstrapError(
                "--update-existing-todo-state can only be used with --write."
            )
        if args.write_draft_marker and not args.write:
            raise FirstAdoptionBootstrapError("--write-draft-marker can only be used with --write.")

        repo_root = sync_candidates.resolve_repo_root(args.repo_root)
        marker_path = resolve_safe_repository_target_path(
            repo_root,
            args.marker,
            field_name="marker path",
        )
        manifest_path = resolve_safe_repository_target_path(
            repo_root,
            args.manifest,
            field_name="manifest path",
        )
        marker_schema_path = resolve_safe_repository_target_path(
            repo_root,
            args.marker_schema,
            field_name="marker schema path",
        )
        manifest_schema_path = resolve_safe_repository_target_path(
            repo_root,
            args.manifest_schema,
            field_name="manifest schema path",
        )
        todo_path = resolve_safe_repository_target_path(
            repo_root,
            args.todo_file,
            field_name="first-adoption checklist path",
        )
        journal_path = resolve_safe_repository_target_path(
            repo_root,
            args.journal_path,
            field_name="adoption journal path",
        )

        inputs = load_bootstrap_inputs(
            repo_root=repo_root,
            marker_path=marker_path,
            manifest_path=manifest_path,
            marker_schema_path=marker_schema_path,
            manifest_schema_path=manifest_schema_path,
            todo_path=todo_path,
            journal_path=journal_path,
            adoption_mode=args.adoption_mode,
            include_github_metadata=args.include_github_metadata,
            github_api_base=args.github_api_base,
        )
        existing_todo_text, existing_state = read_existing_todo_state(todo_path)
        default_state = default_bootstrap_state(
            ledger_rows=inputs.ledger_rows,
            marker_data=inputs.marker_data,
        )
        merged_state = merge_bootstrap_state(default_state, existing_state.state)
        todo_text = format_bootstrap_todo(inputs.discovery, inputs.marker_data, merged_state)

        journal_template_path = resolve_safe_repository_target_path(
            repo_root,
            initialize_adoption_journal.DEFAULT_JOURNAL_TEMPLATE_PATH,
            field_name="journal scaffold path",
        )
        try:
            journal_text = journal_template_path.read_text(encoding="utf-8")
        except OSError as error:
            raise FirstAdoptionBootstrapError(
                f"Unable to read journal scaffold: {os_error_summary(error)}"
            ) from error

        # The preview decision and the write decision are independent: a draft
        # marker may be needed for writing (--write --write-draft-marker) even when
        # the preview is suppressed (--draft-marker never). Compute the marker text
        # when either path needs it, but only surface the preview when printing.
        print_draft_marker = should_print_draft_marker(
            mode=args.draft_marker,
            marker_path=marker_path,
            marker_data=inputs.marker_data,
            manifest_path=manifest_path,
        )
        needs_marker_for_write = args.write and args.write_draft_marker
        marker_text: str | None = None
        if print_draft_marker or needs_marker_for_write:
            marker_document = draft_marker_document(
                source_repo=args.source_repo,
                manifest_modules=inputs.manifest_modules,
                existing_marker_data=inputs.marker_data,
            )
            validate_marker_document(
                marker_document,
                marker_path=marker_path,
                marker_schema_path=marker_schema_path,
                repo_root=repo_root,
            )
            marker_text = format_marker_yaml(marker_document)
        marker_preview_text = marker_text if print_draft_marker else None

        write_result = perform_writes(
            inputs=inputs,
            todo_text=todo_text,
            merged_state=merged_state,
            marker_text=marker_text,
            write=args.write,
            update_existing_todo_state=args.update_existing_todo_state,
            write_draft_marker=args.write_draft_marker,
        )
        print(
            format_report(
                inputs=inputs,
                todo_text=todo_text,
                journal_text=journal_text,
                marker_text=marker_preview_text,
                existing_todo_text=existing_todo_text,
                existing_state=existing_state,
                merged_state=merged_state,
                write_result=write_result,
                full_state=args.full_state,
                max_command_length=args.max_command_length,
            ),
            file=stdout,
        )
    except (
        FirstAdoptionBootstrapError,
        TemplateSyncMaterializationError,
        sync_candidates.CandidateGenerationError,
        first_adoption_checks.FirstAdoptionCheckError,
        initialize_adoption_journal.AdoptionJournalError,
    ) as error:
        fail(str(error))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
