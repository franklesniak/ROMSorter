"""Report first-adoption quality debt without mutating repository files."""

from __future__ import annotations

import argparse
import fnmatch
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
from collections import Counter
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import NoReturn, Protocol, TextIO, cast

import yaml  # type: ignore[import-untyped]
from template_sync_materialization_helpers import (
    TemplateSyncMaterializationError,
    parse_yaml_mapping,
)

DEFAULT_SUPPRESSION_PATH = ".template-sync/first-adoption/quality-suppressions.json"
MARKER_PATH = ".template-sync/marker.yml"
GIT_VISIBLE_FILES_COMMAND = (
    "git",
    "ls-files",
    "-z",
    "--cached",
    "--others",
    "--exclude-standard",
)
GIT_TRACKED_FILES_COMMAND = ("git", "ls-files", "-z", "--cached")
GIT_IGNORED_FILES_COMMAND = (
    "git",
    "ls-files",
    "-z",
    "--others",
    "--ignored",
    "--exclude-standard",
)
GIT_VISIBLE_POWERSHELL_CANDIDATES_COMMAND = (
    "git",
    "ls-files",
    "-z",
    "--full-name",
    "--cached",
    "--others",
    "--exclude-standard",
)
GIT_TRACKED_POWERSHELL_CANDIDATES_COMMAND = (
    "git",
    "ls-files",
    "-z",
    "--full-name",
    "--cached",
)
GIT_IGNORED_POWERSHELL_CANDIDATES_COMMAND = (
    "git",
    "ls-files",
    "-z",
    "--full-name",
    "--others",
    "--ignored",
    "--exclude-standard",
)
PATH_REFERENCE_CATEGORY = "path-reference"
PATH_REFERENCE_CASE_MISMATCH_RULE = "path-reference.case-mismatch"
PATH_REFERENCE_MISSING_TARGET_RULE = "path-reference.missing-target"
KNOWN_PATH_REFERENCE_RULE_IDS = frozenset(
    (PATH_REFERENCE_CASE_MISMATCH_RULE, PATH_REFERENCE_MISSING_TARGET_RULE)
)
KNOWN_PATH_REFERENCE_CATEGORIES = frozenset((PATH_REFERENCE_CATEGORY,))
MARKDOWN_FILE_SUFFIXES = frozenset((".md", ".mdc"))
ISSUE_TEMPLATE_SUFFIXES = frozenset((".md", ".yml", ".yaml"))
CONFIG_SURFACES = frozenset(
    (
        ".gitattributes",
        ".pre-commit-config.yaml",
        ".vscode/settings.json",
        "package.json",
        "pyproject.toml",
    )
)
WORKFLOW_SURFACE_PATTERNS = (".github/workflows/*.yml", ".github/workflows/*.yaml")
# Generic report scan exclusions are intentionally separate from PowerShell
# analyzer candidate policy, which is enforced by the PowerShell helper.
IGNORED_SCAN_DIRECTORIES = frozenset(
    (".cache", ".git", ".pytest_cache", ".ruff_cache", "node_modules")
)
MARKDOWN_PACKAGE_SCRIPT = "lint:md"
GITHUB_ACTIONS_POWERSHELL_CI = ".github/workflows/powershell-ci.yml"
AZURE_PIPELINES_POWERSHELL_CI = ".azuredevops/pipelines/powershell-ci.yml"
PSSCRIPTANALYZER_GATE_MODE_VARIABLE = "PSSCRIPTANALYZER_GATE_MODE"
PSSCRIPTANALYZER_GATE_MODE_PARAMETER = "gateMode"
PSSCRIPTANALYZER_GATE_MODE_PARAMETER_EXPRESSION = "${{ parameters.gateMode }}"
PSSCRIPTANALYZER_CANDIDATE_HELPER = "src/tools/Resolve-PSScriptAnalyzerCandidate.ps1"
PSSCRIPTANALYZER_GATE_HELPER = "src/tools/Resolve-PSScriptAnalyzerGate.ps1"
PSSCRIPTANALYZER_SETTINGS = ".github/linting/PSScriptAnalyzerSettings.psd1"
PSSCRIPTANALYZER_CANDIDATE_SUFFIXES = frozenset((".ps1", ".psm1", ".psd1"))
PSSCRIPTANALYZER_RECOGNIZED_GATE_MODES = frozenset(("strict", "first-adoption"))
UNSAFE_CANDIDATE_EXIT_CODE = 3
LINE_ENDING_RISK_STATES = frozenset(("cr", "crlf", "mixed"))
AZURE_DEVOPS_MODULES = frozenset(
    ("azure-devops-platform", "azure-pipelines", "azure-devops-collaboration")
)
AZURE_HOST_SETUP_FIELDS = (
    ("azure_boards_policy", "Azure Boards intake policy"),
    ("azure_repos_pr_template_policy", "Azure Repos pull request template policy"),
    ("azure_branch_policy_reviewer_guidance", "Branch policy reviewer guidance"),
    ("azure_security_intake_policy", "Security intake policy"),
    ("azure_security_product_enablement", "Security product enablement"),
    ("azure_dependency_update_policy", "Dependency update policy"),
)
PATH_TOKEN_PATTERN = re.compile(
    r"(?P<literal>(?:[A-Za-z0-9_.-]+[\\/])+[A-Za-z0-9_.%+~#?=&,-]+|"
    r"(?:\./|\../)?[A-Za-z0-9_.-]+\."
    r"(?:md|mdc|py|ps1|psm1|psd1|json|jsonc|ya?ml|toml|hcl|tf|tfvars|"
    r"tftpl|tfbackend|mjs|js|txt))"
)
MARKDOWN_LINK_TARGET_PATTERN = re.compile(r"!?\[[^\]]*]\((?P<target>[^)\s]+)(?:\s+[^)]*)?\)")
MARKDOWNLINT_FINDING_PATTERN = re.compile(
    r"^(?P<path>.*?):\d+(?::\d+)?\s+(?P<rule>MD\d{3}(?:/[A-Za-z0-9_-]+)?)\b"
)
GIT_EOL_PATTERN = re.compile(
    r"^i/(?P<index>\S+)\s+w/(?P<worktree>\S+)\s+attr/(?P<attr>.*?)\t(?P<path>.*)$"
)


class CaptureRunner(Protocol):
    """Run a command in a repository root and capture its text output."""

    def __call__(
        self,
        command: Sequence[str],
        repo_root: Path,
        *,
        env: Mapping[str, str] | None = None,
        input: str | None = None,
    ) -> subprocess.CompletedProcess[str]: ...


class FirstAdoptionQualityError(RuntimeError):
    """Raised when a first-adoption report cannot be generated safely."""


@dataclass(frozen=True)
class GitFileCollection:
    """Repository file discovery result."""

    files: tuple[str, ...]
    skipped_non_regular_paths: tuple[str, ...]


@dataclass(frozen=True)
class GitEolRecord:
    """One parsed ``git ls-files --eol`` record."""

    index_eol: str
    worktree_eol: str
    raw_attributes: str


@dataclass(frozen=True)
class LineEndingRecord:
    """Line-ending state for one discovered file."""

    path: str
    index_eol: str
    git_worktree_eol: str
    detected_worktree_eol: str
    attr_text: str
    attr_eol: str
    normalization_risk: bool


@dataclass(frozen=True)
class LineEndingReport:
    """Line-ending inventory and normalization-risk data."""

    records: tuple[LineEndingRecord, ...]
    counts: Counter[str]
    normalization_risk_paths: tuple[str, ...]


@dataclass(frozen=True)
class PathReferenceFinding:
    """One path-reference quality finding."""

    rule_id: str
    category: str
    source_path: str
    line_number: int
    literal: str
    normalized_path: str
    matched_path: str | None
    message: str


@dataclass(frozen=True)
class PathReferenceSuppression:
    """One configured suppression for a path-reference finding."""

    rule_id: str | None
    category: str | None
    path: str | None
    path_glob: str | None
    literal: str | None
    literal_pattern: re.Pattern[str] | None
    reason: str

    def matches(self, finding: PathReferenceFinding) -> bool:
        """Return whether this suppression applies to ``finding``."""
        if self.rule_id is not None and self.rule_id != finding.rule_id:
            return False
        if self.category is not None and self.category != finding.category:
            return False
        if self.path is not None and self.path != finding.source_path:
            return False
        if self.path_glob is not None and not fnmatch.fnmatchcase(
            finding.source_path,
            self.path_glob,
        ):
            return False
        if self.literal is not None and self.literal != finding.literal:
            return False
        if self.literal_pattern is None:
            return True
        return self.literal_pattern.search(finding.literal) is not None


@dataclass(frozen=True)
class QualitySuppressions:
    """Parsed quality report suppression configuration."""

    path_reference: tuple[PathReferenceSuppression, ...]


@dataclass(frozen=True)
class PathReferenceReport:
    """Path-reference report result."""

    scanned_paths: tuple[str, ...]
    findings: tuple[PathReferenceFinding, ...]
    suppressed_count: int


@dataclass(frozen=True)
class MarkdownlintFinding:
    """One parsed markdownlint finding."""

    path: str
    rule: str


@dataclass(frozen=True)
class MarkdownlintReport:
    """Markdownlint report or fixer result."""

    available: bool
    message: str
    return_code: int
    findings: tuple[MarkdownlintFinding, ...]
    changed_files: tuple[str, ...]


@dataclass(frozen=True)
class GateModeStaticValue:
    """One normalized static PSScriptAnalyzer gate-mode value."""

    raw_value: object
    display_value: str
    recognized_mode: str | None
    note: str
    manual_review: bool


@dataclass(frozen=True)
class GateModeSetting:
    """One committed static gate-mode setting found in CI YAML."""

    path: str
    location: str
    specificity: str
    value: GateModeStaticValue
    effective_value: GateModeStaticValue | None = None


@dataclass(frozen=True)
class PowerShellAnalyzerDebtRecord:
    """One structured PSScriptAnalyzer first-adoption debt record."""

    rule_name: str
    severity: str
    normalized_path: str
    line: int | None
    column: int | None
    message: str


@dataclass(frozen=True)
class PowerShellAnalyzerReport:
    """PowerShell analyzer report result."""

    available: bool
    message: str
    candidate_summary_lines: tuple[str, ...]
    unsafe_candidate_count: int
    summary_lines: tuple[str, ...]
    issue_ready_markdown: tuple[str, ...]
    analyzer_debt_records: tuple[PowerShellAnalyzerDebtRecord, ...]
    opt_in_guidance_lines: tuple[str, ...]


@dataclass(frozen=True)
class HostSetupReport:
    """Azure DevOps Services first-adoption setup status."""

    marker_available: bool
    azure_modules_retained: bool
    host_provider: str | None
    setup_items: tuple[tuple[str, str], ...]
    message: str


def default_repo_root() -> Path:
    """Return the repository root implied by this script's committed location."""
    return Path(__file__).resolve().parents[2]


def resolve_repo_root(raw_repo_root: str | None) -> Path:
    """Resolve and validate the repository root argument."""
    repo_root = Path(raw_repo_root).expanduser() if raw_repo_root else default_repo_root()
    resolved = repo_root.resolve()
    if not resolved.is_dir():
        raise FirstAdoptionQualityError("Repository root does not exist or is not a directory.")
    return resolved


def resolve_repo_path(repo_root: Path, relative_path: str) -> Path:
    """Resolve a Git-relative path inside the repository root."""
    if "\\" in relative_path or Path(relative_path).is_absolute():
        raise FirstAdoptionQualityError(f"Git path must be repository-relative: {relative_path}")
    parts = PurePosixPath(relative_path).parts
    if any(part in ("", ".", "..") for part in parts):
        raise FirstAdoptionQualityError(f"Git path must not contain traversal: {relative_path}")
    resolved_path = (repo_root / relative_path).resolve()
    try:
        resolved_path.relative_to(repo_root)
    except ValueError as error:
        raise FirstAdoptionQualityError(
            f"Git path escapes repository root: {relative_path}"
        ) from error
    return resolved_path


def is_present_regular_file(path: Path) -> bool:
    """Return whether ``path`` is a present regular file, excluding symlinks."""
    return not path.is_symlink() and path.is_file()


def load_marker_template_sync(repo_root: Path) -> dict[str, object] | None:
    """Load the marker's template_sync mapping when a marker is present."""
    marker_path = resolve_repo_path(repo_root, MARKER_PATH)
    # resolve_repo_path() collapses a leaf symlink, so check the lexical path for
    # symlink/regular-file status the way the rest of this file's discovery does.
    marker_link = repo_root / MARKER_PATH
    if not marker_path.exists() and not marker_link.is_symlink():
        return None
    if not is_present_regular_file(marker_link):
        raise FirstAdoptionQualityError(f"Expected a regular file: {MARKER_PATH}")
    try:
        marker_document = parse_yaml_mapping(
            marker_path.read_text(encoding="utf-8-sig"), MARKER_PATH
        )
    except TemplateSyncMaterializationError as error:
        raise FirstAdoptionQualityError(str(error)) from error
    except OSError as error:
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        raise FirstAdoptionQualityError(
            f"Unable to read {MARKER_PATH} ({error_summary})."
        ) from error
    template_sync = marker_document.get("template_sync")
    if not isinstance(template_sync, dict):
        raise FirstAdoptionQualityError(f"{MARKER_PATH} must contain template_sync mapping.")
    return cast(dict[str, object], template_sync)


def marker_modules(template_sync: Mapping[str, object]) -> frozenset[str]:
    """Return marker-listed modules from a parsed template_sync mapping."""
    raw_modules = template_sync.get("included_modules")
    if not isinstance(raw_modules, list):
        return frozenset()
    return frozenset(item for item in cast(list[object], raw_modules) if isinstance(item, str))


def build_host_setup_report(repo_root: Path) -> HostSetupReport:
    """Build an Azure DevOps Services first-adoption service setup report."""
    template_sync = load_marker_template_sync(repo_root)
    if template_sync is None:
        return HostSetupReport(
            marker_available=False,
            azure_modules_retained=False,
            host_provider=None,
            setup_items=(),
            message=f"{MARKER_PATH} not found; no host setup marker decisions are available.",
        )

    modules = marker_modules(template_sync)
    azure_modules_retained = bool(modules & AZURE_DEVOPS_MODULES)
    host_provider = template_sync.get("host_provider")
    host_provider_text = host_provider if isinstance(host_provider, str) else None
    if not azure_modules_retained:
        return HostSetupReport(
            marker_available=True,
            azure_modules_retained=False,
            host_provider=host_provider_text,
            setup_items=(),
            message="No Azure DevOps Services host setup tasks are recorded.",
        )

    setup_items = tuple(
        (label, str(template_sync.get(field_name, "not recorded")))
        for field_name, label in AZURE_HOST_SETUP_FIELDS
    )
    return HostSetupReport(
        marker_available=True,
        azure_modules_retained=True,
        host_provider=host_provider_text,
        setup_items=setup_items,
        message=(
            "These are Azure DevOps Services setup follow-ups, not local file "
            "materialization failures or GitHub issue-template findings."
        ),
    )


def print_host_setup_report(report: HostSetupReport, *, stdout: TextIO) -> None:
    """Print a stable Azure DevOps Services setup report."""
    print("First-adoption host setup report:", file=stdout)
    if not report.marker_available or not report.azure_modules_retained:
        print(f"  {report.message}", file=stdout)
        return
    print("  Azure DevOps Services service setup tasks:", file=stdout)
    print(f"    - Host provider: {report.host_provider or 'not recorded'}", file=stdout)
    for label, value in report.setup_items:
        print(f"    - {label}: {value}", file=stdout)
    print(f"  {report.message}", file=stdout)


def format_command(command: Sequence[str]) -> str:
    """Return a shell-like representation of the exact argument vector."""
    command_parts = list(command)
    if os.name == "nt":
        return subprocess.list2cmdline(command_parts)
    return shlex.join(command_parts)


def run_capture(
    command: Sequence[str],
    repo_root: Path,
    *,
    env: Mapping[str, str] | None = None,
    input: str | None = None,
) -> subprocess.CompletedProcess[str]:
    """Run ``command`` in ``repo_root`` and capture text output."""
    try:
        return subprocess.run(
            list(command),
            cwd=repo_root,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            env=env,
            input=input,
        )
    except OSError as error:
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        raise FirstAdoptionQualityError(f"Unable to run {command[0]} ({error_summary}).") from error


def git_capture(repo_root: Path, args: Sequence[str]) -> str:
    """Run a Git command and return stdout, raising on failure."""
    result = run_capture(("git", *args), repo_root)
    if result.returncode != 0:
        message = result.stderr.strip() or result.stdout.strip() or "git command failed"
        raise FirstAdoptionQualityError(f"Unable to inspect repository with git: {message}")
    return result.stdout


def git_capture_bytes(repo_root: Path, args: Sequence[str]) -> bytes:
    """Run a Git command and return raw stdout bytes, raising on failure."""
    try:
        result = subprocess.run(
            ["git", *args],
            cwd=repo_root,
            check=False,
            capture_output=True,
        )
    except OSError as error:
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        raise FirstAdoptionQualityError(f"Unable to run git ({error_summary}).") from error
    if result.returncode != 0:
        stderr = result.stderr.decode("utf-8", errors="replace").strip()
        stdout = result.stdout.decode("utf-8", errors="replace").strip()
        message = stderr or stdout or "git command failed"
        raise FirstAdoptionQualityError(f"Unable to inspect repository with git: {message}")
    return result.stdout


def decode_git_nul_paths(output: bytes) -> tuple[str, ...]:
    """Decode NUL-delimited Git path output as strict UTF-8 records."""
    if not output:
        return ()
    if not output.endswith(b"\0"):
        raise FirstAdoptionQualityError(
            "Git file-list output was not NUL-terminated; manual review is required."
        )

    paths: list[str] = []
    for raw_record in output[:-1].split(b"\0"):
        if not raw_record:
            raise FirstAdoptionQualityError(
                "Git file-list output contained an empty NUL-delimited path record; "
                "manual review is required."
            )
        try:
            paths.append(raw_record.decode("utf-8", errors="strict"))
        except UnicodeDecodeError as error:
            raise FirstAdoptionQualityError(
                "Git returned a PowerShell candidate path that is not valid UTF-8; "
                "manual review is required."
            ) from error
    return tuple(paths)


def collect_git_files(
    repo_root: Path,
    *,
    tracked_only: bool = False,
    include_ignored: bool = False,
    stdout: TextIO | None = None,
) -> GitFileCollection:
    """Collect tracked and adoption-visible files for quality reports."""
    commands: list[tuple[str, ...]] = [
        GIT_TRACKED_FILES_COMMAND if tracked_only else GIT_VISIBLE_FILES_COMMAND
    ]
    if include_ignored and not tracked_only:
        commands.append(GIT_IGNORED_FILES_COMMAND)

    files: list[str] = []
    skipped_non_regular_paths: list[str] = []
    seen_paths: set[str] = set()
    for command in commands:
        if stdout is not None:
            print(f"$ {format_command(command)}", file=stdout, flush=True)
        output = git_capture(repo_root, command[1:])
        for relative_path in output.split("\0"):
            if not relative_path or relative_path in seen_paths:
                continue
            seen_paths.add(relative_path)
            path = resolve_repo_path(repo_root, relative_path)
            if is_present_regular_file(path):
                files.append(relative_path)
            elif path.exists() or path.is_symlink():
                skipped_non_regular_paths.append(relative_path)

    return GitFileCollection(
        files=tuple(sorted(files)),
        skipped_non_regular_paths=tuple(sorted(skipped_non_regular_paths)),
    )


def git_eol_records(
    repo_root: Path,
    *,
    tracked_only: bool = False,
    include_ignored: bool = False,
) -> dict[str, GitEolRecord]:
    """Return parsed ``git ls-files --eol`` records keyed by repository path."""
    commands: list[tuple[str, ...]] = [
        (
            ("git", "ls-files", "--eol", "--cached")
            if tracked_only
            else ("git", "ls-files", "--eol", "--cached", "--others", "--exclude-standard")
        )
    ]
    if include_ignored and not tracked_only:
        commands.append(("git", "ls-files", "--eol", "--others", "--ignored", "--exclude-standard"))

    records: dict[str, GitEolRecord] = {}
    for command in commands:
        output = git_capture(repo_root, command[1:])
        for line in output.splitlines():
            match = GIT_EOL_PATTERN.match(line)
            if match is None:
                continue
            path = match.group("path")
            records[path] = GitEolRecord(
                index_eol=match.group("index"),
                worktree_eol=match.group("worktree"),
                raw_attributes=match.group("attr"),
            )
    return records


def git_attributes(repo_root: Path, files: Sequence[str]) -> dict[str, dict[str, str]]:
    """Return selected Git attributes for ``files``."""
    attributes_by_path: dict[str, dict[str, str]] = {relative_path: {} for relative_path in files}
    if not files:
        return attributes_by_path

    chunk_size = 100
    for index in range(0, len(files), chunk_size):
        chunk = files[index : index + chunk_size]
        output = git_capture(repo_root, ("check-attr", "text", "eol", "--", *chunk))
        for line in output.splitlines():
            try:
                path, attribute, value = line.split(": ", 2)
            except ValueError:
                continue
            attributes_by_path.setdefault(path, {})[attribute] = value
    return attributes_by_path


def detect_file_line_endings(path: Path) -> str:
    """Return a stable line-ending category for ``path``."""
    try:
        data = path.read_bytes()
    except OSError as error:
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        raise FirstAdoptionQualityError(
            f"Unable to read repository file ({error_summary})."
        ) from error

    if b"\0" in data:
        return "binary"

    crlf_count = data.count(b"\r\n")
    lf_count = data.count(b"\n") - crlf_count
    cr_count = data.count(b"\r") - crlf_count

    populated_kinds = sum(1 for count in (crlf_count, lf_count, cr_count) if count > 0)
    if populated_kinds > 1:
        return "mixed"
    if crlf_count > 0:
        return "crlf"
    if lf_count > 0:
        return "lf"
    if cr_count > 0:
        return "cr"
    return "no-newline"


def build_line_ending_report(
    repo_root: Path,
    *,
    tracked_only: bool = False,
    include_ignored: bool = False,
) -> LineEndingReport:
    """Build a line-ending inventory for adoption-visible files."""
    collection = collect_git_files(
        repo_root,
        tracked_only=tracked_only,
        include_ignored=include_ignored,
    )
    eol_records = git_eol_records(
        repo_root,
        tracked_only=tracked_only,
        include_ignored=include_ignored,
    )
    attributes = git_attributes(repo_root, collection.files)

    records: list[LineEndingRecord] = []
    counts: Counter[str] = Counter()
    risk_paths: list[str] = []
    for relative_path in collection.files:
        detected_eol = detect_file_line_endings(resolve_repo_path(repo_root, relative_path))
        counts[detected_eol] += 1
        git_record = eol_records.get(relative_path)
        attr_values = attributes.get(relative_path, {})
        attr_text = attr_values.get("text", "unspecified")
        attr_eol = attr_values.get("eol", "unspecified")
        normalization_risk = (
            attr_eol == "lf" and detected_eol in LINE_ENDING_RISK_STATES and attr_text != "unset"
        )
        if normalization_risk:
            risk_paths.append(relative_path)
        records.append(
            LineEndingRecord(
                path=relative_path,
                index_eol=git_record.index_eol if git_record else "unknown",
                git_worktree_eol=git_record.worktree_eol if git_record else "unknown",
                detected_worktree_eol=detected_eol,
                attr_text=attr_text,
                attr_eol=attr_eol,
                normalization_risk=normalization_risk,
            )
        )

    return LineEndingReport(
        records=tuple(records),
        counts=counts,
        normalization_risk_paths=tuple(sorted(risk_paths)),
    )


def print_line_ending_report(report: LineEndingReport, *, stdout: TextIO) -> None:
    """Print a stable line-ending report."""
    print("First-adoption line-ending report:", file=stdout)
    print(f"  Files inventoried: {len(report.records)}", file=stdout)
    for kind in ("lf", "crlf", "mixed", "cr", "no-newline", "binary"):
        print(f"  {kind}: {report.counts.get(kind, 0)}", file=stdout)
    if report.normalization_risk_paths:
        print(
            "  Normalization risk: files below are matched by Git attributes "
            "that pin LF and currently contain CRLF, CR, or mixed endings.",
            file=stdout,
        )
        print(
            "  Running broad fixers or touching these files may cause reviewable "
            "line-ending normalization; do not run broad renormalization "
            "automatically.",
            file=stdout,
        )
        for path in report.normalization_risk_paths:
            print(f"    - {path}", file=stdout)
    else:
        print("  Normalization risk: none detected.", file=stdout)


def is_excluded_scan_path(relative_path: str) -> bool:
    """Return whether ``relative_path`` lives under a generated or dependency directory."""
    parts = PurePosixPath(relative_path).parts
    return any(part in IGNORED_SCAN_DIRECTORIES for part in parts)


def is_path_reference_surface(relative_path: str) -> bool:
    """Return whether ``relative_path`` belongs to the documented scan surface set."""
    if is_excluded_scan_path(relative_path):
        return False
    path = PurePosixPath(relative_path)
    suffix = path.suffix.lower()
    if suffix in MARKDOWN_FILE_SUFFIXES:
        return True
    if relative_path.startswith(".github/ISSUE_TEMPLATE/") and suffix in ISSUE_TEMPLATE_SUFFIXES:
        return True
    if relative_path in CONFIG_SURFACES:
        return True
    return any(fnmatch.fnmatchcase(relative_path, pattern) for pattern in WORKFLOW_SURFACE_PATTERNS)


def repo_directories(files: Sequence[str]) -> set[str]:
    """Return repository-relative directories implied by ``files``."""
    directories: set[str] = set()
    for file_path in files:
        path = PurePosixPath(file_path)
        for parent in path.parents:
            parent_text = parent.as_posix()
            if parent_text == ".":
                continue
            directories.add(parent_text)
    return directories


def case_index(paths: Sequence[str] | set[str]) -> dict[str, str]:
    """Return a case-folded path index with deterministic canonical values."""
    index: dict[str, str] = {}
    for path in sorted(paths):
        index.setdefault(path.casefold(), path)
    return index


def normalize_reference_literal(
    literal: str,
    *,
    source_path: str,
) -> tuple[str, bool]:
    """Normalize a scanned literal into a repository-relative candidate path."""
    cleaned = literal.strip().strip("`'\"()[]{}<>,;:")
    if not cleaned:
        return "", False
    if re.match(r"^[A-Za-z][A-Za-z0-9+.-]*:", cleaned) or cleaned.startswith("//"):
        return "", False
    if cleaned.startswith(("#", "mailto:")):
        return "", False
    if "@" in cleaned or "$" in cleaned or "*" in cleaned:
        return "", False
    if any(char in cleaned for char in ("{", "}", "<", ">")):
        return "", False
    if re.match(r"^[A-Za-z]:[\\/]", cleaned):
        return "", False

    path_part = cleaned.split("#", maxsplit=1)[0].split("?", maxsplit=1)[0]
    path_part = path_part.replace("\\", "/")
    if not path_part or path_part.startswith("/"):
        return "", False

    source_parent = PurePosixPath(source_path).parent
    try:
        if path_part.startswith(("./", "../")):
            parts: list[str] = []
            for part in (source_parent / path_part).parts:
                if part in ("", "."):
                    continue
                if part == "..":
                    if not parts:
                        return "", False
                    parts.pop()
                else:
                    parts.append(part)
            normalized = PurePosixPath(*parts).as_posix() if parts else ""
        else:
            normalized = PurePosixPath(path_part).as_posix()
    except ValueError:
        return "", False

    if not normalized or normalized == ".":
        return "", False
    return normalized.removesuffix("/"), cleaned.endswith("/")


def is_strong_repo_path(candidate_path: str, known_root_components: set[str]) -> bool:
    """Return whether a missing candidate is strong enough to report."""
    first_part = PurePosixPath(candidate_path).parts[0]
    if first_part in known_root_components:
        return True
    return first_part.startswith(".") and len(PurePosixPath(candidate_path).parts) > 1


def extract_reference_literals(text: str) -> list[tuple[int, str]]:
    """Extract path-like literals from text with source line numbers."""
    literals: list[tuple[int, str]] = []
    for line_number, line in enumerate(text.splitlines(), start=1):
        for match in MARKDOWN_LINK_TARGET_PATTERN.finditer(line):
            literals.append((line_number, match.group("target")))
        for match in PATH_TOKEN_PATTERN.finditer(line):
            literals.append((line_number, match.group("literal")))
    return literals


def _string_property(
    value: object,
    *,
    key: str,
    required: bool = False,
) -> str | None:
    """Validate and return an optional string property from a suppression entry."""
    if value is None:
        if required:
            raise FirstAdoptionQualityError(f"Suppression entry is missing required '{key}'.")
        return None
    if not isinstance(value, str) or not value.strip():
        raise FirstAdoptionQualityError(f"Suppression entry '{key}' must be a non-empty string.")
    return value


def parse_path_reference_suppression(raw_entry: object, *, index: int) -> PathReferenceSuppression:
    """Parse one path-reference suppression entry."""
    if not isinstance(raw_entry, dict):
        raise FirstAdoptionQualityError(
            f"path-reference suppression {index} must be a JSON object."
        )
    raw_entry = cast(dict[str, object], raw_entry)

    allowed_keys = {
        "category",
        "literal",
        "literalPattern",
        "path",
        "pathGlob",
        "reason",
        "ruleId",
    }
    unknown_keys = sorted(set(raw_entry) - allowed_keys)
    if unknown_keys:
        raise FirstAdoptionQualityError(
            f"path-reference suppression {index} has unknown key(s): {', '.join(unknown_keys)}"
        )
    selector_keys = allowed_keys - {"reason"}
    if not selector_keys.intersection(raw_entry):
        raise FirstAdoptionQualityError(
            f"path-reference suppression {index} must include at least one selector."
        )

    rule_id = _string_property(raw_entry.get("ruleId"), key="ruleId")
    if rule_id is not None and rule_id not in KNOWN_PATH_REFERENCE_RULE_IDS:
        raise FirstAdoptionQualityError(
            f"path-reference suppression {index} references unknown ruleId: {rule_id}"
        )

    category = _string_property(raw_entry.get("category"), key="category")
    if category is not None and category not in KNOWN_PATH_REFERENCE_CATEGORIES:
        raise FirstAdoptionQualityError(
            f"path-reference suppression {index} references unknown category: {category}"
        )

    literal_pattern_text = _string_property(raw_entry.get("literalPattern"), key="literalPattern")
    literal_pattern = None
    if literal_pattern_text is not None:
        try:
            literal_pattern = re.compile(literal_pattern_text)
        except re.error as error:
            raise FirstAdoptionQualityError(
                f"path-reference suppression {index} has invalid literalPattern: {error}"
            ) from error

    return PathReferenceSuppression(
        rule_id=rule_id,
        category=category,
        path=_string_property(raw_entry.get("path"), key="path"),
        path_glob=_string_property(raw_entry.get("pathGlob"), key="pathGlob"),
        literal=_string_property(raw_entry.get("literal"), key="literal"),
        literal_pattern=literal_pattern,
        reason=_string_property(raw_entry.get("reason"), key="reason", required=True) or "",
    )


def load_quality_suppressions(repo_root: Path, suppression_path: str) -> QualitySuppressions:
    """Load and validate quality-report suppressions."""
    full_path = resolve_repo_path(repo_root, suppression_path)
    if not full_path.exists():
        return QualitySuppressions(path_reference=())
    if not is_present_regular_file(full_path):
        raise FirstAdoptionQualityError(
            f"Suppression path is not a regular file: {suppression_path}"
        )
    try:
        raw_data = json.loads(full_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as error:
        raise FirstAdoptionQualityError(
            f"{suppression_path} is not valid JSON: {error.msg}"
        ) from error
    except OSError as error:
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        raise FirstAdoptionQualityError(
            f"Unable to read suppression file ({error_summary})."
        ) from error

    if not isinstance(raw_data, dict):
        raise FirstAdoptionQualityError(f"{suppression_path} must contain a JSON object.")
    raw_data = cast(dict[str, object], raw_data)
    allowed_sections = {"path-reference"}
    unknown_sections = sorted(set(raw_data) - allowed_sections)
    if unknown_sections:
        raise FirstAdoptionQualityError(
            f"{suppression_path} has unknown top-level section(s): {', '.join(unknown_sections)}"
        )
    if "path-reference" not in raw_data:
        raise FirstAdoptionQualityError(
            f"{suppression_path} is missing required 'path-reference' section."
        )

    path_reference_section = raw_data["path-reference"]
    if not isinstance(path_reference_section, dict):
        raise FirstAdoptionQualityError("'path-reference' must be a JSON object.")
    path_reference_section = cast(dict[str, object], path_reference_section)
    section_unknown_keys = sorted(set(path_reference_section) - {"suppressions"})
    if section_unknown_keys:
        raise FirstAdoptionQualityError(
            "'path-reference' has unknown key(s): " + ", ".join(section_unknown_keys)
        )
    raw_suppressions = path_reference_section.get("suppressions", [])
    if not isinstance(raw_suppressions, list):
        raise FirstAdoptionQualityError("'path-reference.suppressions' must be an array.")

    suppressions = tuple(
        parse_path_reference_suppression(raw_entry, index=index)
        for index, raw_entry in enumerate(cast(list[object], raw_suppressions), start=1)
    )
    return QualitySuppressions(path_reference=suppressions)


def is_suppressed(
    finding: PathReferenceFinding,
    suppressions: Sequence[PathReferenceSuppression],
) -> bool:
    """Return whether ``finding`` is suppressed."""
    return any(suppression.matches(finding) for suppression in suppressions)


def build_path_reference_report(
    repo_root: Path,
    *,
    tracked_only: bool = False,
    include_ignored: bool = False,
    suppression_path: str = DEFAULT_SUPPRESSION_PATH,
    include_missing_targets: bool = False,
) -> PathReferenceReport:
    """Build a path-reference casing and existence report."""
    collection = collect_git_files(
        repo_root,
        tracked_only=tracked_only,
        include_ignored=include_ignored,
    )
    suppressions = load_quality_suppressions(repo_root, suppression_path)
    file_paths = set(collection.files)
    directory_paths = repo_directories(collection.files)
    all_known_paths = file_paths | directory_paths
    exact_paths = all_known_paths
    folded_paths = case_index(all_known_paths)
    known_root_components = {PurePosixPath(path).parts[0] for path in all_known_paths}

    scanned_paths = tuple(path for path in collection.files if is_path_reference_surface(path))
    findings: list[PathReferenceFinding] = []
    suppressed_count = 0
    seen_findings: set[tuple[str, str, int, str]] = set()
    for source_path in scanned_paths:
        full_path = resolve_repo_path(repo_root, source_path)
        try:
            text = full_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        except OSError as error:
            error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
            raise FirstAdoptionQualityError(
                f"Unable to read scanned surface {source_path} ({error_summary})."
            ) from error

        for line_number, literal in extract_reference_literals(text):
            candidate_path, _had_trailing_slash = normalize_reference_literal(
                literal,
                source_path=source_path,
            )
            if not candidate_path:
                continue
            if candidate_path in exact_paths:
                continue

            folded_match = folded_paths.get(candidate_path.casefold())
            if folded_match is not None and folded_match != candidate_path:
                finding = PathReferenceFinding(
                    rule_id=PATH_REFERENCE_CASE_MISMATCH_RULE,
                    category=PATH_REFERENCE_CATEGORY,
                    source_path=source_path,
                    line_number=line_number,
                    literal=literal,
                    normalized_path=candidate_path,
                    matched_path=folded_match,
                    message=(
                        "Path reference casing differs from the discovered "
                        f"repository path '{folded_match}'."
                    ),
                )
            elif include_missing_targets and is_strong_repo_path(
                candidate_path,
                known_root_components,
            ):
                finding = PathReferenceFinding(
                    rule_id=PATH_REFERENCE_MISSING_TARGET_RULE,
                    category=PATH_REFERENCE_CATEGORY,
                    source_path=source_path,
                    line_number=line_number,
                    literal=literal,
                    normalized_path=candidate_path,
                    matched_path=None,
                    message="Path reference does not match a discovered repository file or directory.",
                )
            else:
                continue

            dedupe_key = (
                finding.rule_id,
                finding.source_path,
                finding.line_number,
                finding.literal,
            )
            if dedupe_key in seen_findings:
                continue
            seen_findings.add(dedupe_key)
            if is_suppressed(finding, suppressions.path_reference):
                suppressed_count += 1
                continue
            findings.append(finding)

    return PathReferenceReport(
        scanned_paths=scanned_paths,
        findings=tuple(findings),
        suppressed_count=suppressed_count,
    )


def print_path_reference_report(report: PathReferenceReport, *, stdout: TextIO) -> None:
    """Print a stable path-reference report."""
    print("First-adoption path-reference report:", file=stdout)
    print("  Scan surfaces:", file=stdout)
    print("    - Markdown files (*.md, *.mdc)", file=stdout)
    print("    - Issue and pull request templates", file=stdout)
    print(
        "    - .vscode/settings.json, pyproject.toml, package.json, "
        ".pre-commit-config.yaml, .github/workflows/*.yml/*.yaml, and .gitattributes",
        file=stdout,
    )
    print(f"  Files scanned: {len(report.scanned_paths)}", file=stdout)
    print(f"  Findings: {len(report.findings)}", file=stdout)
    print(f"  Suppressed findings: {report.suppressed_count}", file=stdout)
    if not report.findings:
        return
    for finding in report.findings:
        print(
            f"  - {finding.rule_id} {finding.source_path}:{finding.line_number} "
            f"`{finding.literal}`",
            file=stdout,
        )
        print(f"    {finding.message}", file=stdout)


def load_package_scripts(repo_root: Path) -> dict[str, object]:
    """Return the root package.json scripts mapping when package.json exists."""
    package_path = repo_root / "package.json"
    if not package_path.exists():
        return {}
    if not is_present_regular_file(package_path):
        raise FirstAdoptionQualityError("Expected a regular file: package.json")
    try:
        package_data = json.loads(package_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as error:
        raise FirstAdoptionQualityError(f"package.json is not valid JSON: {error.msg}") from error
    except OSError as error:
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        raise FirstAdoptionQualityError(
            f"Unable to read package.json ({error_summary})."
        ) from error

    if not isinstance(package_data, dict):
        return {}
    package_data = cast(dict[str, object], package_data)
    scripts = package_data.get("scripts")
    return cast(dict[str, object], scripts) if isinstance(scripts, dict) else {}


def npm_executable() -> str | None:
    """Return an npm executable when one is available."""
    return shutil.which("npm") or shutil.which("npm.cmd")


def parse_markdownlint_findings(output: str) -> tuple[MarkdownlintFinding, ...]:
    """Parse markdownlint-cli2 text findings into stable counters."""
    findings: list[MarkdownlintFinding] = []
    for line in output.splitlines():
        match = MARKDOWNLINT_FINDING_PATTERN.match(line)
        if match is None:
            continue
        findings.append(MarkdownlintFinding(path=match.group("path"), rule=match.group("rule")))
    return tuple(findings)


def git_status_files(repo_root: Path) -> tuple[str, ...]:
    """Return changed Git status paths for fixer summaries."""
    output = git_capture(repo_root, ("status", "--porcelain=v1", "--untracked-files=all"))
    paths: list[str] = []
    for line in output.splitlines():
        if not line.strip():
            continue
        paths.append(line[3:] if len(line) > 3 else line.strip())
    return tuple(sorted(paths))


def build_markdownlint_report(
    repo_root: Path,
    *,
    fix: bool = False,
    runner: CaptureRunner = run_capture,
) -> MarkdownlintReport:
    """Run Markdownlint in report or explicit fixer mode."""
    package_scripts = load_package_scripts(repo_root)
    if MARKDOWN_PACKAGE_SCRIPT not in package_scripts:
        return MarkdownlintReport(
            available=False,
            message="Markdownlint report unavailable: package.json has no lint:md script.",
            return_code=0,
            findings=(),
            changed_files=(),
        )
    npm = npm_executable()
    if npm is None:
        return MarkdownlintReport(
            available=False,
            message="Markdownlint report unavailable: npm was not found on PATH.",
            return_code=0,
            findings=(),
            changed_files=(),
        )

    before_status = git_status_files(repo_root) if fix else ()
    command = (
        (npm, "run", MARKDOWN_PACKAGE_SCRIPT, "--", "--fix")
        if fix
        else (
            npm,
            "run",
            MARKDOWN_PACKAGE_SCRIPT,
        )
    )
    result = runner(command, repo_root)
    combined_output = "\n".join(part for part in (result.stdout, result.stderr) if part)
    findings = parse_markdownlint_findings(combined_output)
    changed_files: tuple[str, ...] = ()
    if fix:
        after_status = git_status_files(repo_root)
        changed_files = tuple(path for path in after_status if path not in before_status)
    return MarkdownlintReport(
        available=True,
        message="Markdownlint fixer completed." if fix else "Markdownlint report completed.",
        return_code=result.returncode,
        findings=findings,
        changed_files=changed_files,
    )


def print_markdownlint_report(
    report: MarkdownlintReport,
    *,
    fix: bool,
    stdout: TextIO,
) -> None:
    """Print Markdownlint report output."""
    heading = (
        "First-adoption Markdownlint fixer report"
        if fix
        else "First-adoption Markdownlint debt report"
    )
    print(f"{heading}:", file=stdout)
    if not report.available:
        print(f"  {report.message}", file=stdout)
        print(
            "  Install npm dependencies intentionally; this helper never installs tools.",
            file=stdout,
        )
        return

    print(f"  Exit code: {report.return_code}", file=stdout)
    print(f"  Findings parsed: {len(report.findings)}", file=stdout)
    rule_counts = Counter(finding.rule for finding in report.findings)
    file_counts = Counter(finding.path for finding in report.findings)
    if rule_counts:
        print("  Rule counts:", file=stdout)
        for rule, count in sorted(rule_counts.items(), key=lambda item: (-item[1], item[0])):
            print(f"    - {rule}: {count}", file=stdout)
    if file_counts:
        print("  File counts:", file=stdout)
        for path, count in sorted(file_counts.items(), key=lambda item: (-item[1], item[0])):
            print(f"    - {path}: {count}", file=stdout)
    if fix:
        print(
            "  Fix mode can touch files that Git attributes normalize to LF; "
            "review the line-ending report before accepting fixer edits.",
            file=stdout,
        )
        if report.changed_files:
            print("  Files changed after Markdown fixer:", file=stdout)
            for path in report.changed_files:
                print(f"    - {path}", file=stdout)
        else:
            print("  Files changed after Markdown fixer: none detected.", file=stdout)


def powershell_executable() -> str | None:
    """Return a PowerShell executable when one is available."""
    return shutil.which("pwsh") or shutil.which("powershell")


def is_powershell_analyzer_candidate_path(path: str) -> bool:
    """Return whether ``path`` has a PSScriptAnalyzer input suffix."""
    normalized_path = path.replace("\\", "/")
    return PurePosixPath(normalized_path).suffix.casefold() in PSSCRIPTANALYZER_CANDIDATE_SUFFIXES


def powershell_files(files: Sequence[str]) -> tuple[str, ...]:
    """Return discovered PowerShell analyzer input paths."""
    return tuple(path for path in files if is_powershell_analyzer_candidate_path(path))


def collect_powershell_candidate_paths(
    repo_root: Path,
    *,
    tracked_only: bool = False,
    include_ignored: bool = False,
    stdout: TextIO | None = None,
) -> tuple[str, ...]:
    """Collect raw Git lexical PowerShell analyzer candidate paths."""
    commands: list[tuple[str, ...]] = [
        (
            GIT_TRACKED_POWERSHELL_CANDIDATES_COMMAND
            if tracked_only
            else GIT_VISIBLE_POWERSHELL_CANDIDATES_COMMAND
        )
    ]
    if include_ignored and not tracked_only:
        commands.append(GIT_IGNORED_POWERSHELL_CANDIDATES_COMMAND)

    # PowerShell analyzer policy is intentionally separate from the generic
    # first-adoption scan exclusions; the PowerShell helper classifies policy
    # exclusions such as node_modules dependency shims.
    candidates: list[str] = []
    seen_paths: set[str] = set()
    for command in commands:
        if stdout is not None:
            print(f"$ {format_command(command)}", file=stdout, flush=True)
        output = git_capture_bytes(repo_root, command[1:])
        for relative_path in decode_git_nul_paths(output):
            if relative_path in seen_paths:
                continue
            seen_paths.add(relative_path)
            if is_powershell_analyzer_candidate_path(relative_path):
                candidates.append(relative_path)

    return tuple(sorted(candidates))


def single_line_text(value: object) -> str:
    """Return a human-facing single-line representation of ``value``."""
    return str(value).replace("\r", r"\r").replace("\n", r"\n")


def _candidate_record_path(record: Mapping[str, object]) -> str:
    """Return the candidate path used in human-facing summaries."""
    path = record.get("RepositoryRelativePath")
    if path is None:
        path = record.get("CandidateFullName", "")
    return single_line_text(path)


def _candidate_records(data: Mapping[str, object], key: str) -> tuple[Mapping[str, object], ...]:
    """Return a validated candidate record array from composite data."""
    raw_records = data.get(key)
    if not isinstance(raw_records, list):
        raise FirstAdoptionQualityError(f"Candidates.{key} must be an array.")
    records: list[Mapping[str, object]] = []
    for index, raw_record in enumerate(raw_records, start=1):
        if not isinstance(raw_record, dict):
            raise FirstAdoptionQualityError(f"Candidates.{key}[{index}] must be an object.")
        records.append(cast(Mapping[str, object], raw_record))
    return tuple(records)


def candidate_summary_lines(candidates_data: Mapping[str, object]) -> tuple[str, ...]:
    """Build compact, escaped candidate summary lines for human output."""
    selected = _candidate_records(candidates_data, "Selected")
    policy_excluded = _candidate_records(candidates_data, "PolicyExcluded")
    unsafe = _candidate_records(candidates_data, "Unsafe")
    summary_counts = candidates_data.get("SummaryCounts")
    if not isinstance(summary_counts, dict):
        raise FirstAdoptionQualityError("Candidates.SummaryCounts must be an object.")
    expected_counts = {
        "Selected": len(selected),
        "PolicyExcluded": len(policy_excluded),
        "Unsafe": len(unsafe),
    }
    for key, expected_count in expected_counts.items():
        if summary_counts.get(key) != expected_count:
            raise FirstAdoptionQualityError(
                f"Candidates.SummaryCounts.{key} does not match the {key} array length."
            )

    lines = [
        (
            "PSScriptAnalyzer candidates: "
            f"{len(selected)} selected; {len(policy_excluded)} policy-excluded; "
            f"{len(unsafe)} unsafe."
        )
    ]
    if not selected and policy_excluded and not unsafe:
        lines.append(
            "No analyzer inputs were selected; "
            f"{len(policy_excluded)} candidate path(s) were excluded by policy."
        )
    if policy_excluded:
        reason_counts = Counter(
            str(record.get("ReasonCode", "Unknown")) for record in policy_excluded
        )
        reason_summary = "; ".join(
            f"{single_line_text(reason)} ({count})"
            for reason, count in sorted(
                reason_counts.items(),
                key=lambda item: (-item[1], item[0]),
            )
        )
        lines.append(f"Policy-excluded reason counts: {reason_summary}.")
        examples = "; ".join(_candidate_record_path(record) for record in policy_excluded[:3])
        lines.append(f"Policy-excluded candidates: {len(policy_excluded)} (examples: {examples}).")
    for record in unsafe:
        reason = single_line_text(record.get("ReasonCode", "Unknown"))
        line = f"Unsafe candidate: {_candidate_record_path(record)}; reason: {reason}"
        if "ResolvedTargetFullName" in record:
            line = f"{line}; resolved target: {single_line_text(record['ResolvedTargetFullName'])}"
        lines.append(line)
    return tuple(lines)


def powershell_candidate_stdin(script_paths: Sequence[str]) -> str:
    """Serialize PowerShell candidate records for the embedded runner stdin."""
    records = [{"RepositoryRelativePath": path} for path in script_paths]
    return json.dumps(records, ensure_ascii=False)


def _string_sequence(value: object) -> tuple[str, ...] | None:
    """Return a tuple of strings when ``value`` is a JSON string array."""
    if isinstance(value, str):
        return tuple(value.splitlines())
    if not isinstance(value, list):
        return None
    if not all(isinstance(item, str) for item in value):
        return None
    return tuple(str(item) for item in value)


def manual_review_powershell_report(
    *,
    message: str,
    candidate_summary_lines: Sequence[str],
    unsafe_candidate_count: int,
    summary_lines: Sequence[str],
    available: bool = False,
) -> PowerShellAnalyzerReport:
    """Return a PowerShell report that requires manual analyzer review."""
    return PowerShellAnalyzerReport(
        available=available,
        message=message,
        candidate_summary_lines=tuple(candidate_summary_lines),
        unsafe_candidate_count=unsafe_candidate_count,
        summary_lines=tuple(summary_lines),
        issue_ready_markdown=(),
        analyzer_debt_records=(),
        opt_in_guidance_lines=(),
    )


def normalize_gate_mode_static_value(value: object) -> GateModeStaticValue:
    """Normalize a committed static PSScriptAnalyzer gate-mode value."""
    if value is None:
        return GateModeStaticValue(
            raw_value=value,
            display_value="unset",
            recognized_mode=None,
            note="missing or empty values resolve to strict at runtime",
            manual_review=False,
        )
    if isinstance(value, bool):
        bool_text = "true" if value else "false"
        return GateModeStaticValue(
            raw_value=value,
            display_value=f"manual-review YAML boolean {bool_text}",
            recognized_mode=None,
            note=(
                "YAML parsed this value as a boolean; quote the intended mode "
                "string before treating it as a gate mode"
            ),
            manual_review=True,
        )
    if not isinstance(value, str):
        return GateModeStaticValue(
            raw_value=value,
            display_value=f"manual-review {type(value).__name__}",
            recognized_mode=None,
            note="non-string values require manual review before gate-mode advice",
            manual_review=True,
        )

    trimmed_value = value.strip()
    if not trimmed_value:
        return GateModeStaticValue(
            raw_value=value,
            display_value="unset",
            recognized_mode=None,
            note="missing or empty values resolve to strict at runtime",
            manual_review=False,
        )

    normalized_mode = trimmed_value.lower()
    if normalized_mode in PSSCRIPTANALYZER_RECOGNIZED_GATE_MODES:
        return GateModeStaticValue(
            raw_value=value,
            display_value=trimmed_value,
            recognized_mode=normalized_mode,
            note="recognized static gate mode",
            manual_review=False,
        )

    return GateModeStaticValue(
        raw_value=value,
        display_value=trimmed_value,
        recognized_mode=None,
        note="unrecognized strings resolve to strict at runtime",
        manual_review=False,
    )


def format_gate_mode_value(value: GateModeStaticValue) -> str:
    """Return a concise human-facing static gate-mode value description."""
    return f"`{single_line_text(value.display_value)}` ({value.note})"


def load_ci_yaml_mapping(
    repo_root: Path,
    relative_path: str,
    *,
    platform_name: str,
) -> tuple[Mapping[str, object] | None, tuple[str, ...]]:
    """Safely load a CI YAML surface as a mapping for non-mutating guidance."""
    path = repo_root / relative_path
    # Mirror load_marker_template_sync(): a broken symlink has ``exists()`` False
    # (it follows the link to a missing target), so guard on ``is_symlink()`` too.
    # Otherwise a symlink placed at the CI YAML path would be silently ignored
    # instead of falling through to the not-a-regular-file manual-review note.
    if not path.exists() and not path.is_symlink():
        return None, ()
    if not is_present_regular_file(path):
        return None, (
            (
                f"{platform_name}: {relative_path} is not a regular YAML file; "
                "manual review is required."
            ),
        )
    try:
        document = yaml.safe_load(path.read_text(encoding="utf-8-sig"))
    except yaml.YAMLError:
        return None, (
            f"{platform_name}: {relative_path} is not valid YAML; manual review is required.",
        )
    except UnicodeDecodeError:
        return None, (
            f"{platform_name}: {relative_path} is not valid UTF-8; manual review is required.",
        )
    except OSError as error:
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        return None, (
            (
                f"{platform_name}: unable to read {relative_path} ({error_summary}); "
                "manual review is required."
            ),
        )

    if not isinstance(document, dict):
        return None, (
            (
                f"{platform_name}: {relative_path} must contain a YAML mapping; "
                "manual review is required."
            ),
        )
    return cast(Mapping[str, object], document), ()


def yaml_contains_gate_mode_consumption(value: object) -> bool:
    """Return whether parsed YAML preserves the analyzer gate-mode call shape."""
    if isinstance(value, str):
        return (
            "Resolve-PSScriptAnalyzerGate" in value
            and f"-Mode $env:{PSSCRIPTANALYZER_GATE_MODE_VARIABLE}" in value
        )
    if isinstance(value, dict):
        return any(yaml_contains_gate_mode_consumption(item) for item in value.values())
    if isinstance(value, list):
        return any(yaml_contains_gate_mode_consumption(item) for item in value)
    return False


def gate_mode_setting_from_env(
    env_data: object,
    *,
    path: str,
    location: str,
    specificity: str,
) -> tuple[GateModeSetting | None, str | None]:
    """Return a gate-mode setting from an ``env`` mapping, if present."""
    if env_data is None:
        return None, None
    if not isinstance(env_data, dict):
        return (
            None,
            f"{path} `{location}` env is not a mapping; manual review is required.",
        )
    env_mapping = cast(Mapping[object, object], env_data)
    if PSSCRIPTANALYZER_GATE_MODE_VARIABLE not in env_mapping:
        return None, None
    return (
        GateModeSetting(
            path=path,
            location=f"{location}.{PSSCRIPTANALYZER_GATE_MODE_VARIABLE}",
            specificity=specificity,
            value=normalize_gate_mode_static_value(
                env_mapping[PSSCRIPTANALYZER_GATE_MODE_VARIABLE]
            ),
        ),
        None,
    )


def github_step_location(job_id: str, step: Mapping[str, object], index: int) -> str:
    """Return a stable display location for a GitHub Actions step."""
    step_name = step.get("name")
    if isinstance(step_name, str) and step_name.strip():
        step_label = single_line_text(step_name.strip())
    else:
        step_label = str(index)
    return f"jobs.{job_id}.steps[{step_label}].env"


def github_actions_gate_mode_settings(
    document: Mapping[str, object],
    *,
    path: str,
) -> tuple[tuple[GateModeSetting, ...], tuple[str, ...], bool]:
    """Return GitHub Actions static gate-mode settings and manual-review notes."""
    notes: list[str] = []
    retained = yaml_contains_gate_mode_consumption(document)
    if not retained:
        return (), (), False

    settings: list[GateModeSetting] = []
    workflow_setting, note = gate_mode_setting_from_env(
        document.get("env"),
        path=path,
        location="env",
        specificity="workflow env",
    )
    if note is not None:
        notes.append(f"GitHub Actions: {note}")
    if workflow_setting is not None:
        settings.append(workflow_setting)

    jobs = document.get("jobs")
    if not isinstance(jobs, dict):
        notes.append("GitHub Actions: jobs is missing or not a mapping; manual review is required.")
        return tuple(settings), tuple(notes), retained

    for raw_job_id, raw_job in cast(Mapping[object, object], jobs).items():
        if not isinstance(raw_job_id, str) or not isinstance(raw_job, dict):
            continue
        job = cast(Mapping[str, object], raw_job)
        if not yaml_contains_gate_mode_consumption(job):
            continue
        job_setting, note = gate_mode_setting_from_env(
            job.get("env"),
            path=path,
            location=f"jobs.{raw_job_id}.env",
            specificity="job env",
        )
        if note is not None:
            notes.append(f"GitHub Actions: {note}")
        if job_setting is not None:
            settings.append(job_setting)

        steps = job.get("steps")
        if not isinstance(steps, list):
            notes.append(
                f"GitHub Actions: {path} `jobs.{raw_job_id}.steps` is not a list; "
                "manual review is required."
            )
            continue
        for index, raw_step in enumerate(steps, start=1):
            if not isinstance(raw_step, dict):
                continue
            step = cast(Mapping[str, object], raw_step)
            if not yaml_contains_gate_mode_consumption(step):
                continue
            step_setting, note = gate_mode_setting_from_env(
                step.get("env"),
                path=path,
                location=github_step_location(raw_job_id, step, index),
                specificity="step env",
            )
            if note is not None:
                notes.append(f"GitHub Actions: {note}")
            if step_setting is not None:
                settings.append(step_setting)

    return tuple(settings), tuple(notes), retained


def azure_parameter_default_result(
    *,
    path: str,
    location: str,
    raw_default: object,
) -> tuple[GateModeSetting | None, tuple[str, ...]]:
    """Return a parameter-default setting, or a manual-review note when no default exists.

    An Azure Pipelines ``parameters`` entry declared without a ``default`` (or with a
    null default) must be supplied when the pipeline runs, so its value is unknown at
    static-analysis time. Reporting it as ``unset`` (which for env vars means "resolves
    to strict at runtime") would be misleading, so surface it for manual review instead.
    """
    if raw_default is None:
        return None, (
            (
                f"Azure Pipelines: {path} `{location}` has no static default; the value must "
                "be supplied when the pipeline runs, so manual review is required."
            ),
        )
    return (
        GateModeSetting(
            path=path,
            location=location,
            specificity="pipeline parameter default",
            value=normalize_gate_mode_static_value(raw_default),
        ),
        (),
    )


def azure_parameter_default_setting(
    document: Mapping[str, object],
    *,
    path: str,
) -> tuple[GateModeSetting | None, tuple[str, ...]]:
    """Return the Azure Pipelines ``gateMode`` parameter default, if present."""
    parameters = document.get("parameters")
    if parameters is None:
        return None, ()
    if isinstance(parameters, list):
        for raw_parameter in parameters:
            if not isinstance(raw_parameter, dict):
                continue
            parameter = cast(Mapping[object, object], raw_parameter)
            if parameter.get("name") != PSSCRIPTANALYZER_GATE_MODE_PARAMETER:
                continue
            return azure_parameter_default_result(
                path=path,
                location=f"parameters[{PSSCRIPTANALYZER_GATE_MODE_PARAMETER}].default",
                raw_default=parameter.get("default"),
            )
        return None, ()
    if isinstance(parameters, dict):
        parameters_mapping = cast(Mapping[object, object], parameters)
        if PSSCRIPTANALYZER_GATE_MODE_PARAMETER not in parameters_mapping:
            return None, ()
        parameter_value = parameters_mapping.get(PSSCRIPTANALYZER_GATE_MODE_PARAMETER)
        if isinstance(parameter_value, dict):
            raw_default = cast(Mapping[object, object], parameter_value).get("default")
        else:
            raw_default = parameter_value
        return azure_parameter_default_result(
            path=path,
            location=f"parameters.{PSSCRIPTANALYZER_GATE_MODE_PARAMETER}.default",
            raw_default=raw_default,
        )
    return None, (
        (
            f"Azure Pipelines: {path} `parameters` is not a mapping or sequence; "
            "manual review is required."
        ),
    )


def azure_variable_settings_from_variables(
    variables: object,
    *,
    path: str,
    location: str,
) -> tuple[tuple[GateModeSetting, ...], tuple[str, ...]]:
    """Return Azure variable settings from mapping or sequence variable syntax."""
    location_prefix = f"{location}." if location else ""
    if variables is None:
        return (), ()
    if isinstance(variables, dict):
        variable_mapping = cast(Mapping[object, object], variables)
        if PSSCRIPTANALYZER_GATE_MODE_VARIABLE not in variable_mapping:
            return (), ()
        return (
            (
                GateModeSetting(
                    path=path,
                    location=(f"{location_prefix}variables.{PSSCRIPTANALYZER_GATE_MODE_VARIABLE}"),
                    specificity="pipeline variable",
                    value=normalize_gate_mode_static_value(
                        variable_mapping[PSSCRIPTANALYZER_GATE_MODE_VARIABLE]
                    ),
                ),
            ),
            (),
        )
    if isinstance(variables, list):
        settings: list[GateModeSetting] = []
        for raw_variable in variables:
            if not isinstance(raw_variable, dict):
                continue
            variable = cast(Mapping[object, object], raw_variable)
            if variable.get("name") != PSSCRIPTANALYZER_GATE_MODE_VARIABLE:
                continue
            settings.append(
                GateModeSetting(
                    path=path,
                    location=(
                        f"{location_prefix}variables"
                        f"[{PSSCRIPTANALYZER_GATE_MODE_VARIABLE}].value"
                    ),
                    specificity="pipeline variable",
                    value=normalize_gate_mode_static_value(variable.get("value")),
                )
            )
        return tuple(settings), ()
    return (), (
        (
            f"Azure Pipelines: {path} `{location_prefix}variables` is not a mapping or sequence; "
            "manual review is required."
        ),
    )


def iter_azure_mapping_children(
    document: Mapping[str, object],
    *,
    location: str,
) -> tuple[tuple[str, Mapping[str, object]], ...]:
    """Return nested Azure Pipelines mappings that may contain variables."""
    children: list[tuple[str, Mapping[str, object]]] = []
    for collection_key in ("stages", "jobs"):
        raw_collection = document.get(collection_key)
        location_prefix = f"{location}." if location else ""
        if isinstance(raw_collection, dict):
            for raw_name, raw_child in cast(Mapping[object, object], raw_collection).items():
                if isinstance(raw_name, str) and isinstance(raw_child, dict):
                    children.append(
                        (
                            f"{location_prefix}{collection_key}[{raw_name}]",
                            cast(Mapping[str, object], raw_child),
                        )
                    )
        elif isinstance(raw_collection, list):
            for index, raw_child in enumerate(raw_collection, start=1):
                if not isinstance(raw_child, dict):
                    continue
                child = cast(Mapping[str, object], raw_child)
                raw_name = child.get("stage") if collection_key == "stages" else child.get("job")
                child_name = single_line_text(raw_name) if isinstance(raw_name, str) else str(index)
                children.append((f"{location_prefix}{collection_key}[{child_name}]", child))
    return tuple(children)


def azure_variable_settings(
    document: Mapping[str, object],
    *,
    path: str,
) -> tuple[tuple[GateModeSetting, ...], tuple[str, ...]]:
    """Return Azure Pipelines gate-mode variable settings across nested scopes."""
    settings: list[GateModeSetting] = []
    notes: list[str] = []
    stack: list[tuple[str, Mapping[str, object]]] = [("", document)]
    while stack:
        location, current = stack.pop()
        scope_settings, scope_notes = azure_variable_settings_from_variables(
            current.get("variables"),
            path=path,
            location=location,
        )
        settings.extend(scope_settings)
        notes.extend(scope_notes)
        stack.extend(reversed(iter_azure_mapping_children(current, location=location)))
    return tuple(settings), tuple(notes)


def resolve_azure_variable_effective_value(
    setting: GateModeSetting,
    parameter_setting: GateModeSetting | None,
) -> GateModeSetting:
    """Resolve Azure variable expressions that point at the gate-mode parameter."""
    if (
        isinstance(setting.value.raw_value, str)
        and setting.value.raw_value.strip() == PSSCRIPTANALYZER_GATE_MODE_PARAMETER_EXPRESSION
        and parameter_setting is not None
    ):
        return GateModeSetting(
            path=setting.path,
            location=setting.location,
            specificity=setting.specificity,
            value=setting.value,
            effective_value=parameter_setting.value,
        )
    return setting


def azure_pipelines_gate_mode_settings(
    document: Mapping[str, object],
    *,
    path: str,
) -> tuple[GateModeSetting | None, tuple[GateModeSetting, ...], tuple[str, ...], bool]:
    """Return Azure Pipelines static gate-mode settings and manual-review notes."""
    retained = yaml_contains_gate_mode_consumption(document)
    if not retained:
        return None, (), (), False

    parameter_setting, parameter_notes = azure_parameter_default_setting(document, path=path)
    variable_settings, variable_notes = azure_variable_settings(document, path=path)
    resolved_variable_settings = tuple(
        resolve_azure_variable_effective_value(setting, parameter_setting)
        for setting in variable_settings
    )
    return (
        parameter_setting,
        resolved_variable_settings,
        (*parameter_notes, *variable_notes),
        retained,
    )


def setting_effective_value(setting: GateModeSetting) -> GateModeStaticValue:
    """Return the value that controls a setting after static indirection."""
    return setting.effective_value or setting.value


def format_static_setting_line(platform_name: str, setting: GateModeSetting) -> str:
    """Format one static gate-mode setting for report output."""
    line = (
        f"{platform_name}: {setting.path} `{setting.location}` "
        f"({setting.specificity}) is {format_gate_mode_value(setting.value)}."
    )
    if setting.effective_value is not None:
        line = (
            f"{line} Effective static default is "
            f"{format_gate_mode_value(setting.effective_value)}."
        )
    return line


def recommended_setting_location(
    settings: Sequence[GateModeSetting],
) -> GateModeSetting | None:
    """Return the most specific setting whose value should be changed."""
    if not settings:
        return None
    for setting in reversed(settings):
        effective_value = setting_effective_value(setting)
        if not effective_value.manual_review:
            return setting
    return settings[-1]


def github_actions_opt_in_guidance(
    repo_root: Path,
    *,
    recommend_first_adoption: bool,
    has_blocking_findings: bool,
) -> tuple[str, ...]:
    """Build GitHub Actions first-adoption opt-in guidance lines."""
    document, notes = load_ci_yaml_mapping(
        repo_root,
        GITHUB_ACTIONS_POWERSHELL_CI,
        platform_name="GitHub Actions",
    )
    if document is None:
        return notes

    settings, setting_notes, retained = github_actions_gate_mode_settings(
        document,
        path=GITHUB_ACTIONS_POWERSHELL_CI,
    )
    if not retained:
        return ()

    lines = [*notes, *setting_notes]
    lines.extend(format_static_setting_line("GitHub Actions", setting) for setting in settings)
    if not settings:
        lines.append(
            "GitHub Actions: retained analyzer step found, but no static "
            f"`{PSSCRIPTANALYZER_GATE_MODE_VARIABLE}` workflow/job/step env entry "
            "was found; missing values resolve to strict at runtime."
        )

    if recommend_first_adoption:
        target_setting = recommended_setting_location(settings)
        if target_setting is None:
            lines.append(
                "GitHub Actions recommendation: manually add a retained static "
                f"`{PSSCRIPTANALYZER_GATE_MODE_VARIABLE}: first-adoption` env value "
                f"in {GITHUB_ACTIONS_POWERSHELL_CI} for the analyzer job in a later PR."
            )
        elif setting_effective_value(target_setting).manual_review:
            lines.append(
                "GitHub Actions recommendation: no opt-in edit is suggested until "
                f"`{target_setting.location}` is reviewed because its committed value "
                "was not a plain mode string."
            )
        elif setting_effective_value(target_setting).recognized_mode == "first-adoption":
            lines.append(
                "GitHub Actions recommendation: the retained static gate-mode value "
                "already resolves to `first-adoption`; no opt-in edit is suggested."
            )
        else:
            lines.append(
                "GitHub Actions recommendation: manually set "
                f"{GITHUB_ACTIONS_POWERSHELL_CI} `{target_setting.location}` to "
                "`first-adoption` in a later implementation PR."
            )
    elif has_blocking_findings:
        lines.append(
            "GitHub Actions note: blocking Error or unknown-severity findings remain "
            "blocking in first-adoption mode; resolve them before expecting CI to pass."
        )

    lines.append(
        "GitHub Actions runtime overrides, including step-written environment changes, "
        "can still change effective behavior; preserve "
        "`Resolve-PSScriptAnalyzerGate -Mode $env:PSSCRIPTANALYZER_GATE_MODE`."
    )
    return tuple(lines)


def azure_pipelines_opt_in_guidance(
    repo_root: Path,
    *,
    recommend_first_adoption: bool,
    has_blocking_findings: bool,
) -> tuple[str, ...]:
    """Build Azure Pipelines first-adoption opt-in guidance lines."""
    document, notes = load_ci_yaml_mapping(
        repo_root,
        AZURE_PIPELINES_POWERSHELL_CI,
        platform_name="Azure Pipelines",
    )
    if document is None:
        return notes

    parameter_setting, variable_settings, setting_notes, retained = (
        azure_pipelines_gate_mode_settings(
            document,
            path=AZURE_PIPELINES_POWERSHELL_CI,
        )
    )
    if not retained:
        return ()

    lines = [*notes, *setting_notes]
    if parameter_setting is not None:
        lines.append(format_static_setting_line("Azure Pipelines", parameter_setting))
    else:
        lines.append(
            "Azure Pipelines: retained analyzer step found, but no static "
            "`gateMode` parameter default was found; manual review is required."
        )

    if variable_settings:
        lines.extend(
            format_static_setting_line("Azure Pipelines", setting) for setting in variable_settings
        )
    else:
        lines.append(
            "Azure Pipelines: retained analyzer step found, but no static "
            f"`{PSSCRIPTANALYZER_GATE_MODE_VARIABLE}` variables entry was found; "
            "manual review is required."
        )

    recommendation_target = parameter_setting
    if recommend_first_adoption:
        if recommendation_target is not None and not recommendation_target.value.manual_review:
            if recommendation_target.value.recognized_mode == "first-adoption":
                lines.append(
                    "Azure Pipelines recommendation: the retained static parameter "
                    "default already resolves to `first-adoption`; no opt-in edit is suggested."
                )
            else:
                lines.append(
                    "Azure Pipelines recommendation: manually set "
                    f"{AZURE_PIPELINES_POWERSHELL_CI} `{recommendation_target.location}` "
                    "to `first-adoption` in a later implementation PR."
                )
        else:
            lines.append(
                "Azure Pipelines recommendation: no opt-in edit is suggested until "
                "the committed `gateMode` parameter default is manually reviewed."
            )
    elif has_blocking_findings:
        lines.append(
            "Azure Pipelines note: blocking Error or unknown-severity findings remain "
            "blocking in first-adoption mode; resolve them before expecting CI to pass."
        )

    lines.append(
        "Azure Pipelines auto-exports non-secret variables as environment variables "
        "with names uppercased and periods changed to underscores, so the analyzer "
        "task can read `$env:PSSCRIPTANALYZER_GATE_MODE` without a task `env` block."
    )
    lines.append(
        "Azure Pipelines queued-run parameters, runtime variables, and branch-policy "
        "build validation can still change effective behavior; preserve "
        "`Resolve-PSScriptAnalyzerGate -Mode $env:PSSCRIPTANALYZER_GATE_MODE`."
    )
    return tuple(lines)


def gate_bool_field(gate_data: Mapping[str, object], key: str) -> bool | None:
    """Return a required Boolean gate field or ``None`` when malformed."""
    value = gate_data.get(key)
    return value if isinstance(value, bool) else None


def gate_string_field(gate_data: Mapping[str, object], key: str) -> str | None:
    """Return a required non-empty string gate field or ``None`` when malformed."""
    value = gate_data.get(key)
    if not isinstance(value, str) or not value.strip():
        return None
    return value.strip()


def positive_integer_from_gate_field(value: object) -> int | None:
    """Return a positive integer gate field value or ``None`` when absent/zero."""
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError("line and column values must be integers when present.")
    return value if value > 0 else None


def required_finding_string(
    finding: Mapping[str, object],
    *,
    key: str,
    index: int,
) -> str:
    """Return a required normalized string field from one gate finding."""
    value = finding.get(key)
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"Gate.Findings[{index}].{key} must be a non-empty string.")
    return single_line_text(value.strip())


def analyzer_debt_records_from_gate_findings(
    raw_findings: object,
) -> tuple[PowerShellAnalyzerDebtRecord, ...] | None:
    """Build analyzer-debt records from structured gate findings."""
    if not isinstance(raw_findings, list):
        return None

    records: list[PowerShellAnalyzerDebtRecord] = []
    try:
        for index, raw_finding in enumerate(raw_findings, start=1):
            if not isinstance(raw_finding, dict):
                raise TypeError(f"Gate.Findings[{index}] must be an object.")
            finding = cast(Mapping[str, object], raw_finding)
            tracked_debt = finding.get("TrackedDebt")
            if not isinstance(tracked_debt, bool):
                raise TypeError(f"Gate.Findings[{index}].TrackedDebt must be a Boolean.")
            if not tracked_debt:
                continue

            severity = required_finding_string(finding, key="Severity", index=index)
            if severity not in ("Warning", "Information"):
                raise ValueError(f"Gate.Findings[{index}].TrackedDebt used unsupported severity.")

            records.append(
                PowerShellAnalyzerDebtRecord(
                    rule_name=required_finding_string(finding, key="RuleName", index=index),
                    severity=severity,
                    normalized_path=required_finding_string(
                        finding,
                        key="ScriptPath",
                        index=index,
                    ),
                    line=positive_integer_from_gate_field(finding.get("Line")),
                    column=positive_integer_from_gate_field(finding.get("Column")),
                    message=required_finding_string(finding, key="Message", index=index),
                )
            )
    except (TypeError, ValueError):
        return None

    return tuple(records)


def analyzer_debt_summary_lines(
    records: Sequence[PowerShellAnalyzerDebtRecord],
) -> tuple[str, ...]:
    """Return per-run debt summary lines aligned with documented placeholders."""
    if not records:
        return ()

    severity_counts = Counter(record.severity for record in records)
    rule_counts = Counter(record.rule_name for record in records)
    affected_rules = "; ".join(f"{rule} ({count})" for rule, count in sorted(rule_counts.items()))
    return (
        (
            "Debt recording summary: "
            f"{len(records)} tracked finding(s); "
            f"Warning={severity_counts['Warning']}; "
            f"Information={severity_counts['Information']}."
        ),
        f"Affected rules: {affected_rules}.",
        "Owner: <owner>; expected removal date: <YYYY-MM-DD>.",
        "Temporary first-adoption mode is a bridge; restore strict mode after cleanup.",
    )


def build_powershell_opt_in_guidance(
    repo_root: Path,
    *,
    analyzer_debt_records: Sequence[PowerShellAnalyzerDebtRecord],
    recommended_mode: str,
    should_fail: bool,
) -> tuple[str, ...]:
    """Build non-mutating first-adoption CI opt-in guidance."""
    if not analyzer_debt_records and recommended_mode != "first-adoption":
        return ()

    recommend_first_adoption = recommended_mode == "first-adoption"
    lines = [
        (
            "These are recommendations to apply manually or in a later implementation PR; "
            "this report did not mutate repository files."
        )
    ]
    if analyzer_debt_records and not recommend_first_adoption and should_fail:
        lines.append(
            "Tracked Warning/Information debt was found, but the gate helper did not "
            "recommend first-adoption for this run because blocking findings are present."
        )

    lines.extend(
        github_actions_opt_in_guidance(
            repo_root,
            recommend_first_adoption=recommend_first_adoption,
            has_blocking_findings=should_fail,
        )
    )
    lines.extend(
        azure_pipelines_opt_in_guidance(
            repo_root,
            recommend_first_adoption=recommend_first_adoption,
            has_blocking_findings=should_fail,
        )
    )
    return tuple(lines) if len(lines) > 1 else ()


def build_powershell_report(
    repo_root: Path,
    *,
    tracked_only: bool = False,
    include_ignored: bool = False,
    runner: CaptureRunner = run_capture,
) -> PowerShellAnalyzerReport:
    """Run an optional PSScriptAnalyzer debt report when tooling is available."""
    script_paths = collect_powershell_candidate_paths(
        repo_root,
        tracked_only=tracked_only,
        include_ignored=include_ignored,
    )
    if not script_paths:
        return PowerShellAnalyzerReport(
            available=True,
            message="No PowerShell files were discovered.",
            candidate_summary_lines=(),
            unsafe_candidate_count=0,
            summary_lines=("No PowerShell files were discovered.",),
            issue_ready_markdown=(),
            analyzer_debt_records=(),
            opt_in_guidance_lines=(),
        )

    settings_path = repo_root / PSSCRIPTANALYZER_SETTINGS
    candidate_helper_path = repo_root / PSSCRIPTANALYZER_CANDIDATE_HELPER
    gate_helper_path = repo_root / PSSCRIPTANALYZER_GATE_HELPER
    if (
        not is_present_regular_file(settings_path)
        or not is_present_regular_file(candidate_helper_path)
        or not is_present_regular_file(gate_helper_path)
    ):
        return PowerShellAnalyzerReport(
            available=False,
            message=(
                "PSScriptAnalyzer report unavailable: retained PowerShell settings "
                "or helper files are missing."
            ),
            candidate_summary_lines=(),
            unsafe_candidate_count=0,
            summary_lines=(),
            issue_ready_markdown=(),
            analyzer_debt_records=(),
            opt_in_guidance_lines=(),
        )

    executable = powershell_executable()
    if executable is None:
        return PowerShellAnalyzerReport(
            available=False,
            message=(
                "PSScriptAnalyzer report unavailable: neither pwsh nor powershell "
                "was found on PATH."
            ),
            candidate_summary_lines=(),
            unsafe_candidate_count=0,
            summary_lines=(),
            issue_ready_markdown=(),
            analyzer_debt_records=(),
            opt_in_guidance_lines=(),
        )

    command_script = r"""
$ErrorActionPreference = 'Stop'
$utf8NoBomEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = $utf8NoBomEncoding

function Write-CompositeJson {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Composite
    )

    $json = $Composite | ConvertTo-Json -Depth 12
    $json = $json -replace "`r`n", "`n"
    $writer = [System.IO.StreamWriter]::new(
        [Console]::OpenStandardOutput(),
        $utf8NoBomEncoding
    )
    try {
        $writer.Write($json)
    } finally {
        $writer.Flush()
        $writer.Dispose()
    }
}

function ConvertTo-RunnerSingleLineText {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    $text = [string]$Value
    $text = $text.Replace("`r", '\r')
    $text = $text.Replace("`n", '\n')
    return $text
}

try {
    $repositoryRoot = (Get-Location).ProviderPath
    . (Join-Path -Path $repositoryRoot -ChildPath 'src/tools/Resolve-PSScriptAnalyzerCandidate.ps1')
    . (Join-Path -Path $repositoryRoot -ChildPath 'src/tools/Resolve-PSScriptAnalyzerGate.ps1')
    $settingsPath = Join-Path -Path $repositoryRoot -ChildPath '.github/linting/PSScriptAnalyzerSettings.psd1'

    $reader = [System.IO.StreamReader]::new(
        [Console]::OpenStandardInput(),
        [System.Text.Encoding]::UTF8
    )
    try {
        $stdinPayload = $reader.ReadToEnd()
    } finally {
        $reader.Dispose()
    }

    if ([string]::IsNullOrWhiteSpace($stdinPayload)) {
        throw 'Candidate JSON stdin is empty.'
    }

    if (-not $stdinPayload.TrimStart().StartsWith('[')) {
        throw 'Candidate JSON stdin must be a top-level array.'
    }

    $rawCandidates = @($stdinPayload | ConvertFrom-Json)
    $candidateRecords = [System.Collections.Generic.List[object]]::new()
    foreach ($rawCandidate in $rawCandidates) {
        if (
            ($null -eq $rawCandidate) -or
            ($null -eq $rawCandidate.PSObject.Properties['RepositoryRelativePath']) -or
            (-not ($rawCandidate.RepositoryRelativePath -is [string])) -or
            [string]::IsNullOrEmpty($rawCandidate.RepositoryRelativePath)
        ) {
            throw 'Candidate JSON contains a malformed candidate record.'
        }

        $repositoryRelativePath = [string]$rawCandidate.RepositoryRelativePath
        [void]($candidateRecords.Add(
                (Resolve-PSScriptAnalyzerCandidate `
                    -RepositoryRoot $repositoryRoot `
                    -CandidatePath $repositoryRelativePath `
                    -RepositoryRelativePath $repositoryRelativePath)
            ))
    }

    $candidateSummary = Get-PSScriptAnalyzerCandidateSummary `
        -Candidate ([object[]]$candidateRecords.ToArray())

    if ($candidateSummary.SummaryCounts.Unsafe -gt 0) {
        $gate = [pscustomobject]@{
            Status = 'unavailable'
            Message = 'Analyzer was not run because unsafe candidates were found.'
            SummaryLines = [string[]]@()
            IssueReadyMarkdown = [string[]]@()
        }
    } elseif ($candidateSummary.SummaryCounts.Selected -eq 0) {
        $gate = Resolve-PSScriptAnalyzerGate `
            -Mode 'first-adoption' `
            -AnalyzerFinding @() `
            -RepositoryRoot $repositoryRoot `
            -AnnotationFormat Plain
        $gate | Add-Member -NotePropertyName Status -NotePropertyValue 'ok' -Force
    } else {
        try {
            Import-Module PSScriptAnalyzer -ErrorAction Stop
            $findings = [System.Collections.Generic.List[object]]::new()
            foreach ($candidate in @($candidateSummary.Selected)) {
                $findings.AddRange(@(
                    Invoke-ScriptAnalyzer `
                        -Path $candidate.EscapedAnalyzerPath `
                        -Settings $settingsPath
                ))
            }
            $gate = Resolve-PSScriptAnalyzerGate `
                -Mode 'first-adoption' `
                -AnalyzerFinding $findings `
                -RepositoryRoot $repositoryRoot `
                -AnnotationFormat Plain
            $gate | Add-Member -NotePropertyName Status -NotePropertyValue 'ok' -Force
        } catch {
            $gate = [pscustomobject]@{
                Status = 'unavailable'
                Message = (
                    'PSScriptAnalyzer report unavailable: run ' +
                    '`Install-Module PSScriptAnalyzer` intentionally if the PowerShell ' +
                    'module is retained. Diagnostic: ' +
                    (ConvertTo-RunnerSingleLineText -Value $_)
                )
                SummaryLines = [string[]]@()
                IssueReadyMarkdown = [string[]]@()
            }
        }
    }

    $composite = [pscustomobject]@{
        Gate = $gate
        Candidates = $candidateSummary
    }
    Write-CompositeJson -Composite $composite
    exit 0
} catch {
    Write-Error -Message (ConvertTo-RunnerSingleLineText -Value $_)
    exit 1
}
"""
    result = runner(
        (executable, "-NoProfile", "-Command", command_script),
        repo_root,
        input=powershell_candidate_stdin(script_paths),
    )

    try:
        result_data = json.loads(result.stdout)
    except json.JSONDecodeError as error:
        if result.returncode != 0:
            diagnostic = (
                result.stderr.strip() or result.stdout.strip() or "PowerShell command failed"
            )
            raise FirstAdoptionQualityError(
                "PSScriptAnalyzer report runner failed before structured output "
                f"was available: {single_line_text(diagnostic)}"
            ) from error
        raise FirstAdoptionQualityError(
            "PSScriptAnalyzer report returned non-JSON output."
        ) from error

    if not isinstance(result_data, dict):
        raise FirstAdoptionQualityError("PSScriptAnalyzer report composite JSON must be an object.")
    result_data = cast(dict[str, object], result_data)
    candidates_data = result_data.get("Candidates")
    if not isinstance(candidates_data, dict):
        raise FirstAdoptionQualityError("PSScriptAnalyzer report is missing Candidates data.")
    candidates_data = cast(dict[str, object], candidates_data)
    rendered_candidate_summary = candidate_summary_lines(candidates_data)
    unsafe_candidate_count = len(_candidate_records(candidates_data, "Unsafe"))
    if unsafe_candidate_count:
        return PowerShellAnalyzerReport(
            available=True,
            message=(
                "PSScriptAnalyzer report stopped before analysis because unsafe "
                "candidate path(s) were found."
            ),
            candidate_summary_lines=rendered_candidate_summary,
            unsafe_candidate_count=unsafe_candidate_count,
            summary_lines=(),
            issue_ready_markdown=(),
            analyzer_debt_records=(),
            opt_in_guidance_lines=(),
        )

    gate_data = result_data.get("Gate")
    if not isinstance(gate_data, dict):
        return manual_review_powershell_report(
            message=(
                "PSScriptAnalyzer report completed, but the Gate object was malformed; "
                "manual review is required."
            ),
            candidate_summary_lines=rendered_candidate_summary,
            unsafe_candidate_count=unsafe_candidate_count,
            summary_lines=("Manual review: PSScriptAnalyzer Gate object was malformed.",),
        )
    gate_data = cast(dict[str, object], gate_data)

    gate_status = gate_data.get("Status")
    if gate_status != "ok" and gate_status != "unavailable":
        return manual_review_powershell_report(
            message=(
                "PSScriptAnalyzer report completed, but Gate.Status was missing or "
                "malformed; manual review is required."
            ),
            candidate_summary_lines=rendered_candidate_summary,
            unsafe_candidate_count=0,
            summary_lines=("Manual review: PSScriptAnalyzer Gate.Status was malformed.",),
        )

    if gate_status == "unavailable":
        message = gate_data.get("Message")
        return PowerShellAnalyzerReport(
            available=False,
            message=str(message or "PSScriptAnalyzer report unavailable."),
            candidate_summary_lines=rendered_candidate_summary,
            unsafe_candidate_count=0,
            summary_lines=(),
            issue_ready_markdown=(),
            analyzer_debt_records=(),
            opt_in_guidance_lines=(),
        )

    summary_lines = _string_sequence(gate_data.get("SummaryLines"))
    issue_ready_lines = _string_sequence(gate_data.get("IssueReadyMarkdown"))
    findings = analyzer_debt_records_from_gate_findings(gate_data.get("Findings"))
    recommended_mode = gate_string_field(gate_data, "RecommendedMode")
    resolved_mode = gate_string_field(gate_data, "Mode")
    should_fail = gate_bool_field(gate_data, "ShouldFail")
    if (
        summary_lines is None
        or issue_ready_lines is None
        or findings is None
        or recommended_mode not in PSSCRIPTANALYZER_RECOGNIZED_GATE_MODES
        or resolved_mode not in PSSCRIPTANALYZER_RECOGNIZED_GATE_MODES
        or should_fail is None
    ):
        return manual_review_powershell_report(
            message=(
                "PSScriptAnalyzer report completed, but required Gate fields were "
                "missing or malformed; manual review is required."
            ),
            candidate_summary_lines=rendered_candidate_summary,
            unsafe_candidate_count=0,
            summary_lines=("Manual review: required PSScriptAnalyzer Gate fields were malformed.",),
        )

    opt_in_guidance_lines = build_powershell_opt_in_guidance(
        repo_root,
        analyzer_debt_records=findings,
        recommended_mode=recommended_mode,
        should_fail=should_fail,
    )

    return PowerShellAnalyzerReport(
        available=True,
        message="PSScriptAnalyzer report completed.",
        candidate_summary_lines=rendered_candidate_summary,
        unsafe_candidate_count=0,
        summary_lines=summary_lines,
        issue_ready_markdown=issue_ready_lines,
        analyzer_debt_records=findings,
        opt_in_guidance_lines=opt_in_guidance_lines,
    )


def format_analyzer_debt_record(record: PowerShellAnalyzerDebtRecord) -> str:
    """Return one single-line analyzer-debt record."""
    location = record.normalized_path
    if record.line is not None:
        location = f"{location}:{record.line}"
        if record.column is not None:
            location = f"{location}:{record.column}"
    return f"{location} [{record.severity}] {record.rule_name}: {single_line_text(record.message)}"


def print_powershell_report(report: PowerShellAnalyzerReport, *, stdout: TextIO) -> None:
    """Print a stable PowerShell analyzer report."""
    print("First-adoption PSScriptAnalyzer debt report:", file=stdout)
    if not report.available:
        print(f"  {report.message}", file=stdout)
        for line in report.candidate_summary_lines:
            print(f"  {line}", file=stdout)
        for line in report.summary_lines:
            print(f"  {line}", file=stdout)
        if not report.summary_lines:
            print("  This helper never installs optional tools automatically.", file=stdout)
        return
    print(f"  {report.message}", file=stdout)
    for line in report.candidate_summary_lines:
        print(f"  {line}", file=stdout)
    for line in report.summary_lines:
        print(f"  {line}", file=stdout)
    if report.analyzer_debt_records:
        print("  Analyzer debt records for manual tracking:", file=stdout)
        for line in analyzer_debt_summary_lines(report.analyzer_debt_records):
            print(f"    - {line}", file=stdout)
        print("    Findings:", file=stdout)
        for record in report.analyzer_debt_records:
            print(f"      - {format_analyzer_debt_record(record)}", file=stdout)
    if report.opt_in_guidance_lines:
        print("  Non-mutating first-adoption opt-in guidance:", file=stdout)
        for line in report.opt_in_guidance_lines:
            print(f"    - {line}", file=stdout)
    if report.issue_ready_markdown:
        print("  Issue-ready Markdown:", file=stdout)
        for line in report.issue_ready_markdown:
            print(f"    {line}", file=stdout)


def add_common_scope_args(parser: argparse.ArgumentParser) -> None:
    """Add common file-discovery scope arguments to ``parser``."""
    parser.add_argument(
        "--tracked-only",
        action="store_true",
        help="Scan only tracked files instead of tracked plus untracked non-ignored files.",
    )
    parser.add_argument(
        "--include-ignored",
        action="store_true",
        help="Also include ignored untracked files. Ignored files are excluded by default.",
    )


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Emit first-adoption quality debt reports without installing tools."
    )
    parser.add_argument(
        "--repo-root",
        default=None,
        help="Repository root. Defaults to the repository root containing this script.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    line_endings_parser = subparsers.add_parser(
        "line-endings",
        help="Inventory line endings and Git-attribute normalization risk.",
    )
    add_common_scope_args(line_endings_parser)

    path_references_parser = subparsers.add_parser(
        "path-references",
        help="Check documented path-reference surfaces for casing and missing targets.",
    )
    add_common_scope_args(path_references_parser)
    path_references_parser.add_argument(
        "--suppressions",
        default=DEFAULT_SUPPRESSION_PATH,
        help=f"Suppression file path. Default: {DEFAULT_SUPPRESSION_PATH}.",
    )
    path_references_parser.add_argument(
        "--fail-on-findings",
        action="store_true",
        help="Exit non-zero when unsuppressed path-reference findings are present.",
    )
    path_references_parser.add_argument(
        "--include-missing-targets",
        action="store_true",
        help=(
            "Also report path-like literals that do not match a discovered file "
            "or directory. By default, only casing mismatches are reported."
        ),
    )

    subparsers.add_parser(
        "host-setup",
        help="Report Azure DevOps Services first-adoption service setup follow-ups.",
    )

    markdown_parser = subparsers.add_parser(
        "markdown",
        help="Report Markdownlint debt or run the explicit Markdown fixer.",
    )
    markdown_parser.add_argument(
        "--fix",
        action="store_true",
        help="Run the lint:md fixer command and report changed files afterward.",
    )

    powershell_parser = subparsers.add_parser(
        "powershell",
        help="Report optional PSScriptAnalyzer debt when PowerShell tooling is available.",
    )
    add_common_scope_args(powershell_parser)

    return parser.parse_args(argv)


def run_report(args: argparse.Namespace, *, stdout: TextIO = sys.stdout) -> int:
    """Run the selected report and return a process exit code."""
    repo_root = resolve_repo_root(args.repo_root)

    if args.command == "line-endings":
        line_ending_report = build_line_ending_report(
            repo_root,
            tracked_only=args.tracked_only,
            include_ignored=args.include_ignored,
        )
        print_line_ending_report(line_ending_report, stdout=stdout)
        return 0

    if args.command == "path-references":
        path_reference_report = build_path_reference_report(
            repo_root,
            tracked_only=args.tracked_only,
            include_ignored=args.include_ignored,
            suppression_path=args.suppressions,
            include_missing_targets=args.include_missing_targets,
        )
        print_path_reference_report(path_reference_report, stdout=stdout)
        if args.fail_on_findings and path_reference_report.findings:
            return 1
        return 0

    if args.command == "host-setup":
        host_setup_report = build_host_setup_report(repo_root)
        print_host_setup_report(host_setup_report, stdout=stdout)
        return 0

    if args.command == "markdown":
        markdown_report = build_markdownlint_report(repo_root, fix=args.fix)
        print_markdownlint_report(markdown_report, fix=args.fix, stdout=stdout)
        return markdown_report.return_code if args.fix and markdown_report.return_code != 0 else 0

    if args.command == "powershell":
        powershell_report = build_powershell_report(
            repo_root,
            tracked_only=args.tracked_only,
            include_ignored=args.include_ignored,
        )
        print_powershell_report(powershell_report, stdout=stdout)
        if powershell_report.unsafe_candidate_count:
            return UNSAFE_CANDIDATE_EXIT_CODE
        return 0

    raise FirstAdoptionQualityError(f"Unsupported report command: {args.command}")


def fail(message: str) -> NoReturn:
    """Print an error and exit non-zero."""
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(1)


def main(argv: Sequence[str] | None = None) -> int:
    """Run the first-adoption quality report CLI."""
    args = parse_args(argv)
    try:
        return run_report(args)
    except FirstAdoptionQualityError as error:
        fail(str(error))


if __name__ == "__main__":
    raise SystemExit(main())
