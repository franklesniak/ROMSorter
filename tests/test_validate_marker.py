"""Exercise the marker-aware downstream sync validation helper."""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT_PATH = REPO_ROOT / ".template-sync" / "scripts" / "validate_marker.py"
MARKER_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-marker.schema.json"
MANIFEST_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-manifest.schema.json"
SOURCE_REPO = "https://github.com/franklesniak/copilot-repo-template.git"
FULL_SHA = "0123456789abcdef0123456789abcdef01234567"
MODULE_DEFINITIONS = {
    "baseline": "Baseline files.",
    "agent-instructions": "Agent instruction files.",
    "template-sync-support": "Template sync support files.",
    "python": "Python files.",
    "schema": "Schema files.",
    "github-actions": "GitHub Actions files.",
    "yaml": "YAML files.",
}


def _write_text(repo_root: Path, relative_path: str, text: str = "placeholder\n") -> None:
    """Write a text file below a fixture repository root."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _write_yaml(repo_root: Path, relative_path: str, data: dict[str, Any]) -> None:
    """Write a YAML fixture below a fixture repository root."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data, sort_keys=False), encoding="utf-8")


def _copy_schemas(repo_root: Path) -> None:
    """Copy the real marker and manifest schemas into a fixture repository."""
    schemas_dir = repo_root / "schemas"
    schemas_dir.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(MARKER_SCHEMA_PATH, schemas_dir / MARKER_SCHEMA_PATH.name)
    shutil.copyfile(MANIFEST_SCHEMA_PATH, schemas_dir / MANIFEST_SCHEMA_PATH.name)


def _manifest(version: int = 2) -> dict[str, Any]:
    """Build a small schema-valid manifest fixture."""
    modules = [
        {"name": name, "description": description}
        for name, description in MODULE_DEFINITIONS.items()
    ]
    if version == 3:
        modules.append({"name": "azure-pipelines", "description": "Azure Pipelines files."})
    path_mappings: list[dict[str, Any]] = [
        {"pattern": "README.md", "requires_all": ["baseline"]},
        {"pattern": "CLAUDE.md", "requires_all": ["agent-instructions"]},
        {"pattern": "GEMINI.md", "requires_all": ["agent-instructions"]},
        {
            "pattern": ".github/copilot-instructions.md",
            "requires_all": ["agent-instructions"],
        },
        {
            "pattern": ".template-sync/marker.yml",
            "requires_all": ["template-sync-support"],
        },
        {
            "pattern": ".template-sync/manifest.yml",
            "requires_all": ["template-sync-support"],
        },
        {
            "pattern": ".template-sync/scripts/**",
            "requires_all": ["template-sync-support"],
        },
        {
            "pattern": "schemas/template-sync-marker.schema.json",
            "requires_all": ["template-sync-support"],
        },
        {
            "pattern": "schemas/template-sync-manifest.schema.json",
            "requires_all": ["template-sync-support"],
        },
        {"pattern": "templates/python/**", "requires_all": ["python"]},
        {"pattern": "tests/test_schema_examples.py", "requires_all": ["schema"]},
    ]
    filtering: dict[str, str] = {
        "default_semantics": "AND",
        "path_matching": "most_specific_match_wins",
        "same_specificity_action": "union_modules",
        "unmapped_action": "surface_for_owner",
    }
    notes: dict[str, Any] = {
        "downstream_retention": "Downstream repositories keep marker data for syncs.",
    }

    compatibility_groups: list[dict[str, Any]] | None = None
    if version == 1:
        notes["known_limitations"] = [
            {
                "path": ".github/workflows/data-ci.yml",
                "description": "Version 1 cannot model alternative module relations.",
                "future_work": "Migrate to manifest version 2 relation semantics.",
            }
        ]
    else:
        path_mappings.append(
            {
                "pattern": ".github/workflows/data-ci.yml",
                "requires_all": ["github-actions"],
                "requires_any": ["yaml", "schema", "template-sync-support"],
            }
        )
        filtering["requires_any_semantics"] = "OR"
        if version == 3:
            path_mappings.append(
                {
                    "pattern": ".azuredevops/pipelines/**",
                    "requires_all": ["azure-pipelines"],
                }
            )
            compatibility_groups = [
                {
                    "name": "ci-host",
                    "description": "CI host modules.",
                    "default_modules": ["github-actions"],
                    "alternatives": [
                        {"host": "github", "modules": ["github-actions"]},
                        {"host": "azure-devops", "modules": ["azure-pipelines"]},
                    ],
                    "mixed_hosts": "unsupported",
                }
            ]

    template_manifest: dict[str, Any] = {
        "version": version,
        "modules": modules,
        "path_mappings": path_mappings,
        "filtering": filtering,
        "notes": notes,
    }
    if compatibility_groups is not None:
        template_manifest["compatibility_groups"] = compatibility_groups
    return {"template_manifest": template_manifest}


def _marker(
    included_modules: list[str],
    *,
    local_overrides: list[dict[str, str]] | None = None,
    local_path_ownership: list[dict[str, str]] | None = None,
    deferred_candidates: list[dict[str, str]] | None = None,
    protected_decisions: list[dict[str, str]] | None = None,
    protected_guide_waivers: list[dict[str, str]] | None = None,
) -> dict[str, Any]:
    """Build a small schema-valid marker fixture."""
    template_sync: dict[str, Any] = {
        "source_repo": SOURCE_REPO,
        "last_reviewed_template_commit": FULL_SHA,
        "included_modules": included_modules,
    }
    if local_overrides is not None:
        template_sync["local_overrides"] = local_overrides
    if local_path_ownership is not None:
        template_sync["local_path_ownership"] = local_path_ownership
    if deferred_candidates is not None:
        template_sync["deferred_protected_candidates"] = deferred_candidates
    if protected_decisions is not None:
        template_sync["protected_file_decisions"] = protected_decisions
    if protected_guide_waivers is not None:
        template_sync["protected_guide_contract_waivers"] = protected_guide_waivers
    return {"template_sync": template_sync}


def _write_manifest(repo_root: Path, version: int = 2) -> None:
    """Write the manifest fixture."""
    _write_yaml(repo_root, ".template-sync/manifest.yml", _manifest(version))


def _write_marker(
    repo_root: Path,
    included_modules: list[str],
    *,
    local_overrides: list[dict[str, str]] | None = None,
    local_path_ownership: list[dict[str, str]] | None = None,
    deferred_candidates: list[dict[str, str]] | None = None,
    protected_decisions: list[dict[str, str]] | None = None,
    protected_guide_waivers: list[dict[str, str]] | None = None,
) -> None:
    """Write the marker fixture."""
    _write_yaml(
        repo_root,
        ".template-sync/marker.yml",
        _marker(
            included_modules,
            local_overrides=local_overrides,
            local_path_ownership=local_path_ownership,
            deferred_candidates=deferred_candidates,
            protected_decisions=protected_decisions,
            protected_guide_waivers=protected_guide_waivers,
        ),
    )


def _run_validator(repo_root: Path, *extra_args: str) -> subprocess.CompletedProcess[str]:
    """Run the helper against a fixture repository."""
    return subprocess.run(
        [
            sys.executable,
            str(SCRIPT_PATH),
            "--repo-root",
            str(repo_root),
            *extra_args,
        ],
        check=False,
        capture_output=True,
        text=True,
    )


def _run_git(repo_root: Path, *args: str) -> None:
    """Run a git command in a fixture repository."""
    subprocess.run(
        ["git", *args],
        cwd=repo_root,
        check=True,
        capture_output=True,
        text=True,
    )


def _section_paths(output: str, heading: str) -> set[str]:
    """Return ``  - <path>`` entries rendered under a named output section."""
    entries: set[str] = set()
    in_section = False
    for line in output.splitlines():
        if line == heading:
            in_section = True
            continue
        if not in_section:
            continue
        if not line.strip():
            break
        if line.startswith("  - "):
            entries.add(line.removeprefix("  - ").strip())
    return entries


def _unrecorded_local_path_lines(output: str) -> set[str]:
    """Return paths listed in the unrecorded local path section."""
    heading = "Git-visible paths not mapped by the manifest or template_sync.local_path_ownership:"
    return _section_paths(output, heading)


def _symlink_or_skip(link_path: Path, target_path: Path, *, is_directory: bool) -> None:
    """Create a symlink, or skip when the platform refuses it."""
    try:
        link_path.symlink_to(target_path, target_is_directory=is_directory)
    except (OSError, NotImplementedError) as error:
        # Avoid interpolating an OSError directly: its default string form and
        # filename attributes can leak absolute local paths. Report the class
        # name plus strerror instead (NotImplementedError has no strerror).
        if isinstance(error, OSError):
            detail = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        else:
            detail = type(error).__name__
        pytest.skip(f"Filesystem does not support symlink creation: {detail}")


@pytest.fixture()
def marker_repo(tmp_path: Path) -> Path:
    """Create a fixture repository with schemas and a manifest."""
    _run_git(tmp_path, "init")
    _copy_schemas(tmp_path)
    _write_manifest(tmp_path)
    return tmp_path


def test_fully_included_template_state_passes(marker_repo: Path) -> None:
    """All retained concrete paths and present glob files are accepted."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support", "python", "schema", "github-actions"],
    )
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, ".template-sync/scripts/validate_marker.py")
    _write_text(marker_repo, "templates/python/app.py")
    _write_text(marker_repo, "tests/test_schema_examples.py")
    _write_text(marker_repo, ".github/workflows/data-ci.yml")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "Marker-aware template sync validation passed." in result.stdout
    assert "No retained-template inconsistencies found." in result.stdout


def test_partial_module_exclusion_without_leftovers_passes(marker_repo: Path) -> None:
    """Excluded glob modules do not imply every possible glob match must exist."""
    _write_marker(marker_repo, ["baseline", "template-sync-support"])
    _write_text(marker_repo, "README.md")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "Marker-aware template sync validation passed." in result.stdout


def test_marker_validation_accepts_azure_provider_fields(marker_repo: Path) -> None:
    """Marker validation accepts durable Azure DevOps Services provider decisions."""
    _write_manifest(marker_repo, version=3)
    _write_yaml(
        marker_repo,
        ".template-sync/marker.yml",
        {
            "template_sync": {
                "source_repo": SOURCE_REPO,
                "last_reviewed_template_commit": FULL_SHA,
                "included_modules": ["baseline", "template-sync-support", "azure-pipelines"],
                "host_provider": "azure-devops-services",
                "azure_devops_organization": "contoso",
                "azure_devops_project": "Platform",
                "azure_devops_project_url": "https://dev.azure.com/contoso/Platform",
                "azure_devops_repository": "downstream-template",
                "azure_devops_repository_url": (
                    "https://dev.azure.com/contoso/Platform/_git/downstream-template"
                ),
                "azure_devops_clone_url": (
                    "https://dev.azure.com/contoso/Platform/_git/downstream-template"
                ),
                "azure_boards_policy": "work-items",
                "azure_repos_pr_template_policy": "materialize",
                "azure_branch_policy_reviewer_guidance": "manual-follow-up",
                "azure_security_intake_policy": "manual-follow-up",
                "azure_security_product_enablement": "github-code-security",
                "azure_dependency_update_policy": "manual-follow-up",
            }
        },
    )
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, ".template-sync/scripts/validate_marker.py")
    _write_text(marker_repo, ".azuredevops/pipelines/azure-pipelines.yml")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "Marker-aware template sync validation passed." in result.stdout


def test_absent_marker_skips_by_default_and_fails_when_required(tmp_path: Path) -> None:
    """A missing marker is a no-op unless CI explicitly requires one."""
    result = _run_validator(tmp_path)
    required_result = _run_validator(tmp_path, "--require-marker")

    assert result.returncode == 0, result.stderr
    assert "No marker found at .template-sync/marker.yml" in result.stdout
    assert required_result.returncode == 1
    assert "Marker is required but was not found" in required_result.stderr


def test_malformed_marker_data_is_rejected(marker_repo: Path) -> None:
    """Present marker files must pass the existing marker schema."""
    _write_yaml(
        marker_repo,
        ".template-sync/marker.yml",
        _marker(["not-a-real-module"]),
    )

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "Schema validation failed for .template-sync/marker.yml" in result.stderr


def test_unsupported_mixed_host_selection_fails_with_actionable_message(
    marker_repo: Path,
) -> None:
    """Unsupported compatibility groups fail before ordinary file consistency checks."""
    _write_manifest(marker_repo, version=3)
    _write_marker(marker_repo, ["baseline", "github-actions", "azure-pipelines"])

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "Unsupported module compatibility selection(s)" in result.stderr
    assert "Unsupported mixed-host selection in compatibility group ci-host" in result.stderr
    assert "Select one host family" in result.stderr


def test_missing_concrete_mapped_file_fails(marker_repo: Path) -> None:
    """Concrete retained mappings must exist on disk."""
    _write_marker(marker_repo, ["baseline", "template-sync-support", "schema"])
    _write_text(marker_repo, "README.md")

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "Concrete mapped files expected for included modules but missing" in result.stdout
    assert "tests/test_schema_examples.py" in result.stdout


def test_untracked_nonignored_concrete_file_counts_as_present(marker_repo: Path) -> None:
    """Untracked retained files are accepted when Git does not ignore them."""
    _write_marker(marker_repo, ["baseline", "template-sync-support", "schema"])
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, "tests/test_schema_examples.py")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "Marker-aware template sync validation passed." in result.stdout


def test_ignored_retained_concrete_file_is_reported_missing(marker_repo: Path) -> None:
    """Ignored files do not satisfy retained concrete manifest mappings."""
    _write_marker(marker_repo, ["baseline", "template-sync-support", "schema"])
    _write_text(marker_repo, ".gitignore", "tests/test_schema_examples.py\n")
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, "tests/test_schema_examples.py")

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "Concrete mapped files expected for included modules but missing" in result.stdout
    assert "tests/test_schema_examples.py" in result.stdout


def test_leftover_file_from_excluded_module_fails(marker_repo: Path) -> None:
    """Existing files mapped only to excluded modules are reported."""
    _write_marker(marker_repo, ["baseline", "template-sync-support"])
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, "templates/python/app.py")

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "Files present on disk but not retained by included modules" in result.stdout
    assert "templates/python/app.py" in result.stdout


def test_local_overrides_skip_matching_paths(marker_repo: Path) -> None:
    """Local override paths are omitted from retained-template consistency checks."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        local_overrides=[
            {
                "path": "templates/python/",
                "reason": "Local project owns this directory.",
                "default_decision": "SKIP",
            }
        ],
    )
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, "templates/python/app.py")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "Local overrides skipped:" in result.stdout
    assert "templates/python/ (SKIP): Local project owns this directory." in result.stdout


def test_local_path_ownership_exact_path_suppresses_unrecorded_suggestion(
    marker_repo: Path,
) -> None:
    """An exact local ownership record covers only the named local path."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        local_path_ownership=[
            {
                "path": "docs/index.md",
                "reason": "Project documentation is downstream-owned.",
            }
        ],
    )
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, "docs/index.md")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "Local path ownership records:" in result.stdout
    assert "path=docs/index.md; reason=Project documentation is downstream-owned." in (
        result.stdout
    )
    assert "docs/index.md" not in _unrecorded_local_path_lines(result.stdout)


def test_local_path_ownership_directory_family_covers_directory_and_descendants(
    marker_repo: Path,
) -> None:
    """Directory-prefix local ownership covers the directory path and descendants."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        local_path_ownership=[
            {
                "path": "docs/",
                "reason": "Project docs are downstream-owned.",
            }
        ],
    )
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, "docs/index.md")
    _write_text(marker_repo, "docs/api/reference.md")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "path=docs/; reason=Project docs are downstream-owned." in result.stdout
    unrecorded = _unrecorded_local_path_lines(result.stdout)
    assert "docs/index.md" not in unrecorded
    assert "docs/api/reference.md" not in unrecorded


def test_duplicate_local_path_ownership_paths_fail(marker_repo: Path) -> None:
    """Local ownership records must not repeat the same normalized path."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        local_path_ownership=[
            {"path": "docs/", "reason": "Project documentation."},
            {"path": "docs/", "reason": "Duplicate documentation record."},
        ],
    )

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "Duplicate local_path_ownership path(s): docs" in result.stderr


def test_local_path_ownership_exact_manifest_collision_fails(marker_repo: Path) -> None:
    """Local ownership cannot override an exact concrete upstream-owned file."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        local_path_ownership=[
            {
                "path": "README.md",
                "reason": "README is local.",
                "overlap_exception_reason": "Attempted exact-file exception.",
            }
        ],
    )
    _write_text(marker_repo, "README.md")

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "README.md exactly collides with concrete manifest-owned path(s): README.md" in (
        result.stderr
    )
    assert "overlap_exception_reason cannot override an exact upstream-owned file collision" in (
        result.stderr
    )


def test_local_path_ownership_broad_manifest_overlap_requires_exception(
    marker_repo: Path,
) -> None:
    """Local ownership inside a broad manifest area requires an exception reason."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        local_path_ownership=[
            {
                "path": "templates/python/local-tool.py",
                "reason": "Local Python utility lives near template starter files.",
            }
        ],
    )
    _write_text(marker_repo, "README.md")

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "templates/python/local-tool.py overlaps broad manifest-owned pattern(s)" in (
        result.stderr
    )
    assert "add overlap_exception_reason" in result.stderr


def test_local_path_ownership_broad_manifest_overlap_with_exception_passes(
    marker_repo: Path,
) -> None:
    """A broad manifest-area local ownership exception may point at a future path."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        local_path_ownership=[
            {
                "path": "templates/python/local-tool.py",
                "reason": "Local Python utility lives near template starter files.",
                "overlap_exception_reason": "This downstream tool is not from the template.",
            }
        ],
    )
    _write_text(marker_repo, "README.md")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "templates/python/local-tool.py" in result.stdout


def test_local_path_ownership_exception_without_broad_overlap_fails(
    marker_repo: Path,
) -> None:
    """An overlap exception reason is invalid when no broad manifest area overlaps."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        local_path_ownership=[
            {
                "path": "docs/local.md",
                "reason": "Project documentation is local.",
                "overlap_exception_reason": "No manifest pattern overlaps this path.",
            }
        ],
    )
    _write_text(marker_repo, "README.md")

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "docs/local.md overlap_exception_reason is only permitted" in result.stderr


def test_local_path_ownership_future_path_is_allowed(marker_repo: Path) -> None:
    """A local ownership record may name a future path whose leaf is absent."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        local_path_ownership=[
            {
                "path": "docs/future.md",
                "reason": "Planned downstream documentation path.",
            }
        ],
    )
    _write_text(marker_repo, "README.md")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "docs/future.md" in result.stdout


def test_local_path_ownership_rejects_existing_symlink_ancestor(
    marker_repo: Path,
    tmp_path: Path,
) -> None:
    """Declared local ownership paths must not traverse symlinked ancestors."""
    outside = tmp_path / "outside"
    outside.mkdir()
    _symlink_or_skip(marker_repo / "local", outside, is_directory=True)
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        local_path_ownership=[
            {
                "path": "local/future.md",
                "reason": "This path traverses a symlink.",
            }
        ],
    )
    _write_text(marker_repo, "README.md")

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "template_sync.local_path_ownership[].path must not traverse a symlink" in (
        result.stderr
    )


def test_git_visible_symlink_local_path_is_surfaced(marker_repo: Path, tmp_path: Path) -> None:
    """Git-visible local symlinks are surfaced without becoming ownership suggestions."""
    outside = tmp_path / "outside.txt"
    outside.write_text("outside\n", encoding="utf-8")
    _symlink_or_skip(marker_repo / "local-link.txt", outside, is_directory=False)
    _write_marker(marker_repo, ["baseline", "template-sync-support"])
    _write_text(marker_repo, "README.md")
    _run_git(marker_repo, "add", "local-link.txt")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "Git-visible local paths that are symlinks or resolve unsafely:" in result.stdout
    assert "local-link.txt" in result.stdout


def test_manifest_directory_symlink_is_a_managed_failure(
    marker_repo: Path,
    tmp_path: Path,
) -> None:
    """A symlinked directory over a manifest-managed tree fails as a managed path.

    A directory replaced by a symlink is stored by Git as a single entry, so a
    glob mapping such as ``templates/python/**`` never matches the directory path
    itself. The validator must still treat it as template-managed so the symlink
    surfaces as a fatal managed-path finding rather than a non-fatal local warning.
    """
    outside = tmp_path / "outside-python"
    outside.mkdir()
    (marker_repo / "templates").mkdir(parents=True, exist_ok=True)
    _symlink_or_skip(marker_repo / "templates" / "python", outside, is_directory=True)
    _write_marker(marker_repo, ["baseline", "template-sync-support", "python"])
    _write_text(marker_repo, "README.md")
    _run_git(marker_repo, "add", "templates/python")

    result = _run_validator(marker_repo)

    assert result.returncode == 1, result.stdout
    managed_paths = _section_paths(
        result.stdout,
        "Template-managed paths that are symlinks or resolve unsafely:",
    )
    local_paths = _section_paths(
        result.stdout,
        "Git-visible local paths that are symlinks or resolve unsafely:",
    )
    assert "templates/python" in managed_paths
    assert "templates/python" not in local_paths


def test_unrecorded_local_paths_surface_copy_ready_suggestion(marker_repo: Path) -> None:
    """Unrecorded local paths are warnings by default and include a YAML stub."""
    _write_marker(marker_repo, ["baseline", "template-sync-support"])
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, "docs/unrecorded.md")

    result = _run_validator(marker_repo)
    strict_result = _run_validator(marker_repo, "--strict-local-path-ownership")

    assert result.returncode == 0, result.stderr
    assert "Git-visible paths not mapped by the manifest" in result.stdout
    assert "docs/unrecorded.md" in result.stdout
    assert "  - path: docs/" in result.stdout
    assert "    reason: Downstream project owns this path family." in result.stdout
    assert "--strict-local-path-ownership to make them fatal" in result.stdout
    assert strict_result.returncode == 1
    assert "Local path ownership findings are fatal because strict mode is enabled." in (
        strict_result.stdout
    )


def test_exact_path_ownership_suggestion_uses_neutral_reason(marker_repo: Path) -> None:
    """An exact-file ownership suggestion avoids directory-family wording.

    The suggested record for a top-level file (no trailing ``/``) should not imply
    directory-family semantics in its reason text.
    """
    _write_marker(marker_repo, ["baseline", "template-sync-support"])
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, "unrecorded.txt")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "  - path: unrecorded.txt" in result.stdout
    assert "    reason: Downstream project owns this path." in result.stdout
    assert "    reason: Downstream project owns this path family." not in result.stdout


def test_unrecorded_local_paths_include_walk_pruned_directories(marker_repo: Path) -> None:
    """Git-tracked files under walk-pruned directories are still reported as unrecorded.

    Discovery follows the documented ``git ls-files`` contract, so a file a
    downstream intentionally tracks under a pruned directory such as ``dist/``
    must not be silently dropped just because the safe walk prunes that directory.
    """
    _write_marker(marker_repo, ["baseline", "template-sync-support"])
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, "dist/app.js")
    _run_git(marker_repo, "add", "dist/app.js")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "dist/app.js" in _unrecorded_local_path_lines(result.stdout)


def test_symlink_under_pruned_directory_is_surfaced_as_unsafe(
    marker_repo: Path,
    tmp_path: Path,
) -> None:
    """A Git-tracked symlink under a walk-pruned directory is still flagged unsafe.

    The safe walk prunes directories such as node_modules/, so such a symlink is
    absent from the walk's skipped list. It must still be derived from the
    Git-visible list and surfaced rather than dropped from every output.
    """
    outside = tmp_path / "outside.txt"
    outside.write_text("outside\n", encoding="utf-8")
    (marker_repo / "node_modules").mkdir(parents=True, exist_ok=True)
    _symlink_or_skip(marker_repo / "node_modules" / "dep.js", outside, is_directory=False)
    _write_marker(marker_repo, ["baseline", "template-sync-support"])
    _write_text(marker_repo, "README.md")
    _run_git(marker_repo, "add", "node_modules/dep.js")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    unsafe_paths = _section_paths(
        result.stdout,
        "Git-visible local paths that are symlinks or resolve unsafely:",
    )
    assert "node_modules/dep.js" in unsafe_paths


def test_ignored_files_do_not_surface_as_local_path_suggestions(marker_repo: Path) -> None:
    """Ignored files are excluded by the Git-visible path convention."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        local_path_ownership=[
            {"path": ".gitignore", "reason": "Ignore rules are downstream-owned."}
        ],
    )
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, ".gitignore", "ignored.log\n")
    _write_text(marker_repo, "ignored.log")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "ignored.log" not in result.stdout


def test_unrecorded_local_path_suggestions_are_bounded(marker_repo: Path) -> None:
    """Large unrecorded local path sets produce deterministic bounded suggestions."""
    _write_marker(marker_repo, ["baseline", "template-sync-support"])
    _write_text(marker_repo, "README.md")
    for index in range(12):
        _write_text(marker_repo, f"local-{index:02d}/file.txt")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "# Showing 10 of 12 suggested record(s)." in result.stdout
    assert "  - path: local-00/" in result.stdout
    assert "  - path: local-09/" in result.stdout
    assert "  - path: local-10/" not in result.stdout


def test_deferred_protected_candidates_are_reported_without_failure(marker_repo: Path) -> None:
    """Deferred protected candidates are visible but do not create failures."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        deferred_candidates=[
            {
                "path": ".github/copilot-instructions.md",
                "source_commit": "89abcdef0123456789abcdef0123456789abcdef",
                "reason": "Owner authorization is still pending.",
            }
        ],
    )
    _write_text(marker_repo, "README.md")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "Deferred protected candidates:" in result.stdout
    assert ".github/copilot-instructions.md" in result.stdout


def test_manifest_v2_requires_any_relation_is_consumed(marker_repo: Path) -> None:
    """Version 2 mappings retain a file when one requires_any module is included."""
    _write_marker(marker_repo, ["baseline", "template-sync-support", "github-actions", "yaml"])
    _write_text(marker_repo, "README.md")
    _write_text(marker_repo, ".github/workflows/data-ci.yml")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "Marker-aware template sync validation passed." in result.stdout


def test_manifest_v1_requires_all_semantics_are_consumed(tmp_path: Path) -> None:
    """Version 1 mappings continue to use requires_all-only semantics."""
    _run_git(tmp_path, "init")
    _copy_schemas(tmp_path)
    _write_manifest(tmp_path, version=1)
    _write_marker(tmp_path, ["baseline", "template-sync-support"])
    _write_text(tmp_path, "README.md")

    result = _run_validator(tmp_path)

    assert result.returncode == 0, result.stderr
    assert "Marker-aware template sync validation passed." in result.stdout


def test_protected_file_decisions_and_remove_local_authorizations_are_reported(
    marker_repo: Path,
) -> None:
    """Protected-file decisions and removal authorizations are visible."""
    _write_marker(
        marker_repo,
        ["template-sync-support"],
        protected_decisions=[
            {
                "path": "CLAUDE.md",
                "decision": "MERGE",
                "adoption_mode": "minimal-preservation",
                "authorization_basis": "Owner authorized protected edits on 2026-05-27.",
                "authorized_scope": "CLAUDE.md only.",
            },
            {
                "path": "GEMINI.md",
                "decision": "REMOVE-LOCAL",
                "authorization_basis": "Owner explicitly authorized removing GEMINI.md.",
                "authorized_scope": "GEMINI.md only.",
                "reason": "Gemini agent not used by this repository.",
            },
        ],
    )

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "Protected file decisions:" in result.stdout
    assert "CLAUDE.md: MERGE" in result.stdout
    assert "REMOVE-LOCAL authorizations:" in result.stdout
    assert "authorization_basis: Owner explicitly authorized removing GEMINI.md." in result.stdout
    assert "authorized_scope: GEMINI.md only." in result.stdout
    assert "reason: Gemini agent not used by this repository." in result.stdout


def test_protected_guide_contract_waivers_are_reported(marker_repo: Path) -> None:
    """Protected-guide waivers are parsed and visible in marker output."""
    _write_marker(
        marker_repo,
        ["baseline", "template-sync-support"],
        protected_guide_waivers=[
            {
                "path": "AGENTS.md",
                "contract_key": "agents-azure-devops-pr-review-protocol",
                "target_module": "azure-devops-collaboration",
                "reason": "GitHub-only fixture keeps the protected guide section visible.",
                "authorization_basis": "Fixture owner authorized this protected-guide waiver.",
            },
            {
                "path": "AGENTS.md",
                "contract_key": "agents-azure-devops-support-guide-path",
                "target_path": "docs/azure-devops-support.md",
                "linked_local_override_path": "AGENTS.md",
                "reason": "GitHub-only fixture keeps the protected guide path visible.",
                "authorization_basis": "Fixture owner authorized this reference waiver.",
            },
        ],
    )
    _write_text(marker_repo, "README.md")

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "Protected guide contract waivers:" in result.stdout
    assert "AGENTS.md: agents-azure-devops-pr-review-protocol" in result.stdout
    assert "target_module: azure-devops-collaboration" in result.stdout
    assert "target_path: docs/azure-devops-support.md" in result.stdout
    assert "linked_local_override_path: AGENTS.md" in result.stdout


def test_duplicate_protected_decision_paths_fail(marker_repo: Path) -> None:
    """Protected-file decisions must not repeat normalized paths."""
    _write_marker(
        marker_repo,
        ["template-sync-support"],
        protected_decisions=[
            {
                "path": "CLAUDE.md",
                "decision": "SKIP",
                "reason": "Claude agent is not used.",
            },
            {
                "path": "CLAUDE.md",
                "decision": "PROTECTED-REVIEW",
                "reason": "Owner authorization is pending.",
            },
        ],
    )

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "Duplicate protected_file_decisions path(s): CLAUDE.md" in result.stderr


def test_directory_style_protected_decision_path_fails(marker_repo: Path) -> None:
    """Protected-file decisions must reference a file, not a directory."""
    _write_marker(
        marker_repo,
        ["template-sync-support"],
        protected_decisions=[
            {
                "path": ".github/instructions/",
                "decision": "MERGE",
                "adoption_mode": "minimal-preservation",
                "authorization_basis": "Owner authorized protected edits on 2026-05-27.",
                "authorized_scope": ".github/instructions/ only.",
            }
        ],
    )

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "Schema validation failed for .template-sync/marker.yml" in result.stderr
    assert (
        ".template_sync.protected_file_decisions[0].path" in result.stderr
        or "template_sync.protected_file_decisions[0].path" in result.stderr
    )


def test_directory_style_deferred_candidate_path_fails(marker_repo: Path) -> None:
    """Deferred protected candidates must reference a file, not a directory."""
    _write_marker(
        marker_repo,
        ["template-sync-support"],
        deferred_candidates=[
            {
                "path": ".github/instructions/",
                "source_commit": FULL_SHA,
                "reason": "Protected governance directory needs explicit owner authorization.",
            }
        ],
    )

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "Schema validation failed for .template-sync/marker.yml" in result.stderr
    assert (
        ".template_sync.deferred_protected_candidates[0].path" in result.stderr
        or "template_sync.deferred_protected_candidates[0].path" in result.stderr
    )


def test_noncontradictory_protected_decision_local_override_overlap_is_reported(
    marker_repo: Path,
) -> None:
    """Same-decision protected decisions and local overrides are shown together."""
    _write_marker(
        marker_repo,
        ["template-sync-support"],
        local_overrides=[
            {
                "path": "CLAUDE.md",
                "reason": "Claude agent is downstream-owned.",
                "default_decision": "SKIP",
            }
        ],
        protected_decisions=[
            {
                "path": "CLAUDE.md",
                "decision": "SKIP",
                "reason": "Claude agent is not used.",
            }
        ],
    )

    result = _run_validator(marker_repo)

    assert result.returncode == 0, result.stderr
    assert "Protected decision overlaps:" in result.stdout
    assert "protected_file_decisions: decision=SKIP; reason=Claude agent is not used." in (
        result.stdout
    )
    assert (
        "local_overrides: path=CLAUDE.md; default_decision=SKIP; "
        "reason=Claude agent is downstream-owned."
    ) in result.stdout


def test_contradictory_protected_decision_local_override_overlap_fails(
    marker_repo: Path,
) -> None:
    """Different same-path protected and local override decisions are contradictory."""
    _write_marker(
        marker_repo,
        ["template-sync-support"],
        local_overrides=[
            {
                "path": "CLAUDE.md",
                "reason": "Claude agent is downstream-owned.",
                "default_decision": "SKIP",
            }
        ],
        protected_decisions=[
            {
                "path": "CLAUDE.md",
                "decision": "MERGE",
                "adoption_mode": "minimal-preservation",
                "authorization_basis": "Owner authorized protected edits on 2026-05-27.",
                "authorized_scope": "CLAUDE.md only.",
            }
        ],
    )

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "Contradictory protected-file marker entries:" in result.stderr
    assert "protected_file_decisions: decision=MERGE" in result.stderr
    assert "local_overrides: path=CLAUDE.md; default_decision=SKIP" in result.stderr


def test_protected_decision_deferred_candidate_overlap_fails(marker_repo: Path) -> None:
    """A current protected decision cannot also be deferred for the same path."""
    _write_marker(
        marker_repo,
        ["template-sync-support"],
        deferred_candidates=[
            {
                "path": "CLAUDE.md",
                "source_commit": "89abcdef0123456789abcdef0123456789abcdef",
                "reason": "Owner authorization is pending.",
            }
        ],
        protected_decisions=[
            {
                "path": "CLAUDE.md",
                "decision": "PROTECTED-REVIEW",
                "reason": "Owner authorization is pending.",
            }
        ],
    )

    result = _run_validator(marker_repo)

    assert result.returncode == 1
    assert "Contradictory protected-file marker entries:" in result.stderr
    assert "deferred_protected_candidates: source_commit=89abcdef" in result.stderr


def test_strict_remove_local_phrasing_is_opt_in_and_token_configurable(
    marker_repo: Path,
) -> None:
    """Strict removal phrasing fails only when opted in and no token matches."""
    _write_marker(
        marker_repo,
        ["template-sync-support"],
        protected_decisions=[
            {
                "path": "GEMINI.md",
                "decision": "REMOVE-LOCAL",
                "authorization_basis": "Owner approved GEMINI.md retirement.",
                "authorized_scope": "GEMINI.md only.",
                "reason": "Gemini agent not used by this repository.",
            }
        ],
    )

    default_result = _run_validator(marker_repo)
    strict_result = _run_validator(marker_repo, "--strict-remove-local-phrasing")
    custom_result = _run_validator(
        marker_repo,
        "--strict-remove-local-phrasing",
        "--remove-local-authorization-token",
        "retire",
    )

    assert default_result.returncode == 0, default_result.stderr
    assert strict_result.returncode == 1
    assert "GEMINI.md REMOVE-LOCAL authorization_basis must contain" in strict_result.stderr
    assert custom_result.returncode == 0, custom_result.stderr
