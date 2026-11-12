"""Exercise the first-adoption bootstrap coordinator."""

from __future__ import annotations

import importlib
import io
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any, cast

import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_PATH = REPO_ROOT / ".template-sync" / "scripts" / "bootstrap_first_adoption.py"
SCRIPT_DIR = SCRIPT_PATH.parent
MARKER_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-marker.schema.json"
MANIFEST_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-manifest.schema.json"
SOURCE_REPO = "https://github.com/franklesniak/copilot-repo-template.git"

if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

bootstrap = cast(Any, importlib.import_module("bootstrap_first_adoption"))
parse_marker_decision_data = importlib.import_module(
    "template_sync_materialization_helpers"
).parse_marker_decision_data


def _run_git(repo_root: Path, *args: str) -> str:
    """Run a Git command in a fixture repository and return stdout."""
    result = subprocess.run(
        ["git", *args],
        cwd=repo_root,
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stderr
    return result.stdout.strip()


def _write_text(repo_root: Path, relative_path: str, text: str = "placeholder\n") -> None:
    """Write a UTF-8 fixture file."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _write_yaml(repo_root: Path, relative_path: str, data: dict[str, Any]) -> None:
    """Write a YAML fixture file."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data, sort_keys=False), encoding="utf-8")


def _copy_schemas(repo_root: Path) -> None:
    """Copy marker and manifest schemas into a fixture repository."""
    schemas_dir = repo_root / "schemas"
    schemas_dir.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(MARKER_SCHEMA_PATH, schemas_dir / MARKER_SCHEMA_PATH.name)
    shutil.copyfile(MANIFEST_SCHEMA_PATH, schemas_dir / MANIFEST_SCHEMA_PATH.name)


def _manifest() -> dict[str, Any]:
    """Return a compact schema-valid template-sync manifest."""
    return {
        "template_manifest": {
            "version": 2,
            "modules": [
                {"name": "baseline", "description": "Baseline files."},
                {"name": "agent-instructions", "description": "Agent instructions."},
                {"name": "template-sync-support", "description": "Template sync support."},
                {"name": "markdown", "description": "Markdown files."},
            ],
            "path_mappings": [
                {"pattern": "README.md", "requires_all": ["baseline"]},
                {"pattern": "AGENTS.md", "requires_all": ["agent-instructions"]},
                {
                    "pattern": ".github/copilot-instructions.md",
                    "requires_all": ["agent-instructions"],
                },
                {
                    "pattern": ".template-sync/manifest.yml",
                    "requires_all": ["template-sync-support"],
                },
                {
                    "pattern": ".template-sync/marker.yml",
                    "requires_all": ["template-sync-support"],
                },
                {
                    "pattern": ".template-sync/scripts/**",
                    "requires_all": ["template-sync-support"],
                },
                {
                    "pattern": "templates/adoption/**",
                    "requires_all": ["template-sync-support"],
                },
            ],
            "filtering": {
                "default_semantics": "AND",
                "requires_any_semantics": "OR",
                "path_matching": "most_specific_match_wins",
                "same_specificity_action": "union_modules",
                "unmapped_action": "surface_for_owner",
            },
            "notes": {
                "downstream_retention": "Fixture repositories keep marker state.",
            },
        }
    }


def _init_fixture_repo(repo_root: Path) -> None:
    """Create a small Git repository with bootstrap prerequisites."""
    _run_git(repo_root, "init")
    _copy_schemas(repo_root)
    _write_yaml(repo_root, ".template-sync/manifest.yml", _manifest())
    _write_text(
        repo_root,
        "templates/adoption/_TEMPLATE-ADOPTION-DIFFICULTIES.md",
        "# Adoption Difficulties Journal\n\n## Entries\n",
    )
    _write_text(repo_root, "README.md")
    _write_text(repo_root, ".template-sync/scripts/validate_marker.py")


def _run_bootstrap(repo_root: Path, *extra_args: str) -> str:
    """Run the bootstrap module entrypoint and return stdout."""
    stdout = io.StringIO()
    result = bootstrap.main(
        [
            "--repo-root",
            str(repo_root),
            "--source-repo",
            SOURCE_REPO,
            *extra_args,
        ],
        stdout=stdout,
    )
    assert result == 0
    return stdout.getvalue()


def _section(output: str, heading: str) -> str:
    """Return the Markdown section under ``heading``."""
    marker = f"## {heading}"
    start = output.index(marker) + len(marker)
    next_heading = output.find("\n## ", start)
    if next_heading == -1:
        return output[start:]
    return output[start:next_heading]


def test_default_mode_prints_artifacts_without_writing(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """The default bootstrap is a read-only report with proposed artifacts."""
    _init_fixture_repo(tmp_path)
    monkeypatch.setattr(
        bootstrap.first_adoption_checks,
        "default_pre_commit_prefix",
        lambda: ("pre-commit", "run", "--files"),
    )

    output = _run_bootstrap(tmp_path)

    assert "# First-Adoption Bootstrap Report" in output
    assert "## Proposed `_TODO-repo-init.md`" in output
    assert "## Proposed `_ADOPTION-DIFFICULTIES.md`" in output
    assert "## Draft `.template-sync/marker.yml`" in output
    assert "not written; default mode is read-only" in output
    assert "[pre-commit] `pre-commit run --files" in output
    assert "README.md" in _section(output, "Runnable Validation Plan")
    assert not (tmp_path / "_TODO-repo-init.md").exists()
    assert not (tmp_path / "_ADOPTION-DIFFICULTIES.md").exists()
    assert not (tmp_path / ".template-sync" / "marker.yml").exists()


def test_write_mode_creates_todo_and_journal_without_marker_by_default(tmp_path: Path) -> None:
    """Explicit write mode creates missing downstream-owned state idempotently."""
    _init_fixture_repo(tmp_path)

    output = _run_bootstrap(tmp_path, "--write")
    second_output = _run_bootstrap(tmp_path, "--write")

    todo_path = tmp_path / "_TODO-repo-init.md"
    journal_path = tmp_path / "_ADOPTION-DIFFICULTIES.md"
    assert "created" in output
    assert todo_path.is_file()
    assert journal_path.is_file()
    assert bootstrap.BOOTSTRAP_STATE_BEGIN in todo_path.read_text(encoding="utf-8")
    assert "unchanged; existing checklist content was preserved" in second_output
    assert "unchanged; existing journal content was preserved" in second_output
    assert not (tmp_path / ".template-sync" / "marker.yml").exists()


def test_existing_todo_state_update_preserves_free_form_notes(tmp_path: Path) -> None:
    """The narrow TODO update path preserves Markdown outside the state block."""
    _init_fixture_repo(tmp_path)
    state = bootstrap.default_bootstrap_state(
        ledger_rows=(),
        marker_data=bootstrap.sync_candidates.empty_marker_data(),
    )
    state["decisions"][0]["status"] = bootstrap.RESOLVED_STATUS
    state["decisions"][0]["answer"] = "Owner chose to leave it disabled."
    state["decisions"][0]["evidence"] = "Maintainer note 2026-06-27"
    todo_text = "\n".join(
        [
            "# Repository Initialization Checklist",
            "",
            "Free-form note before the block.",
            "",
            bootstrap.bootstrap_contract_section(state),
            "",
            "Free-form note after the block.",
            "",
        ]
    )
    _write_text(tmp_path, "_TODO-repo-init.md", todo_text)

    _run_bootstrap(tmp_path, "--write", "--update-existing-todo-state", "--draft-marker", "never")

    updated = (tmp_path / "_TODO-repo-init.md").read_text(encoding="utf-8")
    parsed_state = bootstrap.find_existing_state(updated).state
    assert "Free-form note before the block." in updated
    assert "Free-form note after the block." in updated
    assert parsed_state is not None
    decisions = bootstrap.decision_by_key(parsed_state)
    decision = decisions["manual.github.private_vulnerability_reporting"]
    assert decision["status"] == bootstrap.RESOLVED_STATUS
    assert decision["answer"] == "Owner chose to leave it disabled."
    assert decision["evidence"] == "Maintainer note 2026-06-27"
    assert "manual.github.discussions" in decisions


def test_append_state_block_preserves_trailing_note_spaces() -> None:
    """Appending the state block keeps rendering-significant trailing spaces."""
    state = bootstrap.default_bootstrap_state(
        ledger_rows=(),
        marker_data=bootstrap.sync_candidates.empty_marker_data(),
    )
    # Two trailing spaces form a Markdown hard line break and must survive.
    original = "# Repository Initialization Checklist\n\nFree-form note.  \n"

    updated = bootstrap.replace_or_append_state_block(original, state)

    assert updated.startswith("# Repository Initialization Checklist\n\nFree-form note.  \n\n")
    assert bootstrap.find_existing_state(updated).state is not None


def test_find_existing_state_tolerates_end_marker_inside_json_value() -> None:
    """An end-marker string copied into a JSON value must not truncate parsing."""
    state = bootstrap.default_bootstrap_state(
        ledger_rows=(),
        marker_data=bootstrap.sync_candidates.empty_marker_data(),
    )
    target_key = state["decisions"][0]["key"]
    # A maintainer pastes the end-marker text into a free-form notes field.
    state["decisions"][0]["notes"] = f"see {bootstrap.BOOTSTRAP_STATE_END} here"
    todo_text = bootstrap.bootstrap_contract_section(state)

    parsed = bootstrap.find_existing_state(todo_text).state

    assert parsed is not None
    decisions = bootstrap.decision_by_key(parsed)
    assert decisions[target_key]["notes"] == f"see {bootstrap.BOOTSTRAP_STATE_END} here"


def test_write_todo_rejects_symlinked_checklist_path(tmp_path: Path) -> None:
    """A symlink at the checklist path is rejected before any write is attempted."""
    target = tmp_path / "outside-target.md"
    todo_path = tmp_path / "_TODO-repo-init.md"
    try:
        todo_path.symlink_to(target)  # broken symlink: the target does not exist
    except (OSError, NotImplementedError):
        pytest.skip("Filesystem does not support symlink creation")

    state = bootstrap.default_bootstrap_state(
        ledger_rows=(),
        marker_data=bootstrap.sync_candidates.empty_marker_data(),
    )

    with pytest.raises(bootstrap.FirstAdoptionBootstrapError):
        bootstrap.write_todo_if_requested(
            todo_path=todo_path,
            todo_text="# Repository Initialization Checklist",
            merged_state=state,
            repo_root=tmp_path,
            write=True,
            update_existing_todo_state=False,
        )

    assert todo_path.is_symlink()
    assert not target.exists()  # the symlink target was not written through


def test_resumed_run_uses_structured_state_not_markdown_checkboxes(tmp_path: Path) -> None:
    """Loose Markdown checkbox state never suppresses unresolved questions."""
    _init_fixture_repo(tmp_path)
    _write_text(
        tmp_path,
        "_TODO-repo-init.md",
        "# Repository Initialization Checklist\n\n"
        "- [x] Private vulnerability reporting decision: leave disabled\n",
    )

    loose_output = _run_bootstrap(tmp_path, "--draft-marker", "never")
    loose_active = _section(loose_output, "Active Owner Questions And Blocked State")

    assert "manual.github.private_vulnerability_reporting" in loose_active

    state = bootstrap.default_bootstrap_state(
        ledger_rows=(),
        marker_data=bootstrap.sync_candidates.empty_marker_data(),
    )
    state["decisions"][0]["status"] = bootstrap.RESOLVED_STATUS
    state["decisions"][0]["answer"] = "Owner chose to leave it disabled."
    state["decisions"][0]["evidence"] = "Maintainer note 2026-06-27"
    _write_text(
        tmp_path,
        "_TODO-repo-init.md",
        "# Repository Initialization Checklist\n\n" + bootstrap.bootstrap_contract_section(state),
    )

    structured_output = _run_bootstrap(tmp_path, "--draft-marker", "never")
    structured_active = _section(structured_output, "Active Owner Questions And Blocked State")
    structured_resolved = _section(structured_output, "Resolved Structured Decisions")
    questionnaire = _section(structured_output, "Maintainer Questionnaire")

    assert "manual.github.private_vulnerability_reporting" not in structured_active
    assert "manual.github.private_vulnerability_reporting" in structured_resolved
    assert "Should private vulnerability reporting" not in questionnaire
    assert "Resolved maintainer questions suppressed" in questionnaire


def test_protected_questions_reuse_preflight_output(tmp_path: Path) -> None:
    """Protected-file prompts come from the existing preflight helper."""
    _init_fixture_repo(tmp_path)

    output = _run_bootstrap(tmp_path, "--draft-marker", "never")
    protected_section = _section(output, "Protected-File Authorization Questions")

    assert (
        "- [ ] Does the maintainer authorize the selected action for "
        "`AGENTS.md`? Ledger decision:"
    ) in protected_section


def test_write_draft_marker_validates_and_creates_absent_marker(tmp_path: Path) -> None:
    """The explicit marker write path creates a schema-valid integrity-checked marker."""
    _init_fixture_repo(tmp_path)

    output = _run_bootstrap(tmp_path, "--write", "--write-draft-marker")

    marker_path = tmp_path / ".template-sync" / "marker.yml"
    marker_text = marker_path.read_text(encoding="utf-8")
    marker = yaml.safe_load(marker_text)
    assert "created from validated draft marker" in output
    assert marker["template_sync"]["source_repo"] == SOURCE_REPO
    assert marker["template_sync"]["included_modules"] == [
        "agent-instructions",
        "baseline",
        "markdown",
        "template-sync-support",
    ]
    assert marker_text == bootstrap.format_marker_yaml(marker)
    parse_marker_decision_data(marker, validate_protected_decision_integrity=True)


def test_existing_marker_is_preserved_by_write_draft_marker(tmp_path: Path) -> None:
    """Marker write mode refuses to overwrite existing marker content."""
    _init_fixture_repo(tmp_path)
    existing_marker: dict[str, Any] = {
        "template_sync": {
            "source_repo": SOURCE_REPO,
            "included_modules": ["baseline"],
            "host_provider": "github",
        }
    }
    _write_yaml(tmp_path, ".template-sync/marker.yml", existing_marker)
    before = (tmp_path / ".template-sync" / "marker.yml").read_text(encoding="utf-8")

    try:
        _run_bootstrap(
            tmp_path,
            "--write",
            "--write-draft-marker",
            "--draft-marker",
            "always",
        )
    except SystemExit as error:
        assert error.code == 1
    else:  # pragma: no cover - defensive assertion
        raise AssertionError("Expected marker overwrite refusal")
    assert (tmp_path / ".template-sync" / "marker.yml").read_text(encoding="utf-8") == before


def test_write_draft_marker_without_preview_still_creates_marker(tmp_path: Path) -> None:
    """--write-draft-marker writes the marker even when --draft-marker never hides the preview."""
    _init_fixture_repo(tmp_path)

    output = _run_bootstrap(tmp_path, "--write", "--write-draft-marker", "--draft-marker", "never")

    marker_path = tmp_path / ".template-sync" / "marker.yml"
    assert marker_path.is_file()
    assert "created from validated draft marker" in output
    # The preview is suppressed, so the draft-marker preview section is not printed.
    assert "## Draft `.template-sync/marker.yml`" not in output
    marker = yaml.safe_load(marker_path.read_text(encoding="utf-8"))
    parse_marker_decision_data(marker, validate_protected_decision_integrity=True)


def test_existing_durable_path_status_flags_symlinks(tmp_path: Path) -> None:
    """A symlinked or broken-symlink durable path is reported as a symlink, not found/missing."""
    target = tmp_path / "real-marker.yml"
    target.write_text("template_sync:\n  source_repo: x\n", encoding="utf-8")
    linked = tmp_path / "linked-marker.yml"
    broken = tmp_path / "broken-marker.yml"
    try:
        linked.symlink_to(target)
        broken.symlink_to(tmp_path / "does-not-exist.yml")
    except (OSError, NotImplementedError):
        pytest.skip("Filesystem does not support symlink creation")

    symlink_status = "present as a symlink (write mode will refuse it)"
    assert bootstrap.existing_durable_path_status(linked) == symlink_status
    assert bootstrap.existing_durable_path_status(broken) == symlink_status
    assert bootstrap.existing_durable_path_status(target) == "found"
    assert bootstrap.existing_durable_path_status(tmp_path / "absent.yml") == "missing"


def test_read_existing_todo_state_rejects_symlinked_checklist(tmp_path: Path) -> None:
    """A symlink (even broken) at the checklist path is rejected on read, not treated as missing."""
    todo_path = tmp_path / "_TODO-repo-init.md"
    try:
        todo_path.symlink_to(tmp_path / "missing-target.md")  # broken symlink
    except (OSError, NotImplementedError):
        pytest.skip("Filesystem does not support symlink creation")

    with pytest.raises(bootstrap.FirstAdoptionBootstrapError):
        bootstrap.read_existing_todo_state(todo_path)
