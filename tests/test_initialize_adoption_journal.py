"""Exercise adoption difficulties journal initialization."""

from __future__ import annotations

import importlib
import subprocess
import sys
from pathlib import Path
from typing import Any, cast

from tests._pytest_compat import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT_PATH = REPO_ROOT / ".template-sync" / "scripts" / "initialize_adoption_journal.py"
SCRIPT_DIR = SCRIPT_PATH.parent

if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

initialize_adoption_journal = cast(Any, importlib.import_module("initialize_adoption_journal"))


def _write_scaffold(repo_root: Path, text: str = "# Adoption Difficulties Journal\n") -> Path:
    """Create the scaffold file expected by the helper in a fixture repository."""
    scaffold = repo_root / "templates" / "adoption" / "_TEMPLATE-ADOPTION-DIFFICULTIES.md"
    scaffold.parent.mkdir(parents=True, exist_ok=True)
    scaffold.write_text(text, encoding="utf-8")
    return scaffold


def _run_helper(repo_root: Path, *extra_args: str) -> subprocess.CompletedProcess[str]:
    """Run the adoption journal helper against a fixture repository."""
    return subprocess.run(
        [
            sys.executable,
            str(SCRIPT_PATH),
            "--repo-root",
            str(repo_root),
            *extra_args,
        ],
        cwd=repo_root,
        check=False,
        capture_output=True,
        text=True,
    )


def test_cli_creates_default_journal_name_when_absent(tmp_path: Path) -> None:
    """The documented command creates `_ADOPTION-DIFFICULTIES.md` by default."""
    _write_scaffold(
        tmp_path,
        "# Adoption Difficulties Journal\n\nProcess and difficulty evidence.\n",
    )

    result = _run_helper(tmp_path)

    journal = tmp_path / "_ADOPTION-DIFFICULTIES.md"
    assert result.returncode == 0, result.stderr
    assert "Created adoption difficulties journal: `_ADOPTION-DIFFICULTIES.md`" in result.stdout
    assert journal.is_file()
    assert journal.name == "_ADOPTION-DIFFICULTIES.md"
    assert "_TEMPLATE" not in journal.name
    assert "Process and difficulty evidence." in journal.read_text(encoding="utf-8")


def test_journal_path_override_creates_parent_directories(tmp_path: Path) -> None:
    """The helper accepts a repository-contained journal path override."""
    scaffold = _write_scaffold(tmp_path, "# Adoption Difficulties Journal\n")
    journal = tmp_path / "notes" / "adoption-difficulties.md"

    result = initialize_adoption_journal.create_adoption_journal(
        repo_root=tmp_path,
        journal_path=journal,
        template_path=scaffold,
    )

    assert result.created is True
    assert result.path == journal
    assert journal.read_text(encoding="utf-8") == "# Adoption Difficulties Journal\n"


def test_existing_journal_is_preserved(tmp_path: Path) -> None:
    """Existing journal content is never overwritten."""
    scaffold = _write_scaffold(tmp_path, "# Scaffold\n")
    journal = tmp_path / "_ADOPTION-DIFFICULTIES.md"
    journal.write_text("local evidence\n", encoding="utf-8")

    result = initialize_adoption_journal.create_adoption_journal(
        repo_root=tmp_path,
        journal_path=journal,
        template_path=scaffold,
    )

    assert result.created is False
    assert journal.read_text(encoding="utf-8") == "local evidence\n"


def test_directory_at_journal_path_is_rejected(tmp_path: Path) -> None:
    """A directory at the journal path is surfaced, not reported as existing."""
    scaffold = _write_scaffold(tmp_path, "# Scaffold\n")
    journal = tmp_path / "_ADOPTION-DIFFICULTIES.md"
    journal.mkdir()

    with pytest.raises(initialize_adoption_journal.AdoptionJournalError):
        initialize_adoption_journal.create_adoption_journal(
            repo_root=tmp_path,
            journal_path=journal,
            template_path=scaffold,
        )


def test_symlinked_journal_path_is_rejected(tmp_path: Path) -> None:
    """A symlink at the journal path is treated as an error, not an existing journal."""
    scaffold = _write_scaffold(tmp_path, "# Scaffold\n")
    target = tmp_path / "real-journal.md"
    target.write_text("real evidence\n", encoding="utf-8")
    journal = tmp_path / "_ADOPTION-DIFFICULTIES.md"
    try:
        journal.symlink_to(target)
    except (OSError, NotImplementedError):
        pytest.skip("Filesystem does not support symlink creation")

    with pytest.raises(initialize_adoption_journal.AdoptionJournalError):
        initialize_adoption_journal.create_adoption_journal(
            repo_root=tmp_path,
            journal_path=journal,
            template_path=scaffold,
        )

    assert target.read_text(encoding="utf-8") == "real evidence\n"
