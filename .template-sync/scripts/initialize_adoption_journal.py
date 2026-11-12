"""Create the first-adoption difficulties journal when it is absent."""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import NoReturn

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from first_adoption_state import DEFAULT_ADOPTION_JOURNAL_PATH  # noqa: E402
from template_sync_materialization_helpers import (  # noqa: E402
    TemplateSyncMaterializationError,
    os_error_summary,
    repository_relative_path,
    resolve_repo_path,
    resolve_repo_root,
)

# Public alias kept for callers such as bootstrap_first_adoption and the tests,
# which catch ``initialize_adoption_journal.AdoptionJournalError``.
AdoptionJournalError = TemplateSyncMaterializationError

DEFAULT_JOURNAL_TEMPLATE_PATH = "templates/adoption/_TEMPLATE-ADOPTION-DIFFICULTIES.md"


@dataclass(frozen=True)
class AdoptionJournalResult:
    """Result of an adoption journal initialization attempt."""

    path: Path
    created: bool


def parse_args(argv: list[str]) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Create a first-adoption difficulties journal from the checked-in scaffold. "
            "The helper is idempotent and never overwrites existing content."
        )
    )
    parser.add_argument(
        "--repo-root",
        default=None,
        help=(
            "Repository root to update. Defaults to the parent of the .template-sync "
            "directory that contains this script."
        ),
    )
    parser.add_argument(
        "--journal-path",
        default=DEFAULT_ADOPTION_JOURNAL_PATH,
        help=(
            "Journal path relative to the repository root. "
            f"Default: {DEFAULT_ADOPTION_JOURNAL_PATH}"
        ),
    )
    return parser.parse_args(argv)


def create_adoption_journal(
    *,
    repo_root: Path,
    journal_path: Path,
    template_path: Path,
) -> AdoptionJournalResult:
    """Create ``journal_path`` from ``template_path`` only when absent.

    A pre-existing journal is preserved only when it is a regular file that is
    not a symlink. Any other existing path kind (directory, symlink, or broken
    symlink) is an ambiguous state that is surfaced as an error instead of being
    silently reported as an already-initialized journal.
    """
    if journal_path.is_symlink() or journal_path.exists():
        if journal_path.is_file() and not journal_path.is_symlink():
            return AdoptionJournalResult(path=journal_path, created=False)
        journal_relative = repository_relative_path(journal_path, repo_root)
        raise AdoptionJournalError(
            f"Adoption journal path {journal_relative} exists but is not a regular file; "
            "remove it or pass a different --journal-path to initialize the journal."
        )

    try:
        template_text = template_path.read_text(encoding="utf-8")
    except OSError as error:
        template_relative = repository_relative_path(template_path, repo_root)
        raise AdoptionJournalError(
            f"Unable to read journal scaffold {template_relative}: {os_error_summary(error)}"
        ) from error

    try:
        journal_path.parent.mkdir(parents=True, exist_ok=True)
        journal_path.write_text(template_text, encoding="utf-8")
    except OSError as error:
        journal_relative = repository_relative_path(journal_path, repo_root)
        raise AdoptionJournalError(
            f"Unable to write adoption journal {journal_relative}: {os_error_summary(error)}"
        ) from error

    return AdoptionJournalResult(path=journal_path, created=True)


def fail(message: str) -> NoReturn:
    """Print an error and exit non-zero."""
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(1)


def main(argv: list[str] | None = None) -> int:
    """Create the adoption difficulties journal if needed."""
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        repo_root = resolve_repo_root(args.repo_root)
        journal_path = resolve_repo_path(repo_root, args.journal_path)
        template_path = resolve_repo_path(repo_root, DEFAULT_JOURNAL_TEMPLATE_PATH)
        result = create_adoption_journal(
            repo_root=repo_root,
            journal_path=journal_path,
            template_path=template_path,
        )
    except AdoptionJournalError as error:
        fail(str(error))

    journal_relative = repository_relative_path(result.path, repo_root)
    if result.created:
        print(f"Created adoption difficulties journal: `{journal_relative}`")
    else:
        print(f"Adoption difficulties journal already exists: `{journal_relative}`")
        print("Existing content was left unchanged.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
