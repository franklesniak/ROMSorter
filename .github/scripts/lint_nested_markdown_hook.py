"""Run the root nested-Markdown tool for filenames selected by pre-commit."""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    """Preserve argument boundaries and the linter's native failure status."""
    filenames = sys.argv[1:] if argv is None else argv
    if not filenames or any(not name for name in filenames):
        print("Nested Markdown hook requires nonempty filenames.", file=sys.stderr)
        return 1
    node = shutil.which("node")
    if node is None:
        print(
            "Nested Markdown requires Node.js. Install a supported Node.js version, "
            "then run npm ci --ignore-scripts at the repository root.",
            file=sys.stderr,
        )
        return 1
    repo_root = Path(__file__).resolve().parents[2]
    try:
        result = subprocess.run(
            [node, str(repo_root / ".github/scripts/lint-nested-markdown.js"), *filenames],
            cwd=repo_root,
            check=False,
        )
    except OSError as error:
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        print(f"Cannot start nested Markdown lint: {error_summary}", file=sys.stderr)
        return 1
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
