"""Assert invalid template-placeholder schema examples are rejected."""

from __future__ import annotations

import argparse
import importlib
import json
import sys
from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import Any, cast

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_SCHEMA_PATH = "schemas/template-placeholders.schema.json"
DEFAULT_INVALID_EXAMPLES_DIR = "schemas/examples/template-placeholders/invalid"
DIAGNOSTIC_LIMIT = 10


class PlaceholderExampleValidationError(RuntimeError):
    """Raised when placeholder example validation cannot complete."""


def load_json_object(path: Path, display_path: str) -> dict[str, Any]:
    """Load a JSON file that must contain an object."""
    try:
        parsed: object = json.loads(path.read_text(encoding="utf-8"))
    except OSError as error:
        error_summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        raise PlaceholderExampleValidationError(
            f"{display_path}: unable to read file ({error_summary})."
        ) from error
    except json.JSONDecodeError as error:
        raise PlaceholderExampleValidationError(
            f"{display_path}: invalid JSON ({error})."
        ) from error
    if not isinstance(parsed, dict):
        raise PlaceholderExampleValidationError(f"{display_path}: JSON document must be an object.")
    return cast(dict[str, Any], parsed)


def repository_path(repo_root: Path, raw_path: str, field_name: str) -> Path:
    """Resolve a repository-relative path without allowing traversal."""
    candidate = repo_root / raw_path
    resolved = candidate.resolve()
    try:
        resolved.relative_to(repo_root.resolve())
    except ValueError as error:
        raise PlaceholderExampleValidationError(
            f"{field_name} must stay inside the repository: {raw_path}"
        ) from error
    return resolved


def iter_invalid_examples(examples_dir: Path, repo_root: Path) -> tuple[Path, ...]:
    """Return regular invalid example files below ``examples_dir``."""
    if not examples_dir.is_dir():
        raise PlaceholderExampleValidationError(
            f"Invalid example directory does not exist: {examples_dir.relative_to(repo_root).as_posix()}"
        )
    examples: list[Path] = []
    for path in sorted(examples_dir.glob("*.json")):
        if path.is_symlink() or not path.is_file():
            continue
        examples.append(path)
    if not examples:
        raise PlaceholderExampleValidationError(
            f"No invalid template-placeholder examples found in {examples_dir.relative_to(repo_root).as_posix()}."
        )
    return tuple(examples)


def import_jsonschema_validator() -> Any:
    """Import the optional JSON Schema validator used by the hook environment."""
    try:
        jsonschema_module = cast(Any, importlib.import_module("jsonschema"))
    except ImportError as error:
        raise PlaceholderExampleValidationError(
            "jsonschema is unavailable. Install jsonschema, or run this through "
            "the pre-commit hook, which declares the validator dependency."
        ) from error
    return jsonschema_module.Draft202012Validator


def validation_errors_for(
    validator: Any,
    example: dict[str, Any],
) -> tuple[str, ...]:
    """Return bounded validation errors for one invalid example."""
    referencing_errors = cast(Any, importlib.import_module("referencing.exceptions"))
    try:
        errors = sorted(validator.iter_errors(example), key=lambda error: error.json_path)
    except referencing_errors.Unresolvable as error:
        raise PlaceholderExampleValidationError(
            "Unable to resolve a placeholder example schema reference."
        ) from error
    return tuple(f"{error.json_path}: {error.message}" for error in errors[:DIAGNOSTIC_LIMIT])


def validate_invalid_examples(
    *,
    repo_root: Path,
    schema_path: Path,
    examples_dir: Path,
) -> tuple[str, ...]:
    """Return failures for invalid examples that unexpectedly validate."""
    validator_class = import_jsonschema_validator()
    schema = load_json_object(schema_path, schema_path.relative_to(repo_root).as_posix())
    referencing_module = cast(Any, importlib.import_module("referencing"))
    validator = validator_class(schema, registry=referencing_module.Registry())
    failures: list[str] = []
    for example_path in iter_invalid_examples(examples_dir, repo_root):
        relative_example_path = example_path.relative_to(repo_root).as_posix()
        example = load_json_object(example_path, relative_example_path)
        errors = validation_errors_for(validator, example)
        if not errors:
            failures.append(f"{relative_example_path}: expected rejection but validation passed.")
    return tuple(failures)


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Assert invalid template-placeholder schema examples are rejected."
    )
    parser.add_argument("--repo-root", default=None, help="Repository root.")
    parser.add_argument(
        "--schema",
        default=DEFAULT_SCHEMA_PATH,
        help=f"Schema path relative to the repository root. Default: {DEFAULT_SCHEMA_PATH}",
    )
    parser.add_argument(
        "--invalid-examples",
        default=DEFAULT_INVALID_EXAMPLES_DIR,
        help=(
            "Invalid examples directory relative to the repository root. "
            f"Default: {DEFAULT_INVALID_EXAMPLES_DIR}"
        ),
    )
    return parser.parse_args(argv)


def print_failures(failures: Iterable[str]) -> None:
    """Print validation failures."""
    failure_items = tuple(failures)
    if not failure_items:
        print("Invalid template-placeholder examples were rejected.")
        return
    print("Invalid template-placeholder example validation failed:")
    for failure in failure_items:
        print(f"  - {failure}")


def main(argv: Sequence[str] | None = None) -> int:
    """Run invalid-example validation."""
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        repo_root = Path(args.repo_root).expanduser().resolve() if args.repo_root else REPO_ROOT
        if not repo_root.is_dir():
            raise PlaceholderExampleValidationError("Repository root does not exist.")
        schema_path = repository_path(repo_root, args.schema, "--schema")
        examples_dir = repository_path(repo_root, args.invalid_examples, "--invalid-examples")
        failures = validate_invalid_examples(
            repo_root=repo_root,
            schema_path=schema_path,
            examples_dir=examples_dir,
        )
    except PlaceholderExampleValidationError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1

    print_failures(failures)
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
