r"""Validate schema example files with ``check-jsonschema``.

This is the **active, canonical** schema-example test for this
repository. It auto-discovers schema/example pairs under ``schemas/``
and verifies that:

- Every file under ``schemas/examples/<schema-name>/valid/`` validates
  successfully against ``schemas/<schema-name>.schema.json``
  (``check-jsonschema`` exits ``0``).
- Every file under ``schemas/examples/<schema-name>/invalid/`` is
  rejected (``check-jsonschema`` exits non-zero).

JSON Schema reads ``pattern`` in the ECMA-262 dialect, which
``check-jsonschema`` uses by default. python-jsonschema, which the
template's own scripts use, reads it in Python's ``re`` dialect, which
``--regex-variant python`` selects. The examples of the schemas in
``PYTHON_DIALECT_SCHEMAS`` run in both dialects. Every other schema,
such as one an adopter adds, runs in ECMA-262 only, because it can use
ECMA-262 syntax that Python's ``re`` reads differently. Two guards make
a pattern read the same way in both dialects:

- In Python's dialect ``$`` also matches before a final line break, so a
  whole-value pattern ends with ``$(?![\s\S])``. The
  ``invalid/trailing-newline-*`` examples hold one value that ends in a
  line break.
- In Python's dialect ``.`` also matches a carriage return, U+2028, and
  U+2029, so a path pattern uses ``[^\n\r\u2028\u2029]``, which is
  exactly ECMA-262's ``.``. The ``invalid/trailing-carriage-return-*``
  examples hold one path that ends in a carriage return.

A second test removes one guard from a copy of the schema and checks
that Python's dialect then accepts each example that depends on it,
which proves that the examples detect a missing guard. A third test
fails any pattern in a Python-dialect schema that ends with a bare
``$``, so a lost end guard fails even where no example exists. A
fourth test fails when a template script names a schema that is not in
``PYTHON_DIALECT_SCHEMAS``. It reads only the scripts that
``.template-sync/manifest.yml`` maps, so it never reads an adopter's
own scripts, and it skips when no manifest is present.

Discovery rules:

- Schemas are found via the glob ``schemas/*.schema.json``.
- For each schema, the example directory name is derived by removing
  the trailing ``.schema.json`` from the file name. For example,
  ``schemas/example-config.schema.json`` maps to
  ``schemas/examples/example-config/``.
- Only regular files under ``valid/`` and ``invalid/`` are exercised;
  directories and other non-file entries are ignored.

Paths are resolved from the repository root, which is derived from
this file's known location at ``<repo>/tests/test_schema_examples.py``
(via ``Path(__file__).resolve().parent.parent``) rather than the
process current working directory, so the test behaves the same
regardless of where ``pytest`` is invoked from.

Invalid examples are intentionally NOT wired into a normal
``check-jsonschema`` pre-commit hook, because a failing exit code from
the validator would be reported as a hook failure. This test exists in
part to prove that the schema actually rejects them.

A starter version of this test is available at
``templates/python/tests/test_schema_examples.py`` for downstream
consumers of the template. The two files share the same discovery
and ECMA-262 validation. The starter does not run Python's regex
dialect or the guard tests, because an adopter's schemas can use
ECMA-262-only syntax.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
from collections.abc import Callable
from fnmatch import fnmatchcase
from importlib.util import find_spec
from pathlib import Path

from tests._pytest_compat import pytest


def _check_jsonschema_command() -> list[str] | None:
    """Resolve the preferred ``check-jsonschema`` invocation for this environment."""
    executable = shutil.which("check-jsonschema")
    if executable is not None:
        return [executable]
    if find_spec("check_jsonschema") is not None:
        return [sys.executable, "-m", "check_jsonschema"]
    return None


CHECK_JSONSCHEMA_COMMAND = _check_jsonschema_command()

# Repository root, relative to this file: ``<repo>/tests/test_schema_examples.py``.
REPO_ROOT = Path(__file__).resolve().parent.parent
SCHEMAS_DIR = REPO_ROOT / "schemas"
EXAMPLES_DIR = SCHEMAS_DIR / "examples"
SCHEMA_SUFFIX = ".schema.json"
END_GUARD = r"(?![\s\S])"
# Exactly ECMA-262's ``.``: any character except LF, CR, U+2028, and U+2029.
LINE_CHARACTER_CLASS = r"[^\n\r\u2028\u2029]"
# The schemas that the template's own Python scripts read with python-jsonschema. Only their
# examples also run in Python's regex dialect (``--regex-variant python``). A schema that an
# adopter adds stays ECMA-262 only: it can use ECMA-262 syntax, such as ``\p{Lu}``, that
# Python's ``re`` rejects or reads differently.
PYTHON_DIALECT_SCHEMAS = frozenset(
    {
        "instruction-contracts.schema.json",
        "instruction-profile.schema.json",
        "template-placeholders.schema.json",
        "template-sync-instruction-contracts.schema.json",
        "template-sync-manifest.schema.json",
        "template-sync-marker.schema.json",
        "workflow-security-contract.schema.json",
    }
)
# The template script directories, and a schema path as a script names it. The list test reads
# only the scripts in these directories that the template-sync manifest maps: the template's
# own scripts, never an adopter's.
TEMPLATE_SCRIPT_DIRS = (REPO_ROOT / ".github" / "scripts", REPO_ROOT / ".template-sync" / "scripts")
SCRIPT_SCHEMA_REFERENCE = re.compile(r"schemas/([a-z0-9-]+\.schema\.json)")
TEMPLATE_MANIFEST = REPO_ROOT / ".template-sync" / "manifest.yml"


def _is_within_root(candidate: Path, root: Path) -> bool:
    """Return ``True`` iff ``candidate`` is reachable from ``root`` with no symlink ancestors.

    Walks upward from ``candidate`` to ``root`` (inclusive), checking
    ``Path.is_symlink()`` at each step. Any symlink found on that path
    causes a refusal, which closes the symlink-rooted-discovery
    escape: a malicious symlink at, for example,
    ``schemas/examples/<schema-name>``, ``schemas/examples/``,
    ``schemas/``, or even ``schemas/<schema-name>.schema.json`` itself
    would otherwise survive a downstream ``os.walk(followlinks=False)``
    or a yielded-path-only ``Path.resolve().relative_to(base)`` check,
    because those checks examine paths *under* ``candidate`` rather
    than the symlink chain *to* ``candidate``.

    Also returns ``False`` if ``candidate.resolve()`` does not stay
    inside ``root.resolve()``, or if any ``OSError`` / ``ValueError``
    is raised while resolving (callers should treat such errors as a
    hard refusal rather than a recoverable condition, which is why
    they are caught here and surfaced as a boolean).

    This helper does **not** enforce a mount-point or device boundary
    (``Path.is_symlink()`` does not detect bind mounts). Callers that
    require that level of isolation should layer an explicit
    ``os.stat().st_dev`` (or mount-table) check on top of this helper.

    Args:
        candidate: Path that the caller wants to use as a discovery
            root or as a discovered fixture. May be a file or a
            directory; need not exist (``is_symlink`` is well-defined
            on dangling symlinks).
        root: Trusted ancestor that ``candidate`` must live inside.
            Typically a constant such as ``REPO_ROOT`` or
            ``SCHEMAS_DIR`` derived from ``Path(__file__)`` and
            therefore not user-controlled.

    Returns:
        ``True`` only when ``candidate`` resolves under ``root`` *and*
        every path component from ``candidate`` up to and including
        ``root`` is a regular (non-symlink) entry; ``False`` otherwise.
    """
    try:
        candidate.resolve().relative_to(root.resolve())
        cur = candidate
        while True:
            if cur.is_symlink():
                return False
            if cur == root:
                return True
            parent = cur.parent
            if parent == cur:
                return False
            cur = parent
    except (OSError, ValueError):
        return False


def _iter_safe_files(directory: Path, root: Path) -> list[Path]:
    """List regular files under ``directory``, refusing symlink escapes.

    Caller's contract: ``directory`` must be intended to live inside
    ``root``. This helper enforces that contract before walking by
    calling :func:`_is_within_root`, which rejects discovery if
    ``directory`` (or any ancestor up to ``root``) is itself a
    symlink, or if ``directory.resolve()`` escapes ``root.resolve()``.
    Without this pre-walk check, an attacker who controls a path
    component (e.g., a symlink at ``schemas/examples/<schema-name>``
    pointing outside the repository) could coerce the test suite into
    walking and validating files outside ``root`` even though
    ``os.walk(followlinks=False)`` and the per-yielded-path
    boundary check would each, in isolation, appear to "succeed".

    Once the discovery root is accepted, walks ``directory`` with
    ``os.walk(followlinks=False)`` so that symlinked subdirectories
    are never traversed, drops any yielded entry that is itself a
    symlink, and re-checks each remaining path with
    ``Path.resolve()`` and ``Path.relative_to`` to guarantee its
    fully-resolved path string still lives inside ``directory``'s
    resolved location, which catches symlink and ``..``-style
    traversals that survive the initial walk-time skip.

    The repository constitution requires that file-discovery callers
    "reject path traversal and symlink escapes"; this helper
    centralizes that policy for example-fixture discovery so the test
    suite cannot be coerced into validating files outside the
    schemas/examples tree by a malicious or accidentally-introduced
    symlink at any level.

    This helper does **not** enforce a mount-point or device boundary.
    A bind mount (or any other mount) located *under* ``directory``
    will still expose its target content because ``Path.resolve()``
    returns a path that, from the filesystem's perspective, remains
    inside ``directory``. Callers that require mount-boundary
    isolation should layer an explicit ``os.stat().st_dev`` (or
    mount-table) check on top of this helper.

    Args:
        directory: Directory whose regular-file descendants should be
            returned. May not exist; if absent, an empty list is
            returned.
        root: Trusted ancestor of ``directory`` (typically
            ``REPO_ROOT``). Discovery is refused if ``directory`` or
            any ancestor up to ``root`` is a symlink, or if
            ``directory`` resolves outside ``root``.

    Returns:
        A list of regular-file ``Path`` objects below ``directory``,
        in the order produced by ``os.walk`` (callers that need a
        deterministic order should sort the result). Returns ``[]``
        when ``directory`` is missing, fails the symlink-ancestor /
        containment check against ``root``, or contains no eligible
        regular files.
    """
    if not directory.is_dir():
        return []
    if not _is_within_root(directory, root):
        return []
    base = directory.resolve()
    discovered: list[Path] = []
    for current_root, _dir_names, file_names in os.walk(directory, followlinks=False):
        for file_name in file_names:
            file_path = Path(current_root) / file_name
            if file_path.is_symlink():
                continue
            try:
                file_path.resolve().relative_to(base)
            except (OSError, ValueError):
                continue
            discovered.append(file_path)
    return discovered


def _discover_cases() -> list[tuple[Path, Path, bool]]:
    """Discover ``(schema, example, expected_to_pass)`` triples.

    Returns:
        A list of ``(schema_path, example_path, expected_to_pass)``
        tuples covering every regular file under
        ``schemas/examples/<schema-name>/valid/`` (expected to pass)
        and ``schemas/examples/<schema-name>/invalid/`` (expected to
        fail). Schema files and example directories whose path chain
        from ``REPO_ROOT`` includes a symlink (or that resolve outside
        ``REPO_ROOT``) are rejected via :func:`_is_within_root` so a
        malicious or accidentally-introduced symlink at any level
        cannot coerce the suite into validating files outside the
        repository. Returns an empty list if no schemas or examples
        are present, so the test suite degrades gracefully rather
        than hard-failing on a count assertion.
    """
    cases: list[tuple[Path, Path, bool]] = []
    if not SCHEMAS_DIR.is_dir():
        return cases
    for schema_path in sorted(SCHEMAS_DIR.glob(f"*{SCHEMA_SUFFIX}")):
        if not _is_within_root(schema_path, REPO_ROOT):
            continue
        schema_name = schema_path.name[: -len(SCHEMA_SUFFIX)]
        schema_examples_dir = EXAMPLES_DIR / schema_name
        for kind, expected_to_pass in (("valid", True), ("invalid", False)):
            kind_dir = schema_examples_dir / kind
            if not kind_dir.is_dir():
                continue
            for example_path in sorted(_iter_safe_files(kind_dir, REPO_ROOT)):
                if not example_path.is_file():
                    continue
                cases.append((schema_path, example_path, expected_to_pass))
    return cases


def _case_id(case: tuple[Path, Path, bool]) -> str:
    """Build a readable parametrize ID for a discovered case.

    The ID identifies both the schema and the example path relative to
    the repository root, plus the expected outcome, so failing cases
    are easy to locate from pytest output.
    """
    schema_path, example_path, expected_to_pass = case
    schema_rel = schema_path.relative_to(REPO_ROOT).as_posix()
    example_rel = example_path.relative_to(REPO_ROOT).as_posix()
    outcome = "valid" if expected_to_pass else "invalid"
    return f"{schema_rel}::{outcome}::{example_rel}"


_CASES = _discover_cases()
# Every example runs in ECMA-262 (``default``). The examples of the Python-dialect schemas also
# run in Python's ``re`` dialect (``python``).
_VARIANT_CASES = [(*case, "default") for case in _CASES] + [
    (*case, "python") for case in _CASES if case[0].name in PYTHON_DIALECT_SCHEMAS
]


@pytest.mark.skipif(
    CHECK_JSONSCHEMA_COMMAND is None,
    reason="check-jsonschema is not installed in this environment",
)
@pytest.mark.skipif(
    not _CASES,
    reason="No schema example files found under schemas/examples/",
)
@pytest.mark.parametrize(
    ("schema_path", "example_path", "expected_to_pass", "regex_variant"),
    _VARIANT_CASES,
    ids=[f"{case[3]}-{_case_id(case[:3])}" for case in _VARIANT_CASES],
)
def test_schema_example(
    schema_path: Path,
    example_path: Path,
    expected_to_pass: bool,
    regex_variant: str,
) -> None:
    """Validate one ``(schema, example)`` pair against its labeled outcome.

    Args:
        schema_path: Absolute path to a ``*.schema.json`` file under
            ``schemas/``.
        example_path: Absolute path to an example file under either
            ``schemas/examples/<schema-name>/valid/`` or
            ``schemas/examples/<schema-name>/invalid/``.
        expected_to_pass: ``True`` when the example lives under
            ``valid/`` (must validate cleanly), ``False`` when it
            lives under ``invalid/`` (must be rejected).
        regex_variant: The ``check-jsonschema`` ``--regex-variant``
            value: ``default`` for ECMA-262 or ``python`` for Python's
            ``re`` dialect. ``python`` cases exist only for the schemas
            in ``PYTHON_DIALECT_SCHEMAS``.

    Raises:
        AssertionError: If a valid example is rejected, or an invalid
            example is accepted, by ``check-jsonschema`` in either
            regex dialect.
    """
    validator_command = CHECK_JSONSCHEMA_COMMAND
    # The command is non-None here because of the skipif guard above;
    # assert for type-checkers and as a defensive runtime check.
    assert validator_command is not None

    result = subprocess.run(
        [
            *validator_command,
            "--regex-variant",
            regex_variant,
            "--schemafile",
            str(schema_path),
            str(example_path),
        ],
        check=False,
        capture_output=True,
        text=True,
    )

    if expected_to_pass:
        assert result.returncode == 0, (
            f"Valid example {example_path} was unexpectedly rejected by "
            f"{schema_path} in the {regex_variant} regex dialect."
            f"\nstdout:\n{result.stdout}\nstderr:\n{result.stderr}"
        )
    else:
        assert result.returncode != 0, (
            f"Invalid example {example_path} was unexpectedly accepted by "
            f"{schema_path} in the {regex_variant} regex dialect; the schema may be "
            f"too permissive or the example is no longer invalid."
            f"\nstdout:\n{result.stdout}\nstderr:\n{result.stderr}"
        )


def _cut_end_guard(pattern: str) -> str:
    r"""Cut a final ``$(?![\s\S])`` back to ``$``; return other patterns unchanged."""
    if pattern.endswith("$" + END_GUARD):
        return pattern[: -len(END_GUARD)]
    return pattern


def _cut_line_character_class(pattern: str) -> str:
    r"""Cut each ``[^\n\r\u2028\u2029]`` back to ``.``; return other patterns unchanged."""
    return pattern.replace(LINE_CHARACTER_CLASS, ".")


# Each guard that makes a pattern read the same way in both regex dialects: the file-name
# prefix of the invalid examples that depend on it, and a cut that removes it from a pattern.
DIALECT_GUARDS: dict[str, tuple[str, Callable[[str], str]]] = {
    "end-guard": ("trailing-newline-", _cut_end_guard),
    "line-character-class": ("trailing-carriage-return-", _cut_line_character_class),
}


def _rewrite_patterns(node: object, cut: Callable[[str], str]) -> object:
    """Return a copy of a schema node with ``cut`` applied to every pattern.

    Args:
        node: A parsed JSON Schema value.
        cut: A function that rewrites one regular expression.

    Returns:
        A copy of ``node`` in which every ``pattern`` value and every
        ``patternProperties`` key is replaced by ``cut(pattern)``. Other
        values are unchanged.
    """
    if isinstance(node, dict):
        copied: dict[str, object] = {}
        for key, value in node.items():
            if key == "pattern" and isinstance(value, str):
                copied[key] = cut(value)
            elif key == "patternProperties" and isinstance(value, dict):
                copied[key] = {
                    cut(name): _rewrite_patterns(sub, cut) for name, sub in value.items()
                }
            else:
                copied[key] = _rewrite_patterns(value, cut)
        return copied
    if isinstance(node, list):
        return [_rewrite_patterns(item, cut) for item in node]
    return node


def _discover_guard_cases() -> list[tuple[Path, str, tuple[Path, ...]]]:
    """Group the invalid examples that depend on a dialect guard by schema and guard.

    Returns:
        A list of ``(schema_path, guard_name, example_paths)`` tuples, one
        for each schema and :data:`DIALECT_GUARDS` entry that has at least
        one invalid example whose file name starts with that guard's
        prefix. Only schemas in ``PYTHON_DIALECT_SCHEMAS`` count.
        Discovery reuses :func:`_discover_cases`, so the same symlink
        and containment checks apply.
    """
    grouped: dict[tuple[Path, str], list[Path]] = {}
    for schema_path, example_path, expected_to_pass in _CASES:
        if expected_to_pass or schema_path.name not in PYTHON_DIALECT_SCHEMAS:
            continue
        for guard_name, (prefix, _cut) in DIALECT_GUARDS.items():
            if example_path.name.startswith(prefix):
                grouped.setdefault((schema_path, guard_name), []).append(example_path)
    return [
        (schema_path, guard_name, tuple(example_paths))
        for (schema_path, guard_name), example_paths in sorted(grouped.items())
    ]


def _rejected_examples(
    schema_path: Path, example_paths: tuple[Path, ...], regex_variant: str
) -> set[Path]:
    """Return the examples that ``check-jsonschema`` rejects in one regex dialect.

    Args:
        schema_path: The schema file to validate against.
        example_paths: The instance files to validate in one run.
        regex_variant: The ``--regex-variant`` value.

    Returns:
        The subset of ``example_paths`` that have at least one
        validation error.

    Raises:
        AssertionError: If the validator output is not the expected JSON
            report, or if an example cannot be parsed.
    """
    assert CHECK_JSONSCHEMA_COMMAND is not None
    result = subprocess.run(
        [
            *CHECK_JSONSCHEMA_COMMAND,
            "--output-format",
            "JSON",
            "--regex-variant",
            regex_variant,
            "--schemafile",
            str(schema_path),
            *[str(path) for path in example_paths],
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    try:
        report = json.loads(result.stdout)
    except json.JSONDecodeError as error:
        raise AssertionError(
            f"check-jsonschema did not print a JSON report.\n"
            f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
        ) from error
    assert not report.get("parse_errors"), report
    return {Path(error["filename"]) for error in report["errors"]}


_GUARD_CASES = _discover_guard_cases()


@pytest.mark.skipif(
    CHECK_JSONSCHEMA_COMMAND is None,
    reason="check-jsonschema is not installed in this environment",
)
@pytest.mark.skipif(
    not _GUARD_CASES,
    reason="No invalid schema examples that depend on a dialect guard were found",
)
@pytest.mark.parametrize(
    ("schema_path", "guard_name", "example_paths"),
    _GUARD_CASES,
    ids=[
        f"{schema.relative_to(REPO_ROOT).as_posix()}::{guard_name}"
        for schema, guard_name, _ in _GUARD_CASES
    ],
)
def test_dialect_guard_examples_depend_on_their_guard(
    schema_path: Path,
    guard_name: str,
    example_paths: tuple[Path, ...],
    tmp_path: Path,
) -> None:
    r"""Remove one dialect guard and check that only Python's dialect then accepts the examples.

    This is the failure-injection control for :func:`test_schema_example`.
    A copy of the schema loses one guard from every pattern:

    - ``end-guard``: each ``$(?![\s\S])`` is cut back to ``$``. Python's
      ``$`` also matches before a final line break, so Python's dialect
      must accept each ``trailing-newline-*`` example. ECMA-262's ``$``
      matches only at the end of the text, so ECMA-262 must still reject
      each one.
    - ``line-character-class``: each ``[^\n\r\u2028\u2029]`` is cut back
      to ``.``. Python's ``.`` also matches a carriage return, so Python's
      dialect must accept each ``trailing-carriage-return-*`` example.
      ECMA-262's ``.`` refuses all four line terminators, so ECMA-262 must
      still reject each one.

    The examples therefore detect a missing guard.

    Args:
        schema_path: A schema with invalid examples that depend on the guard.
        guard_name: A :data:`DIALECT_GUARDS` key.
        example_paths: Those examples.
        tmp_path: Pytest's per-test temporary directory.

    Raises:
        AssertionError: If the schema has no such guard, if ECMA-262
            accepts an example without the guard, or if Python's dialect
            rejects an example without the guard.
    """
    prefix, cut = DIALECT_GUARDS[guard_name]
    schema = json.loads(schema_path.read_text(encoding="utf-8"))
    unguarded = _rewrite_patterns(schema, cut)
    assert unguarded != schema, f"{schema_path} has no pattern with the {guard_name}"
    unguarded_path = tmp_path / schema_path.name
    unguarded_path.write_text(json.dumps(unguarded, indent=2), encoding="utf-8")

    rejected_by_ecma = _rejected_examples(unguarded_path, example_paths, "default")
    rejected_by_python = _rejected_examples(unguarded_path, example_paths, "python")

    assert rejected_by_ecma == set(example_paths), (
        f"Without the {guard_name}, ECMA-262 accepted "
        f"{sorted(str(p) for p in set(example_paths) - rejected_by_ecma)}; "
        f"these examples do not isolate a difference between the dialects."
    )
    assert not rejected_by_python, (
        f"Without the {guard_name}, Python's dialect still rejected "
        f"{sorted(str(p) for p in rejected_by_python)}; each {prefix}* example "
        f"must differ from a valid document only by one character that the "
        f"{guard_name} refuses."
    )


def _iter_patterns(node: object) -> list[str]:
    """Return every ``pattern`` value and ``patternProperties`` key in a schema node.

    Args:
        node: A parsed JSON Schema value.

    Returns:
        The regular expressions in ``node``, in document order.
    """
    found: list[str] = []
    if isinstance(node, dict):
        for key, value in node.items():
            if key == "pattern" and isinstance(value, str):
                found.append(value)
            elif key == "patternProperties" and isinstance(value, dict):
                found.extend(value)
            found.extend(_iter_patterns(value))
    elif isinstance(node, list):
        for item in node:
            found.extend(_iter_patterns(item))
    return found


def _ends_with_bare_dollar(pattern: str) -> bool:
    """Return ``True`` when ``pattern`` ends with an unescaped ``$``.

    Args:
        pattern: A regular expression.

    Returns:
        ``True`` when the final ``$`` is an anchor, that is, when an even
        number of backslashes precedes it.
    """
    if not pattern.endswith("$"):
        return False
    body = pattern[:-1]
    return (len(body) - len(body.rstrip("\\"))) % 2 == 0


@pytest.mark.parametrize(
    ("pattern", "expected"),
    [
        ("^a$", True),
        ("^a\\$", False),
        ("^a\\\\$", True),
        (r"^a$(?![\s\S])", False),
        (r"[\r\n]", False),
    ],
)
def test_ends_with_bare_dollar(pattern: str, expected: bool) -> None:
    """The bare-``$`` detector separates an anchor from an escaped dollar sign."""
    assert _ends_with_bare_dollar(pattern) is expected


_PYTHON_DIALECT_SCHEMA_PATHS = [
    path
    for path in sorted(SCHEMAS_DIR.glob(f"*{SCHEMA_SUFFIX}"))
    if path.name in PYTHON_DIALECT_SCHEMAS and _is_within_root(path, REPO_ROOT)
]


@pytest.mark.skipif(
    not _PYTHON_DIALECT_SCHEMA_PATHS,
    reason="No Python-dialect schema is present under schemas/",
)
@pytest.mark.parametrize(
    "schema_path", _PYTHON_DIALECT_SCHEMA_PATHS, ids=[p.name for p in _PYTHON_DIALECT_SCHEMA_PATHS]
)
def test_python_dialect_schema_has_no_bare_final_dollar(schema_path: Path) -> None:
    r"""Every whole-value pattern in a Python-dialect schema ends with ``$(?![\s\S])``.

    This static check covers patterns that no example isolates, such as
    the manifest's ``moduleName`` and ``hostName``.

    Args:
        schema_path: A schema in ``PYTHON_DIALECT_SCHEMAS``.

    Raises:
        AssertionError: If a pattern ends with a bare ``$``.
    """
    schema = json.loads(schema_path.read_text(encoding="utf-8"))
    bare = [pattern for pattern in _iter_patterns(schema) if _ends_with_bare_dollar(pattern)]
    assert not bare, f"{schema_path.name}: end each pattern with $(?![\\s\\S]): {bare}"


def _manifest_patterns(document: object) -> list[str]:
    """Return every path-mapping pattern in a loaded template-sync manifest.

    Args:
        document: The parsed ``.template-sync/manifest.yml``.

    Returns:
        The ``pattern`` of each ``template_manifest.path_mappings`` entry.

    Raises:
        AssertionError: If the document has no ``path_mappings`` list.
    """
    manifest = document.get("template_manifest") if isinstance(document, dict) else None
    mappings = manifest.get("path_mappings") if isinstance(manifest, dict) else None
    assert isinstance(mappings, list), "The template-sync manifest has no path_mappings list."
    return [
        entry["pattern"]
        for entry in mappings
        if isinstance(entry, dict) and isinstance(entry.get("pattern"), str)
    ]


def _manifest_maps(patterns: list[str], relative_path: str) -> bool:
    """Return whether a manifest pattern covers ``relative_path``.

    This mirrors ``manifest_pattern_matches_path`` in
    ``.template-sync/scripts/template_sync_materialization_helpers.py``: a
    pattern with ``*``, ``?``, or ``[`` matches with ``fnmatchcase``, and any
    other pattern must equal the path.

    Args:
        patterns: Manifest path-mapping patterns.
        relative_path: A POSIX path relative to the repository root.

    Returns:
        ``True`` when at least one pattern covers the path.
    """
    return any(
        (
            fnmatchcase(relative_path, pattern)
            if any(wildcard in pattern for wildcard in "*?[")
            else pattern == relative_path
        )
        for pattern in patterns
    )


@pytest.mark.parametrize(
    ("relative_path", "expected"),
    [
        (".github/scripts/validate_workflow_security.py", True),
        (".template-sync/scripts/validate_marker.py", True),
        (".github/scripts/check-x-not-y.py", False),
        (".github/scripts/validate_workflow_security.py.bak", False),
    ],
)
def test_manifest_maps(relative_path: str, expected: bool) -> None:
    """The manifest matcher covers exact and glob mappings, and nothing else."""
    patterns = [".github/scripts/validate_workflow_security.py", ".template-sync/scripts/**"]
    assert _manifest_maps(patterns, relative_path) is expected


@pytest.mark.skipif(
    not (TEMPLATE_MANIFEST.is_file() and _is_within_root(TEMPLATE_MANIFEST, REPO_ROOT)),
    reason="No .template-sync/manifest.yml, so no template script is known",
)
def test_python_dialect_schemas_include_every_schema_a_script_names() -> None:
    """Every schema that a template script names also runs in Python's dialect.

    Only the scripts that ``.template-sync/manifest.yml`` maps are read, so
    an adopter's own scripts and schemas never reach this check.

    Raises:
        AssertionError: If a mapped script under ``.github/scripts/`` or
            ``.template-sync/scripts/`` names a schema that is not in
            ``PYTHON_DIALECT_SCHEMAS``.
    """
    yaml = pytest.importorskip("yaml")
    patterns = _manifest_patterns(yaml.safe_load(TEMPLATE_MANIFEST.read_text(encoding="utf-8")))
    named: dict[str, set[str]] = {}
    for directory in TEMPLATE_SCRIPT_DIRS:
        for script in sorted(_iter_safe_files(directory, REPO_ROOT)):
            relative = script.relative_to(REPO_ROOT).as_posix()
            if script.suffix != ".py" or not _manifest_maps(patterns, relative):
                continue
            for name in SCRIPT_SCHEMA_REFERENCE.findall(script.read_text(encoding="utf-8")):
                named.setdefault(name, set()).add(relative)
    missing = {
        name: sorted(scripts)
        for name, scripts in sorted(named.items())
        if name not in PYTHON_DIALECT_SCHEMAS
    }
    assert (
        not missing
    ), f"Template scripts name schemas missing from PYTHON_DIALECT_SCHEMAS: {missing}"
