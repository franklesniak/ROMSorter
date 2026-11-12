"""Reject ambiguous policy inputs without requiring downstream YAML linting."""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / ".github/scripts"))
sys.path.insert(0, str(ROOT / ".template-sync/scripts"))

import instruction_contract_core as core  # noqa: E402
import instruction_contract_support as support  # noqa: E402
import materialize_downstream_adoption as materializer  # noqa: E402

pytestmark = pytest.mark.upstream_template_only

RUNTIME = (
    ".github/scripts/instruction_contract_core.py",
    ".github/scripts/instruction_contract_support.py",
    ".github/scripts/validate_instruction_profile.py",
    "schemas/instruction-profile.schema.json",
    "schemas/instruction-contracts.schema.json",
)
PROFILE = (
    "version: 1\n"
    "mode: standalone\n"
    "modules: [agent-instructions, instruction-enforcement, github-actions, agent-codex]\n"
    "exceptions: []\n"
)
CATALOG = (
    "instruction_contracts:\n"
    "  - path: AGENTS.md\n"
    "    requires_modules: [agent-codex]\n"
    '    required_phrases: ["Agents MUST validate.", "Agents MUST preserve authority."]\n'
)
WEAK_CATALOG = (
    "instruction_contracts:\n"
    "  - path: AGENTS.md\n"
    "    requires_modules: [agent-codex]\n"
    '    required_phrases: ["Agents MUST validate."]\n'
)


def write(root: Path, relative: str, text: str) -> Path:
    """Write a fixture file with deterministic encoding."""
    path = root / relative
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")
    return path


def deploy(root: Path) -> None:
    """Copy only the retained standalone runtime, without sync or baseline files."""
    for relative in RUNTIME:
        target = root / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(ROOT / relative, target)
    write(root, ".github/instruction-profile.yml", PROFILE)
    write(root, ".github/instruction-contracts.yml", CATALOG)
    write(root, "AGENTS.md", "Agents MUST validate.\nAgents MUST preserve authority.\n")
    assert not (root / ".template-sync").exists()
    assert not (root / ".pre-commit-config.yaml").exists()
    assert not (root / "pyproject.toml").exists()


def run_profile(root: Path) -> subprocess.CompletedProcess[str]:
    """Run the physically independent validator and preserve its native result."""
    return subprocess.run(
        [sys.executable, str(root / RUNTIME[2]), "--repo-root", str(root)],
        check=False,
        capture_output=True,
        text=True,
    )


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("one: {value: 1}\ntwo: {value: 2}\n", {"one": {"value": 1}, "two": {"value": 2}}),
        ("one: &name value\ntwo: *name\n", {"one": "value", "two": "value"}),
        (
            "base: &base {value: 1}\ncopy: {<<: *base, value: 2}\n",
            {"base": {"value": 1}, "copy": {"value": 2}},
        ),
        (
            "base: &base {value: 1}\ncopy: {value: 2, <<: *base}\n",
            {"base": {"value": 1}, "copy": {"value": 2}},
        ),
        (
            "one: &one {value: 1}\ntwo: &two {value: 2, other: 3}\ncopy: {<<: [*one, *two]}\n",
            {
                "one": {"value": 1},
                "two": {"value": 2, "other": 3},
                "copy": {"value": 1, "other": 3},
            },
        ),
        (
            (
                "base: &base {value: 1}\nnext: &next {<<: *base, value: 2}\n"
                "copy: {<<: *next}\nalias: *next\n"
            ),
            {
                "base": {"value": 1},
                "next": {"value": 2},
                "copy": {"value": 2},
                "alias": {"value": 2},
            },
        ),
        (
            'base: &base {value: 1}\ncopy: {<<: *base, "<<": literal}\n',
            {"base": {"value": 1}, "copy": {"value": 1, "<<": "literal"}},
        ),
        ("=: value\n", {"=": "value"}),
        ("items: [{value: 1}, {value: 2}]\n", {"items": [{"value": 1}, {"value": 2}]}),
    ],
)
def test_yaml_preserves_safe_mapping_semantics(text: str, expected: dict[str, Any]) -> None:
    """Fixed outputs protect merge precedence, aliases and independent mappings."""
    assert support.parse_yaml_mapping(text, "policy.yml") == expected


@pytest.mark.parametrize(
    "text",
    [
        "value: 1\nvalue: 2\n",
        "value: 1\nvalue: 1\n",
        'value: 1\n"value": 2\n',
        "outer:\n  value: 1\n  value: 2\n",
        "{value: 1, value: 2}",
        "items:\n  - value: 1\n    value: 2\n",
        "base: &base {value: 1, value: 2}\ncopy: {<<: *base}\n",
        "copy: {<<: {value: 1, value: 2}}\n",
        "copy: {<<: {one: 1}, <<: {two: 2}}\n",
        "true: first\n1: second\n",
        "0xB: first\n11: second\n",
        "&name value: first\n*name: second\n",
        '=: first\n"=": second\n',
    ],
)
def test_yaml_rejects_duplicate_explicit_keys(text: str) -> None:
    """Duplicates fail before a later key can discard the first value."""
    with pytest.raises(
        support.TemplateSyncMaterializationError, match="duplicate explicit"
    ) as error:
        support.parse_yaml_mapping(text, "policy.yml")
    assert "policy.yml" in str(error.value)
    assert "line " in str(error.value)
    assert isinstance(error.value.__cause__, yaml.YAMLError)


@pytest.mark.parametrize(
    ("text", "message"),
    [
        ("? [one, two]\n: value\n", "unhashable key"),
        ("value: !Custom value\n", "constructor"),
        ("value: !!python/object/apply:builtins.str []\n", "constructor"),
        ("value: [\n", "Invalid YAML"),
        ("value: 1\n---\nvalue: 2\n", "Invalid YAML"),
        ("", "must contain a YAML mapping"),
        ("[]", "must contain a YAML mapping"),
        ("scalar", "must contain a YAML mapping"),
    ],
)
def test_yaml_retains_safe_loader_and_shape_failures(text: str, message: str) -> None:
    """A duplicate check must not enable unsafe tags or bypass mapping validation."""
    with pytest.raises(support.TemplateSyncMaterializationError, match=message):
        support.parse_yaml_mapping(text, "policy.yml")


def test_yaml_loader_isolation_and_input_limit(tmp_path: Path) -> None:
    """Custom parsing leaves global loaders unchanged and retains pre-decode bounds."""
    assert yaml.safe_load("value: 1\nvalue: 2\n") == {"value": 2}
    text = "\ufeffvalue: 1\n"
    path = write(tmp_path, "policy.yml", text)
    length = len(text.encode("utf-8"))
    assert support.load_yaml_mapping(path, tmp_path, maximum_bytes=length) == {"value": 1}
    with pytest.raises(support.TemplateSyncMaterializationError, match="input limit"):
        support.load_yaml_mapping(path, tmp_path, maximum_bytes=length - 1)
    with pytest.raises(support.TemplateSyncMaterializationError, match="Unable to read"):
        support.load_yaml_mapping(tmp_path / "missing.yml", tmp_path)


def test_standalone_positive_and_missing_obligation(tmp_path: Path) -> None:
    """An independent native control proves that the retained obligation matters."""
    deploy(tmp_path)
    valid = run_profile(tmp_path)
    assert valid.returncode == 0, valid.stderr
    assert "Contracts checked: 1" in valid.stdout
    write(tmp_path, "AGENTS.md", "Agents MUST validate.\n")
    invalid = run_profile(tmp_path)
    assert invalid.returncode == 1, invalid.stderr
    assert "missing required phrase: Agents MUST preserve authority." in invalid.stdout


@pytest.mark.parametrize("case", ["modules", "catalog", "nested-phrases"])
def test_standalone_duplicate_guard_and_restoration_mutant(tmp_path: Path, case: str) -> None:
    """Restoring the old loader makes the fixed native failure oracle detect a bypass."""
    deploy(tmp_path)
    write(tmp_path, "AGENTS.md", "Agents MUST validate.\n")
    if case == "modules":
        write(
            tmp_path,
            ".github/instruction-profile.yml",
            PROFILE + "modules: [agent-instructions, instruction-enforcement, github-actions]\n",
        )
    elif case == "catalog":
        write(tmp_path, ".github/instruction-contracts.yml", CATALOG + WEAK_CATALOG)
    else:
        write(
            tmp_path,
            ".github/instruction-contracts.yml",
            CATALOG + '    required_phrases: ["Agents MUST validate."]\n',
        )
    invalid = run_profile(tmp_path)
    assert invalid.returncode == 1, (invalid.stdout, invalid.stderr)
    assert "duplicate explicit mapping key" in invalid.stderr

    path = tmp_path / RUNTIME[1]
    original = path.read_text(encoding="utf-8")
    guarded = "yaml.load(text, Loader=UniqueKeySafeLoader)"
    assert original.count(guarded) == 1
    path.write_text(original.replace(guarded, "yaml.safe_load(text)"), encoding="utf-8")
    mutant = run_profile(tmp_path)
    assert mutant.returncode == 0, (mutant.stdout, mutant.stderr)
    with pytest.raises(AssertionError):
        assert mutant.returncode == 1, "Duplicate policy must fail."


def test_schema_checks_remain_after_yaml_loading(tmp_path: Path) -> None:
    """Unique schema-invalid inputs fail, while malformed YAML fails before the schema read."""
    deploy(tmp_path)
    write(tmp_path, ".github/instruction-profile.yml", PROFILE.replace("version: 1", "version: 2"))
    result = run_profile(tmp_path)
    assert result.returncode == 1
    assert "Schema validation failed" in result.stderr
    path = write(tmp_path, "duplicate.yml", "key: one\nkey: two\n")
    with pytest.raises(support.TemplateSyncMaterializationError, match="duplicate explicit"):
        core.load_schema_validated_yaml(path, tmp_path / "absent-schema.json", tmp_path)


def test_marker_consumer_rejects_duplicate_modules(tmp_path: Path) -> None:
    """The shared marker ingress rejects ambiguity without a YAML lint invocation."""
    text = (
        "template_sync:\n"
        "  source_repo: https://github.com/franklesniak/copilot-repo-template.git\n"
        "  included_modules: [baseline]\n"
    )
    marker = write(tmp_path, ".template-sync/marker.yml", text)
    schema = tmp_path / "schemas/template-sync-marker.schema.json"
    schema.parent.mkdir()
    shutil.copyfile(ROOT / "schemas/template-sync-marker.schema.json", schema)
    document = core.load_schema_validated_yaml(marker, schema, tmp_path)
    assert document["template_sync"]["included_modules"] == ["baseline"]
    marker.write_text(text + "  included_modules: [github-actions]\n", encoding="utf-8")
    result = subprocess.run(
        [
            sys.executable,
            str(ROOT / ".template-sync/scripts/validate_marker.py"),
            "--repo-root",
            str(tmp_path),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 1, (result.stdout, result.stderr)
    assert "duplicate explicit mapping key" in result.stderr


@pytest.mark.parametrize("field", ["included_modules", "included_modules_csv"])
def test_materializer_yaml_duplicates_fail_before_writes(tmp_path: Path, field: str) -> None:
    """Native argument parsing cannot silently discard a selected enforcement module."""
    value = "[baseline]" if field == "included_modules" else "baseline"
    path = write(tmp_path, "operator.yaml", f"{field}: {value}\n{field}: {value}\n")
    target = tmp_path / "uncreated-target"
    result = subprocess.run(
        [
            sys.executable,
            str(ROOT / ".template-sync/scripts/materialize_downstream_adoption.py"),
            "--args-file",
            str(path),
            "--target-root",
            str(target),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 1, (result.stdout, result.stderr)
    assert "--args-file" in result.stderr
    assert "duplicate explicit mapping key" in result.stderr
    assert not target.exists()


def test_materializer_yaml_args_preserve_bom_paths_and_cli_precedence(tmp_path: Path) -> None:
    """External operator files keep their read contract and family-level CLI override."""
    path = write(tmp_path, "operator.yaml", "\ufeffincluded_modules: [baseline]\n")
    assert materializer.load_yaml_args_file(path) == {"included_modules": ["baseline"]}
    parsed = materializer.parse_args(
        ["--args-file", str(path), "--included-module", "github-actions"]
    )
    assert parsed.included_modules == ["github-actions"]


@pytest.mark.parametrize(
    "text",
    [
        '{"value":1,"value":2}',
        '{"outer":{"value":1,"value":2}}',
        '{"items":[{"value":1,"value":2}]}',
        '{"value":1,"v\\u0061lue":2}',
    ],
)
def test_json_rejects_duplicate_decoded_names(text: str) -> None:
    """JSON object pairs are checked before a later name overwrites the first."""
    with pytest.raises(support.TemplateSyncMaterializationError, match="duplicate object key"):
        support.parse_json_mapping(text, "schema.json")


@pytest.mark.parametrize(
    ("text", "message"),
    [
        ('{"value":1,}', "Invalid JSON"),
        ("value: 1", "Invalid JSON"),
        ('{"value": /* comment */ 1}', "Invalid JSON"),
        ("[]", "must contain a JSON object"),
        ("null", "must contain a JSON object"),
    ],
)
def test_json_preserves_syntax_and_object_errors(text: str, message: str) -> None:
    """The YAML adapter must not broaden the strict JSON input dialect."""
    with pytest.raises(support.TemplateSyncMaterializationError, match=message):
        support.parse_json_mapping(text, "schema.json")


def test_json_valid_keys_order_and_read_bound(tmp_path: Path) -> None:
    """Independent objects, escaped names and byte limits retain their meaning."""
    text = '{"second":{"value":2},"first":{"value":1},"\\u0061":true}'
    expected = {"second": {"value": 2}, "first": {"value": 1}, "a": True}
    assert support.parse_json_mapping(text, "schema.json") == expected
    assert list(support.parse_json_mapping(text, "schema.json")) == ["second", "first", "a"]
    path = write(tmp_path, "schema.json", text)
    assert support.load_json_mapping(path, tmp_path, maximum_bytes=len(text)) == expected
    with pytest.raises(support.TemplateSyncMaterializationError, match="input limit"):
        support.load_json_mapping(path, tmp_path, maximum_bytes=len(text) - 1)


def test_standalone_duplicate_schema_guard_and_restoration_mutant(tmp_path: Path) -> None:
    """The real standalone schema path cannot silently discard required constraints."""
    deploy(tmp_path)
    schema = write(
        tmp_path,
        "schemas/instruction-contracts.schema.json",
        '{"type":"object","required":["mandatory"]}',
    )
    missing = run_profile(tmp_path)
    assert missing.returncode == 1
    assert "Schema validation failed" in missing.stderr
    schema.write_text('{"type":"object","required":["mandatory"],"required":[]}', encoding="utf-8")
    duplicate = run_profile(tmp_path)
    assert duplicate.returncode == 1
    assert "duplicate object key 'required'" in duplicate.stderr
    path = tmp_path / RUNTIME[1]
    original = path.read_text(encoding="utf-8")
    guarded = "json.loads(text, object_pairs_hook=unique_object)"
    assert original.count(guarded) == 1
    path.write_text(original.replace(guarded, "json.loads(text)"), encoding="utf-8")
    mutant = run_profile(tmp_path)
    assert mutant.returncode == 0, (mutant.stdout, mutant.stderr)
    with pytest.raises(AssertionError):
        assert mutant.returncode == 1, "Duplicate schema constraints must fail."


@pytest.mark.parametrize("field", ["included_modules", "included_modules_csv"])
def test_materializer_json_duplicates_fail_before_writes(tmp_path: Path, field: str) -> None:
    """The JSON argument route must reject the same ambiguous module selection."""
    value = '["baseline"]' if field == "included_modules" else '"baseline"'
    path = write(tmp_path, "operator.json", f'{{"{field}":{value},"{field}":{value}}}')
    target = tmp_path / "uncreated-target"
    result = subprocess.run(
        [
            sys.executable,
            str(ROOT / ".template-sync/scripts/materialize_downstream_adoption.py"),
            "--args-file",
            str(path),
            "--target-root",
            str(target),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 1, (result.stdout, result.stderr)
    assert "--args-file" in result.stderr
    assert "duplicate object key" in result.stderr
    assert not target.exists()


def test_materializer_json_args_preserve_bom_paths_and_cli_precedence(tmp_path: Path) -> None:
    """JSON operator files keep their read contract and family-level CLI override."""
    path = write(tmp_path, "operator.json", '\ufeff{"included_modules":["baseline"]}')
    assert materializer.load_json_args_file(path) == {"included_modules": ["baseline"]}
    parsed = materializer.parse_args(
        ["--args-file", str(path), "--included-module", "github-actions"]
    )
    assert parsed.included_modules == ["github-actions"]


@pytest.mark.parametrize("file_format", ["yaml", "json"])
def test_materializer_args_loader_restoration_mutant(tmp_path: Path, file_format: str) -> None:
    """An independent selected-module oracle detects either shared decoder's removal."""
    scripts = tmp_path / ".template-sync/scripts"
    shutil.copytree(
        ROOT / ".template-sync/scripts", scripts, ignore=shutil.ignore_patterns("__pycache__")
    )
    for relative in RUNTIME[:2]:
        write(tmp_path, relative, (ROOT / relative).read_text(encoding="utf-8"))
    retained = "[agent-instructions, instruction-enforcement, github-actions, agent-codex]"
    omitted = "[agent-instructions, github-actions]"
    if file_format == "yaml":
        text = f"included_modules: {retained}\nincluded_modules: {omitted}\n"
        guarded = "yaml.load(text, Loader=UniqueKeySafeLoader)"
        restored = "yaml.safe_load(text)"
    else:
        text = (
            '{"included_modules":["agent-instructions","instruction-enforcement",'
            '"github-actions","agent-codex"],'
            '"included_modules":["agent-instructions","github-actions"]}'
        )
        guarded = "json.loads(text, object_pairs_hook=unique_object)"
        restored = "json.loads(text)"
    path = write(tmp_path, f"operator.{file_format}", text)
    command = [
        sys.executable,
        "-c",
        (
            "import json,sys; sys.path.insert(0,sys.argv[1]); "
            "import materialize_downstream_adoption as m; "
            "print(json.dumps(m.parse_args(['--args-file',sys.argv[2]]).included_modules))"
        ),
        str(scripts),
        str(path),
    ]
    rejected = subprocess.run(command, capture_output=True, text=True, check=False)
    assert rejected.returncode == 1
    assert "duplicate" in rejected.stderr
    support_path = tmp_path / RUNTIME[1]
    original = support_path.read_text(encoding="utf-8")
    assert original.count(guarded) == 1
    support_path.write_text(original.replace(guarded, restored), encoding="utf-8")
    mutant = subprocess.run(command, capture_output=True, text=True, check=False)
    assert mutant.returncode == 0, mutant.stderr
    assert json.loads(mutant.stdout) == ["agent-instructions", "github-actions"]
    with pytest.raises(AssertionError):
        assert mutant.returncode == 1, "Ambiguous applicability must fail."


@pytest.mark.parametrize("kind", ["json", "yaml"])
def test_caller_resolved_external_mapping_preserves_loading_and_guards(
    tmp_path: Path, kind: str
) -> None:
    """External shared schemas still load, while duplicate rejection remains active."""
    marker_root = tmp_path / "marker-repository"
    marker_root.mkdir()
    source = tmp_path / f"external-schema.{kind}"
    loader = support.load_json_mapping if kind == "json" else support.load_yaml_mapping
    valid = '{"required": ["mandatory"]}' if kind == "json" else "required: [mandatory]\n"
    source.write_text(valid, encoding="utf-8")
    assert loader(source, marker_root) == {"required": ["mandatory"]}
    duplicate = (
        '{"required": ["mandatory"], "required": []}'
        if kind == "json"
        else "required: [mandatory]\nrequired: []\n"
    )
    source.write_text(duplicate, encoding="utf-8")
    with pytest.raises(support.TemplateSyncMaterializationError, match="duplicate") as caught:
        loader(source, marker_root)
    assert source.name in str(caught.value)
    assert str(tmp_path) not in str(caught.value)
