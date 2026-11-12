"""Exercise each workflow invariant and independent failure-injection oracles."""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import re
import shlex
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any, cast

import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest

pytestmark = pytest.mark.upstream_template_only
ROOT = Path(__file__).resolve().parents[1]
OWNED_WORKFLOW_COUNT = len(list((ROOT / ".github/workflows").glob("*.yml")))
SPEC = importlib.util.spec_from_file_location(
    "workflow_security_tests", ROOT / ".github/scripts/validate_workflow_security.py"
)
assert SPEC is not None and SPEC.loader is not None
policy = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(policy)


def copy_policy(root: Path) -> dict[str, Any]:
    """Copy the real policy surface, including copyable documentation examples."""
    contract = policy.load_contract(ROOT)
    for relative in [
        policy.CONTRACT,
        policy.SCHEMA,
        ".github/scripts/instruction_contract_support.py",
        *contract["workflows"],
        *contract["examples"],
    ]:
        target = root / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(ROOT / relative, target)
    return cast(dict[str, Any], contract)


@pytest.mark.parametrize(
    ("relative", "safe"),
    [
        ("README.md", True),
        ("docs/guide.md", True),
        (".github/instructions/yaml.instructions.md", True),
        ("docs/.hidden/guide.md", True),
        ("docs/a-b_2.md", True),
        (".cursor/rules/example.mdc", True),
        (".cursor/rules/nested/a-b_2.mdc", True),
        ("/tmp/example.md", False),
        ("./guide.md", False),
        ("docs//guide.md", False),
        ("docs/./guide.md", False),
        ("docs/../guide.md", False),
        ("C:/guide.md", False),
        ("docs\\guide.md", False),
        ("/tmp/example.mdc", False),
        ("./guide.mdc", False),
        (".cursor//guide.mdc", False),
        (".cursor/./guide.mdc", False),
        (".cursor/../guide.mdc", False),
        ("C:/guide.mdc", False),
        (".cursor\\guide.mdc", False),
    ],
)
def test_example_path_schema_matches_reader(tmp_path: Path, relative: str, safe: bool) -> None:
    """Schema path components agree with the reader on independently labeled boundaries."""
    schema = policy.parse_yaml(policy.read_text(ROOT, policy.SCHEMA))
    contract = copy.deepcopy(policy.load_contract(ROOT))
    contract["examples"] = [relative]
    if safe:
        target = tmp_path / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("example\n", encoding="utf-8")
        policy.jsonschema.validate(contract, schema)
        assert policy.read_text(tmp_path, relative) == "example\n"
    else:
        with pytest.raises(policy.jsonschema.ValidationError):
            policy.jsonschema.validate(contract, schema)
        with pytest.raises(policy.PolicyError, match="Unsafe repository path"):
            policy.read_text(tmp_path, relative)


@pytest.mark.parametrize(
    "relative", ["/tmp/example.md", "./guide.md", "docs//guide.md", "docs/./guide.md"]
)
def test_example_path_pattern_regression(relative: str) -> None:
    """Restoring the previous pattern admits fixed unsafe cases rejected by the current schema."""
    schema = policy.parse_yaml(policy.read_text(ROOT, policy.SCHEMA))
    contract = copy.deepcopy(policy.load_contract(ROOT))
    contract["examples"] = [relative]
    with pytest.raises(policy.jsonschema.ValidationError):
        policy.jsonschema.validate(contract, schema)
    schema["properties"]["examples"]["items"]["pattern"] = r"^(?!.*\.\.)[A-Za-z0-9_./-]+\.md$"
    policy.jsonschema.validate(contract, schema)


@pytest.mark.parametrize("relative", ["guide.mdx", "guide.MDC", "guide.mdcc", "guide.mdc.yml"])
def test_example_path_schema_rejects_unsupported_suffixes(relative: str) -> None:
    """Only the two declared Markdown formats enter the explicit contract."""
    schema = policy.parse_yaml(policy.read_text(ROOT, policy.SCHEMA))
    contract = copy.deepcopy(policy.load_contract(ROOT))
    contract["examples"] = [relative]
    with pytest.raises(policy.jsonschema.ValidationError):
        policy.jsonschema.validate(contract, schema)


def test_cursor_example_schema_old_suffix_mutant() -> None:
    """The old suffix predicate breaks the fixed safe Cursor declaration oracle."""
    schema = policy.parse_yaml(policy.read_text(ROOT, policy.SCHEMA))
    contract = copy.deepcopy(policy.load_contract(ROOT))
    contract["examples"] = [".cursor/rules/example.mdc"]
    policy.jsonschema.validate(contract, schema)
    schema["properties"]["examples"]["items"][
        "pattern"
    ] = r"^(?!.*\.\.)(?:(?!\./)[A-Za-z0-9_.-]+/)*[A-Za-z0-9_.-]+\.md$"
    with pytest.raises(policy.jsonschema.ValidationError):
        policy.jsonschema.validate(contract, schema)


@pytest.mark.parametrize(
    ("text", "count"),
    [
        ("Ordinary documentation without action examples.", 0),
        ("```yaml\n- uses: owner/action@" + "a" * 40 + " # v1.2.3\n```", 1),
        ('- {"uses": "owner/action@' + "a" * 40 + '"} # v1.2.3', 1),
        ("# - uses: owner/action@" + "a" * 40 + " # v1.2.3", 1),
        (
            "- uses: owner/action@"
            + "a" * 40
            + " # v1.2.3\n"
            + "- uses: owner/other@"
            + "b" * 40
            + " # v2.3.4",
            2,
        ),
    ],
)
def test_example_reference_count(text: str, count: int) -> None:
    """Discovery counts semantic references once across fences, fragments and comments."""
    assert policy.check_examples(text) == count


@pytest.mark.parametrize("path", sorted(policy.load_contract(ROOT)["workflows"]))
def test_each_owned_workflow_passes(path: str) -> None:
    """Every declared workflow passes universal and required-control validation."""
    expected = policy.load_contract(ROOT)["workflows"][path]
    assert policy.validate_workflow(policy.read_text(ROOT, path)) == expected


def test_standalone_required_check_names_are_distinct() -> None:
    """Fixed job identities remain independently selectable across owned workflows."""
    expected = {
        ".github/workflows/instruction-contracts.yml": "Instruction Contracts",
        ".github/workflows/workflow-security.yml": "Workflow Security",
    }
    identities: dict[str, list[tuple[str, str]]] = {}
    for path in sorted((ROOT / ".github/workflows").glob("*.yml")):
        relative = path.relative_to(ROOT).as_posix()
        document = policy.parse_yaml(path.read_text(encoding="utf-8"))
        for key, job in document["jobs"].items():
            identities.setdefault(job.get("name", key), []).append((relative, key))
        if relative in expected:
            assert document["jobs"]["validate"]["name"] == expected[relative]
    for relative, name in expected.items():
        assert identities[name] == [(relative, "validate")]


@pytest.mark.parametrize(
    ("relative", "name", "other_name"),
    [
        ("instruction-contracts.yml", "Instruction Contracts", "Workflow Security"),
        ("workflow-security.yml", "Workflow Security", "Instruction Contracts"),
    ],
)
@pytest.mark.parametrize("mutation", ["remove", "duplicate"])
def test_standalone_check_name_drift_fails_natively(
    tmp_path: Path, relative: str, name: str, other_name: str, mutation: str
) -> None:
    """Removing or colliding a name fails against the unchanged reviewed contract."""
    copy_policy(tmp_path)
    path = tmp_path / ".github/workflows" / relative
    text = path.read_text(encoding="utf-8")
    anchor = f"    name: {name}\n"
    assert text.count(anchor) == 1
    command = [
        sys.executable,
        "-B",
        str(ROOT / ".github/scripts/validate_workflow_security.py"),
        "--repo-root",
        str(tmp_path),
    ]
    accepted = subprocess.run(command, check=False, capture_output=True, text=True, timeout=30)
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    replacement = "" if mutation == "remove" else f"    name: {other_name}\n"
    path.write_text(text.replace(anchor, replacement), encoding="utf-8")
    rejected = subprocess.run(command, check=False, capture_output=True, text=True, timeout=30)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "execution controls" in rejected.stderr
    assert "Traceback" not in rejected.stderr


@pytest.mark.parametrize(
    ("before", "after", "message"),
    [
        (
            "{checkout}",
            "actions/checkout@v7",
            "full SHA",
        ),
        ("{sha}", "abcdef0", "full SHA"),
        ("{annotation}", "", "annotation"),
        ("{annotation}", " # v7", "annotation"),
        ("contents: read", "contents: write", "permissions"),
        ("persist-credentials: false", "persist-credentials: true", "persist-credentials"),
        ("  pull_request:", "  pull_request_target:", "privileged"),
        ("  pull_request:", "  workflow_run:", "privileged"),
        ("name: Workflow Security", "name: first\nname: second", "Duplicate"),
        ("name: Workflow Security", "name: [", "Invalid YAML"),
        ("name: Workflow Security", "name: &name Workflow Security", "anchors"),
        (
            "{checkout}",
            "./.github/actions/local",
            "external",
        ),
        (
            "{checkout}",
            "docker://image:latest",
            "external",
        ),
    ],
)
def test_universal_rule_mutations(before: str, after: str, message: str) -> None:
    """Each isolated unsafe input must fail before contract comparison."""
    source = policy.read_text(ROOT, ".github/workflows/workflow-security.yml")
    checkout_line = next(line for line in source.splitlines() if "uses: actions/checkout@" in line)
    reference = checkout_line.strip().split()[1]
    before = (
        before.replace("{checkout}", reference)
        .replace("{sha}", reference.split("@", 1)[1])
        .replace("{annotation}", " # " + checkout_line.split("#", 1)[1].strip())
    )
    assert before in source
    with pytest.raises(policy.PolicyError, match=message):
        policy.validate_workflow(source.replace(before, after, 1))


def test_job_permission_and_inline_reference_bypass() -> None:
    """Missing job permissions and flow-style unannotated uses are rejected."""
    text = policy.read_text(ROOT, ".github/workflows/workflow-security.yml")
    for source in [
        text.replace("    permissions:\n      contents: read\n", ""),
        text.replace(
            "    permissions:\n      contents: read\n", "    permissions:\n      contents: write\n"
        ),
        text.replace(
            "        uses: actions/checkout@", "        uses: >-\n          actions/checkout@"
        ),
    ]:
        with pytest.raises(policy.PolicyError):
            policy.validate_workflow(source)


@pytest.mark.parametrize("mutation", ["delete", "advisory", "skip", "mask", "aggregate"])
def test_required_execution_controls(tmp_path: Path, mutation: str) -> None:
    """Deleted gates, advisory conditions, and masked aggregate failures fail closed."""
    copy_policy(tmp_path)
    path = tmp_path / ".github/workflows/markdownlint.yml"
    text = path.read_text(encoding="utf-8")
    if mutation == "delete":
        text = text.replace("      - name: Check validation results", "      - name: Deleted check")
    elif mutation == "advisory":
        text = text.replace(
            "      - name: Check validation results",
            "      - name: Check validation results\n        continue-on-error: true",
        )
    elif mutation == "skip":
        text = text.replace("steps.lint-outer.outcome == 'failure'", "false")
    elif mutation == "mask":
        text = text.replace("run: npm run lint:md\n", "run: npm run lint:md || true\n")
    else:
        text = text.replace("          exit 1", "          exit 0")
    path.write_text(text, encoding="utf-8")
    with pytest.raises(policy.PolicyError, match="execution controls"):
        policy.validate_repository(tmp_path)


@pytest.mark.parametrize("scope", ["workflow", "job"])
def test_required_permissions_drift(tmp_path: Path, scope: str) -> None:
    """Removing reviewed contents access is drift even though empty permissions are safe."""
    copy_policy(tmp_path)
    path = tmp_path / ".github/workflows/workflow-security.yml"
    text = path.read_text(encoding="utf-8")
    before, after = (
        ("permissions:\n  contents: read\n", "permissions: {}\n")
        if scope == "workflow"
        else ("    permissions:\n      contents: read\n", "    permissions: {}\n")
    )
    assert before in text
    changed = text.replace(before, after, 1)
    policy.validate_workflow(changed)
    path.write_text(changed, encoding="utf-8")
    with pytest.raises(policy.PolicyError, match="execution controls"):
        policy.validate_repository(tmp_path)


@pytest.mark.parametrize("changed", ["  print(1)\n", "print(1)  \n", "print(1)\n\n"])
def test_run_fingerprint_preserves_whitespace(changed: str) -> None:
    """Preserve shell-significant scalar whitespace while allowing CRLF transport."""
    document = policy.parse_yaml(
        "on: push\npermissions: {}\njobs:\n  check:\n    permissions: {}\n"
        "    steps:\n      - shell: python\n        run: |\n          print(1)\n"
    )
    original = policy.describe_workflow(document)
    step = document["jobs"]["check"]["steps"][0]
    step["run"] = "print(1)\r\n"
    assert policy.describe_workflow(document) == original
    step["run"] = changed
    assert policy.describe_workflow(document) != original


@pytest.mark.parametrize(
    "guard",
    [
        "workflow_permissions",
        "job_permissions",
        "job_name",
        "workflow_concurrency",
        "job_concurrency",
        "run_whitespace",
    ],
)
def test_descriptor_guard_removal(tmp_path: Path, guard: str) -> None:
    """Independent inputs expose restored permission omission or whitespace stripping."""
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    before, after = {
        "workflow_permissions": ('        "permissions": document["permissions"],\n', ""),
        "job_permissions": ('                    "permissions",\n', ""),
        "job_name": ('                    "name",\n', ""),
        "workflow_concurrency": (
            '        **{field: document[field] for field in ("concurrency",) if field in document},\n',
            "",
        ),
        "job_concurrency": ('                    "concurrency",\n', ""),
        "run_whitespace": (
            'step["run"].replace("\\r\\n", "\\n").encode("utf-8")',
            'step["run"].replace("\\r\\n", "\\n").strip().encode("utf-8")',
        ),
    }[guard]
    assert source.count(before) == 1
    mutant_path = tmp_path / "descriptor_mutant.py"
    mutant_path.write_text(source.replace(before, after), encoding="utf-8")
    spec = importlib.util.spec_from_file_location("descriptor_mutant", mutant_path)
    assert spec is not None and spec.loader is not None
    mutant = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mutant)
    original = policy.parse_yaml(
        "on: push\npermissions: {contents: read}\njobs:\n  check:\n"
        "    permissions: {contents: read}\n    steps:\n      - run: print(1)\n"
    )
    changed = copy.deepcopy(original)
    if guard.endswith("concurrency"):
        original_scope = original if guard.startswith("workflow") else original["jobs"]["check"]
        changed_scope = changed if guard.startswith("workflow") else changed["jobs"]["check"]
        original_scope["concurrency"] = {"group": "required", "cancel-in-progress": False}
        changed_scope["concurrency"] = {"group": "required", "cancel-in-progress": True}
    elif guard == "job_name":
        original["jobs"]["check"]["name"] = "Required check"
        changed["jobs"]["check"]["name"] = "Different check"
    elif guard == "workflow_permissions":
        changed["permissions"] = {}
    elif guard == "job_permissions":
        changed["jobs"]["check"]["permissions"] = {}
    else:
        changed["jobs"]["check"]["steps"][0]["run"] = "  print(1)"
    assert policy.describe_workflow(original) != policy.describe_workflow(changed)
    assert mutant.describe_workflow(original) == mutant.describe_workflow(changed)


@pytest.mark.parametrize("scope", ["workflow", "job"])
@pytest.mark.parametrize("change", ["add", "delete", "group", "cancel", "string"])
def test_concurrency_controls_require_review(tmp_path: Path, scope: str, change: str) -> None:
    """Both concurrency scopes preserve reviewed scheduling and detect every drift form."""
    contract = copy_policy(tmp_path)
    path = ".github/workflows/workflow-security.yml"
    document = policy.parse_yaml(policy.read_text(tmp_path, path))
    target = document if scope == "workflow" else document["jobs"]["validate"]
    if change != "add":
        target["concurrency"] = {"group": "required", "cancel-in-progress": False}
    # Review only the isolated fixture descriptor; do not rewrite executable annotations.
    original = copy.deepcopy(policy.describe_workflow(document))
    contract["workflows"][path] = original
    (tmp_path / policy.CONTRACT).write_text(yaml.safe_dump(contract), encoding="utf-8")
    policy.load_contract(tmp_path)
    if change == "delete":
        del target["concurrency"]
    elif change == "group":
        target["concurrency"]["group"] = "other-workflow"
    elif change == "cancel":
        target["concurrency"]["cancel-in-progress"] = True
    elif change == "string":
        target["concurrency"] = "shared-group"
    else:
        target["concurrency"] = {"group": "shared-group", "cancel-in-progress": True}
    assert policy.describe_workflow(document) != original
    contract["workflows"][path] = policy.describe_workflow(document)
    (tmp_path / policy.CONTRACT).write_text(yaml.safe_dump(contract), encoding="utf-8")
    policy.load_contract(tmp_path)
    target["concurrency"] = 42
    contract["workflows"][path] = policy.describe_workflow(document)
    (tmp_path / policy.CONTRACT).write_text(yaml.safe_dump(contract), encoding="utf-8")
    with pytest.raises(policy.jsonschema.ValidationError):
        policy.load_contract(tmp_path)


def test_existing_concurrency_cancellation_change_is_rejected(tmp_path: Path) -> None:
    """The current auto-fix control and absent controls retain their distinct semantics."""
    copy_policy(tmp_path)
    path = tmp_path / ".github/workflows/auto-fix-precommit.yml"
    text = path.read_text(encoding="utf-8")
    expected = {"group": "auto-fix-precommit-${{ github.ref }}", "cancel-in-progress": False}
    assert policy.validate_workflow(text)["concurrency"] == expected
    assert policy.validate_repository(tmp_path) == OWNED_WORKFLOW_COUNT
    path.write_text(
        text.replace("cancel-in-progress: false", "cancel-in-progress: true"), encoding="utf-8"
    )
    with pytest.raises(policy.PolicyError, match="execution controls"):
        policy.validate_repository(tmp_path)
    other = policy.validate_workflow(
        policy.read_text(ROOT, ".github/workflows/workflow-security.yml")
    )
    assert "concurrency" not in other
    assert "concurrency" not in other["jobs"]["validate"]["controls"]


@pytest.mark.parametrize("change", ["rename", "delete", "add", "matrix", "fallback"])
def test_job_check_identity_drift(tmp_path: Path, change: str) -> None:
    """Explicit names and fallback IDs cannot silently change an owned check identity."""
    contract = copy_policy(tmp_path)
    filename, before, after = {
        "rename": ("precommit-ci", "    name: Pre-commit\n", "    name: Renamed check\n"),
        "delete": ("precommit-ci", "    name: Pre-commit\n", ""),
        "add": ("markdownlint", "  markdownlint:\n", "  markdownlint:\n    name: New check\n"),
        "matrix": ("python-ci", "    name: Test\n", "    name: Changed test\n"),
        "fallback": ("markdownlint", "  markdownlint:\n", "  different-id:\n"),
    }[change]
    workflow_path = f".github/workflows/{filename}.yml"
    path = tmp_path / workflow_path
    text = path.read_text(encoding="utf-8")
    if change in {"add", "fallback"}:
        assert "name" not in policy.parse_yaml(text)["jobs"]["markdownlint"]
        assert (
            "name" not in contract["workflows"][workflow_path]["jobs"]["markdownlint"]["controls"]
        )
    assert policy.validate_repository(tmp_path) == OWNED_WORKFLOW_COUNT
    assert text.count(before) == 1
    modified = text.replace(before, after)
    policy.validate_workflow(modified)
    path.write_text(modified, encoding="utf-8")
    with pytest.raises(policy.PolicyError, match="execution controls"):
        policy.validate_repository(tmp_path)


def test_workflow_display_label_is_not_a_required_check_identity(tmp_path: Path) -> None:
    """Changing only the Actions-tab label preserves the reviewed job controls."""
    copy_policy(tmp_path)
    path = tmp_path / ".github/workflows/precommit-ci.yml"
    text = path.read_text(encoding="utf-8")
    assert text.count("name: Pre-commit CI") == 1
    path.write_text(text.replace("name: Pre-commit CI", "name: Renamed workflow"), encoding="utf-8")
    assert policy.validate_repository(tmp_path) == OWNED_WORKFLOW_COUNT


@pytest.mark.parametrize("coordinate", ["actions/checkout", "Actions/Checkout", "ACTIONS/CHECKOUT"])
@pytest.mark.parametrize("credentials", ["false", "true", "'false'", "missing", "empty"])
def test_checkout_identity_is_case_insensitive(coordinate: str, credentials: str) -> None:
    """Every spelling of the checkout repository requires literal false credentials."""
    text = policy.read_text(ROOT, ".github/workflows/workflow-security.yml")
    text = text.replace("actions/checkout@", coordinate + "@")
    if credentials == "missing":
        text = text.replace("        with:\n          persist-credentials: false\n", "")
    elif credentials == "empty":
        text = text.replace(
            "        with:\n          persist-credentials: false\n", "        with: {}\n"
        )
    else:
        text = text.replace("persist-credentials: false", "persist-credentials: " + credentials)
    if credentials == "false":
        assert policy.validate_workflow(text)
    else:
        with pytest.raises(policy.PolicyError, match="persist-credentials"):
            policy.validate_workflow(text)


@pytest.mark.parametrize(
    "coordinate", ["actions/checkout-extra", "actions/checkout/path", "other/checkout"]
)
def test_checkout_near_matches_do_not_inherit_its_inputs(coordinate: str) -> None:
    """Different external actions remain outside checkout's specific input contract."""
    text = policy.read_text(ROOT, ".github/workflows/workflow-security.yml")
    text = text.replace("actions/checkout@", coordinate + "@")
    text = text.replace("        with:\n          persist-credentials: false\n", "")
    assert policy.validate_workflow(text)


def test_strict_adopter_checkout_casing_and_guard_mutant(tmp_path: Path) -> None:
    """Strict scope catches the bypass and a case-sensitive mutant restores it."""
    copy_policy(tmp_path)
    bad = (
        policy.read_text(ROOT, ".github/workflows/workflow-security.yml")
        .replace("actions/checkout@", "Actions/Checkout@")
        .replace("persist-credentials: false", "persist-credentials: true")
    )
    (tmp_path / ".github/workflows/adopter.yml").write_text(bad, encoding="utf-8")
    assert policy.validate_repository(tmp_path) == OWNED_WORKFLOW_COUNT
    with pytest.raises(policy.PolicyError, match="persist-credentials"):
        policy.validate_repository(tmp_path, strict=True)
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    guard = 'reference.split("@", 1)[0].casefold()'
    assert source.count(guard) == 1
    mutant = load_policy_mutant(tmp_path, source.replace(guard, 'reference.split("@", 1)[0]'))
    assert mutant.validate_repository(tmp_path, strict=True) == OWNED_WORKFLOW_COUNT


def load_policy_mutant(root: Path, source: str) -> Any:
    """Load an isolated deliberately weakened validator for an independent input oracle."""
    shutil.copyfile(
        ROOT / ".github/scripts/instruction_contract_support.py",
        root / "instruction_contract_support.py",
    )
    path = root / "policy_mutant.py"
    path.write_text(source, encoding="utf-8")
    spec = importlib.util.spec_from_file_location("policy_mutant", path)
    assert spec is not None and spec.loader is not None
    mutant = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mutant)
    return mutant


@pytest.mark.parametrize(
    "prefix", ["> - ", "   > # ", ">> - ", "> # - ", "  > > - ", "- > - ", "> 1. > # - "]
)
@pytest.mark.parametrize("reference", ["pinned", "tag", "short", "missing", "malformed"])
def test_quoted_document_action_policy(tmp_path: Path, prefix: str, reference: str) -> None:
    """Copyable quoted actions use the same pin and annotation rules as plain examples."""
    contract = copy_policy(tmp_path)
    line = {
        "pinned": "owner/action@" + "a" * 40 + " # v1.2.3",
        "tag": "owner/action@v1 # v1.2.3",
        "short": "owner/action@abcdef0 # v1.2.3",
        "missing": "owner/action@" + "a" * 40,
        "malformed": "owner/action@" + "a" * 40 + " # v1",
    }[reference]
    path = tmp_path / contract["examples"][0]
    with path.open("a", encoding="utf-8") as stream:
        stream.write("\n> [!NOTE]\n>\n> ```yaml\n" + prefix + "uses: " + line + "\n> ```\n")
    if reference == "pinned":
        assert policy.validate_repository(tmp_path) == OWNED_WORKFLOW_COUNT
    else:
        with pytest.raises(policy.PolicyError, match="full SHA"):
            policy.validate_repository(tmp_path)


@pytest.mark.parametrize(
    "line", [r"\> - uses: owner/action@v1", "prose > - uses: owner/action@v1", "> [!NOTE]"]
)
def test_markdown_prefix_near_matches(line: str) -> None:
    """Escaped or prose quote markers do not manufacture action lines."""
    assert policy.USES_LINE.match(policy.markdown_example_content(line)) is None


@pytest.mark.parametrize("mutation", ["bypass", "one-level"])
def test_quoted_document_guard_removal(tmp_path: Path, mutation: str) -> None:
    """A fixed nested-quote counterexample exposes either removed container guard."""
    contract = copy_policy(tmp_path)
    path = tmp_path / contract["examples"][0]
    with path.open("a", encoding="utf-8") as stream:
        stream.write("\n>> - uses: owner/action@v1\n")
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.validate_repository(tmp_path)
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    before, after = (
        (
            "lines = [markdown_example_content(line) for line in raw_lines]",
            "lines = raw_lines.copy()",
        )
        if mutation == "bypass"
        else (
            "while (match := MARKDOWN_QUOTE_PREFIX.match(line, offset))",
            "if (match := MARKDOWN_QUOTE_PREFIX.match(line, offset))",
        )
    )
    assert source.count(before) == 1
    mutant = load_policy_mutant(tmp_path, source.replace(before, after))
    assert mutant.validate_repository(tmp_path) == OWNED_WORKFLOW_COUNT


def test_quoted_example_release_resolution() -> None:
    """Peeling preserves the exact repository, release annotation, and digest oracle."""
    line = policy.markdown_example_content("> 1. > - uses: owner/action@" + "a" * 40 + " # v1.2.3")
    reference = policy.USES_LINE.match(line)[1]
    calls = []

    def resolve(repository: str, release: str) -> str:
        """Record exact inputs without a network request."""
        calls.append((repository, release))
        return "a" * 40

    policy.check_reference(reference, line, resolve)
    assert calls == [("owner/action", "v1.2.3")]
    with pytest.raises(policy.PolicyError, match="Misleading"):
        policy.check_reference(reference, line, lambda _repository, _release: "b" * 40)


def test_misleading_annotation_uses_upstream_resolution() -> None:
    """A valid-looking release comment must match the independently resolved commit."""
    ref = "owner/action@" + "a" * 40
    policy.check_reference(ref, ref + " # v1.2.3", lambda _repo, _tag: "a" * 40)
    with pytest.raises(policy.PolicyError, match="Misleading"):
        policy.check_reference(ref, ref + " # v9.9.9", lambda _repo, _tag: "b" * 40)


def test_reusable_and_subpath_references() -> None:
    """The external grammar includes reusable workflows and action subdirectories."""
    for ref in ["owner/repo/.github/workflows/build.yml@", "owner/repo/action/path@"]:
        ref += "a" * 40
        policy.check_reference(ref, ref + " # v1.0.0")


@pytest.mark.parametrize("path", ["../escape.yml", "/outside", "C:/outside", "a/../b", "a\\b"])
def test_unsafe_input_paths(tmp_path: Path, path: str) -> None:
    """Reject lexical escapes independently of existence."""
    with pytest.raises(policy.PolicyError, match="Unsafe"):
        policy.read_text(tmp_path, path)


def test_bounds_and_yaml_shape(tmp_path: Path) -> None:
    """Input size, nesting, multi-document YAML, and custom tags fail safely."""
    path = tmp_path / "large.yml"
    path.write_bytes(b"x" * (policy.LIMIT + 1))
    with pytest.raises(policy.PolicyError, match="1 MiB"):
        policy.read_text(tmp_path, path.name)
    for text in [
        "x" * (policy.LIMIT + 1),
        "x: " + "[" * 65 + "]" * 65,
        "x: 1\n---\nx: 2",
        "x: !!python/object:builtins.object {}",
    ]:
        with pytest.raises(policy.PolicyError):
            policy.parse_yaml(text)
    assert "on" in policy.parse_yaml("on: [push]")


@pytest.mark.parametrize("suffix", [".yml", ".md", ".mdc"])
def test_symlink_is_rejected(tmp_path: Path, suffix: str) -> None:
    """A symlink must not provide a policy input even inside the repository."""
    source = tmp_path / f"real{suffix}"
    source.write_text("name: test", encoding="utf-8")
    link = tmp_path / f"link{suffix}"
    try:
        link.symlink_to(source)
    except OSError:
        pytest.skip("Host does not permit symlink creation")
    with pytest.raises(policy.PolicyError, match="Symlink"):
        policy.read_text(tmp_path, link.name)


def test_owned_scope_and_strict_opt_in(tmp_path: Path) -> None:
    """Adopter workflows are untouched and ignored until strict checking is requested."""
    copy_policy(tmp_path)
    local = tmp_path / ".github/workflows/adopter.yml"
    local.write_text("on: pull_request_target\njobs: {}\n", encoding="utf-8")
    before = local.read_bytes()
    assert policy.validate_repository(tmp_path) == OWNED_WORKFLOW_COUNT
    with pytest.raises(policy.PolicyError, match="privileged"):
        policy.validate_repository(tmp_path, strict=True)
    assert local.read_bytes() == before


def test_contract_schema_and_inventory(tmp_path: Path) -> None:
    """Schema bounds policy paths and an independent inventory detects dropped owners."""
    contract = copy_policy(tmp_path)
    assert set(contract["workflows"]) == {
        path.relative_to(ROOT).as_posix() for path in (ROOT / ".github/workflows").glob("*.yml")
    }
    malformed = copy.deepcopy(contract)
    malformed["workflows"]["../escape.yml"] = malformed["workflows"].pop(
        next(iter(malformed["workflows"]))
    )
    (tmp_path / policy.CONTRACT).write_text(yaml.safe_dump(malformed), encoding="utf-8")
    assert policy.main(["--repo-root", str(tmp_path)]) == 1


def test_schema_cannot_trigger_external_resolution(tmp_path: Path) -> None:
    """Offline validation rejects remote and filesystem refs before schema evaluation."""
    copy_policy(tmp_path)
    schema = tmp_path / policy.SCHEMA
    original = schema.read_text(encoding="utf-8")
    for reference in [
        "https://example.invalid/schema.json",
        "file:///outside.json",
        "../outside.json",
    ]:
        schema.write_text(
            original.replace('"$ref": "#/$defs/workflow"', f'"$ref": "{reference}"'),
            encoding="utf-8",
        )
        with pytest.raises(policy.PolicyError, match="Only local"):
            policy.validate_repository(tmp_path)


def test_guard_removal_is_detected_by_negative_oracle(tmp_path: Path) -> None:
    """Removing the credential guard resurrects the unsafe fixture independently."""
    source_path = ROOT / ".github/scripts/validate_workflow_security.py"
    source = source_path.read_text(encoding="utf-8")
    guard = 'step.get("with", {}).get("persist-credentials") is not False'
    assert source.count(guard) == 1
    mutant_path = tmp_path / "mutant.py"
    mutant_path.write_text(source.replace(guard, "False"), encoding="utf-8")
    spec = importlib.util.spec_from_file_location("credential_guard_mutant", mutant_path)
    assert spec is not None and spec.loader is not None
    mutant = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mutant)
    bad = policy.read_text(ROOT, ".github/workflows/workflow-security.yml").replace(
        "persist-credentials: false", "persist-credentials: true"
    )
    with pytest.raises(policy.PolicyError, match="persist-credentials"):
        policy.validate_workflow(bad)
    assert mutant.validate_workflow(bad)


@pytest.mark.parametrize(
    ("before", "after"),
    [
        ("runs-on: ubuntu-latest", "runs-on: self-hosted"),
        ("runs-on: ubuntu-latest", "runs-on: ubuntu-latest\n    container: attacker/image"),
        ("actions/checkout@", "attacker/checkout@"),
        ("python-version: '3.13'", "python-version: '2.7'"),
    ],
)
def test_execution_surface_substitution_fails(tmp_path: Path, before: str, after: str) -> None:
    """Reviewed runner, action coordinate, and inputs cannot be replaced silently."""
    copy_policy(tmp_path)
    path = tmp_path / ".github/workflows/workflow-security.yml"
    text = path.read_text(encoding="utf-8")
    assert before in text
    path.write_text(text.replace(before, after), encoding="utf-8")
    with pytest.raises(policy.PolicyError, match="execution controls"):
        policy.validate_repository(tmp_path)


def test_comment_cannot_authorize_flow_style_executable_pin() -> None:
    """A descriptive comment on another line cannot annotate an executable scalar."""
    text = policy.read_text(ROOT, ".github/workflows/workflow-security.yml")
    checkout_line = next(line for line in text.splitlines() if "uses: actions/checkout@" in line)
    reference = checkout_line.strip().split()[1]
    annotation = checkout_line.split("#", 1)[1].strip()
    original = (
        "      - name: Checkout repository\n" + checkout_line + "\n"
        "        with:\n          persist-credentials: false"
    )
    changed = (
        f"      # uses: {reference} # {annotation}\n"
        f"      - {{name: Checkout repository, uses: {reference}, with: {{persist-credentials: false}}}}"
    )
    assert original in text
    with pytest.raises(policy.PolicyError, match="own literal"):
        policy.validate_workflow(text.replace(original, changed))


def test_configured_workflow_hook_native_failure(tmp_path: Path) -> None:
    """The actual configured hook propagates a real unsafe checkout's native exit."""
    copy_policy(tmp_path)
    config = policy.parse_yaml(policy.read_text(ROOT, ".pre-commit-config.yaml"))
    hook = next(
        h
        for repo in config["repos"]
        for h in repo["hooks"]
        if h["id"] == "validate-workflow-security"
    )
    command = shlex.split(hook["entry"])
    path = tmp_path / ".github/workflows/workflow-security.yml"
    path.write_text(
        path.read_text(encoding="utf-8").replace(
            "persist-credentials: false", "persist-credentials: true"
        ),
        encoding="utf-8",
    )
    result = subprocess.run(
        [sys.executable, str(ROOT / command[1]), *command[2:], "--repo-root", str(tmp_path)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 1
    assert "persist-credentials" in result.stderr


@pytest.mark.parametrize("operation", ["resolve", "open"])
@pytest.mark.parametrize(
    ("error_kind", "summary"),
    [
        ("permission", "PermissionError: Permission denied"),
        ("missing-cause", "OSError: I/O error"),
        ("opaque", "OSError: I/O error"),
    ],
)
def test_native_os_error_diagnostics(operation: str, error_kind: str, summary: str) -> None:
    """Filesystem failures stay nonzero without exposing either exception filename."""
    script = """
import importlib.util
import sys
from pathlib import Path
from unittest.mock import patch

source, root, operation, error_kind = sys.argv[1:]
spec = importlib.util.spec_from_file_location("workflow_diagnostic_probe", source)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
root = Path(root)
first = "C:/private-diagnostic-canary/source.yml"
second = "C:/private-diagnostic-canary/destination.yml"
if error_kind == "permission":
    error = PermissionError(13, "Permission denied", first, None, second)
elif error_kind == "missing-cause":
    error = OSError(5, None, first, None, second)
else:
    error = OSError("C:/private-diagnostic-canary/opaque-message")
original = getattr(Path, operation)
target = root if operation == "resolve" else root / module.CONTRACT

def fail_selected(path, *args, **kwargs):
    if path == target:
        raise error
    return original(path, *args, **kwargs)

with patch.object(Path, operation, fail_selected):
    raise SystemExit(module.main(["--repo-root", str(root)]))
"""
    result = subprocess.run(
        [
            sys.executable,
            "-c",
            script,
            str(ROOT / ".github/scripts/validate_workflow_security.py"),
            str(ROOT),
            operation,
            error_kind,
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )
    assert result.returncode == 1, result.stdout + result.stderr
    assert result.stdout == ""
    assert result.stderr == f"Workflow security validation failed: {summary}\n"
    assert "private-diagnostic-canary" not in result.stderr


@pytest.mark.parametrize(
    "template",
    [
        "- uses : REF",
        "- uses   : REF",
        "- 'uses': REF",
        '- "uses": "REF"',
        r'- "u\u0073es": REF',
        "- !!str uses: REF",
        "- !inert uses: REF",
        "- uses: !!str REF",
        "- {uses: REF}",
        '- {"uses":"REF"}',
        "- {? uses : REF}",
        "# - 'uses' : REF",
        "> 1. > # - 'uses' : REF",
    ],
)
@pytest.mark.parametrize("pinned", [False, True])
def test_example_semantic_spellings(template: str, pinned: bool) -> None:
    """Decoded YAML keys and values preserve pin policy across presentation forms."""
    reference = "owner/action@" + ("a" * 40 if pinned else "v1")
    text = template.replace("REF", reference) + " # v1.2.3"
    if pinned:
        policy.check_examples(text)
    else:
        with pytest.raises(policy.PolicyError, match="full SHA"):
            policy.check_examples(text)


@pytest.mark.parametrize("fence", ["```yaml", "~~~yml", "````yaml", "```"])
@pytest.mark.parametrize("pinned", [False, True])
def test_multiline_example_semantics(fence: str, pinned: bool) -> None:
    """Whole fenced fragments expose flow mappings whose physical lines cannot parse alone."""
    revision = "a" * 40 if pinned else "v1"
    body = '- {name: Checkout,\n   "u\\u0073es" : "owner\\/action\\u0040' + revision
    body += '", # v1.2.3\n   with: {}}\n'
    close = fence.rstrip("yaml") if "yaml" in fence else fence.rstrip("yml")
    text = fence + "\n" + body + close + "\n"
    if pinned:
        policy.check_examples(text)
    else:
        with pytest.raises(policy.PolicyError, match="full SHA"):
            policy.check_examples(text)


@pytest.mark.parametrize(
    ("body", "message"),
    [
        ("- uses:", "physical line"),
        ("- uses: []", "physical line"),
        ("- uses: {}", "physical line"),
        ("- uses:\n    REF # v1.2.3", "physical line"),
        ("- {uses: REF, uses: REF} # v1.2.3", "Multiple"),
        ("- {uses: owner/action@v1,", "Invalid YAML"),
        ("- uses:owner/action@v1", "full SHA"),
    ],
)
def test_example_ambiguous_or_malformed_forms(body: str, message: str) -> None:
    """Missing values, ambiguous annotations, and incomplete action fences fail closed."""
    text = "```yaml\n" + body.replace("REF", "owner/action@" + "a" * 40) + "\n```"
    with pytest.raises(policy.PolicyError, match=message):
        policy.check_examples(text)


@pytest.mark.parametrize(
    "text",
    [
        "uses-extra: owner/action@v1",
        "Uses: owner/action@v1",
        "Prose uses an action.",
        r"\> - uses: owner/action@v1",
        "https://example.invalid/uses",
        "```yaml\nnot: [a complete fragment\n```",
        "````markdown\n```yaml\nnot: [a complete fragment\n```\n````",
    ],
)
def test_example_semantic_near_matches(text: str) -> None:
    """Prose and non-action fragments do not manufacture uses mappings."""
    policy.check_examples(text)


@pytest.mark.parametrize("guard", ["line", "fence", "comment", "legacy", "physical"])
def test_example_semantic_guard_removal(tmp_path: Path, guard: str) -> None:
    """Independent bypass fixtures expose removal of each semantic extraction boundary."""
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    replacement, text, message = {
        "line": (
            ("example_references(fragment)", "[]"),
            "- uses : owner/action@v1 # v1.2.3",
            "full SHA",
        ),
        "fence": (
            ("example_references(block)", "[]"),
            '```yaml\n- {name: Checkout,\n "u\\u0073es": "owner\\/action\\u0040v1", # v1.2.3\n with: {}}\n```',
            "full SHA",
        ),
        "comment": (
            (
                'fragment = re.sub(r"^(\\s*)#\\s?", r"\\1", candidate, count=1)',
                "fragment = candidate",
            ),
            "# - 'uses' : owner/action@v1 # v1.2.3",
            "full SHA",
        ),
        "legacy": (
            ("references.add((number, match[1]))", "pass"),
            "uses:owner/action@v1 # v1.2.3",
            "full SHA",
        ),
        "physical": (
            ("or key.start_mark.line != value.start_mark.line", "or False"),
            "```yaml\n- uses:\n    owner/action@" + "a" * 40 + " # v1.2.3\n```",
            "physical line",
        ),
    }[guard]
    before, after = replacement
    assert source.count(before) == 1
    mutant = load_policy_mutant(tmp_path, source.replace(before, after))
    with pytest.raises(policy.PolicyError, match=message):
        policy.check_examples(text)
    mutant.check_examples(text)


@pytest.mark.parametrize("key", ["uses :", "'uses':", '"uses":'])
def test_document_spelling_support_does_not_relax_executable_binding(key: str) -> None:
    """Documentation normalization cannot authorize alternate executable key spellings."""
    text = policy.read_text(ROOT, ".github/workflows/workflow-security.yml")
    with pytest.raises(policy.PolicyError, match="own literal"):
        policy.validate_workflow(text.replace("uses:", key, 1))


@pytest.mark.parametrize(
    ("scope", "field", "before", "after"),
    [
        ("workflow", "cache-mode", "read", "write"),
        ("job", "cache-mode", "read", "write"),
        (
            "job",
            "outputs",
            {"verdict": "${{ steps.gate.outputs.verdict }}"},
            {"verdict": "success"},
        ),
        ("job", "snapshot", "reviewed-image", {"image-name": "changed-image"}),
        ("step", "background", False, True),
        ("step", "wait", "gate", ["other"]),
        ("step", "wait-all", None, "absent"),
        ("step", "cancel", "gate", "other"),
    ],
)
def test_additional_execution_control_guards(
    tmp_path: Path, scope: str, field: str, before: Any, after: Any
) -> None:
    """Known execution controls require review, and removing each binding restores drift."""
    document = policy.parse_yaml(
        "on: push\npermissions: {}\njobs:\n  check:\n    permissions: {}\n"
        "    steps:\n      - run: echo check\n  consumer:\n    permissions: {}\n"
        "    needs: check\n    if: needs.check.outputs.verdict == 'success'\n    steps: []\n"
    )
    target = {
        "workflow": document,
        "job": document["jobs"]["check"],
        "step": document["jobs"]["check"]["steps"][0],
    }[scope]
    target[field] = before
    baseline = copy.deepcopy(document)
    reviewed = policy.describe_workflow(document)
    # Schema validation of positive and absent forms is independent of the descriptor predicate.
    schema = policy.parse_yaml(policy.read_text(ROOT, policy.SCHEMA))
    policy.jsonschema.validate(
        {"version": 1, "workflows": {".github/workflows/test.yml": reviewed}, "examples": []},
        schema,
    )
    if after == "absent":
        del target[field]
    else:
        target[field] = after
    assert policy.describe_workflow(document) != reviewed
    code = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    if scope == "workflow":
        guard = '        **{field: document[field] for field in ("cache-mode",) if field in document},\n'
    else:
        guard = ("                    " if scope == "job" else "    ") + f'"{field}",\n'
    assert code.count(guard) == 1
    mutant = load_policy_mutant(tmp_path, code.replace(guard, ""))
    assert mutant.describe_workflow(baseline) == mutant.describe_workflow(document)


def test_parallel_groups_fail_until_recursive_policy_exists(tmp_path: Path) -> None:
    """Unsupported nested steps cannot evade pin, command, or failure-control review."""
    template = "on: push\npermissions: {}\njobs:\n  check:\n    permissions: {}\n    steps:\n      - parallel:\n          - run: echo VALUE\n"
    for text in [template, template.replace("run: echo VALUE", "'uses': 'owner/action@v1'")]:
        with pytest.raises(policy.PolicyError, match="Parallel"):
            policy.validate_workflow(text)
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    guard = 'if "parallel" in step:'
    assert source.count(guard) == 1
    mutant = load_policy_mutant(tmp_path, source.replace(guard, "if False:"))
    assert mutant.validate_workflow(template.replace("VALUE", "one")) == mutant.validate_workflow(
        template.replace("VALUE", "two")
    )


@pytest.mark.parametrize("pinned", [False, True])
@pytest.mark.parametrize("flow", [False, True])
def test_commented_workflow_examples_use_semantic_keys(pinned: bool, flow: bool) -> None:
    """Commented copyable workflows obey the same identity policy without becoming executable."""
    ref = "owner/action@" + ("a" * 40 if pinned else "v1")
    example = (
        f'# - {{name: Checkout,\n#    "uses" : "{ref}", # v1.2.3\n#    with: {{}}}}'
        if flow
        else f"# - uses : {ref} # v1.2.3"
    )
    text = policy.read_text(ROOT, ".github/workflows/workflow-security.yml") + "\n" + example + "\n"
    if pinned:
        policy.validate_workflow(text)
    else:
        with pytest.raises(policy.PolicyError, match="full SHA"):
            policy.validate_workflow(text)


def test_commented_workflow_guard_removal(tmp_path: Path) -> None:
    """Removing semantic comment checking restores the alternate-key bypass independently."""
    text = policy.read_text(ROOT, ".github/workflows/workflow-security.yml")
    text += "\n# - uses : owner/action@v1 # v1.2.3\n"
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.validate_workflow(text)
    code = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    guard = "    check_commented_examples(lines, resolver)\n"
    assert code.count(guard) == 1
    mutant = load_policy_mutant(tmp_path, code.replace(guard, ""))
    assert mutant.validate_workflow(text)


def test_unclosed_escaped_yaml_example_is_not_ignored() -> None:
    """Escaping the action coordinate cannot hide decoded uses in an unfinished fence."""
    text = '```yaml\n? "u\\\n  ses"\n: "owner\\/action\\u0040v1" # v1.2.3\n'
    with pytest.raises(policy.PolicyError):
        policy.check_examples(text)


def test_pinned_multiline_plain_flow_example() -> None:
    """A flow comma belongs to YAML syntax and is not part of the decoded action identity."""
    text = "```yaml\n- {name: Checkout,\n uses : owner/action@" + "a" * 40
    text += ", # v1.2.3\n with: {}}\n```"
    policy.check_examples(text)


def test_example_composition_bounds_and_inert_cycles(tmp_path: Path) -> None:
    """Recursive nodes terminate and an independent deep input detects removed depth bounds."""
    assert policy.example_references("&cycle [*cycle]") == []
    assert policy.example_references("!!python/object/apply:os.system [not-executed]") == []
    text = "[" * 70 + "x" + "]" * 70
    with pytest.raises(policy.PolicyError, match="nesting"):
        policy.example_references(text)
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    start = source.index("class ExampleLoader")
    prefix, example_code = source[:start], source[start:]
    assert example_code.count("if depth >= 64:") == 1
    mutant = load_policy_mutant(
        tmp_path, prefix + example_code.replace("if depth >= 64:", "if False:", 1)
    )
    assert mutant.example_references(text) == []


@pytest.mark.parametrize(
    ("scope", "field", "invalid"),
    [
        ("workflow", "cache-mode", "bogus"),
        ("job", "cache-mode", "bogus"),
        ("job", "outputs", []),
        ("job", "snapshot", 42),
        ("step", "background", "true"),
        ("step", "wait", [1]),
        ("step", "wait-all", False),
        ("step", "cancel", ["gate"]),
    ],
)
def test_optional_control_schema_guard_removal(scope: str, field: str, invalid: Any) -> None:
    """Independent invalid values expose removal of optional schema shape restrictions."""
    descriptor = policy.describe_workflow(
        {"on": "push", "permissions": {}, "jobs": {"check": {"permissions": {}, "steps": [{}]}}}
    )
    target = {
        "workflow": descriptor,
        "job": descriptor["jobs"]["check"]["controls"],
        "step": descriptor["jobs"]["check"]["steps"][0],
    }[scope]
    target[field] = invalid
    contract = {
        "version": 1,
        "workflows": {".github/workflows/test.yml": descriptor},
        "examples": [],
    }
    schema = policy.parse_yaml(policy.read_text(ROOT, policy.SCHEMA))
    with pytest.raises(policy.jsonschema.ValidationError):
        policy.jsonschema.validate(contract, schema)
    workflow_properties = schema["$defs"]["workflow"]["properties"]
    job_properties = workflow_properties["jobs"]["additionalProperties"]["properties"]
    properties = {
        "workflow": workflow_properties,
        "job": job_properties["controls"]["properties"],
        "step": job_properties["steps"]["items"]["properties"],
    }[scope]
    properties[field] = {}
    policy.jsonschema.validate(contract, schema)


@pytest.mark.parametrize("language", ["text", "bash", "json", "markdown", "python"])
@pytest.mark.parametrize(
    "body",
    [
        "uses: owner/action@v1",
        "# uses: owner/action@v1",
        "- 'uses': owner/action@v1",
        "- {uses: owner/action@v1}",
        "uses:owner/action@v1",
    ],
)
def test_non_yaml_fences_are_literal(language: str, body: str) -> None:
    """Literal samples neither fail pin policy nor contribute governed references."""
    calls: list[tuple[str, str]] = []

    def resolver(repository: str, release: str) -> str:
        calls.append((repository, release))
        return "a" * 40

    assert policy.check_examples(f"```{language}\n{body}\n```", resolver) == 0
    assert policy.check_examples(f"```{language}\nuses: owner/action@{'a' * 40} # v1.2.3\n```") == 0
    assert calls == []


@pytest.mark.parametrize(
    "text",
    [
        "```text\nuses: owner/action@v1",
        "````text\n```yaml\nuses: owner/action@v1\n```\n````",
        "~~~text\n```\nuses: owner/action@v1\n~~~~",
        "> ```text\n> uses: owner/action@v1\n> ```",
        "- ```text\n  uses: owner/action@v1\n  ```",
        "- > ```text\n  > uses: owner/action@v1\n  > ```",
        "> 1. > ```text\n>    > uses: owner/action@v1\n>    > ```",
        "- > " * 80 + "```text\n" + "  > " * 80 + "uses: owner/action@v1\n",
    ],
)
def test_literal_fence_container_boundaries(text: str) -> None:
    """Matched or unfinished literal fences preserve ordered containing blocks."""
    assert policy.check_examples(text) == 0


@pytest.mark.parametrize(
    "prefix",
    [
        "```te`xt\n",
        "> ```text\n> literal\n",
        "- ```text\n  literal\n\n",
        "- > ```text\n  > literal\n",
        "> 1. > ```text\n>    > literal\n",
        "- > " * 80 + "```text\n" + "  > " * 80 + "literal\n",
        "```text\nuses: owner/action@v1\n```\n",
    ],
)
def test_literal_fences_do_not_hide_live_following_examples(prefix: str) -> None:
    """Invalid openers, ended containers and closed literals grant no exemption."""
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.check_examples(prefix + "uses: owner/action@v1\n")


@pytest.mark.parametrize("language", ["yaml title=example", "YML extra", "yaml", "yml", ""])
@pytest.mark.parametrize("ending", ["closed", "unclosed", "malformed"])
def test_governed_multiline_fences_with_info_metadata(language: str, ending: str) -> None:
    """The first info token retains complete YAML and unfinished-block enforcement."""
    body = '- {name: Checkout,\n "u\\u0073es": "owner/action@' + "a" * 40
    body += '", # v1.2.3\n with: {}}\n'
    text = f"```{language}\n" + body
    if ending == "closed":
        text += "```\n"
        assert policy.check_examples(text) == 1
        with pytest.raises(policy.PolicyError, match="full SHA"):
            policy.check_examples(text.replace("a" * 40, "v1"))
    else:
        if ending == "malformed":
            text = text.replace("with: {}}", "with: {") + "```\n"
        with pytest.raises(policy.PolicyError, match="Unclosed|Invalid YAML"):
            policy.check_examples(text)


@pytest.mark.parametrize("guard", ["semantic", "literal", "quote", "list", "pin"])
def test_non_yaml_opacity_guard_mutations(tmp_path: Path, guard: str) -> None:
    """Independent literal-count and live-failure oracles detect each restored defect."""
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    before, after, text, literal = {
        "semantic": (
            "        if number in opaque_lines:\n            continue\n        if number in fenced_lines",
            "        if number in fenced_lines",
            "```text\n- 'uses': owner/action@v1\n```",
            True,
        ),
        "literal": (
            "number not in opaque_lines\n            and number not in semantic_lines",
            "number not in semantic_lines",
            "```text\nuses:owner/action@v1\n```",
            True,
        ),
        "quote": (
            "if depth < width:",
            "if False:",
            "> ```text\n> harmless\nuses: owner/action@v1",
            False,
        ),
        "list": (
            "if line_is_outside_list_item(line[offset:], width):",
            "if False:",
            "- ```text\n  harmless\nuses: owner/action@v1",
            False,
        ),
        "pin": (
            "check_reference(reference, lines[number], resolver)",
            "pass",
            "uses: owner/action@v1",
            False,
        ),
    }[guard]
    assert source.count(before) == 1
    mutant = load_policy_mutant(tmp_path, source.replace(before, after))
    if literal:
        assert policy.check_examples(text) == 0
        with pytest.raises(mutant.PolicyError, match="full SHA"):
            mutant.check_examples(text)
    else:
        with pytest.raises(policy.PolicyError, match="full SHA"):
            policy.check_examples(text)
        assert mutant.check_examples(text) == (1 if guard == "pin" else 0)


def test_non_yaml_blank_continuation_resource_mutant(tmp_path: Path) -> None:
    """Near-limit nested lists and blanks cannot cause repeated full-stack scans."""
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    fixture = tmp_path / "literal.md"
    fixture.write_text(
        "- " * 300_000 + "```text\n" + "\n" * 300_000, encoding="utf-8", newline="\n"
    )
    assert 850_000 < fixture.stat().st_size < policy.LIMIT
    code = (
        "import importlib.util,sys; from pathlib import Path; "
        "spec=importlib.util.spec_from_file_location('resource_policy',sys.argv[1]); "
        "module=importlib.util.module_from_spec(spec); spec.loader.exec_module(module); "
        "count=module.check_examples(Path(sys.argv[2]).read_text(encoding='utf-8')); "
        "assert count == 0, count; print(count)"
    )
    command = [sys.executable, "-E", "-c", code]
    current = ROOT / ".github/scripts/validate_workflow_security.py"
    result = subprocess.run(
        [*command, str(current), str(fixture)],
        cwd=tmp_path,
        capture_output=True,
        text=True,
        timeout=10,
        check=False,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert result.stdout.strip() == "0"
    start = source.index(
        '    if not line.strip(" \\t"):', source.index("def example_fence_content(")
    )
    end = source.index("\n\ndef opaque_example_lines", start)
    restored = """    content = line
    for kind, width in containers:
        if kind == "quote":
            depth, offset = consume_blockquote_prefix(content, max_depth=width)
            if depth < width:
                return None
            content = content[offset:]
        else:
            if line_is_outside_list_item(content, width):
                return None
            content = content[width:]
    return active_fence_content(content, fence)
"""
    mutant = tmp_path / "resource_mutant.py"
    mutant.write_text(source[:start] + restored + source[end:], encoding="utf-8")
    shutil.copyfile(
        ROOT / ".github/scripts/instruction_contract_support.py",
        tmp_path / "instruction_contract_support.py",
    )
    with pytest.raises(subprocess.TimeoutExpired):
        subprocess.run(
            [*command, str(mutant), str(fixture)],
            cwd=tmp_path,
            capture_output=True,
            text=True,
            timeout=10,
            check=False,
        )


@pytest.mark.parametrize("info", ["te`xt", "text`"])
def test_non_yaml_invalid_opener_guard_mutation(tmp_path: Path, info: str) -> None:
    """Permitting backticks in an info string would hide a fixed live reference."""
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    anchor = "parse_markdown_fence_open(line[offset:], MARKDOWN_FENCE_CONTEXT)"
    assert source.count(anchor) == 1
    mutant = load_policy_mutant(
        tmp_path,
        source.replace(
            anchor,
            f'parse_markdown_fence_open(line[offset:].replace({info!r}, "text"), MARKDOWN_FENCE_CONTEXT)',
        ),
    )
    text = f"```{info}\nuses: owner/action@v1\n"
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.check_examples(text)
    assert mutant.check_examples(text) == 0


def test_non_yaml_large_literal_preserves_live_pin_check(tmp_path: Path) -> None:
    """A large literal remains inert without excusing an adjacent actual example."""
    literal = "```text\n" + "uses: owner/action@v1\n" * 40_000 + "```\n"
    assert 800_000 < len(literal.encode("utf-8")) < policy.LIMIT
    path = tmp_path / "literal.md"
    path.write_text(literal, encoding="utf-8", newline="\n")
    assert policy.check_examples(policy.read_text(tmp_path, "literal.md")) == 0
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.check_examples(literal + "uses: owner/action@v1\n")
    pinned = "uses: owner/action@" + "a" * 40 + " # v1.2.3\n"
    assert policy.check_examples(literal + pinned) == 1
    path.write_bytes(b"x" * (policy.LIMIT + 1))
    with pytest.raises(policy.PolicyError, match="1 MiB"):
        policy.read_text(tmp_path, "literal.md")


@pytest.mark.parametrize("which", [0, 1])
def test_real_terraform_authoring_examples_remain_governed(which: int) -> None:
    """Both intentionally copyable workflow snippets retain their fixed pin obligations."""
    path = "docs/terraform/TERRAFORM_COPILOT_INSTRUCTIONS_GUIDE.md"
    text = (ROOT / path).read_text(encoding="utf-8")
    assert path in policy.load_contract(ROOT)["examples"]
    assert policy.check_examples(text) == 10
    anchor = "```yaml\n# .github/workflows/terraform-ci.yml\n"
    starts = [match.start() for match in re.finditer(re.escape(anchor), text)]
    assert len(starts) == 2
    start = starts[which]
    end = text.index("\n```", start + len(anchor))
    block = text[start:end]
    assert policy.check_examples(block + "\n```") == (2 if which == 0 else 8)
    floating = re.sub(r"(?<=@)[0-9a-f]{40}", "v1", block, count=1)
    assert floating != block
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.check_examples(text[:start] + floating + text[end:])


def test_old_terraform_literal_wrapping_breaks_fixed_coverage_oracle() -> None:
    """Restoring the old outer text fences loses all ten intentional references."""
    text = (ROOT / "docs/terraform/TERRAFORM_COPILOT_INSTRUCTIONS_GUIDE.md").read_text(
        encoding="utf-8"
    )
    assert policy.check_examples(text) == 10
    anchor = "\n```\n\n```yaml\n# .github/workflows/terraform-ci.yml\n"
    assert text.count(anchor) == 2
    for _block in range(2):
        start = text.index(anchor)
        end = text.index("\n```", start + len(anchor))
        body = text[start:end].replace(
            anchor, "\n\\`\\`\\`yaml\n# .github/workflows/terraform-ci.yml\n"
        )
        text = text[:start] + body + "\n\\`\\`\\`" + text[end:]
    assert policy.check_examples(text) == 0


@pytest.mark.parametrize(
    ("opener", "prefix"),
    [
        ("- 1. ", "     "),
        ("1. - 2. ", "        "),
        ("> - 1. ", ">      "),
        ("- > 1. ", "  >    "),
        ("- > - > ", "  >   > "),
        ("- " * 80, " " * 160),
    ],
)
@pytest.mark.parametrize("language", ["yaml", "YML metadata", ""])
def test_governed_nested_container_semantics(opener: str, prefix: str, language: str) -> None:
    """Nested Markdown containers neither hide YAML nodes nor count as YAML depth."""
    reference = "owner/action@" + "a" * 40
    body = f'- {{name: Checkout,\n "uses": {reference}, # v1.2.3\n with: {{}}}}'
    text = opener + "```" + language + "\n"
    text += "\n".join(prefix + line for line in body.splitlines()) + "\n" + prefix + "```"
    calls: list[tuple[str, str]] = []

    def resolve(repository: str, version: str) -> str:
        """Record each physically bound pin check without network access."""
        calls.append((repository, version))
        return "a" * 40

    assert policy.check_examples(text, resolve) == 1
    assert calls == [("owner/action", "v1.2.3")]
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.check_examples(text.replace("a" * 40, "v1"))
    for pinned in (False, True):
        split = '- {\n ? "uses"\n : owner/action@' + ("a" * 40 if pinned else "v1") + "\n}"
        split_text = opener + "```" + language + "\n"
        split_text += "\n".join(prefix + line for line in split.splitlines())
        with pytest.raises(policy.PolicyError, match="physical line"):
            policy.check_examples(split_text + "\n" + prefix + "```")


@pytest.mark.parametrize(
    ("before", "prefix"),
    [
        ("1. item\n\n", "    "),
        ("- Outer\n  1. Inner\n\n", "     "),
        ("123. item\n\n", "     "),
        ("> - Outer\n>   1. Inner\n>\n", ">      "),
        ("", "    "),
    ],
)
def test_governed_indented_continuation_compatibility(before: str, prefix: str) -> None:
    """Existing indented governed examples retain complete-block validation."""
    body = '- {name: Checkout,\n "uses": owner/action@' + "a" * 40
    body += ", # v1.2.3\n with: {}}"
    text = before + prefix + "```yaml\n"
    text += "\n".join(prefix + line for line in body.splitlines()) + "\n" + prefix + "```"
    assert policy.check_examples(text) == 1
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.check_examples(text.replace("a" * 40, "v1"))
    with pytest.raises(policy.PolicyError, match="Unclosed"):
        policy.check_examples(text.rsplit("\n", 1)[0])


@pytest.mark.parametrize("ending", ["eof", "dedent", "quote-end", "wrong-close", "short-close"])
def test_governed_nested_unclosed_boundaries(ending: str) -> None:
    """Container ends and nonmatching closers cannot bless unfinished action examples."""
    prefix = ">      " if ending == "quote-end" else "     "
    opener = "> - 1. " if ending == "quote-end" else "- 1. "
    text = opener + "````yaml\n" + prefix + "uses: owner/action@" + "a" * 40 + " # v1.2.3\n"
    text += {
        "eof": "",
        "dedent": "outside\n",
        "quote-end": "outside\n",
        "wrong-close": prefix + "~~~~\n",
        "short-close": prefix + "```\n",
    }[ending]
    with pytest.raises(policy.PolicyError, match="Unclosed|Invalid YAML"):
        policy.check_examples(text)


def test_governed_nested_comments_and_annotation_lines() -> None:
    """Comments remain governed, and a release annotation cannot move to another line."""
    text = "- 1. ~~~yaml\n     # uses: owner/action@" + "a" * 40 + " # v1.2.3\n     ~~~~"
    assert policy.check_examples(text) == 1
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.check_examples(
            text.replace("# uses:", "uses:").replace(" # v1.2.3", "\n     # v1.2.3")
        )
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.check_examples(text + "\n# uses: owner/action@v1\n")


@pytest.mark.parametrize("guard", ["block", "physical", "malformed", "unclosed", "delimiter"])
def test_governed_nested_guard_removal(tmp_path: Path, guard: str) -> None:
    """Fixed nested-block failure and count oracles detect precise guard removal."""
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    pin = "owner/action@" + "a" * 40
    split = '- 1. ```yaml\n     - {"uses"\n     : owner/action@v1}\n     ```'
    before, after, text, expected = {
        "block": ("example_references(block)", "[]", split, "Invalid YAML"),
        "physical": (
            "or key.start_mark.line != value.start_mark.line",
            "or False",
            f"- 1. ```yaml\n     - uses:\n         {pin} # v1.2.3\n     ```",
            "physical line",
        ),
        "malformed": (
            "if EXAMPLE_ACTION.search(block):",
            "if False:",
            '- 1. ```yaml\n     - {"uses"\n     : owner/action@v1\n     ```',
            "Invalid YAML",
        ),
        "unclosed": (
            "if unclosed and (found or EXAMPLE_ACTION.search(block)):",
            "if False:",
            f"- 1. ```yaml\n     uses: {pin} # v1.2.3",
            "Unclosed",
        ),
        "delimiter": (
            "                fenced_lines.add(number)\n    if fence is not None and governed:",
            "                pass\n    if fence is not None and governed:",
            "- " * 80 + "```yaml\n" + " " * 160 + f"uses: {pin} # v1.2.3\n" + " " * 160 + "```",
            "",
        ),
    }[guard]
    assert source.count(before) == 1
    mutant = load_policy_mutant(tmp_path, source.replace(before, after))
    if expected:
        with pytest.raises(policy.PolicyError, match=expected):
            policy.check_examples(text)
        mutant.check_examples(text)
    else:
        assert policy.check_examples(text) == 1
        with pytest.raises(mutant.PolicyError, match="nesting"):
            mutant.check_examples(text)


def test_governed_nested_old_collector_restoration(tmp_path: Path) -> None:
    """The old single-marker collector fails the unchanged split-key rejection oracle."""
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    start = source.index(
        "    for number, raw_line in enumerate(raw_lines):", source.index("def check_examples(")
    )
    end = source.index("    block_reference_lines =", start)
    old = """    old_fence = None
    for number, candidate in enumerate(lines):
        if number in opaque_lines:
            continue
        match = EXAMPLE_FENCE.match(candidate)
        if not match:
            continue
        delimiter, info = match.groups()
        if old_fence is None:
            old_fence = (delimiter[0], len(delimiter), number + 1,
                         not info.strip() or info.strip().lower().split(maxsplit=1)[0] in {"yaml", "yml"})
        elif delimiter[0] == old_fence[0] and len(delimiter) >= old_fence[1] and not info.strip():
            if old_fence[3]:
                collect_block(old_fence[2], number)
            old_fence = None
    if old_fence is not None and old_fence[3]:
        collect_block(old_fence[2], len(lines), unclosed=True)
"""
    mutant = load_policy_mutant(tmp_path, source[:start] + old + source[end:])
    text = '- 1. ```yaml\n     - {"uses"\n     : owner/action@v1}\n     ```'
    with pytest.raises(policy.PolicyError, match="Invalid YAML"):
        policy.check_examples(text)
    assert mutant.check_examples(text) == 0


def test_governed_deep_container_resource_bound(tmp_path: Path) -> None:
    """A near-limit governed block processes deep containers and blank lines once."""
    fixture = tmp_path / "governed.md"
    prefix = " " * 240_000
    fixture.write_text(
        "- " * 120_000
        + "```yaml\n"
        + "\n" * 200_000
        + prefix
        + "uses: owner/action@"
        + "a" * 40
        + " # v1.2.3\n"
        + prefix
        + "```\n",
        encoding="utf-8",
        newline="\n",
    )
    assert 900_000 < fixture.stat().st_size < policy.LIMIT
    code = (
        "import importlib.util,sys; from pathlib import Path; "
        "spec=importlib.util.spec_from_file_location('resource_policy',sys.argv[1]); "
        "module=importlib.util.module_from_spec(spec); spec.loader.exec_module(module); "
        "count=module.check_examples(Path(sys.argv[2]).read_text(encoding='utf-8')); "
        "assert count == 1, count; print(count)"
    )
    result = subprocess.run(
        [
            sys.executable,
            "-E",
            "-c",
            code,
            str(ROOT / ".github/scripts/validate_workflow_security.py"),
            str(fixture),
        ],
        cwd=tmp_path,
        capture_output=True,
        text=True,
        timeout=10,
        check=False,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert result.stdout.strip() == "1"


@pytest.mark.parametrize("pinned", [False, True])
def test_governed_nested_malformed_implicit_key(pinned: bool) -> None:
    """The reported implicit split key fails complete YAML syntax even with a pin."""
    reference = "owner/action@" + ("a" * 40 if pinned else "v1")
    text = f'- 1. ```yaml\n     - {{"uses"\n     : {reference}}}\n     ```'
    with pytest.raises(policy.PolicyError, match="Invalid YAML"):
        policy.check_examples(text)


def test_governed_indented_non_yaml_compatibility_boundary() -> None:
    """Legacy non-YAML delimiters neither consume following YAML nor gain opacity."""
    before = "    ```text\n    harmless literal\n    ```\n\n"
    governed = "```yaml\nuses: owner/action@" + "a" * 40 + " # v1.2.3\n```\n"
    assert policy.check_examples(before + governed) == 1
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.check_examples(
            before.replace("harmless literal", "uses: owner/action@v1") + governed
        )
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.check_examples(before + governed.replace("a" * 40, "v1"))


def workflow_with_scalar_step(body: str) -> str:
    """Build a fixed workflow with caller-supplied step YAML, without a policy oracle."""
    return (
        "on: push\npermissions: {}\njobs:\n  sample:\n    permissions: {}\n"
        "    runs-on: ubuntu-latest\n    steps:\n      - " + body
    )


@pytest.mark.parametrize("style", ["|", "|-", "|+", "|2-", ">", ">-", ">+"])
@pytest.mark.parametrize("prefix", ["", "# "])
def test_workflow_scalar_action_text(style: str, prefix: str) -> None:
    """Action-looking block scalar content remains data with its original fingerprint."""
    content = prefix + "uses: owner/action@v1"
    text = workflow_with_scalar_step(f"run: {style}\n          {content}\n")
    expected = content + ("" if style.endswith("-") else "\n")
    actual = policy.validate_workflow(text)
    assert actual["jobs"]["sample"]["steps"] == [
        {"run_sha256": hashlib.sha256(expected.encode()).hexdigest()}
    ]


@pytest.mark.parametrize("quote", ["'", '"'])
@pytest.mark.parametrize("prefix", ["", "# "])
def test_workflow_quoted_scalar_action_text(quote: str, prefix: str) -> None:
    """Quoted continuation and closing rows stay within their composed scalar."""
    content = prefix + "uses: owner/action@v1"
    text = workflow_with_scalar_step(f"run: {quote}echo\n          {content}{quote}\n")
    actual = policy.validate_workflow(text)
    assert actual["jobs"]["sample"]["steps"] == [
        {"run_sha256": hashlib.sha256(f"echo {content}".encode()).hexdigest()}
    ]


@pytest.mark.parametrize(
    ("body", "expected"),
    [
        ("run: echo\n          uses:owner/action@v1\n", "echo uses:owner/action@v1"),
        ("run: |\n          # uses: owner/action@v1", "# uses: owner/action@v1"),
        (
            "run: |\n          cat <<'SAMPLE'\n          uses: owner/action@v1\n          SAMPLE\n",
            "cat <<'SAMPLE'\nuses: owner/action@v1\nSAMPLE\n",
        ),
        (
            "run: |\n          cat <<'SAMPLE'\n          # uses: owner/action@v1\n          SAMPLE\n",
            "cat <<'SAMPLE'\n# uses: owner/action@v1\nSAMPLE\n",
        ),
    ],
)
def test_workflow_scalar_plain_eof_and_heredoc(body: str, expected: str) -> None:
    """Plain continuation, EOF and shell fixture text do not create action identities."""
    actual = policy.validate_workflow(workflow_with_scalar_step(body))
    assert actual["jobs"]["sample"]["steps"] == [
        {"run_sha256": hashlib.sha256(expected.encode()).hexdigest()}
    ]


def test_workflow_nonexecutable_uses_input_and_environment() -> None:
    """An action input named uses and an environment scalar are ordinary data."""
    pin = "owner/action@" + "a" * 40 + " # v1.2.3"
    text = workflow_with_scalar_step(
        f"uses: {pin}\n        with:\n          uses: owner/action@v1\n"
        "        env:\n          SAMPLE: |\n            uses: owner/action@v1\n"
        "            # uses: owner/action@v1\n"
    )
    actual = policy.validate_workflow(text)
    assert actual["jobs"]["sample"]["steps"] == [
        {
            "action": "owner/action",
            "with": {"uses": "owner/action@v1"},
            "env": {"SAMPLE": "uses: owner/action@v1\n# uses: owner/action@v1\n"},
        }
    ]


@pytest.mark.parametrize(
    "body",
    [
        "run: |\n          echo harmless\n",
        "run: |-\n          echo harmless\n",
        "run: |+\n          echo harmless\n\n",
        "run: |2-\n          echo harmless\n",
        "run: >\n          echo harmless\n",
        "run: 'echo\n          harmless'\n",
        'run: "echo\n          harmless"\n',
        "run: echo\n          harmless\n",
    ],
)
@pytest.mark.parametrize("suffix", ["      # uses: ", "      - uses: "])
def test_workflow_scalar_end_preserves_real_references(body: str, suffix: str) -> None:
    """The first true YAML comment or action after a scalar remains governed."""
    text = workflow_with_scalar_step(body) + suffix
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.validate_workflow(text + "owner/action@v1\n")
    policy.validate_workflow(text + "owner/action@" + "a" * 40 + " # v1.2.3\n")


def test_scalar_opacity_does_not_allow_invalid_yaml_or_executable_forms() -> None:
    """Literal data cannot authorize a real mapping, folded action or borrowed annotation."""
    pin = "owner/action@" + "a" * 40
    for body in [
        "run: echo\n          uses: owner/action@v1\n",
        f"uses: >-\n          {pin} # v1.2.3\n",
        f"uses: '{pin}' # v1.2.3\n",
        f"{{uses: {pin}}} # v1.2.3\n",
        f"run: |\n          uses: {pin} # v1.2.3\n      - uses: {pin}\n",
    ]:
        with pytest.raises(policy.PolicyError):
            policy.validate_workflow(workflow_with_scalar_step(body))
    reusable = (
        "on: push\npermissions: {}\njobs:\n  sample:\n    permissions: {}\n"
        "    uses: owner/repo/.github/workflows/build.yml@"
    )
    with pytest.raises(policy.PolicyError, match="full SHA"):
        policy.validate_workflow(reusable + "v1\n")
    policy.validate_workflow(reusable + "a" * 40 + " # v1.2.3\n")


@pytest.mark.parametrize("guard", ["raw-scan", "scalar-exclusion", "end-boundary", "executable"])
def test_workflow_scalar_guard_mutations(tmp_path: Path, guard: str) -> None:
    """Independent positive and negative fixtures detect each classification regression."""
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    if guard == "raw-scan":
        before = "    for job in jobs.values():\n"
        after = (
            "    for line in text.splitlines():\n"
            "        if not line.lstrip().startswith('#') and (match := USES_LINE.match(line)):\n"
            "            check_reference(match[1], line, resolver)\n" + before
        )
        text = workflow_with_scalar_step("run: |\n          uses: owner/action@v1\n")
    elif guard == "scalar-exclusion":
        before = "    lines = workflow_comment_lines(tree, text.splitlines())\n"
        after = "    lines = text.splitlines()\n"
        text = workflow_with_scalar_step("run: |\n          # uses: owner/action@v1\n")
    elif guard == "end-boundary":
        before = "            end = node.end_mark.line + bool(node.end_mark.column)\n"
        after = "            end = node.end_mark.line + 1\n"
        text = workflow_with_scalar_step("run: |\n          echo harmless\n")
        text += "      # uses: owner/action@v1\n"
    else:
        before = "    tree = check_executable_annotations(text, resolver)\n"
        after = "    tree = yaml.compose(text, Loader=WorkflowLoader)\n"
        text = workflow_with_scalar_step("uses: owner/action@v1\n")
    assert source.count(before) == 1
    mutant = load_policy_mutant(tmp_path, source.replace(before, after))
    if guard in {"raw-scan", "scalar-exclusion"}:
        policy.validate_workflow(text)
        with pytest.raises(mutant.PolicyError, match="full SHA"):
            mutant.validate_workflow(text)
    else:
        with pytest.raises(policy.PolicyError, match="full SHA"):
            policy.validate_workflow(text)
        mutant.validate_workflow(text)


def test_workflow_scalar_graph_identity_guard(tmp_path: Path) -> None:
    """Identity tracking handles shared/cyclic nodes without granting public YAML aliases."""
    source = (ROOT / ".github/scripts/validate_workflow_security.py").read_text(encoding="utf-8")
    start = source.index("def workflow_comment_lines(")
    end = source.index("\ndef validate_workflow(", start)
    helper = source[start:end]
    guard = "        if id(node) in seen:\n"
    assert helper.count(guard) == 1
    changed = source[:start] + helper.replace(guard, "        if False:\n") + source[end:]
    mutant = load_policy_mutant(tmp_path, changed)
    script = """
import importlib.util
import sys

spec = importlib.util.spec_from_file_location("scalar_graph", sys.argv[1])
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
text = "run: |\\n  # uses: owner/action@v1\\n# real comment\\n"
tree = module.yaml.compose(text, Loader=module.WorkflowLoader)
key, value = tree.value[0]
tree.value.extend([(key, value), (key, tree)])
assert module.workflow_comment_lines(tree, text.splitlines()) == ["", "", "# real comment"]
print("shared-cycle-pass")
"""
    command = [sys.executable, "-c", script]
    result = subprocess.run(
        [*command, str(ROOT / ".github/scripts/validate_workflow_security.py")],
        check=False,
        capture_output=True,
        text=True,
        timeout=10,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert result.stdout.strip() == "shared-cycle-pass"
    with pytest.raises(subprocess.TimeoutExpired):
        subprocess.run(
            [*command, mutant.__file__], check=False, capture_output=True, text=True, timeout=2
        )
    with pytest.raises(policy.PolicyError, match="anchors"):
        policy.validate_workflow(workflow_with_scalar_step("run: &body echo\n        env: *body\n"))


def test_configured_workflow_hook_scalar_boundary_native(tmp_path: Path) -> None:
    """Actual hook argv accepts owned scalar data and rejects a real comment after it."""
    contract = copy_policy(tmp_path)
    path = ".github/workflows/scalar.yml"
    run = "cat <<'SAMPLE'\nuses: owner/action@v1\n# uses: owner/action@v1\nSAMPLE\n"
    text = workflow_with_scalar_step(
        "run: |\n" + "".join("          " + line + "\n" for line in run.splitlines())
    )
    # Fixed contract structure and command bytes are independent of describe_workflow.
    contract["workflows"][path] = {
        "events": "push",
        "permissions": {},
        "defaults": {},
        "env": {},
        "jobs": {
            "sample": {
                "controls": {"permissions": {}, "runs-on": "ubuntu-latest"},
                "steps": [{"run_sha256": hashlib.sha256(run.encode()).hexdigest()}],
            }
        },
    }
    (tmp_path / policy.CONTRACT).write_text(yaml.safe_dump(contract), encoding="utf-8")
    workflow = tmp_path / path
    workflow.write_text(text, encoding="utf-8")
    config = policy.parse_yaml(policy.read_text(ROOT, ".pre-commit-config.yaml"))
    hook = next(
        h
        for repo in config["repos"]
        for h in repo["hooks"]
        if h["id"] == "validate-workflow-security"
    )
    configured = shlex.split(hook["entry"])
    command = [
        sys.executable,
        str(ROOT / configured[1]),
        *configured[2:],
        "--repo-root",
        str(tmp_path),
    ]
    for suffix, expected in [("", 0), ("      # uses: owner/action@v1\n", 1)]:
        workflow.write_text(text + suffix, encoding="utf-8")
        result = subprocess.run(command, check=False, capture_output=True, text=True, timeout=30)
        assert result.returncode == expected, result.stdout + result.stderr
        if expected:
            assert "full SHA" in result.stderr
    del contract["workflows"][path]
    (tmp_path / policy.CONTRACT).write_text(yaml.safe_dump(contract), encoding="utf-8")
    # Local unowned files still require explicit strict opt-in.
    result = subprocess.run(command, check=False, capture_output=True, text=True, timeout=30)
    assert result.returncode == 0, result.stdout + result.stderr
    result = subprocess.run(
        [*command, "--strict"], check=False, capture_output=True, text=True, timeout=30
    )
    assert result.returncode == 1 and "full SHA" in result.stderr
