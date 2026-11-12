"""Accept workflow governance through actual module adoption and synchronization."""

from __future__ import annotations

import importlib.util
import json
import shlex
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

import yaml  # type: ignore[import-untyped]

from tests import test_materialize_downstream_adoption as lifecycle
from tests._pytest_compat import pytest
from tests.test_contract_wiring import assert_ci, assert_hook
from tests.test_workflow_security_contract import ROOT, policy

pytestmark = pytest.mark.upstream_template_only
YAML_EXAMPLE = ".github/instructions/yaml.instructions.md"
ONBOARDING_EXAMPLE = "OPTIONAL_CONFIGURATIONS.md"
TERRAFORM_EXAMPLES = (
    "docs/terraform/TERRAFORM_COPILOT_INSTRUCTIONS_GUIDE.md",
    "docs/terraform/TERRAFORM_LINTING_GUIDE.md",
    "docs/terraform/TERRAFORM_TESTING_GUIDE.md",
)
CURSOR_EXAMPLE = ".cursor/rules/repository-instructions.mdc"


def cursor_example_source(source: Path, *, declared: bool, floating: bool = False) -> None:
    """Copy reviewed policy inputs and add one private Cursor documentation example."""
    from tests.test_workflow_security_contract import copy_policy

    contract = copy_policy(source)
    for relative in (CURSOR_EXAMPLE, ".github/scripts/validate_workflow_security.py"):
        destination = source / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(ROOT / relative, destination)
    reference = "actions/checkout@" + ("v7" if floating else "a" * 40)
    with (source / CURSOR_EXAMPLE).open("a", encoding="utf-8", newline="\n") as stream:
        stream.write(f"\n```yaml\n- uses: {reference} # v7.0.1\n```\n")
    if declared:
        contract["examples"].append(CURSOR_EXAMPLE)
    lifecycle.write_yaml(source / policy.CONTRACT, contract)


@pytest.mark.parametrize("floating", [False, True])
def test_cursor_example_discovery_and_old_suffix_mutant(
    tmp_path: Path, monkeypatch: Any, floating: bool
) -> None:
    """The fixed unlisted Cursor oracle detects restoration of the old suffix filter."""
    source, stage = tmp_path / "source", tmp_path / "stage"
    cursor_example_source(source, declared=False, floating=floating)
    stage.mkdir()
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    modules = ("github-actions", "agent-instructions", "agent-cursor")
    arguments = {
        "template_root": source,
        "staging_root": stage,
        "mappings": mappings,
        "included_modules": modules,
        "summary": lifecycle.materializer.Summary(list(modules), [], "copy"),
    }
    with pytest.raises(
        lifecycle.materializer.MaterializationError,
        match="full SHA" if floating else "Retained workflow example inventory",
    ):
        lifecycle.materializer.write_staged_candidate(**arguments)
    assert snapshot(stage) == {}
    code = (ROOT / ".template-sync/scripts/materialize_downstream_adoption.py").read_text(
        encoding="utf-8"
    )
    guard = 'if PurePosixPath(path).suffix not in {".md", ".mdc"}:'
    assert code.count(guard) == 1
    mutant_path = tmp_path / "cursor_discovery_mutant.py"
    mutant_path.write_text(
        code.replace(guard, 'if PurePosixPath(path).suffix != ".md":'), encoding="utf-8"
    )
    spec = importlib.util.spec_from_file_location("cursor_discovery_mutant", mutant_path)
    assert spec is not None and spec.loader is not None
    mutant = importlib.util.module_from_spec(spec)
    monkeypatch.setitem(sys.modules, mutant.__name__, mutant)
    spec.loader.exec_module(mutant)
    monkeypatch.setattr(mutant, "TRUSTED_TOOL_ROOT", ROOT)
    arguments["summary"] = mutant.Summary(list(modules), [], "copy")
    mutant.write_staged_candidate(**arguments)
    assert (stage / CURSOR_EXAMPLE).is_file()
    assert CURSOR_EXAMPLE not in policy.load_contract(stage)["examples"]
    assert policy.validate_repository(stage) >= 1


@pytest.mark.parametrize("omitted", ["agent-cursor", "agent-instructions"])
@pytest.mark.parametrize("declared", [False, True])
def test_cursor_example_exclusion_and_source_validation(
    tmp_path: Path, omitted: str, declared: bool
) -> None:
    """Omitted agents stay inactive; declared source examples remain checked first."""
    source, stage = tmp_path / "source", tmp_path / "stage"
    cursor_example_source(source, declared=declared)
    modules = {"github-actions", "agent-instructions", "agent-cursor"} - {omitted}
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    lifecycle.materializer.write_staged_candidate(
        template_root=source,
        staging_root=stage,
        mappings=mappings,
        included_modules=modules,
        summary=lifecycle.materializer.Summary(sorted(modules), [], "copy"),
    )
    assert not (stage / CURSOR_EXAMPLE).exists()
    assert CURSOR_EXAMPLE not in policy.load_contract(stage)["examples"]
    assert policy.validate_repository(stage) >= 1
    path = source / CURSOR_EXAMPLE
    path.write_text(path.read_text(encoding="utf-8").replace("a" * 40, "v7"), encoding="utf-8")
    if declared:
        with pytest.raises(lifecycle.materializer.MaterializationError, match="full SHA"):
            lifecycle.materializer.render_workflow_contract(source, mappings, modules)
    else:
        assert lifecycle.materializer.render_workflow_contract(source, mappings, modules)


def test_cursor_example_actual_adoption_and_standalone_failures(tmp_path: Path) -> None:
    """Reviewed Cursor examples survive real adoption, no-op and sync-free validation."""
    source, target = tmp_path / "source", tmp_path / "target"
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    paths, skipped = lifecycle.materializer.iter_safe_repository_files(ROOT)
    assert not skipped
    for relative in paths:
        if lifecycle.materializer.selected_relation_for_path(relative, mappings) is not None:
            destination = source / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(ROOT / relative, destination)
    cursor_example_source(source, declared=True)
    target.mkdir()
    modules = ("github-actions", "agent-instructions", "agent-cursor", "yaml")
    write_decisions(
        target,
        modules,
        protected_file_decisions=lifecycle.protected_take_decisions_for_modules(modules),
    )
    command = [
        sys.executable,
        str(ROOT / ".template-sync/scripts/materialize_downstream_adoption.py"),
        "--template-root",
        str(source),
        "--target-root",
        str(target),
        "--decisions-file",
        "decisions.yml",
    ]
    adopted = subprocess.run(command, check=False, capture_output=True, text=True, timeout=60)
    assert adopted.returncode == 0, adopted.stdout + adopted.stderr
    assert not (target / ".template-sync").exists()
    assert policy.load_contract(target)["examples"] == [YAML_EXAMPLE, CURSOR_EXAMPLE]
    assert (target / CURSOR_EXAMPLE).read_text(encoding="utf-8").startswith("---\n")
    validate = [
        sys.executable,
        str(target / ".github/scripts/validate_workflow_security.py"),
        "--repo-root",
        str(target),
    ]
    accepted = subprocess.run(validate, check=False, capture_output=True, text=True, timeout=20)
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    before = snapshot(target)
    repeated = subprocess.run(command, check=False, capture_output=True, text=True, timeout=60)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert snapshot(target) == before
    path = target / CURSOR_EXAMPLE
    original = path.read_text(encoding="utf-8")
    path.write_text(original.replace("a" * 40, "v7"), encoding="utf-8")
    rejected = subprocess.run(validate, check=False, capture_output=True, text=True, timeout=20)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "full SHA" in rejected.stderr
    path.unlink()
    missing = subprocess.run(validate, check=False, capture_output=True, text=True, timeout=20)
    assert missing.returncode == 1, missing.stdout + missing.stderr
    path.write_text(original, encoding="utf-8", newline="\n")
    remaining = tuple(module for module in modules if module != "agent-cursor")
    decisions = lifecycle.protected_take_decisions_for_modules(remaining)
    decisions.append(
        {
            "path": CURSOR_EXAMPLE,
            "decision": "REMOVE-LOCAL",
            "authorization_basis": "Fixture owner approves later Cursor cleanup",
            "authorized_scope": CURSOR_EXAMPLE,
            "reason": "Cursor module omitted",
        }
    )
    write_decisions(target, remaining, protected_file_decisions=decisions)
    removed = subprocess.run(command, check=False, capture_output=True, text=True, timeout=60)
    assert removed.returncode == 0, removed.stdout + removed.stderr
    assert policy.load_contract(target)["examples"] == [YAML_EXAMPLE]
    # Exclusion updates the contract; reviewed protected cleanup remains explicit.
    assert path.read_text(encoding="utf-8") == original
    checked = subprocess.run(validate, check=False, capture_output=True, text=True, timeout=20)
    assert checked.returncode == 0, checked.stdout + checked.stderr


@pytest.mark.parametrize(
    ("modules", "expected_examples"),
    [
        (("github-actions",), ()),
        (("baseline", "github-actions", "github-platform"), ()),
        (("baseline", "github-actions", "markdown"), ()),
        (("github-actions", "python"), ()),
        (("github-actions", "powershell", "terraform"), TERRAFORM_EXAMPLES),
        (
            (
                "baseline",
                "github-actions",
                "template-sync-support",
                "agent-instructions",
                "yaml",
            ),
            (YAML_EXAMPLE,),
        ),
        (("github-actions", "template-onboarding"), (ONBOARDING_EXAMPLE,)),
    ],
)
def test_retained_profiles_pass_without_unrelated_stacks(
    tmp_path: Path,
    modules: tuple[str, ...],
    expected_examples: tuple[str, ...],
) -> None:
    """Selected language workflows and policy runtime survive real materialization."""
    target = lifecycle.materialize_module_fixture(tmp_path, modules, authorize_protected_files=True)
    assert policy.validate_repository(target) >= 1
    contract = policy.load_contract(target)
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    expected = {
        path
        for file in (ROOT / ".github/workflows").iterdir()
        if file.suffix in {".yml", ".yaml"}
        and (path := file.relative_to(ROOT).as_posix())
        and (relation := lifecycle.materializer.selected_relation_for_path(path, mappings))
        is not None
        and relation.is_retained_by(modules)
    }
    assert set(contract["workflows"]) == expected
    assert contract["examples"] == list(expected_examples)
    for module, filename in [
        ("markdown", "markdownlint"),
        ("python", "python-ci"),
        ("powershell", "powershell-ci"),
        ("terraform", "terraform-ci"),
    ]:
        assert (target / f".github/workflows/{filename}.yml").exists() == (module in modules)
    assert (target / "package.json").exists() == ("markdown" in modules)
    if "baseline" in modules:
        config = yaml.safe_load((target / ".pre-commit-config.yaml").read_text(encoding="utf-8"))
        assert_hook(config, "validate-workflow-security")
        data_ci = yaml.safe_load(
            (target / ".github/workflows/data-ci.yml").read_text(encoding="utf-8")
        )
        assert_ci(data_ci, "validate-workflow-security")
        if "template-sync-support" in modules:
            for mode in ("upstream", "downstream"):
                assert_hook(config, f"validate-instruction-contracts-{mode}")
                assert_ci(data_ci, f"validate-instruction-contracts-{mode}")
    if "github-platform" in modules:
        assert "github-actions" in lifecycle.dependabot_update_ecosystems(
            target / ".github/dependabot.yml"
        )


def snapshot(root: Path) -> dict[str, bytes]:
    """Capture downstream files without transient caches."""
    return {
        p.relative_to(root).as_posix(): p.read_bytes()
        for p in root.rglob("*")
        if p.is_file() and "__pycache__" not in p.parts and ".git" not in p.parts
    }


@pytest.mark.parametrize("mutation", ["entry", "contract", "unmapped", "empty"])
def test_retained_workflow_inventory_fails_before_staging(tmp_path: Path, mutation: str) -> None:
    """A retained manifest file cannot escape its contract or leave partial staged output."""
    from tests.test_workflow_security_contract import copy_policy

    source, stage = tmp_path / "source", tmp_path / "stage"
    contract = copy_policy(source)
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    modules = ("baseline", "github-actions")
    if mutation == "entry":
        omitted = ".github/workflows/auto-fix-precommit.yml"
        del contract["workflows"][omitted]
        (source / policy.CONTRACT).write_text(yaml.safe_dump(contract), encoding="utf-8")
        (source / omitted).write_text(
            "on: pull_request_target\npermissions: write-all\njobs: {}\n", encoding="utf-8"
        )
    elif mutation == "contract":
        (source / policy.CONTRACT).unlink()
    elif mutation == "unmapped":
        mappings = tuple(mapping for mapping in mappings if mapping.pattern != policy.CONTRACT)
    else:
        for path in (source / ".github/workflows").iterdir():
            path.unlink()
    stage.mkdir()
    with pytest.raises(
        lifecycle.materializer.MaterializationError, match="Retained workflow inventory"
    ):
        lifecycle.materializer.write_staged_candidate(
            template_root=source,
            staging_root=stage,
            mappings=mappings,
            included_modules=modules,
            summary=lifecycle.materializer.Summary(list(modules), [], "copy"),
        )
    assert snapshot(stage) == {}


@pytest.mark.parametrize(
    ("floating", "message"),
    [(False, "Retained workflow example inventory"), (True, "full SHA")],
)
def test_retained_example_inventory_fails_before_staging(
    tmp_path: Path, floating: bool, message: str
) -> None:
    """A retained governed document cannot escape the rendered contract."""
    from tests.test_workflow_security_contract import copy_policy

    source, stage = tmp_path / "source", tmp_path / "stage"
    contract = copy_policy(source)
    contract["examples"].remove(YAML_EXAMPLE)
    (source / policy.CONTRACT).write_text(
        yaml.safe_dump(contract, sort_keys=False), encoding="utf-8"
    )
    if floating:
        with (source / YAML_EXAMPLE).open("a", encoding="utf-8") as stream:
            stream.write("\n```yaml\n- uses: actions/checkout@v7 # v7.0.1\n```\n")
    stage.mkdir()
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    with pytest.raises(lifecycle.materializer.MaterializationError, match=message):
        lifecycle.materializer.write_staged_candidate(
            template_root=source,
            staging_root=stage,
            mappings=mappings,
            included_modules=("github-actions", "agent-instructions", "yaml"),
            summary=lifecycle.materializer.Summary(
                ["github-actions", "agent-instructions", "yaml"], [], "copy"
            ),
        )
    assert snapshot(stage) == {}


def test_stale_retained_example_entry_fails_before_staging(tmp_path: Path) -> None:
    """A declared retained document cannot overstate governed example coverage."""
    from tests.test_workflow_security_contract import copy_policy

    source, stage = tmp_path / "source", tmp_path / "stage"
    copy_policy(source)
    (source / YAML_EXAMPLE).write_text(
        "# Fixture document with no governed action references.\n", encoding="utf-8"
    )
    stage.mkdir()
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    with pytest.raises(
        lifecycle.materializer.MaterializationError,
        match=r"unexpected: \.github/instructions/yaml\.instructions\.md",
    ):
        lifecycle.materializer.write_staged_candidate(
            template_root=source,
            staging_root=stage,
            mappings=mappings,
            included_modules=("github-actions", "agent-instructions", "yaml"),
            summary=lifecycle.materializer.Summary(
                ["github-actions", "agent-instructions", "yaml"], [], "copy"
            ),
        )
    assert snapshot(stage) == {}


@pytest.mark.parametrize("guard", ["comparison", "semantic-discovery"])
def test_example_inventory_guard_removal(tmp_path: Path, monkeypatch: Any, guard: str) -> None:
    """Independent mutants restore safe and floating omitted-example bypasses."""
    from tests.test_workflow_security_contract import copy_policy

    source, stage = tmp_path / "source", tmp_path / "stage"
    contract = copy_policy(source)
    contract["examples"].remove(YAML_EXAMPLE)
    (source / policy.CONTRACT).write_text(
        yaml.safe_dump(contract, sort_keys=False), encoding="utf-8"
    )
    if guard == "semantic-discovery":
        with (source / YAML_EXAMPLE).open("a", encoding="utf-8") as stream:
            stream.write("\n```yaml\n- uses: actions/checkout@v7 # v7.0.1\n```\n")
        message = "full SHA"
        predicate = "if validator.check_examples(text):"
    else:
        message = "Retained workflow example inventory"
        predicate = "if set(examples) != expected_examples:"
    modules = ("github-actions", "agent-instructions", "yaml")
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    arguments = {
        "template_root": source,
        "staging_root": stage,
        "mappings": mappings,
        "included_modules": modules,
        "summary": lifecycle.materializer.Summary(list(modules), [], "copy"),
    }
    stage.mkdir()
    with pytest.raises(lifecycle.materializer.MaterializationError, match=message):
        lifecycle.materializer.write_staged_candidate(**arguments)
    assert snapshot(stage) == {}

    code = (ROOT / ".template-sync/scripts/materialize_downstream_adoption.py").read_text(
        encoding="utf-8"
    )
    assert code.count(predicate) == 1
    mutant_path = tmp_path / "example_inventory_mutant.py"
    mutant_path.write_text(code.replace(predicate, "if False:"), encoding="utf-8")
    spec = importlib.util.spec_from_file_location("example_inventory_mutant", mutant_path)
    assert spec is not None and spec.loader is not None
    mutant = importlib.util.module_from_spec(spec)
    monkeypatch.setitem(sys.modules, mutant.__name__, mutant)
    spec.loader.exec_module(mutant)
    monkeypatch.setattr(mutant, "TRUSTED_TOOL_ROOT", ROOT)
    arguments["summary"] = mutant.Summary(list(modules), [], "copy")
    mutant.write_staged_candidate(**arguments)
    rendered = policy.load_contract(stage)
    assert YAML_EXAMPLE not in rendered["examples"]
    assert (stage / YAML_EXAMPLE).is_file()
    if guard == "semantic-discovery":
        assert "actions/checkout@v7" in (stage / YAML_EXAMPLE).read_text(encoding="utf-8")
    assert policy.validate_repository(stage) >= 1


@pytest.mark.parametrize("explicit", [False, True])
def test_direct_renderer_rejects_empty_inventory(explicit: bool) -> None:
    """Both direct-call paths reject a contract that would violate its own schema."""
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    arguments: dict[str, set[str]] = {"expected_workflows": set()} if explicit else {}
    with pytest.raises(
        lifecycle.materializer.MaterializationError, match="without retained workflows"
    ):
        lifecycle.materializer.render_workflow_contract(
            ROOT, mappings, ("template-sync-support",), **arguments
        )


def test_excluded_workflow_and_unmapped_source_do_not_expand_inventory(tmp_path: Path) -> None:
    """The source manifest selects ownership; omitted modules and local workflows stay outside."""
    from tests.test_workflow_security_contract import copy_policy

    source, stage = tmp_path / "source", tmp_path / "stage"
    contract = copy_policy(source)
    omitted = ".github/workflows/auto-fix-precommit.yml"
    del contract["workflows"][omitted]
    (source / policy.CONTRACT).write_text(yaml.safe_dump(contract), encoding="utf-8")
    (source / ".github/workflows/local.yml").write_text(
        "on: pull_request_target\njobs: {}", encoding="utf-8"
    )
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    lifecycle.materializer.write_staged_candidate(
        template_root=source,
        staging_root=stage,
        mappings=mappings,
        included_modules=("github-actions",),
        summary=lifecycle.materializer.Summary(["github-actions"], [], "copy"),
    )
    assert policy.validate_repository(stage) == 1
    assert not (stage / omitted).exists()
    assert not (stage / ".github/workflows/local.yml").exists()


@pytest.mark.parametrize("guard", ["entry", "contract"])
def test_workflow_inventory_guard_removal(tmp_path: Path, monkeypatch: Any, guard: str) -> None:
    """Independent pre-stage mutants restore omitted-contract and omitted-entry bypasses."""
    from tests.test_workflow_security_contract import copy_policy

    source, stage = tmp_path / "source", tmp_path / "stage"
    contract = copy_policy(source)
    omitted = ".github/workflows/auto-fix-precommit.yml"
    if guard == "entry":
        del contract["workflows"][omitted]
        (source / policy.CONTRACT).write_text(yaml.safe_dump(contract), encoding="utf-8")
        (source / omitted).write_text("on: pull_request_target\njobs: {}\n", encoding="utf-8")
    else:
        (source / policy.CONTRACT).unlink()
    script = ROOT / ".template-sync/scripts/materialize_downstream_adoption.py"
    code = script.read_text(encoding="utf-8")
    predicate = (
        "if set(rendered) != expected_workflows:"
        if guard == "entry"
        else "if expected_workflows or contract_retained:"
    )
    assert code.count(predicate) == 1
    mutant_path = tmp_path / "inventory_mutant.py"
    mutant_path.write_text(code.replace(predicate, "if False:"), encoding="utf-8")
    spec = importlib.util.spec_from_file_location("inventory_mutant", mutant_path)
    assert spec is not None and spec.loader is not None
    mutant = importlib.util.module_from_spec(spec)
    monkeypatch.setitem(sys.modules, mutant.__name__, mutant)
    spec.loader.exec_module(mutant)
    monkeypatch.setattr(mutant, "TRUSTED_TOOL_ROOT", ROOT)
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    arguments = {
        "template_root": source,
        "staging_root": stage,
        "mappings": mappings,
        "included_modules": ("baseline", "github-actions"),
        "summary": mutant.Summary(["baseline", "github-actions"], [], "copy"),
    }
    with pytest.raises(
        lifecycle.materializer.MaterializationError, match="Retained workflow inventory"
    ):
        lifecycle.materializer.write_staged_candidate(**arguments)
    mutant.write_staged_candidate(**arguments)
    if guard == "entry":
        assert (stage / omitted).read_bytes() == (source / omitted).read_bytes()
        assert policy.validate_repository(stage) == 4
    else:
        assert (stage / omitted).is_file()
        assert not (stage / policy.CONTRACT).exists()


def test_selected_validator_is_inert_and_trust_anchor_failure_is_detected(
    tmp_path: Path, monkeypatch: Any
) -> None:
    """Selected Python bytes never execute; restoring that trust boundary trips a marker."""
    from tests.test_workflow_security_contract import copy_policy

    source = tmp_path / "source"
    copy_policy(source)
    sentinel = source / "candidate-executed.txt"
    selected_script = source / ".github/scripts/validate_workflow_security.py"
    selected_script.parent.mkdir(parents=True, exist_ok=True)
    selected_script.write_text(
        "from pathlib import Path\n"
        "Path(__file__).parents[2].joinpath('candidate-executed.txt').write_text('executed')\n"
        "raise RuntimeError('candidate execution sentinel')\n",
        encoding="utf-8",
    )
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    rendered = lifecycle.materializer.render_workflow_contract(
        source, mappings, ("github-actions",)
    )
    assert ".github/workflows/workflow-security.yml" in rendered
    assert not sentinel.exists()
    monkeypatch.setattr(lifecycle.materializer, "TRUSTED_TOOL_ROOT", source)
    with pytest.raises(RuntimeError, match="candidate execution sentinel"):
        lifecycle.materializer.render_workflow_contract(source, mappings, ("github-actions",))
    assert sentinel.read_text(encoding="utf-8") == "executed"


def test_installed_support_tools_can_add_actions_and_require_shared_helpers(tmp_path: Path) -> None:
    """A support-only installation carries its trusted executable and data closure."""
    target = lifecycle.materialize_module_fixture(
        tmp_path, ("template-sync-support",), authorize_protected_files=True
    )
    shared = [
        ".github/scripts/validate_workflow_security.py",
        ".github/scripts/replace-template-placeholders.py",
        ".github/template-placeholders.json",
        "schemas/template-placeholders.schema.json",
        policy.SCHEMA,
    ]
    for relative in shared:
        assert (target / relative).is_file(), relative
    assert not (target / policy.CONTRACT).exists()
    assert (target / policy.SCHEMA).is_file()
    modules = ("template-sync-support", "github-actions")
    write_decisions(
        target,
        modules,
        protected_file_decisions=lifecycle.protected_take_decisions_for_modules(modules),
        local_overrides=[
            {
                "path": path,
                "default_decision": "TAKE",
                "reason": "Fixture owner approves added Actions wiring.",
            }
            for path in ("TEMPLATE_UPDATE_PROCEDURE.md", "tests/test_contract_wiring.py")
        ],
    )
    command = [
        sys.executable,
        str(target / ".template-sync/scripts/materialize_downstream_adoption.py"),
        "--template-root",
        str(ROOT),
        "--target-root",
        str(target),
        "--decisions-file",
        "decisions.yml",
        "--repository",
        "octo/widget",
        "--security-contact",
        "security@example.com",
    ]
    result = subprocess.run(command, capture_output=True, text=True, check=False)
    assert result.returncode == 0, result.stdout + result.stderr
    assert policy.validate_repository(target) == 1
    installed = subprocess.run(
        [sys.executable, str(target / shared[0]), "--repo-root", str(target)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert installed.returncode == 0, installed.stdout + installed.stderr
    for relative in [*shared[:2], policy.SCHEMA]:
        helper = target / relative
        original = helper.read_bytes()
        helper.unlink()
        try:
            failed = subprocess.run(command, capture_output=True, text=True, check=False)
            assert failed.returncode == 1, failed.stdout + failed.stderr
            assert "trusted helper is unavailable" in failed.stderr
            assert relative in failed.stderr
        finally:
            helper.write_bytes(original)


@pytest.mark.parametrize("mutation", ["entry", "contract"])
def test_installed_materializer_rejects_incomplete_source_inventory(
    tmp_path: Path, mutation: str
) -> None:
    """Installed support tools reject missing ownership before changing an adopter tree."""
    target = lifecycle.materialize_module_fixture(
        tmp_path, ("template-sync-support",), authorize_protected_files=True
    )
    source = tmp_path / "source"
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    paths, _ = lifecycle.materializer.iter_safe_repository_files(ROOT)
    for relative in paths:
        if lifecycle.materializer.selected_relation_for_path(relative, mappings) is not None:
            destination = source / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(ROOT / relative, destination)
    if mutation == "entry":
        contract = policy.load_contract(source)
        del contract["workflows"][".github/workflows/workflow-security.yml"]
        (source / policy.CONTRACT).write_text(yaml.safe_dump(contract), encoding="utf-8")
    else:
        (source / policy.CONTRACT).unlink()
    modules = ("template-sync-support", "github-actions")
    write_decisions(
        target,
        modules,
        protected_file_decisions=lifecycle.protected_take_decisions_for_modules(modules),
    )
    before = snapshot(target)
    result = subprocess.run(
        [
            sys.executable,
            str(target / ".template-sync/scripts/materialize_downstream_adoption.py"),
            "--template-root",
            str(source),
            "--target-root",
            str(target),
            "--decisions-file",
            "decisions.yml",
            "--repository",
            "octo/widget",
            "--security-contact",
            "security@example.com",
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Retained workflow inventory" in result.stderr
    assert "Traceback" not in result.stderr
    assert snapshot(target) == before


@pytest.mark.parametrize(
    "example",
    [
        "- uses : owner/action@v1 # v1.2.3",
        '- {name: Checkout,\n "u\\u0073es": "owner\\/action\\u0040v1", # v1.2.3\n with: {}}',
    ],
)
def test_installed_validator_checks_semantic_document_examples(
    tmp_path: Path, example: str
) -> None:
    """The retained helper and governed example inventory enforce decoded YAML downstream."""
    target = lifecycle.materialize_module_fixture(
        tmp_path, ("github-actions", "yaml", "agent-instructions"), authorize_protected_files=True
    )
    script = target / ".github/scripts/validate_workflow_security.py"
    command = [sys.executable, str(script), "--repo-root", str(target)]
    positive = subprocess.run(command, capture_output=True, text=True, check=False)
    assert positive.returncode == 0, positive.stdout + positive.stderr
    path = target / ".github/instructions/yaml.instructions.md"
    with path.open("a", encoding="utf-8") as stream:
        stream.write("\n```yaml\n" + example + "\n```\n")
    negative = subprocess.run(command, capture_output=True, text=True, check=False)
    assert negative.returncode == 1, negative.stdout + negative.stderr
    assert "full SHA" in negative.stderr


@pytest.mark.parametrize("candidate_schema", ["sentinel", "malformed", "missing"])
def test_candidate_workflow_schema_is_inert(tmp_path: Path, candidate_schema: str) -> None:
    """Selected schema bytes cannot change the running renderer's validation authority."""
    from tests.test_workflow_security_contract import copy_policy

    source = tmp_path / "source"
    copy_policy(source)
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    expected = lifecycle.materializer.render_workflow_contract(
        source, mappings, ("github-actions",)
    )
    path = source / policy.SCHEMA
    if candidate_schema == "sentinel":
        schema = json.loads(path.read_text(encoding="utf-8"))
        schema["required"].append("candidate_schema_sentinel")
        path.write_text(json.dumps(schema), encoding="utf-8")
    elif candidate_schema == "malformed":
        path.write_text("not a schema", encoding="utf-8")
    else:
        path.unlink()
    assert (
        lifecycle.materializer.render_workflow_contract(source, mappings, ("github-actions",))
        == expected
    )


def test_installed_workflow_schema_is_authoritative(tmp_path: Path, monkeypatch: Any) -> None:
    """The same harmless schema sentinel is enforced when installed in the trusted bundle."""
    trusted = tmp_path / "trusted"
    for relative in [".github/scripts/validate_workflow_security.py", policy.SCHEMA]:
        path = trusted / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(ROOT / relative, path)
    path = trusted / policy.SCHEMA
    schema = json.loads(path.read_text(encoding="utf-8"))
    schema["required"].append("trusted_schema_sentinel")
    path.write_text(json.dumps(schema), encoding="utf-8")
    monkeypatch.setattr(lifecycle.materializer, "TRUSTED_TOOL_ROOT", trusted)
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    with pytest.raises(
        lifecycle.materializer.MaterializationError, match="trusted_schema_sentinel"
    ):
        lifecycle.materializer.render_workflow_contract(ROOT, mappings, ("github-actions",))


@pytest.mark.parametrize("call", ["validate_repository", "load_contract"])
def test_workflow_schema_authority_guard_removal(
    tmp_path: Path, monkeypatch: Any, call: str
) -> None:
    """Removing either trusted-schema argument restores candidate control independently."""
    from tests.test_workflow_security_contract import copy_policy

    source = tmp_path / "source"
    copy_policy(source)
    schema_path = source / policy.SCHEMA
    schema = json.loads(schema_path.read_text(encoding="utf-8"))
    schema["required"].append("candidate_schema_sentinel")
    schema_path.write_text(json.dumps(schema), encoding="utf-8")
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    assert lifecycle.materializer.render_workflow_contract(source, mappings, ("github-actions",))
    script = ROOT / ".template-sync/scripts/materialize_downstream_adoption.py"
    code = script.read_text(encoding="utf-8")
    before = f"validator.{call}(template_root, schema_root=TRUSTED_TOOL_ROOT)"
    assert code.count(before) == 1
    mutant_path = tmp_path / "materializer_schema_mutant.py"
    mutant_path.write_text(
        code.replace(before, f"validator.{call}(template_root)"), encoding="utf-8"
    )
    spec = importlib.util.spec_from_file_location("materializer_schema_mutant", mutant_path)
    assert spec is not None and spec.loader is not None
    mutant = importlib.util.module_from_spec(spec)
    monkeypatch.setitem(sys.modules, mutant.__name__, mutant)
    spec.loader.exec_module(mutant)
    monkeypatch.setattr(mutant, "TRUSTED_TOOL_ROOT", ROOT)
    with pytest.raises(mutant.MaterializationError, match="candidate_schema_sentinel"):
        mutant.render_workflow_contract(source, mappings, ("github-actions",))


@pytest.mark.parametrize("mode", ["help", "materialize", "render", "eager-import-mutant"])
def test_installed_support_schema_dependency_is_lazy(tmp_path: Path, mode: str) -> None:
    """Help needs no schema runtime, while actual validation fails with a clear diagnostic."""
    target = lifecycle.materialize_module_fixture(
        tmp_path, ("template-sync-support",), authorize_protected_files=True
    )
    assert not (target / "pyproject.toml").exists()
    assert not (target / policy.CONTRACT).exists()
    script = target / ".template-sync/scripts/materialize_downstream_adoption.py"
    if mode == "eager-import-mutant":
        source = script.read_text(encoding="utf-8")
        anchor = "\nimport sys\n"
        assert source.count(anchor) == 1
        script.write_text(source.replace(anchor, "\nimport jsonschema" + anchor), encoding="utf-8")
    runner = (
        "import runpy, sys\n"
        "from pathlib import Path\n"
        "sys.modules['jsonschema'] = None\n"
        "script = sys.argv.pop(1)\n"
    )
    if mode == "render":
        runner += (
            "namespace = runpy.run_path(script)\n"
            "try:\n"
            "    namespace['render_workflow_contract'](Path(sys.argv[1]), (), ('github-actions',))\n"
            "except namespace['MaterializationError'] as error:\n"
            "    print(error, file=sys.stderr)\n"
            "    sys.exit(1)\n"
        )
        arguments = [str(ROOT)]
    else:
        runner += "runpy.run_path(script, run_name='__main__')\n"
        arguments = ["--help"]
        if mode == "materialize":
            output = tmp_path / "output"
            output.mkdir()
            arguments = [
                "--template-root",
                str(ROOT),
                "--target-root",
                str(output),
                "--included-module",
                "template-sync-support",
            ]
    result = subprocess.run(
        [sys.executable, "-c", runner, str(script), *arguments],
        capture_output=True,
        text=True,
        check=False,
    )
    if mode == "help":
        assert result.returncode == 0, result.stderr
        assert "--template-root" in result.stdout
    elif mode == "eager-import-mutant":
        assert result.returncode == 1
        assert "ModuleNotFoundError" in result.stderr
    else:
        assert result.returncode == 1
        expected = (
            "jsonschema is unavailable"
            if mode == "materialize"
            else "Workflow contract rendering requires jsonschema"
        )
        assert expected in result.stderr
        assert "Traceback" not in result.stderr


def write_decisions(target: Path, modules: tuple[str, ...], **fields: Any) -> None:
    """Record explicit fixture authorization using the normal marker schema."""
    reviewed = fields.pop("reviewed_commit", None)
    document = lifecycle.marker_document(list(modules), **fields)
    if reviewed:
        document["template_sync"]["last_reviewed_template_commit"] = reviewed
    lifecycle.write_yaml(target / "decisions.yml", document)


def test_existing_adoption_noop_customization_and_protection(tmp_path: Path) -> None:
    """First adoption, repeat sync, and protected conflicts preserve adopter content."""
    target = tmp_path / "existing"
    local = target / ".github/workflows/custom.yml"
    local.parent.mkdir(parents=True)
    local.write_text("# Adopter content\non: pull_request_target\njobs: {}\n", encoding="utf-8")
    modules = ("github-actions",)
    write_decisions(target, modules)
    blocked = lifecycle.run_materialize(ROOT, target, "--decisions-file", "decisions.yml")
    assert blocked.returncode == 2, blocked.stdout + blocked.stderr
    assert not (target / policy.CONTRACT).exists()
    write_decisions(
        target,
        modules,
        protected_file_decisions=lifecycle.protected_take_decisions_for_modules(modules),
    )
    adopted = lifecycle.run_materialize(ROOT, target, "--decisions-file", "decisions.yml")
    assert adopted.returncode == 0, adopted.stdout + adopted.stderr
    assert policy.validate_repository(target) == 1
    before = snapshot(target)
    repeated = lifecycle.run_materialize(ROOT, target, "--decisions-file", "decisions.yml")
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert snapshot(target) == before
    contract = target / policy.CONTRACT
    contract.write_text(
        contract.read_text(encoding="utf-8") + "# Local policy customization\n", encoding="utf-8"
    )
    write_decisions(target, modules)
    tampered = contract.read_bytes()
    blocked = lifecycle.run_materialize(ROOT, target, "--decisions-file", "decisions.yml")
    assert blocked.returncode == 2, blocked.stdout + blocked.stderr
    assert contract.read_bytes() == tampered
    assert local.read_bytes() == before[local.relative_to(target).as_posix()]
    contract.unlink()
    assert policy.main(["--repo-root", str(target)]) == 1


@pytest.mark.parametrize("change", ["command", "job-name", "concurrency"])
def test_template_update_requires_review_and_preserves_adopter_workflow(
    tmp_path: Path, change: str
) -> None:
    """A reviewed template revision updates required behavior only with path decisions."""
    source = tmp_path / "source"
    source.mkdir()
    # Copy only manifest-owned files, excluding dependencies, scratch, and Git internals.
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    paths, _ = lifecycle.materializer.iter_safe_repository_files(ROOT)
    for relative in paths:
        if lifecycle.materializer.selected_relation_for_path(relative, mappings) is not None:
            destination = source / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(ROOT / relative, destination)
    target = tmp_path / "target"
    target.mkdir()
    modules = ("github-actions", "template-sync-support")
    revision_a = lifecycle.commit_fixture_template(source)
    decisions = lifecycle.protected_take_decisions_for_modules(modules)
    write_decisions(target, modules, protected_file_decisions=decisions, reviewed_commit=revision_a)
    assert (
        lifecycle.run_materialize(source, target, "--decisions-file", "decisions.yml").returncode
        == 0
    )
    marker_path = target / ".template-sync/marker.yml"
    assert (
        yaml.safe_load(marker_path.read_text(encoding="utf-8"))["template_sync"][
            "last_reviewed_template_commit"
        ]
        == revision_a
    )
    local = target / ".github/workflows/local.yml"
    local_content = "on: push\npermissions: {}\njobs:\n  local:\n    permissions: {}\n    runs-on: ubuntu-latest\n    steps:\n      - run: echo local\n"
    local.write_text(local_content, encoding="utf-8")
    workflow_path = ".github/workflows/workflow-security.yml"
    workflow = source / workflow_path
    original_control, updated_control = {
        "command": (
            "run: python .github/scripts/validate_workflow_security.py",
            "run: python .github/scripts/validate_workflow_security.py --strict",
        ),
        "job-name": ("    name: Workflow Security\n", "    name: Reviewed workflow check\n"),
        "concurrency": ("jobs:\n", "concurrency: reviewed-${{ github.ref }}\n\njobs:\n"),
    }[change]
    workflow_text = workflow.read_text(encoding="utf-8")
    assert workflow_text.count(original_control) == 1
    workflow.write_text(workflow_text.replace(original_control, updated_control), encoding="utf-8")
    # Unreviewed source drift must not be blessed by rendering new fingerprints.
    rejected = lifecycle.run_materialize(source, target, "--decisions-file", "decisions.yml")
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    contract = policy.load_contract(source)
    contract["workflows"][workflow_path] = policy.validate_workflow(
        workflow.read_text(encoding="utf-8")
    )
    (source / policy.CONTRACT).write_text(
        yaml.safe_dump(contract, sort_keys=False), encoding="utf-8"
    )
    lifecycle.run_git(source, "add", ".")
    lifecycle.run_git(
        source,
        "-c",
        "user.name=Template Tester",
        "-c",
        "user.email=template@example.com",
        "commit",
        "-q",
        "-m",
        "Reviewed workflow policy update",
    )
    revision_b = lifecycle.run_git(source, "rev-parse", "HEAD").stdout.strip()
    assert revision_b != revision_a
    write_decisions(target, modules, reviewed_commit=revision_b)
    blocked = lifecycle.run_materialize(source, target, "--decisions-file", "decisions.yml")
    assert blocked.returncode == 2, blocked.stdout + blocked.stderr
    assert (
        yaml.safe_load(marker_path.read_text(encoding="utf-8"))["template_sync"][
            "last_reviewed_template_commit"
        ]
        == revision_a
    )
    write_decisions(
        target,
        modules,
        protected_file_decisions=decisions,
        reviewed_commit=revision_b,
        local_overrides=[
            {
                "path": workflow_path,
                "default_decision": "TAKE",
                "reason": "Fixture owner reviewed the workflow control change.",
            }
        ],
    )
    updated = lifecycle.run_materialize(source, target, "--decisions-file", "decisions.yml")
    assert updated.returncode == 0, updated.stdout + updated.stderr
    assert updated_control in (target / workflow_path).read_text(encoding="utf-8")
    assert policy.validate_repository(target, strict=True) == 1
    installed = policy.parse_yaml((target / workflow_path).read_text(encoding="utf-8"))
    gate = installed["jobs"]["validate"]["steps"][-1]["run"]
    command = shlex.split(gate)
    executed = subprocess.run(
        [sys.executable, *command[1:]], cwd=target, capture_output=True, text=True, check=False
    )
    assert executed.returncode == 0, executed.stdout + executed.stderr
    assert local.read_text(encoding="utf-8") == local_content
    marker = yaml.safe_load(marker_path.read_text(encoding="utf-8"))["template_sync"]
    assert marker["last_reviewed_template_commit"] == revision_b
    assert any(item["path"] == policy.CONTRACT for item in marker["protected_file_decisions"])
    before = snapshot(target)
    repeated = lifecycle.run_materialize(source, target, "--decisions-file", "decisions.yml")
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert snapshot(target) == before


def test_omission_and_reviewed_removal_leave_no_policy_orphans(tmp_path: Path) -> None:
    """Removal preserves old files pending explicit cleanup, then matches clean omission."""
    retained = ("baseline", "github-actions", "github-platform", "template-sync-support")
    omitted = ("baseline", "github-platform", "template-sync-support")
    (tmp_path / "retained").mkdir()
    (tmp_path / "omitted").mkdir()
    target = lifecycle.materialize_module_fixture(
        tmp_path / "retained", retained, authorize_protected_files=True
    )
    clean = lifecycle.materialize_module_fixture(
        tmp_path / "omitted", omitted, authorize_protected_files=True
    )
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    original = snapshot(target)
    overrides = []
    removals = []
    for relative in original:
        relation = lifecycle.materializer.selected_relation_for_path(relative, mappings)
        if relation is None:
            continue
        if relation.is_retained_by(omitted):
            overrides.append(
                {
                    "path": relative,
                    "default_decision": "TAKE",
                    "reason": "Fixture owner approved module removal cleanup.",
                }
            )
        else:
            removals.append(relative)
    write_decisions(
        target,
        omitted,
        local_overrides=overrides,
        protected_file_decisions=lifecycle.protected_take_decisions_for_modules(omitted),
    )
    result = lifecycle.run_materialize(ROOT, target, "--decisions-file", "decisions.yml")
    assert result.returncode == 0, result.stdout + result.stderr
    assert (
        target / policy.CONTRACT
    ).exists(), "Materialization must not silently delete excluded files"
    lifecycle.run_git(target, "init", "-q")
    lifecycle.run_git(target, "add", ".")
    before_cleanup = lifecycle.run_downstream_adoption_validator(target)
    assert before_cleanup.returncode == 1, before_cleanup.stdout + before_cleanup.stderr
    assert policy.CONTRACT in before_cleanup.stdout
    for relative in removals:
        (target / relative).unlink()
    lifecycle.run_git(target, "add", "-A")
    lifecycle.run_git(clean, "init", "-q")
    lifecycle.run_git(clean, "add", ".")
    for tree in (target, clean):
        lifecycle.finish_review_profile_link_cleanup(tree, "neither")
    after_cleanup = lifecycle.run_downstream_adoption_validator(target)
    assert after_cleanup.returncode == 0, after_cleanup.stdout + after_cleanup.stderr
    ignored = {"decisions.yml", ".template-sync/marker.yml"}
    assert {p: b for p, b in snapshot(target).items() if p not in ignored} == {
        p: b for p, b in snapshot(clean).items() if p not in ignored
    }
    assert not (target / policy.CONTRACT).exists()
    assert (target / policy.SCHEMA).is_file()
    assert (target / ".github/scripts/validate_workflow_security.py").is_file()
    for relative in [
        ".github/workflows/workflow-security.yml",
        "docs/workflow-security.md",
        "schemas/examples/workflow-security-contract/valid/contract.yml",
        "schemas/examples/workflow-security-contract/invalid/contract.yml",
        "tests/test_workflow_security_contract.py",
        "tests/test_workflow_security_lifecycle.py",
        ".github/workflows/data-ci.yml",
    ]:
        assert not (target / relative).exists(), relative
    assert "validate-workflow-security" not in (target / ".pre-commit-config.yaml").read_text(
        encoding="utf-8"
    )
    assert "github-actions" not in lifecycle.dependabot_update_ecosystems(
        target / ".github/dependabot.yml"
    )
    assert "workflow-security" not in (target / "tests/test_contract_wiring.py").read_text(
        encoding="utf-8"
    )


@pytest.mark.parametrize("relative", ["docs/workflow-security.md", CURSOR_EXAMPLE])
def test_literal_fence_materializer_inventory(tmp_path: Path, relative: str) -> None:
    """Opaque Markdown and Cursor samples never acquire governed inventory entries."""
    from tests.test_workflow_security_contract import copy_policy

    source = tmp_path / "source"
    contract = copy_policy(source)
    path = source / relative
    path.parent.mkdir(parents=True, exist_ok=True)
    literal = "```text\nuses: owner/action@v1\n```\n"
    path.write_text(literal, encoding="utf-8")
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    modules = {"github-actions", "agent-instructions", "agent-cursor", "yaml"}

    def stage(name: str) -> Path:
        destination = tmp_path / name
        lifecycle.materializer.write_staged_candidate(
            template_root=source,
            staging_root=destination,
            mappings=mappings,
            included_modules=modules,
            summary=lifecycle.materializer.Summary(sorted(modules), [], "copy"),
        )
        return destination

    first = stage("literal")
    assert relative not in policy.load_contract(first)["examples"]
    assert (first / relative).read_text(encoding="utf-8") == literal
    assert snapshot(stage("repeat")) == snapshot(first)
    pinned = f"```yaml\nuses: owner/action@{'a' * 40} # v1.2.3\n```\n"
    path.write_text(literal + pinned, encoding="utf-8")
    with pytest.raises(
        lifecycle.materializer.MaterializationError, match="Retained workflow example"
    ):
        stage("undeclared")
    assert snapshot(tmp_path / "undeclared") == {}
    contract["examples"].append(relative)
    lifecycle.write_yaml(source / policy.CONTRACT, contract)
    mixed = stage("mixed")
    assert relative in policy.load_contract(mixed)["examples"]
    assert policy.validate_repository(mixed) >= 1
    path.write_text(literal, encoding="utf-8")
    with pytest.raises(
        lifecycle.materializer.MaterializationError, match="Retained workflow example"
    ):
        stage("stale")
    assert snapshot(tmp_path / "stale") == {}
    # Source validation remains authoritative even if this document is omitted.
    path.write_text(pinned.replace("a" * 40, "v1"), encoding="utf-8")
    with pytest.raises(lifecycle.materializer.MaterializationError, match="full SHA"):
        lifecycle.materializer.render_workflow_contract(source, mappings, {"github-actions"})


def test_actions_only_literal_examples_and_shared_helper_closure(tmp_path: Path) -> None:
    """An isolated deployed CLI needs only Actions assets and fails if its helper vanishes."""
    target = lifecycle.materialize_module_fixture(
        tmp_path, ("github-actions",), authorize_protected_files=True
    )
    for omitted in (
        ".template-sync",
        ".pre-commit-config.yaml",
        "pyproject.toml",
        "AGENTS.md",
        ".github/scripts/instruction_contract_core.py",
        ".github/scripts/validate_instruction_profile.py",
        ".github/instruction-profile.yml",
        "schemas/instruction-profile.schema.json",
    ):
        assert not (target / omitted).exists(), omitted
    helper = target / ".github/scripts/instruction_contract_support.py"
    assert helper.is_file()
    script = target / ".github/scripts/validate_workflow_security.py"
    command = [sys.executable, "-E", str(script), "--repo-root", str(target)]

    def invoke(expected: int, message: str = "") -> None:
        result = subprocess.run(command, cwd=tmp_path, text=True, capture_output=True, check=False)
        assert result.returncode == expected, result.stdout + result.stderr
        assert message in result.stdout + result.stderr

    invoke(0)
    contract = policy.load_contract(target)
    contract["examples"].append("docs/literal.md")
    lifecycle.write_yaml(target / policy.CONTRACT, contract)
    example = target / "docs/literal.md"
    example.parent.mkdir(parents=True, exist_ok=True)
    literal = "```text\nuses: owner/action@v1\n```\n"
    example.write_text(literal, encoding="utf-8")
    invoke(0)
    example.write_text(literal.replace("```text", "```yaml"), encoding="utf-8")
    invoke(1, "full SHA")
    example.write_text(literal + "uses: owner/action@v1\n", encoding="utf-8")
    invoke(1, "full SHA")
    nested = "- 1. ```yaml\n     - {name: Example,\n      uses: owner/action@"
    nested += "a" * 40 + ", # v1.2.3\n      with: {}}\n     ```\n"
    example.write_text(nested, encoding="utf-8")
    invoke(0)
    example.write_text(
        '- 1. ```yaml\n     - {"uses"\n     : owner/action@v1}\n     ```\n',
        encoding="utf-8",
    )
    invoke(1, "Invalid YAML")
    example.write_text(nested.replace("with: {}}", "with: {"), encoding="utf-8")
    invoke(1, "Invalid YAML")
    example.write_text(nested.rsplit("     ```", 1)[0], encoding="utf-8")
    invoke(1, "Unclosed")
    example.write_text(literal, encoding="utf-8")
    helper.unlink()
    invoke(1, "instruction_contract_support")


def test_actions_helper_previous_mapping_mutant(tmp_path: Path) -> None:
    """The old relation produces an actual incomplete Actions-only runtime."""
    from template_sync_materialization_helpers import parse_manifest_mappings

    manifest = yaml.safe_load((ROOT / ".template-sync/manifest.yml").read_text(encoding="utf-8"))
    relation = next(
        item
        for item in manifest["template_manifest"]["path_mappings"]
        if item["pattern"] == ".github/scripts/instruction_contract_support.py"
    )
    assert set(relation["requires_any"]) == {
        "instruction-enforcement",
        "template-sync-support",
        "github-actions",
    }
    relation["requires_any"].remove("github-actions")
    _, mappings = parse_manifest_mappings(manifest)
    stage = tmp_path / "old-mapping"
    lifecycle.materializer.write_staged_candidate(
        template_root=ROOT,
        staging_root=stage,
        mappings=mappings,
        included_modules={"github-actions"},
        summary=lifecycle.materializer.Summary(["github-actions"], [], "copy"),
    )
    assert not (stage / ".github/scripts/instruction_contract_support.py").exists()
    result = subprocess.run(
        [
            sys.executable,
            "-E",
            str(stage / ".github/scripts/validate_workflow_security.py"),
            "--repo-root",
            str(stage),
        ],
        cwd=tmp_path,
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 1
    assert "instruction_contract_support" in result.stderr


@pytest.mark.parametrize("relative", ["docs/workflow-security.md", CURSOR_EXAMPLE])
def test_nested_governed_fence_materializer_inventory(tmp_path: Path, relative: str) -> None:
    """Actual staging preserves nested references and validates source before exclusions."""
    from tests.test_workflow_security_contract import copy_policy

    source = tmp_path / "source"
    contract = copy_policy(source)
    path = source / relative
    path.parent.mkdir(parents=True, exist_ok=True)
    pinned = "- 1. ```yaml\n     - {name: Example,\n      uses: owner/action@"
    pinned += "a" * 40 + ", # v1.2.3\n      with: {}}\n     ```\n"
    path.write_text(pinned, encoding="utf-8")
    _, _, mappings = lifecycle.materializer.load_validated_manifest_context(ROOT)
    modules = {"github-actions", "agent-instructions", "agent-cursor", "yaml"}

    def stage(name: str) -> Path:
        """Call the installed materializer against privately authored source data."""
        destination = tmp_path / name
        lifecycle.materializer.write_staged_candidate(
            template_root=source,
            staging_root=destination,
            mappings=mappings,
            included_modules=modules,
            summary=lifecycle.materializer.Summary(sorted(modules), [], "copy"),
        )
        return destination

    with pytest.raises(
        lifecycle.materializer.MaterializationError, match="Retained workflow example"
    ):
        stage("undeclared")
    assert snapshot(tmp_path / "undeclared") == {}
    contract["examples"].append(relative)
    lifecycle.write_yaml(source / policy.CONTRACT, contract)
    first = stage("declared")
    assert relative in policy.load_contract(first)["examples"]
    assert policy.check_examples((first / relative).read_text(encoding="utf-8")) == 1
    assert policy.validate_repository(first) >= 1
    assert snapshot(stage("repeat")) == snapshot(first)
    path.write_text("- 1. ```text\n     uses: owner/action@v1\n     ```\n", encoding="utf-8")
    with pytest.raises(
        lifecycle.materializer.MaterializationError, match="Retained workflow example"
    ):
        stage("stale")
    assert snapshot(tmp_path / "stale") == {}
    # Even a soon-excluded Cursor document is checked before selection removes it.
    path.write_text(
        '- 1. ```yaml\n     - {"uses"\n     : owner/action@v1}\n     ```\n', encoding="utf-8"
    )
    with pytest.raises(lifecycle.materializer.MaterializationError, match="Invalid YAML"):
        lifecycle.materializer.render_workflow_contract(source, mappings, {"github-actions"})
