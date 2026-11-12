"""Independent downstream deployment and failure controls for instruction profiles."""

from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest

pytestmark = pytest.mark.upstream_template_only

ROOT = Path(__file__).resolve().parents[1]
RUNTIME = (
    ".github/scripts/instruction_contract_core.py",
    ".github/scripts/instruction_contract_support.py",
    ".github/scripts/validate_instruction_profile.py",
    "schemas/instruction-profile.schema.json",
    "schemas/instruction-contracts.schema.json",
)


def write(root: Path, path: str, value: str) -> None:
    """Write one bounded fixture input."""
    target = root / path
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(value, encoding="utf-8", newline="\n")


def profile(root: Path, modules: list[str] | None = None) -> dict[str, Any]:
    """Create a standalone deployment containing no sync files or Python project."""
    for path in RUNTIME:
        write(root, path, (ROOT / path).read_text(encoding="utf-8"))
    document: dict[str, Any] = {
        "version": 1,
        "mode": "standalone",
        "modules": modules or ["agent-instructions", "instruction-enforcement", "baseline"],
        "exceptions": [],
    }
    write(root, ".github/instruction-profile.yml", yaml.safe_dump(document))
    write(
        root,
        ".github/instruction-contracts.yml",
        yaml.safe_dump(
            {
                "instruction_contracts": [
                    {
                        "path": "AGENTS.md",
                        "requires_modules": ["agent-instructions"],
                        "required_phrases": [
                            "Agents MUST validate.",
                            "Agents MUST preserve authority.",
                        ],
                    }
                ],
            }
        ),
    )
    write(root, "AGENTS.md", "Agents MUST validate.\nAgents MUST preserve authority.\n")
    return document


def run(root: Path) -> subprocess.CompletedProcess[str]:
    """Execute the deployed CLI so missing imports and native failures are visible."""
    return subprocess.run(
        [sys.executable, str(root / RUNTIME[2]), "--repo-root", str(root)],
        check=False,
        capture_output=True,
        text=True,
    )


@pytest.mark.parametrize(
    "excluded",
    [
        {"template-sync-support", "python"},
        {"template-sync-support", "python", "baseline", "azure-pipelines"},
        {"template-sync-support", "python", "baseline", "github-actions"},
        {"template-sync-support", "python", "powershell"},
        {"template-sync-support", "python", "terraform"},
        {
            "template-sync-support",
            "python",
            "powershell",
            "terraform",
            "agent-claude",
            "agent-cursor",
            "agent-gemini",
            "agent-hermes",
            "agent-copilot",
        },
    ],
)
def test_standalone_runs_after_sync_owned_files_are_physically_deleted(
    tmp_path: Path, excluded: set[str]
) -> None:
    """A real selected tree survives deletion of every support-owned path."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    import materialize_downstream_adoption as adoption
    from instruction_profile_migration import render_instruction_profile
    from template_sync_materialization_helpers import (
        iter_safe_repository_files,
        parse_manifest_mappings,
        selected_relation_for_path,
    )

    manifest = yaml.safe_load((ROOT / ".template-sync/manifest.yml").read_text(encoding="utf-8"))
    modules, mappings = parse_manifest_mappings(manifest)
    selected = set(modules) - excluded
    summary = adoption.Summary(
        retained_modules=sorted(selected),
        excluded_modules=["template-sync-support", "python"],
        default_adoption_mode="minimal-preservation",
    )
    adoption.write_staged_candidate(
        template_root=ROOT,
        staging_root=tmp_path,
        mappings=mappings,
        included_modules=selected,
        summary=summary,
    )
    render_instruction_profile(
        staging_root=tmp_path,
        target_root=tmp_path / "empty-target",
        marker_document={"template_sync": {"included_modules": sorted(selected)}},
    )
    # Independently copy and then physically remove all concrete support-owned
    # files, so success cannot be explained by a retained hidden dependency.
    support_paths: list[str] = []
    for relative in iter_safe_repository_files(ROOT)[0]:
        relation = selected_relation_for_path(relative, mappings)
        if (
            relation is not None
            and not relation.is_retained_by(selected)
            and relation.is_retained_by(selected | {"template-sync-support"})
        ):
            destination = tmp_path / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(ROOT / relative, destination)
            support_paths.append(relative)
    assert "schemas/template-sync-marker.schema.json" in support_paths
    assert "tests/test_validate_instruction_contracts.py" in support_paths
    for relative in support_paths:
        (tmp_path / relative).unlink()
    shutil.rmtree(tmp_path / ".template-sync")
    assert not (tmp_path / ".template-sync").exists()
    assert not any((tmp_path / relative).exists() for relative in support_paths)
    for path in tmp_path.rglob("*"):
        if path.is_file():
            relation = selected_relation_for_path(path.relative_to(tmp_path).as_posix(), mappings)
            assert relation is None or relation.is_retained_by(selected)
    result = run(tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    assert not (tmp_path / "pyproject.toml").exists()
    assert "Mode: standalone" in result.stdout
    agents_path = tmp_path / "AGENTS.md"
    if agents_path.exists():
        intact = agents_path.read_text(encoding="utf-8")
        anchor = "## PR Review Workflow (Codex-adapted)"
        assert intact.count(anchor) == 1
        agents_path.write_text(intact.replace(anchor, "## Removed protocol"), encoding="utf-8")
        weakened = run(tmp_path)
        assert weakened.returncode == 1, weakened.stdout + weakened.stderr
        assert anchor in weakened.stdout


@pytest.mark.parametrize(
    "bad",
    [
        "missing",
        "malformed",
        "no-modules",
        "unknown-module",
        "conflicting-modules",
        "conflicting-marker",
        "no-runner",
    ],
)
def test_invalid_applicability_is_native_failure(tmp_path: Path, bad: str) -> None:
    """No missing applicability input silently chooses all or no obligations."""
    document = profile(tmp_path)
    path = tmp_path / ".github/instruction-profile.yml"
    if bad == "missing":
        path.unlink()
    elif bad == "malformed":
        path.write_text("[", encoding="utf-8")
    else:
        if bad == "no-modules":
            del document["modules"]
        elif bad == "unknown-module":
            document["modules"].append("unknown")
        elif bad == "conflicting-modules":
            document["modules"].append("template-sync-support")
        elif bad == "conflicting-marker":
            write(tmp_path, ".template-sync/marker.yml", "invalid but conflicting\n")
        elif bad == "no-runner":
            document["modules"].remove("baseline")
        path.write_text(yaml.safe_dump(document), encoding="utf-8")
    result = run(tmp_path)
    assert result.returncode == 1, result.stdout + result.stderr


def test_marker_context_requires_missing_marker(tmp_path: Path) -> None:
    """An explicit downstream context never falls back to upstream-template."""
    profile(tmp_path)
    for source in (ROOT / ".template-sync/scripts").glob("*.py"):
        write(tmp_path, ".template-sync/scripts/" + source.name, source.read_text(encoding="utf-8"))
    write(
        tmp_path,
        ".github/instruction-profile.yml",
        "version: 1\nmode: marker\ncontext: downstream\n",
    )
    result = run(tmp_path)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Marker is required" in result.stderr


@pytest.mark.parametrize("change", ["delete", "weaken", "unrelated", "unsafe", "oversized"])
def test_scoped_exception_cannot_hide_unrelated_failures(tmp_path: Path, change: str) -> None:
    """Fixed required clauses and a content-bound waiver have independent outcomes."""
    document = profile(tmp_path)
    text = "Agents MAY validate.\nAgents MUST preserve authority.\n"
    write(tmp_path, "AGENTS.md", text)
    assert run(tmp_path).returncode == 1
    document["exceptions"] = [
        {
            "path": "AGENTS.md",
            "anchor": "Agents MUST validate.",
            "content_sha256": hashlib.sha256(text.encode()).hexdigest(),
            "reason": "Owner-reviewed fixture",
            "authorization_basis": "Fixture declaration",
        }
    ]
    write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
    accepted = run(tmp_path)
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    assert "Applied local exception" in accepted.stdout
    assert "not independent proof" in accepted.stdout
    if change == "delete":
        (tmp_path / "AGENTS.md").unlink()
    elif change == "weaken":
        write(tmp_path, "AGENTS.md", text.replace("MUST preserve", "MAY preserve"))
    elif change == "unrelated":
        write(tmp_path, "AGENTS.md", text + "Unrelated new content.\n")
    elif change == "unsafe":
        document["exceptions"][0]["path"] = "../outside.md"
        write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
    else:
        write(tmp_path, ".github/instruction-profile.yml", " " * (1024 * 1024 + 1))
    assert run(tmp_path).returncode == 1


def test_content_scope_guard_removal_is_detected_by_fixed_oracle(tmp_path: Path) -> None:
    """A guard mutant must recreate false acceptance of an unrelated file edit."""
    document = profile(tmp_path)
    text = "Agents MAY validate.\nAgents MUST preserve authority.\n"
    document["exceptions"] = [
        {
            "path": "AGENTS.md",
            "anchor": "Agents MUST validate.",
            "content_sha256": hashlib.sha256(text.encode()).hexdigest(),
            "reason": "Fixture",
            "authorization_basis": "Fixture declaration",
        }
    ]
    write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
    write(tmp_path, "AGENTS.md", text + "Unrelated content.\n")
    assert run(tmp_path).returncode == 1
    script = tmp_path / RUNTIME[2]
    source = script.read_text(encoding="utf-8")
    guard = 'declaration["content_sha256"] != file_digest(root, path)'
    assert source.count(guard) == 1
    script.write_text(source.replace(guard, "False"), encoding="utf-8")
    assert run(tmp_path).returncode == 0


def test_reviewed_seed_and_module_taxonomy_do_not_drift() -> None:
    """The seed is reviewed source data, never a standalone runtime dependency."""
    assert (ROOT / ".github/instruction-contracts.yml").read_bytes() == (
        ROOT / ".template-sync/instruction-contracts.yml"
    ).read_bytes()
    schema = json.loads((ROOT / "schemas/instruction-profile.schema.json").read_text())
    manifest = yaml.safe_load((ROOT / ".template-sync/manifest.yml").read_text(encoding="utf-8"))
    assert set(schema["$defs"]["moduleName"]["enum"]) == {
        item["name"] for item in manifest["template_manifest"]["modules"]
    }


def candidate_installation_marker(marker: dict[str, Any]) -> dict[str, Any]:
    """Record the candidate profile/catalog installation used by synthetic fixtures."""
    selected: dict[str, Any] = json.loads(json.dumps(marker))
    decisions = selected["template_sync"].setdefault("protected_file_decisions", [])
    for path in (".github/instruction-profile.yml", ".github/instruction-contracts.yml"):
        if not any(item["path"] == path for item in decisions):
            decisions.append(
                {
                    "path": path,
                    "decision": "TAKE",
                    "adoption_mode": "minimal-preservation",
                    "authorization_basis": "Fixture owner installs candidate input",
                    "authorized_scope": path,
                    "reason": "Explicit synthetic candidate installation",
                }
            )
    return selected


def test_migration_preserves_scoped_declarations_and_repeated_noop(tmp_path: Path) -> None:
    """Migrated decisions bind local content and do not duplicate on repeat adoption."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from instruction_profile_migration import render_instruction_profile

    target, stage = tmp_path / "target", tmp_path / "stage"
    document = profile(target)
    profile(stage)
    text = "Agents MAY validate.\nAgents MUST preserve authority.\n"
    write(target, "AGENTS.md", text)
    marker = {
        "template_sync": {
            "included_modules": document["modules"],
            "instruction_contract_waivers": [
                {
                    "path": "AGENTS.md",
                    "anchor": "Agents MUST validate.",
                    "reason": "Fixture",
                    "authorization_basis": "Reviewed fixture declaration",
                }
            ],
        }
    }
    marker = candidate_installation_marker(marker)
    render_instruction_profile(staging_root=stage, target_root=target, marker_document=marker)
    first = (stage / ".github/instruction-profile.yml").read_bytes()
    (target / ".github/instruction-profile.yml").write_bytes(first)
    assert run(target).returncode == 0
    render_instruction_profile(staging_root=stage, target_root=target, marker_document=marker)
    assert first == (stage / ".github/instruction-profile.yml").read_bytes()
    migrated = yaml.safe_load(first)
    assert len(migrated["exceptions"]) == 1
    assert (
        migrated["source_decisions"]["instruction_contract_waivers"]
        == marker["template_sync"]["instruction_contract_waivers"]
    )


def test_later_agent_removal_retires_exception_as_nonoperative_evidence(tmp_path: Path) -> None:
    """Reviewed module removal preserves history without retaining active waivers."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from instruction_profile_migration import render_instruction_profile

    target, stage = tmp_path / "target", tmp_path / "stage"
    document = profile(target)
    profile(stage)
    text = "Agents MAY validate.\nAgents MUST preserve authority.\n"
    write(target, "AGENTS.md", text)
    document["exceptions"] = [
        {
            "path": "AGENTS.md",
            "anchor": "Agents MUST validate.",
            "content_sha256": hashlib.sha256(text.encode()).hexdigest(),
            "reason": "Fixture",
            "authorization_basis": "Reviewed fixture declaration",
        }
    ]
    write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
    for root in (target, stage):
        catalog = yaml.safe_load((root / ".github/instruction-contracts.yml").read_text())
        catalog["instruction_contracts"][0]["requires_modules"].append("agent-codex")
        write(root, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
    render_instruction_profile(
        staging_root=stage,
        target_root=target,
        marker_document=candidate_installation_marker(
            {"template_sync": {"included_modules": document["modules"]}}
        ),
    )
    result = yaml.safe_load((stage / ".github/instruction-profile.yml").read_text())
    assert result["exceptions"] == []
    assert result["source_decisions"]["retired_exceptions"] == document["exceptions"]
    assert run(stage).returncode == 0


@pytest.mark.parametrize(
    "modules",
    [
        {"instruction-enforcement", "baseline"},
        {"instruction-enforcement", "agent-instructions"},
        {"agent-codex", "baseline"},
    ],
)
def test_materializer_rejects_contradictory_selections(modules: set[str]) -> None:
    """Selection validity is enforced before any candidate file is written."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from template_sync_materialization_helpers import validate_module_compatibility

    assert validate_module_compatibility(modules, ())


def test_legacy_agent_selection_requires_deliberate_migration(tmp_path: Path) -> None:
    """Old all-agent markers cannot silently delete existing agent entry points."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from materialize_downstream_adoption import (
        MaterializationError,
        validate_agent_selection_migration,
    )
    from template_sync_materialization_helpers import parse_marker_decision_data

    write(tmp_path, "AGENTS.md", "Existing local instructions.\n")
    marker = parse_marker_decision_data(
        {"template_sync": {"included_modules": ["agent-instructions"]}}
    )
    with pytest.raises(MaterializationError, match="Legacy agent selection"):
        validate_agent_selection_migration(tmp_path, {"agent-instructions"}, marker)
    validate_agent_selection_migration(tmp_path, {"agent-instructions", "agent-codex"}, marker)


@pytest.mark.parametrize(
    "path",
    [
        ".github/instruction-contracts.yml",
        "schemas/instruction-contracts.schema.json",
        "schemas/instruction-profile.schema.json",
    ],
)
@pytest.mark.parametrize("change", ["missing", "malformed"])
def test_standalone_missing_or_malformed_companion_inputs_fail(
    tmp_path: Path, path: str, change: str
) -> None:
    """Every required local input fails visibly, independently of profile selection."""
    profile(tmp_path)
    if change == "missing":
        (tmp_path / path).unlink()
    else:
        (tmp_path / path).write_text("[", encoding="utf-8")
    result = run(tmp_path)
    assert result.returncode == 1, result.stdout + result.stderr


@pytest.mark.parametrize("modules", [set(), {"baseline"}, {"agent-instructions", "baseline"}])
def test_policy_only_selection_has_no_enforcement_runtime_or_routes(modules: set[str]) -> None:
    """Instruction retention never silently makes optional enforcement mandatory."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from template_sync_materialization_helpers import (
        parse_manifest_mappings,
        selected_relation_for_path,
    )

    manifest = yaml.safe_load((ROOT / ".template-sync/manifest.yml").read_text(encoding="utf-8"))
    _known, mappings = parse_manifest_mappings(manifest)
    for path in (
        *RUNTIME,
        ".github/instruction-profile.yml",
        ".github/instruction-contracts.yml",
        ".github/workflows/instruction-contracts.yml",
        ".azuredevops/pipelines/instruction-contracts.yml",
    ):
        relation = selected_relation_for_path(path, mappings)
        assert relation is not None and not relation.is_retained_by(modules)


def test_current_decision_authorizes_legacy_agent_removal(tmp_path: Path) -> None:
    """An explicit current protected removal is accepted without inferring retention."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from materialize_downstream_adoption import validate_agent_selection_migration
    from template_sync_materialization_helpers import parse_marker_decision_data

    write(tmp_path, "AGENTS.md", "Existing local instructions.\n")
    marker = parse_marker_decision_data(
        {
            "template_sync": {
                "included_modules": ["agent-instructions"],
                "protected_file_decisions": [
                    {
                        "path": "AGENTS.md",
                        "decision": "REMOVE-LOCAL",
                        "authorization_basis": "Fixture owner approved removing AGENTS.md",
                        "authorized_scope": "Remove AGENTS.md",
                        "reason": "Codex omitted",
                    }
                ],
            }
        }
    )
    validate_agent_selection_migration(tmp_path, {"agent-instructions"}, marker)


def run_migration_schema_control(
    stage: Path,
    target: Path,
    *,
    redirect_trust: bool = False,
    omit_root_normalization: bool = False,
    omit_scoped_retirement: bool = False,
    omit_reference_validation: bool = False,
    restore_prior_path_prefilter: bool = False,
    reverse_section_precedence: bool = False,
    restore_content_preference: str | None = None,
    omit_retained_check: str | None = None,
    marker_document: dict[str, Any] | None = None,
) -> subprocess.CompletedProcess[str]:
    """Run the real migration or an exact schema-provenance mutant natively."""
    source_path = ROOT / ".template-sync/scripts/instruction_profile_migration.py"
    source = source_path.read_text(encoding="utf-8")
    if omit_retained_check is not None:
        guards = {
            "anchor": "(path, anchor) not in failure_keys[content_root]",
            "digest": "original_digest != digest",
        }
        for key, guard in guards.items():
            if omit_retained_check in (key, "both"):
                assert source.count(guard) == 1
                source = source.replace(guard, "False")
    if restore_content_preference:
        guard = '    """Select reconciled bytes without granting authority to unresolved decisions."""\n'
        assert source.count(guard) == 1
        source = source.replace(
            guard,
            guard
            + f"    if relative_path == {restore_content_preference!r}:\n"
            + "        return (target_root if (target_root / relative_path).exists() else staging_root), False\n",
        )
    if redirect_trust:
        guard = "TRUSTED_TOOL_ROOT = Path(__file__).resolve().parents[2]"
        assert source.count(guard) == 1
        source = source.replace(guard, "TRUSTED_TOOL_ROOT = Path(sys.argv[1])")
    if omit_root_normalization:
        for name in ("staging_root", "target_root"):
            guard = f"    {name} = {name}.resolve()\n"
            assert source.count(guard) == 1
            source = source.replace(guard, "")
    if omit_scoped_retirement:
        guard = "return core.section_applies(section, modules)"
        assert source.count(guard) == 1
        source = source.replace(guard, "return True")
    if omit_reference_validation:
        guard = "                protected_guide_reference_obligations=reference_obligations,\n"
        assert source.count(guard) == 1
        source = source.replace(guard, "")
    if restore_prior_path_prefilter:
        guard = '                exceptions.extend(local.get("exceptions", []))'
        assert source.count(guard) == 1
        source = source.replace(
            guard,
            "                active_paths = {item.path for item in contracts if set(item.requires_modules) <= modules}\n"
            '                for declaration in local.get("exceptions", []):\n'
            '                    if declaration["path"] in active_paths:\n'
            "                        exceptions.append(declaration)\n"
            "                    else:\n"
            "                        retired_exceptions.append(declaration)",
        )
    if reverse_section_precedence:
        guard = "key=lambda item: len(item.heading), reverse=True"
        assert source.count(guard) == 1
        source = source.replace(guard, "key=lambda item: len(item.heading), reverse=False")
    marker_document = candidate_installation_marker(
        marker_document
        or {
            "template_sync": {
                "included_modules": ["agent-instructions", "instruction-enforcement", "baseline"]
            }
        }
    )
    program = (
        "import sys\n"
        "from pathlib import Path\n"
        f"sys.path.insert(0, {str(ROOT / '.template-sync/scripts')!r})\n"
        f"namespace = {{'__file__': {str(source_path)!r}, '__name__': 'schema_control'}}\n"
        f"exec(compile({source!r}, {str(source_path)!r}, 'exec'), namespace)\n"
        "namespace['render_instruction_profile'](\n"
        "    staging_root=Path(sys.argv[1]), target_root=Path(sys.argv[2]),\n"
        f"    marker_document={marker_document!r})\n"
    )
    return subprocess.run(
        [sys.executable, "-c", program, str(stage), str(target)],
        check=False,
        capture_output=True,
        text=True,
    )


def test_migration_ignores_rejecting_candidate_schemas_and_detects_trust_mutant(
    tmp_path: Path,
) -> None:
    """Valid local data uses the installed tool schema even if candidates reject all."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    profile(target)
    for path in RUNTIME[3:]:
        write(stage, path, json.dumps({"type": "object", "not": {}}))
    accepted = run_migration_schema_control(stage, target)
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    mutant = run_migration_schema_control(stage, target, redirect_trust=True)
    assert mutant.returncode == 1, mutant.stdout + mutant.stderr
    assert "Schema validation failed" in mutant.stderr


@pytest.mark.parametrize("invalid_input", ["prior-profile", "candidate-catalog"])
def test_permissive_candidate_schema_cannot_excuse_invalid_data(
    tmp_path: Path, invalid_input: str
) -> None:
    """Fixed forbidden fields remain errors unless trusted schema provenance is removed."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    profile(target)
    # Preserve the module definition needed by the renderer while removing all
    # candidate validation constraints. The reviewed installed schemas stay intact.
    reviewed = json.loads((ROOT / RUNTIME[3]).read_text(encoding="utf-8"))
    write(stage, RUNTIME[3], json.dumps({"$defs": reviewed["$defs"]}))
    write(stage, RUNTIME[4], "{}")
    root, path = (
        (target, ".github/instruction-profile.yml")
        if invalid_input == "prior-profile"
        else (stage, ".github/instruction-contracts.yml")
    )
    document = yaml.safe_load((root / path).read_text(encoding="utf-8"))
    document["unreviewed_schema_bypass"] = True
    write(root, path, yaml.safe_dump(document))
    rejected = run_migration_schema_control(stage, target)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "unreviewed_schema_bypass" in rejected.stderr
    mutant = run_migration_schema_control(stage, target, redirect_trust=True)
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr


def test_installed_support_only_tool_later_enables_standalone_enforcement(tmp_path: Path) -> None:
    """Exercise installed dependency closure through the native adoption pipeline."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from template_sync_materialization_helpers import (
        is_protected_instruction_path,
        iter_safe_repository_files,
        parse_manifest_mappings,
        selected_relation_for_path,
    )

    target = tmp_path / "installed"
    target.mkdir()
    manifest = yaml.safe_load((ROOT / ".template-sync/manifest.yml").read_text(encoding="utf-8"))
    _known, mappings = parse_manifest_mappings(manifest)
    source_paths = iter_safe_repository_files(ROOT)[0]
    script_relative = ".template-sync/scripts/materialize_downstream_adoption.py"

    def adopt(
        script_root: Path, modules: set[str], *, source_root: Path = ROOT
    ) -> subprocess.CompletedProcess[str]:
        decisions = []
        for relative in source_paths:
            relation = selected_relation_for_path(relative, mappings)
            if (
                relation is not None
                and relation.is_retained_by(modules)
                and is_protected_instruction_path(relative)
            ):
                decisions.append(
                    {
                        "path": relative,
                        "decision": (
                            "SKIP" if source_root != ROOT and relative in RUNTIME[3:] else "TAKE"
                        ),
                        "adoption_mode": "minimal-preservation",
                        "authorization_basis": "Fixture owner authorizes this exact file",
                        "authorized_scope": relative,
                        "reason": "Fixture keeps reviewed schemas and selects retained files",
                    }
                )
        write(
            target,
            "decisions.yml",
            yaml.safe_dump(
                {
                    "template_sync": {
                        "source_repo": "https://github.com/franklesniak/copilot-repo-template",
                        "included_modules": sorted(modules),
                        "protected_file_decisions": decisions,
                    }
                }
            ),
        )
        return subprocess.run(
            [
                sys.executable,
                str(script_root / script_relative),
                "--template-root",
                str(source_root),
                "--target-root",
                str(target),
                "--decisions-file",
                "decisions.yml",
            ],
            check=False,
            capture_output=True,
            text=True,
        )

    initial = adopt(ROOT, {"template-sync-support"})
    assert initial.returncode == 0, initial.stdout + initial.stderr
    assert not (target / RUNTIME[2]).exists()
    assert not (target / ".github/instruction-profile.yml").exists()
    assert all((target / path).is_file() for path in RUNTIME[3:])
    help_result = subprocess.run(
        [sys.executable, str(target / script_relative), "--help"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert help_result.returncode == 0, help_result.stdout + help_result.stderr
    assert "--included-modules" in help_result.stdout
    installed_migration = target / ".template-sync/scripts/instruction_profile_migration.py"
    intact_migration = installed_migration.read_bytes()
    installed_migration.write_bytes(
        intact_migration + b"\nfrom validate_instruction_profile import file_digest\n"
    )
    missing_cli = subprocess.run(
        [sys.executable, str(target / script_relative), "--help"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert missing_cli.returncode == 1, missing_cli.stdout + missing_cli.stderr
    assert "ModuleNotFoundError" in missing_cli.stderr
    installed_migration.write_bytes(intact_migration)
    selected = {"baseline", "agent-instructions", "instruction-enforcement"}
    installed_schema = target / RUNTIME[3]
    reviewed_schema = installed_schema.read_bytes()
    installed_schema.unlink()
    missing_schema = adopt(target, selected)
    assert missing_schema.returncode == 1, missing_schema.stdout + missing_schema.stderr
    installed_schema.write_bytes(reviewed_schema)
    enabled = adopt(target, selected)
    assert enabled.returncode == 0, enabled.stdout + enabled.stderr
    generated = yaml.safe_load((target / ".github/instruction-profile.yml").read_text())
    assert generated["mode"] == "standalone"
    assert set(generated["modules"]) == selected
    candidate = tmp_path / "candidate"
    for relative in source_paths:
        destination = candidate / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(ROOT / relative, destination)
    for relative in RUNTIME[3:]:
        write(candidate, relative, json.dumps({"type": "object", "not": {}}))
    # Explicit SKIP declarations preserve installed reviewed schemas while the
    # selected source carries rejecting candidates. Migration must not execute them.
    poisoned = adopt(target, selected, source_root=candidate)
    assert poisoned.returncode == 0, poisoned.stdout + poisoned.stderr
    assert all((target / path).read_bytes() == (ROOT / path).read_bytes() for path in RUNTIME[3:])
    before = {
        path.relative_to(target).as_posix(): path.read_bytes()
        for path in target.rglob("*")
        if path.is_file() and "__pycache__" not in path.parts
    }
    repeated = adopt(target, selected, source_root=candidate)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    after = {
        path.relative_to(target).as_posix(): path.read_bytes()
        for path in target.rglob("*")
        if path.is_file() and "__pycache__" not in path.parts
    }
    assert before == after
    # Materialization reports excluded surfaces; cleanup is an explicit owner
    # step. Remove support-only files after its installed tool has finished.
    for path in tuple(target.rglob("*")):
        if not path.is_file():
            continue
        relation = selected_relation_for_path(path.relative_to(target).as_posix(), mappings)
        if relation is not None and not relation.is_retained_by(selected):
            path.unlink()
    shutil.rmtree(target / ".template-sync")
    assert not (target / ".template-sync").exists()
    validated = run(target)
    assert validated.returncode == 0, validated.stdout + validated.stderr


@pytest.mark.parametrize("root_form", ["canonical", "parent", "symlink"])
@pytest.mark.parametrize("aliased_root", ["stage", "target"])
def test_migration_normalizes_caller_roots_without_changing_profile(
    tmp_path: Path, root_form: str, aliased_root: str
) -> None:
    """Trusted caller roots may use aliases while input containment stays enforced."""
    real = tmp_path / "real"
    stage, target = real / "stage", real / "target"
    profile(stage)
    document = profile(target)
    roots = {"stage": stage, "target": target}
    if root_form == "parent":
        (real / "unused").mkdir()
        roots[aliased_root] = real / "unused" / ".." / aliased_root
    elif root_form == "symlink":
        alias = tmp_path / "alias"
        try:
            alias.symlink_to(real, target_is_directory=True)
        except OSError as error:
            pytest.skip(f"Directory symlink creation unavailable: {type(error).__name__}")
        roots[aliased_root] = alias / aliased_root
    accepted = run_migration_schema_control(roots["stage"], roots["target"])
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    generated = yaml.safe_load(
        (stage / ".github/instruction-profile.yml").read_text(encoding="utf-8")
    )
    assert generated["mode"] == "standalone"
    assert set(generated["modules"]) == set(document["modules"])
    assert generated["exceptions"] == []
    if root_form != "canonical":
        mutant = run_migration_schema_control(
            roots["stage"], roots["target"], omit_root_normalization=True
        )
        assert mutant.returncode == 1, mutant.stdout + mutant.stderr
        assert "Path escapes the repository root" in mutant.stderr


@pytest.mark.parametrize("escaped_input", ["catalog", "previous-profile"])
@pytest.mark.parametrize("root_form", ["canonical", "parent", "symlink"])
def test_migration_root_normalization_keeps_external_symlink_escape_rejected(
    tmp_path: Path, escaped_input: str, root_form: str
) -> None:
    """Normalizing trusted roots never makes sibling-prefix input targets trusted."""
    real = tmp_path / "real"
    stage, target = real / "stage", real / "target"
    profile(stage)
    profile(target)
    destination = stage / ".github/instruction-profile.yml"
    before = destination.read_bytes()
    escaped_path = (
        stage / ".github/instruction-contracts.yml"
        if escaped_input == "catalog"
        else target / ".github/instruction-profile.yml"
    )
    outside = real / ("stage-outside" if escaped_input == "catalog" else "target-outside")
    outside.mkdir()
    outside_file = outside / escaped_path.name
    outside_file.write_bytes(escaped_path.read_bytes())
    escaped_path.unlink()
    try:
        escaped_path.symlink_to(outside_file)
        if root_form == "symlink":
            alias = tmp_path / "alias"
            alias.symlink_to(real, target_is_directory=True)
            stage, target = alias / "stage", alias / "target"
    except OSError as error:
        pytest.skip(f"Symlink creation unavailable: {type(error).__name__}")
    if root_form == "parent":
        (real / "unused").mkdir()
        stage, target = real / "unused/../stage", real / "unused/../target"
    rejected = run_migration_schema_control(stage, target)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "Path escapes the repository root" in rejected.stderr
    assert destination.read_bytes() == before


@pytest.mark.parametrize("source", ["standalone", "marker"])
@pytest.mark.parametrize("heading", ["## Azure review", "## Azure: review"])
def test_migration_retires_scoped_section_declaration_and_detects_mutant(
    tmp_path: Path, source: str, heading: str
) -> None:
    """A retained file cannot keep a declaration for an excluded section scope."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    document["modules"].append("azure-devops-collaboration")
    paragraph = "Owner MUST review Azure changes."
    anchor = f"section:{heading}:paragraph:{hashlib.sha256(paragraph.encode()).hexdigest()}"
    for root in (stage, target):
        catalog_path = root / ".github/instruction-contracts.yml"
        catalog = yaml.safe_load(catalog_path.read_text())
        catalog["instruction_contracts"][0]["required_sections"] = [
            {
                "heading": heading,
                "next_heading": None,
                "required_paragraphs": [paragraph],
                "requires_modules": ["azure-devops-collaboration"],
            }
        ]
        write(root, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
        with (root / "AGENTS.md").open("a", encoding="utf-8") as stream:
            stream.write(f"\n{heading}\n")
    declaration = {
        "path": "AGENTS.md",
        "anchor": anchor,
        "content_sha256": hashlib.sha256(
            (target / "AGENTS.md").read_text(encoding="utf-8").encode()
        ).hexdigest(),
        "reason": "Owner retains local section",
        "authorization_basis": "Fixture owner approval",
    }
    document["exceptions"] = [declaration]
    write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
    before = run(target)
    assert before.returncode == 0, before.stdout + before.stderr
    marker: dict[str, Any] = {
        "template_sync": {
            "included_modules": ["baseline", "agent-instructions", "instruction-enforcement"]
        }
    }
    if source == "marker":
        document["exceptions"] = []
        write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
        marker["template_sync"]["instruction_contract_waivers"] = [
            {key: value for key, value in declaration.items() if key != "content_sha256"}
        ]
    migrated = run_migration_schema_control(stage, target, marker_document=marker)
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    destination = stage / ".github/instruction-profile.yml"
    generated = yaml.safe_load(destination.read_text())
    assert generated["exceptions"] == []
    assert generated["source_decisions"]["retired_exceptions"] == [declaration]
    validated = run(stage)
    assert validated.returncode == 0, validated.stdout + validated.stderr

    before_mutation = destination.read_bytes()
    mutant = run_migration_schema_control(
        stage, target, marker_document=marker, omit_scoped_retirement=True
    )
    if source == "standalone":
        assert mutant.returncode == 1, mutant.stdout + mutant.stderr
        assert "Existing standalone exception conflicts" in mutant.stderr
    else:
        assert mutant.returncode == 1, mutant.stdout + mutant.stderr
        assert "Instruction waiver conflicts with selected preserved content" in mutant.stderr
        assert destination.read_bytes() == before_mutation
        with pytest.raises(AssertionError):
            assert mutant.returncode == 0
    restored = run_migration_schema_control(stage, target, marker_document=marker)
    assert restored.returncode == 0, restored.stdout + restored.stderr
    before_bytes = destination.read_bytes()
    write(target, ".github/instruction-profile.yml", before_bytes.decode())
    repeated = run_migration_schema_control(stage, target, marker_document=marker)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert destination.read_bytes() == before_bytes


@pytest.mark.parametrize("control", ["active", "unknown", "changed-content"])
def test_migration_keeps_active_and_unrecognized_declarations(tmp_path: Path, control: str) -> None:
    """Retirement must not hide invalid declarations by inspecting current failures."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    for root in (stage, target):
        write(root, "AGENTS.md", "Agents MUST preserve authority.\n")
    declaration = {
        "path": "AGENTS.md",
        "anchor": "unknown-anchor" if control == "unknown" else "Agents MUST validate.",
        "content_sha256": (
            "0" * 64
            if control == "changed-content"
            else hashlib.sha256(
                (target / "AGENTS.md").read_text(encoding="utf-8").encode()
            ).hexdigest()
        ),
        "reason": "Owner retains local text",
        "authorization_basis": "Fixture owner approval",
    }
    document["exceptions"] = [declaration]
    write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
    before = (stage / ".github/instruction-profile.yml").read_bytes()
    migrated = run_migration_schema_control(stage, target)
    if control != "active":
        assert migrated.returncode == 1, migrated.stdout + migrated.stderr
        assert "Existing standalone exception conflicts" in migrated.stderr
        assert (stage / ".github/instruction-profile.yml").read_bytes() == before
        assert yaml.safe_load((target / ".github/instruction-profile.yml").read_text()) == document
        return
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    generated = yaml.safe_load((stage / ".github/instruction-profile.yml").read_text())
    assert generated["exceptions"] == [declaration]
    assert "retired_exceptions" not in generated["source_decisions"]
    validated = run(stage)
    assert validated.returncode == 0, validated.stdout + validated.stderr


def test_migration_retires_stale_section_declaration_when_module_returns(tmp_path: Path) -> None:
    """Restoring the target module retires its previously permitted stale section."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    for root in (stage, target):
        catalog_path = root / ".github/instruction-contracts.yml"
        catalog = yaml.safe_load(catalog_path.read_text())
        catalog["protected_guide_section_obligations"] = [
            {
                "key": "azure-section",
                "path": "AGENTS.md",
                "target_modules": ["azure-devops-collaboration"],
                "stale_headings": ["## Azure review"],
            }
        ]
        write(root, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
        with (root / "AGENTS.md").open("a", encoding="utf-8") as stream:
            stream.write("\n## Azure review\n")
    declaration = {
        "path": "AGENTS.md",
        "anchor": "stale:azure-section:heading:## Azure review",
        "content_sha256": hashlib.sha256(
            (target / "AGENTS.md").read_text(encoding="utf-8").encode()
        ).hexdigest(),
        "reason": "Owner retains local text",
        "authorization_basis": "Fixture owner approval",
    }
    document["exceptions"] = [declaration]
    write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
    before = run(target)
    assert before.returncode == 0, before.stdout + before.stderr
    marker = {
        "template_sync": {"included_modules": document["modules"] + ["azure-devops-collaboration"]}
    }
    migrated = run_migration_schema_control(stage, target, marker_document=marker)
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    generated = yaml.safe_load((stage / ".github/instruction-profile.yml").read_text())
    assert generated["exceptions"] == []
    assert generated["source_decisions"]["retired_exceptions"] == [declaration]
    after = run(stage)
    assert after.returncode == 0, after.stdout + after.stderr


@pytest.mark.parametrize(
    "kind,target_text,content",
    [
        ("prose-reference", "Azure workflow", "Use Azure workflow."),
        ("absolute-url", "https://example.invalid/azure", "Read https://example.invalid/azure."),
        (
            "markdown-relative-link",
            "docs/azure%20guide.md#review",
            "Read [Azure](docs/azure%20guide.md#review).",
        ),
    ],
)
@pytest.mark.parametrize("selector", ["matching", "wrong-key", "wrong-module", "wrong-path"])
def test_migration_reference_waiver_scope_and_module_return(
    tmp_path: Path, kind: str, target_text: str, content: str, selector: str
) -> None:
    """Only exact reviewed reference failures migrate; later scope changes retire them."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    obligation: dict[str, Any] = {
        "key": "azure-reference",
        "path": "AGENTS.md",
        "reference_kind": kind,
        "target_modules": ["azure-devops-collaboration"],
    }
    if kind == "markdown-relative-link":
        obligation["target_path"] = "docs/azure guide.md"
    else:
        obligation["tokens"] = [target_text]
    for root in (stage, target):
        catalog_path = root / ".github/instruction-contracts.yml"
        catalog = yaml.safe_load(catalog_path.read_text())
        catalog["protected_guide_reference_obligations"] = [obligation]
        write(root, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
        old_text = (root / "AGENTS.md").read_text(encoding="utf-8")
        write(root, "AGENTS.md", old_text + content + "\n")
    waiver: dict[str, str] = {
        "path": "CLAUDE.md" if selector == "wrong-path" else "AGENTS.md",
        "contract_key": "wrong-key" if selector == "wrong-key" else "azure-reference",
        "target_module": "python" if selector == "wrong-module" else "azure-devops-collaboration",
        "reason": "Owner retains exact reference",
        "authorization_basis": "Fixture owner approval",
    }
    # Relative-link declarations may use the reviewed target-path selector.
    if kind == "markdown-relative-link" and selector == "matching":
        waiver.pop("target_module")
        waiver["target_path"] = "docs/azure guide.md"
    marker = {
        "template_sync": {
            "included_modules": document["modules"],
            "protected_guide_contract_waivers": [waiver],
        }
    }
    before = run(stage)
    assert before.returncode == 1, before.stdout + before.stderr
    migrated = run_migration_schema_control(stage, target, marker_document=marker)
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    destination = stage / ".github/instruction-profile.yml"
    generated = yaml.safe_load(destination.read_text())
    assert generated["source_decisions"]["protected_guide_contract_waivers"] == [waiver]
    if selector != "matching":
        assert generated["exceptions"] == []
        rejected = run(stage)
        assert rejected.returncode == 1, rejected.stdout + rejected.stderr
        return
    declaration = {
        "path": "AGENTS.md",
        "anchor": f"reference:azure-reference:{kind}:{target_text}",
        "content_sha256": hashlib.sha256(
            (target / "AGENTS.md").read_text(encoding="utf-8").encode()
        ).hexdigest(),
        "reason": waiver["reason"],
        "authorization_basis": waiver["authorization_basis"],
    }
    assert generated["exceptions"] == [declaration]
    validated = run(stage)
    assert validated.returncode == 0, validated.stdout + validated.stderr
    mutant = run_migration_schema_control(
        stage, target, marker_document=marker, omit_reference_validation=True
    )
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    rejected = run(stage)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    restored = run_migration_schema_control(stage, target, marker_document=marker)
    assert restored.returncode == 0, restored.stdout + restored.stderr
    before_bytes = destination.read_bytes()
    write(target, ".github/instruction-profile.yml", before_bytes.decode())
    repeated = run_migration_schema_control(stage, target, marker_document=marker)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert destination.read_bytes() == before_bytes
    marker["template_sync"]["included_modules"] = document["modules"] + [
        "azure-devops-collaboration"
    ]
    retired = run_migration_schema_control(stage, target, marker_document=marker)
    assert retired.returncode == 0, retired.stdout + retired.stderr
    generated = yaml.safe_load(destination.read_text())
    assert generated["exceptions"] == []
    assert generated["source_decisions"]["retired_exceptions"] == [declaration]
    validated = run(stage)
    assert validated.returncode == 0, validated.stdout + validated.stderr


@pytest.mark.parametrize("kind", ["section", "reference"])
@pytest.mark.parametrize("declaration_source", ["profile", "instruction-waiver", "removal"])
def test_migration_preserves_obligation_only_paths_and_retires_only_known_scope(
    tmp_path: Path, kind: str, declaration_source: str
) -> None:
    """Catalog obligations do not require a parallel instruction-contract row."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    obligation: dict[str, Any] = {
        "key": "azure-only",
        "path": "REFERENCE.md",
        "target_modules": ["azure-devops-collaboration"],
    }
    if kind == "section":
        obligation["stale_headings"] = ["## Azure"]
        anchor = "stale:azure-only:heading:## Azure"
        content = "## Azure\n"
    else:
        obligation.update(reference_kind="prose-reference", tokens=["Azure workflow"])
        anchor = "reference:azure-only:prose-reference:Azure workflow"
        content = "Azure workflow\n"
    for root in (stage, target):
        catalog_path = root / ".github/instruction-contracts.yml"
        catalog = yaml.safe_load(catalog_path.read_text())
        catalog[f"protected_guide_{kind}_obligations"] = [obligation]
        write(root, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
        if declaration_source != "removal":
            write(root, "REFERENCE.md", content)
    declaration = {
        "path": "REFERENCE.md",
        "anchor": "file:absent" if declaration_source == "removal" else anchor,
        "content_sha256": (
            "absent"
            if declaration_source == "removal"
            else hashlib.sha256(content.encode()).hexdigest()
        ),
        "reason": "Owner retains exact local decision",
        "authorization_basis": "Fixture owner approval",
    }
    document["exceptions"] = [declaration]
    write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
    before = run(target)
    assert before.returncode == 0, before.stdout + before.stderr
    marker: dict[str, Any] = {"template_sync": {"included_modules": document["modules"]}}
    if declaration_source != "profile":
        document["exceptions"] = []
        write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
        if declaration_source == "instruction-waiver":
            marker["template_sync"]["instruction_contract_waivers"] = [
                {key: value for key, value in declaration.items() if key != "content_sha256"}
            ]
        else:
            marker["template_sync"]["protected_file_decisions"] = [
                {
                    "path": "REFERENCE.md",
                    "decision": "REMOVE-LOCAL",
                    "authorized_scope": "REFERENCE.md",
                    "reason": declaration["reason"],
                    "authorization_basis": declaration["authorization_basis"],
                }
            ]
    migrated = run_migration_schema_control(stage, target, marker_document=marker)
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    destination = stage / ".github/instruction-profile.yml"
    generated = yaml.safe_load(destination.read_text())
    assert generated["exceptions"] == [declaration]
    validated = run(stage)
    assert validated.returncode == 0, validated.stdout + validated.stderr
    saved = destination.read_bytes()
    write(target, ".github/instruction-profile.yml", saved.decode())
    repeated = run_migration_schema_control(stage, target, marker_document=marker)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert destination.read_bytes() == saved
    if declaration_source == "profile":
        mutant = run_migration_schema_control(
            stage, target, marker_document=marker, restore_prior_path_prefilter=True
        )
        assert mutant.returncode == 0, mutant.stdout + mutant.stderr
        rejected = run(stage)
        assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    marker["template_sync"]["included_modules"] = document["modules"] + [
        "azure-devops-collaboration"
    ]
    retired = run_migration_schema_control(stage, target, marker_document=marker)
    assert retired.returncode == 0, retired.stdout + retired.stderr
    generated = yaml.safe_load(destination.read_text())
    assert generated["exceptions"] == []
    assert generated["source_decisions"]["retired_exceptions"] == [declaration]
    validated = run(stage)
    assert validated.returncode == 0, validated.stdout + validated.stderr


@pytest.mark.parametrize("anchor", ["unknown-anchor", "file:absent"])
def test_migration_preserves_unknown_path_failure_and_detects_false_success_mutant(
    tmp_path: Path, anchor: str
) -> None:
    """An unknown path is not evidence that its declaration is legitimately retired."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    declaration = {
        "path": "UNKNOWN.md",
        "anchor": anchor,
        "content_sha256": "absent",
        "reason": "Unknown fixture declaration",
        "authorization_basis": "Fixture owner approval",
    }
    document["exceptions"] = [declaration]
    write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
    before = run(target)
    assert before.returncode == 1, before.stdout + before.stderr
    before_profile = (stage / ".github/instruction-profile.yml").read_bytes()
    migrated = run_migration_schema_control(stage, target)
    assert migrated.returncode == 1, migrated.stdout + migrated.stderr
    assert "Existing standalone exception conflicts" in migrated.stderr
    assert (stage / ".github/instruction-profile.yml").read_bytes() == before_profile
    assert yaml.safe_load((target / ".github/instruction-profile.yml").read_text()) == document
    mutant = run_migration_schema_control(stage, target, restore_prior_path_prefilter=True)
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    false_success = run(stage)
    assert false_success.returncode == 0, false_success.stdout + false_success.stderr


def test_migration_does_not_invent_exception_for_unknown_protected_removal(tmp_path: Path) -> None:
    """Ordinary protected removal history does not manufacture an instruction failure."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    removal = {
        "path": "UNKNOWN.md",
        "decision": "REMOVE-LOCAL",
        "authorized_scope": "UNKNOWN.md",
        "authorization_basis": "Fixture owner approval",
        "reason": "Ordinary file removal",
    }
    marker = {
        "template_sync": {
            "included_modules": document["modules"],
            "protected_file_decisions": [removal],
        }
    }
    migrated = run_migration_schema_control(stage, target, marker_document=marker)
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    generated = yaml.safe_load((stage / ".github/instruction-profile.yml").read_text())
    assert generated["exceptions"] == []
    expected_decisions = candidate_installation_marker(marker)["template_sync"][
        "protected_file_decisions"
    ]
    assert expected_decisions[0] == removal
    assert generated["source_decisions"]["protected_file_decisions"] == expected_decisions
    validated = run(stage)
    assert validated.returncode == 0, validated.stdout + validated.stderr


def test_migration_direct_phrase_can_resemble_excluded_structured_section(tmp_path: Path) -> None:
    """A direct phrase takes precedence over a syntactically identical section key."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    anchor = "section:## Azure: review"
    for root in (stage, target):
        catalog_path = root / ".github/instruction-contracts.yml"
        catalog = yaml.safe_load(catalog_path.read_text())
        contract = catalog["instruction_contracts"][0]
        contract["required_phrases"].append(anchor)
        contract["required_sections"] = [
            {
                "heading": "## Azure: review",
                "next_heading": None,
                "required_paragraphs": ["Owner MUST review."],
                "requires_modules": ["azure-devops-collaboration"],
            }
        ]
        write(root, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
    declaration = {
        "path": "AGENTS.md",
        "anchor": anchor,
        "content_sha256": hashlib.sha256((target / "AGENTS.md").read_bytes()).hexdigest(),
        "reason": "Fixture direct phrase",
        "authorization_basis": "Fixture owner approval",
    }
    document["exceptions"] = [declaration]
    write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
    migrated = run_migration_schema_control(stage, target)
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    generated = yaml.safe_load((stage / ".github/instruction-profile.yml").read_text())
    assert generated["exceptions"] == [declaration]
    validated = run(stage)
    assert validated.returncode == 0, validated.stdout + validated.stderr


def test_migration_uses_longest_overlapping_colon_heading_and_detects_mutant(
    tmp_path: Path,
) -> None:
    """A full heading must outrank a shorter heading with a plausible suffix."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    heading = "## Azure:paragraph:" + "a" * 64
    anchor = "section:" + heading
    for root in (stage, target):
        catalog_path = root / ".github/instruction-contracts.yml"
        catalog = yaml.safe_load(catalog_path.read_text())
        catalog["instruction_contracts"][0]["required_sections"] = [
            {
                "heading": "## Azure",
                "next_heading": heading,
                "required_paragraphs": ["Owner MUST review."],
                "requires_modules": ["azure-devops-collaboration"],
            },
            {
                "heading": heading,
                "next_heading": None,
                "required_paragraphs": ["Owner MUST validate."],
                "requires_modules": ["baseline"],
            },
        ]
        write(root, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
    declaration = {
        "path": "AGENTS.md",
        "anchor": anchor,
        "content_sha256": hashlib.sha256((target / "AGENTS.md").read_bytes()).hexdigest(),
        "reason": "Fixture missing long heading",
        "authorization_basis": "Fixture owner approval",
    }
    # Fixed missing-heading, boundary and paragraph identities are independently
    # asserted; do not derive expected failures from the validator under test.
    declarations = [declaration] + [
        {**declaration, "anchor": anchor + suffix}
        for suffix in (
            ":boundary:66d48f215ef64df932d7f6213bfe29d3276997c3455f54ab40955aa9938c7a78",
            ":paragraph:1ae271af103a21057c05f2d53240e5694f34b8b461031f30c47b38619fa4f995",
        )
    ]
    document["exceptions"] = declarations
    write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
    before = run(target)
    assert before.returncode == 0, before.stdout + before.stderr
    migrated = run_migration_schema_control(stage, target)
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    generated = yaml.safe_load((stage / ".github/instruction-profile.yml").read_text())
    assert generated["exceptions"] == declarations
    validated = run(stage)
    assert validated.returncode == 0, validated.stdout + validated.stderr
    mutant = run_migration_schema_control(stage, target, reverse_section_precedence=True)
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    rejected = run(stage)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr


@pytest.mark.parametrize("path", ["AGENTS.md", "docs/review.md"])
@pytest.mark.parametrize("kind", ["direct", "section", "reference"])
@pytest.mark.parametrize("decision", ["TAKE", "SKIP"])
@pytest.mark.parametrize("candidate_failure", [True, False])
def test_migration_binds_only_selected_content(
    tmp_path: Path, path: str, kind: str, decision: str, candidate_failure: bool
) -> None:
    """Scoped declarations describe reconciled bytes, not a replaced local file."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    intact = "Agents MUST validate.\nAgents MUST preserve authority.\n"
    damaged = intact.replace("Agents MUST validate.\n", "") if kind == "direct" else intact
    if kind == "section":
        damaged += "\n## Azure review\n"
    elif kind == "reference":
        damaged += "\nUse Azure workflow.\n"
    for root in (stage, target):
        catalog = yaml.safe_load((root / ".github/instruction-contracts.yml").read_text())
        catalog["instruction_contracts"][0]["path"] = path
        if kind == "section":
            catalog["protected_guide_section_obligations"] = [
                {
                    "path": path,
                    "key": "azure-section",
                    "stale_headings": ["## Azure review"],
                    "target_modules": ["azure-devops-collaboration"],
                }
            ]
        elif kind == "reference":
            catalog["protected_guide_reference_obligations"] = [
                {
                    "path": path,
                    "key": "azure-reference",
                    "reference_kind": "prose-reference",
                    "tokens": ["Azure workflow"],
                    "target_modules": ["azure-devops-collaboration"],
                }
            ]
        write(root, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
    write(target, path, "<!-- Local bytes. -->\n" + damaged)
    write(stage, path, damaged if candidate_failure else intact)
    marker: dict[str, Any] = {"template_sync": {"included_modules": document["modules"]}}
    fields = marker["template_sync"]
    if path == "AGENTS.md":
        fields["protected_file_decisions"] = [
            {
                "path": path,
                "decision": decision,
                "adoption_mode": "minimal-preservation",
                "authorization_basis": "Fixture owner",
                "authorized_scope": path,
                "reason": "Reviewed content selection",
            }
        ]
    else:
        fields["local_overrides"] = [
            {"path": path, "default_decision": decision, "reason": "Owner selection"}
        ]
    waiver = {
        "path": path,
        "reason": "Exact reviewed failure",
        "authorization_basis": "Fixture owner",
    }
    if kind == "direct":
        waiver["anchor"] = "Agents MUST validate."
        fields["instruction_contract_waivers"] = [waiver]
    else:
        waiver.update(
            {"contract_key": "azure-" + kind, "target_module": "azure-devops-collaboration"}
        )
        fields["protected_guide_contract_waivers"] = [waiver]
    before = (stage / ".github/instruction-profile.yml").read_bytes()
    selected = (stage if decision == "TAKE" else target) / path
    selected_text = selected.read_text(encoding="utf-8")
    migrated = run_migration_schema_control(stage, target, marker_document=marker)
    if kind == "direct" and decision == "TAKE" and not candidate_failure:
        assert migrated.returncode == 1, migrated.stdout + migrated.stderr
        assert "Instruction waiver conflicts with selected TAKE content" in migrated.stderr
        assert (stage / ".github/instruction-profile.yml").read_bytes() == before
        assert (target / path).read_text(encoding="utf-8").startswith("<!-- Local bytes.")
        return
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    generated = yaml.safe_load((stage / ".github/instruction-profile.yml").read_text())
    expected_count = int(decision == "SKIP" or candidate_failure)
    assert len(generated["exceptions"]) == expected_count
    if expected_count:
        assert (
            generated["exceptions"][0]["content_sha256"]
            == hashlib.sha256(selected_text.encode()).hexdigest()
        )
    # Reconcile the one fixture guide, then exercise the real installed CLI.
    write(stage, path, selected_text)
    validated = run(stage)
    assert validated.returncode == 0, validated.stdout + validated.stderr
    assert not (stage / ".template-sync").exists()
    if decision == "TAKE" and candidate_failure:
        mutant = run_migration_schema_control(
            stage, target, marker_document=marker, restore_content_preference=path
        )
        assert mutant.returncode == 0, mutant.stdout + mutant.stderr
        rejected = run(stage)
        assert rejected.returncode == 1, rejected.stdout + rejected.stderr
        assert "Exception does not match" in rejected.stderr
        restored = run_migration_schema_control(stage, target, marker_document=marker)
        assert restored.returncode == 0, restored.stdout + restored.stderr
    write(
        stage, path, selected_text.replace("Agents MUST preserve authority.", "Unreviewed removal.")
    )
    unrelated = run(stage)
    assert unrelated.returncode == 1, unrelated.stdout + unrelated.stderr


@pytest.mark.parametrize("selection", ["SKIP", "TAKE"])
def test_migration_selection_preserves_absence_and_old_bindings(
    tmp_path: Path, selection: str
) -> None:
    """SKIP absence stays absent; TAKE never renews prior standalone declarations."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    marker = {
        "template_sync": {
            "included_modules": document["modules"],
            "protected_file_decisions": [
                {
                    "path": "AGENTS.md",
                    "decision": selection,
                    "adoption_mode": "minimal-preservation",
                    "authorization_basis": "Fixture owner",
                    "authorized_scope": "AGENTS.md",
                    "reason": "Reviewed selection",
                }
            ],
        }
    }
    if selection == "SKIP":
        (target / "AGENTS.md").unlink()
        document["exceptions"] = [
            {
                "path": "AGENTS.md",
                "anchor": "file:absent",
                "content_sha256": "absent",
                "reason": "Reviewed absence",
                "authorization_basis": "Owner",
            }
        ]
    else:
        write(target, "AGENTS.md", "Agents MUST preserve authority.\n")
        write(
            stage, "AGENTS.md", "<!-- Selected replacement. -->\nAgents MUST preserve authority.\n"
        )
        document["exceptions"] = [
            {
                "path": "AGENTS.md",
                "anchor": "Agents MUST validate.",
                "content_sha256": hashlib.sha256((target / "AGENTS.md").read_bytes()).hexdigest(),
                "reason": "Old local binding",
                "authorization_basis": "Owner",
            }
        ]
    write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
    before_profile = (stage / ".github/instruction-profile.yml").read_bytes()
    migrated = run_migration_schema_control(stage, target, marker_document=marker)
    if selection == "TAKE":
        assert migrated.returncode == 1, migrated.stdout + migrated.stderr
        assert "Existing standalone exception conflicts" in migrated.stderr
        assert (stage / ".github/instruction-profile.yml").read_bytes() == before_profile
        assert yaml.safe_load((target / ".github/instruction-profile.yml").read_text()) == document
        return
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    generated = yaml.safe_load((stage / ".github/instruction-profile.yml").read_text())
    assert generated["exceptions"] == document["exceptions"]
    if selection == "SKIP":
        (stage / "AGENTS.md").unlink()
    validated = run(stage)
    assert validated.returncode == 0, validated.stdout + validated.stderr


@pytest.mark.parametrize(
    "overrides,expected",
    [
        ([("docs/", "SKIP"), ("docs/review.md", "TAKE")], "TAKE"),
        ([("docs/review.md", "TAKE"), ("docs/", "SKIP")], "TAKE"),
        ([("docs/review.md", "TAKE"), ("docs/review.md", "SKIP")], "SKIP"),
        ([("docs/review.md", "SKIP"), ("docs/review.md", "TAKE")], "TAKE"),
        ([("docs/", "SKIP"), ("other/", "TAKE")], "SKIP"),
        ([("docs/review.md/", "SKIP"), ("docs/review.md", "TAKE")], "TAKE"),
    ],
)
def test_migration_override_selection_matches_reconciliation(
    tmp_path: Path, overrides: list[tuple[str, str]], expected: str
) -> None:
    """Fixed precedence expectations independently match the real caller selector."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from instruction_profile_migration import selected_content_root
    from materialize_downstream_adoption import most_specific_local_override
    from template_sync_materialization_helpers import parse_marker_decision_data

    stage, target = tmp_path / "stage", tmp_path / "target"
    for root in (stage, target):
        write(root, "docs/review.md", "Reviewed guide\n")
        write(root, "AGENTS.md", "Reviewed agent\n")
    marker = {
        "included_modules": ["baseline"],
        "local_overrides": [
            {"path": path, "default_decision": decision, "reason": "Reviewed selection"}
            for path, decision in overrides
        ],
    }
    parsed = parse_marker_decision_data({"template_sync": marker})
    actual = most_specific_local_override("docs/review.md", parsed.local_overrides)
    assert actual is not None and actual.default_decision == expected
    chosen, taken = selected_content_root("docs/review.md", stage, target, marker)
    assert chosen == (stage if expected == "TAKE" else target)
    assert taken == (expected == "TAKE")
    protected_marker = {
        **marker,
        "local_overrides": [
            {
                "path": "AGENTS.md",
                "default_decision": "TAKE",
                "reason": "Cannot override protected SKIP",
            }
        ],
        "protected_file_decisions": [{"path": "AGENTS.md", "decision": "SKIP"}],
    }
    assert selected_content_root("AGENTS.md", stage, target, protected_marker) == (target, False)


def test_materialized_protected_waivers_follow_take_skip_and_noop(tmp_path: Path) -> None:
    """Actual reconciliation and installed validation agree on selected guide bytes."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from template_sync_materialization_helpers import (
        is_protected_instruction_path,
        iter_safe_repository_files,
        parse_manifest_mappings,
        selected_relation_for_path,
    )

    modules = {
        "baseline",
        "agent-instructions",
        "instruction-enforcement",
        "agent-codex",
        "github-actions",
    }
    manifest = yaml.safe_load((ROOT / ".template-sync/manifest.yml").read_text(encoding="utf-8"))
    _, mappings = parse_manifest_mappings(manifest)
    protected = []
    for relative in iter_safe_repository_files(ROOT)[0]:
        relation = selected_relation_for_path(relative, mappings)
        if (
            relation is not None
            and relation.is_retained_by(modules)
            and is_protected_instruction_path(relative)
        ):
            protected.append(relative)
    for case in ("empty-take", "existing-take", "existing-skip"):
        target = tmp_path / case
        previous_text = ""
        if case != "empty-take":
            previous_text = "<!-- Local annotation. -->\n" + (
                tmp_path / "empty-take/AGENTS.md"
            ).read_text(encoding="utf-8")
            write(target, "AGENTS.md", previous_text)
        decisions = [
            {
                "path": path,
                "decision": "SKIP" if case == "existing-skip" and path == "AGENTS.md" else "TAKE",
                "adoption_mode": "minimal-preservation",
                "authorization_basis": "Fixture owner selects this exact path",
                "authorized_scope": path,
                "reason": "Native content binding control",
            }
            for path in protected
        ]
        marker = {
            "template_sync": {
                "source_repo": "https://github.com/franklesniak/copilot-repo-template",
                "included_modules": sorted(modules),
                "protected_file_decisions": decisions,
                "protected_guide_contract_waivers": [
                    {
                        "path": "AGENTS.md",
                        "contract_key": "agents-azure-devops-pr-review-protocol",
                        "target_module": "azure-devops-collaboration",
                        "reason": "Retain reviewed protocol",
                        "authorization_basis": "Fixture owner retains this protocol",
                    }
                ],
            }
        }
        write(target, "decisions.yml", yaml.safe_dump(marker))
        command = [
            sys.executable,
            str(ROOT / ".template-sync/scripts/materialize_downstream_adoption.py"),
            "--template-root",
            str(ROOT),
            "--target-root",
            str(target),
            "--decisions-file",
            "decisions.yml",
        ]
        adopted = subprocess.run(command, check=False, capture_output=True, text=True, timeout=60)
        assert adopted.returncode == 0, adopted.stdout + adopted.stderr
        validated = run(target)
        assert validated.returncode == 0, validated.stdout + validated.stderr
        assert not (target / ".template-sync").exists()
        selected_text = (target / "AGENTS.md").read_text(encoding="utf-8")
        assert selected_text == (
            previous_text
            if case == "existing-skip"
            else (tmp_path / "empty-take/AGENTS.md").read_text(encoding="utf-8")
        )
        generated_path = target / ".github/instruction-profile.yml"
        before = generated_path.read_bytes()
        generated = yaml.safe_load(before)
        assert (
            generated["exceptions"][0]["content_sha256"]
            == hashlib.sha256(selected_text.encode()).hexdigest()
        )
        repeated = subprocess.run(command, check=False, capture_output=True, text=True, timeout=60)
        assert repeated.returncode == 0, repeated.stdout + repeated.stderr
        assert generated_path.read_bytes() == before


@pytest.mark.parametrize("case", ["absent-skip", "missing-candidate", "external-candidate"])
def test_selected_content_keeps_absence_and_path_guards(tmp_path: Path, case: str) -> None:
    """Neither candidate absence nor an escaping symlink invents selected authority."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from instruction_profile_migration import selected_content_root
    from template_sync_materialization_helpers import TemplateSyncMaterializationError

    stage, target = tmp_path / "stage", tmp_path / "target"
    stage.mkdir()
    target.mkdir()
    marker = {
        "included_modules": ["baseline"],
        "protected_file_decisions": [
            {"path": "AGENTS.md", "decision": "SKIP" if case == "absent-skip" else "TAKE"}
        ],
    }
    if case == "absent-skip":
        write(stage, "AGENTS.md", "Candidate must not fill absent SKIP\n")
        assert selected_content_root("AGENTS.md", stage, target, marker) == (target, False)
    elif case == "missing-candidate":
        write(target, "AGENTS.md", "Retained local bytes\n")
        assert selected_content_root("AGENTS.md", stage, target, marker) == (target, False)
    else:
        outside = tmp_path / "outside.md"
        outside.write_text("External bytes\n", encoding="utf-8")
        (stage / "AGENTS.md").symlink_to(outside)
        with pytest.raises(TemplateSyncMaterializationError, match="escapes the repository root"):
            selected_content_root("AGENTS.md", stage, target, marker)


@pytest.mark.parametrize("guard", ["anchor", "digest", "both"])
def test_retained_declaration_guards_reject_before_candidate_write(
    tmp_path: Path, guard: str
) -> None:
    """Removing each compatibility condition restores a false migration success."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    text = "Agents MUST preserve authority.\n"
    for root in (stage, target):
        write(root, "AGENTS.md", text)
    declaration = {
        "path": "AGENTS.md",
        "anchor": "Agents MUST validate.",
        "content_sha256": hashlib.sha256(text.encode()).hexdigest(),
        "reason": "Reviewed local text",
        "authorization_basis": "Fixture owner",
    }
    if guard == "anchor":
        declaration["anchor"] = "Agents MUST preserve authority."
    else:
        declaration["content_sha256"] = "0" * 64
    document["exceptions"] = [declaration]
    write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
    destination = stage / ".github/instruction-profile.yml"
    before = destination.read_bytes()
    retained = (target / ".github/instruction-profile.yml").read_bytes()
    rejected = run_migration_schema_control(stage, target)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert (
        "Existing standalone exception conflicts with selected content: AGENTS.md:"
        in rejected.stderr
    )
    assert "Review or remove the declaration" in rejected.stderr
    assert destination.read_bytes() == before
    assert (target / ".github/instruction-profile.yml").read_bytes() == retained
    mutant = run_migration_schema_control(stage, target, omit_retained_check=guard)
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    assert destination.read_bytes() != before
    with pytest.raises(AssertionError):
        assert mutant.returncode == 1
    with pytest.raises(AssertionError):
        assert destination.read_bytes() == before


def test_retained_declarations_fail_before_real_materializer_writes(tmp_path: Path) -> None:
    """Real adoption rejects stale retained bindings without changing the target."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from template_sync_materialization_helpers import (
        is_protected_instruction_path,
        iter_safe_repository_files,
        parse_manifest_mappings,
        selected_relation_for_path,
    )

    modules = {
        "baseline",
        "agent-instructions",
        "instruction-enforcement",
        "agent-codex",
        "github-actions",
    }
    manifest = yaml.safe_load((ROOT / ".template-sync/manifest.yml").read_text(encoding="utf-8"))
    _, mappings = parse_manifest_mappings(manifest)
    decisions = []
    for relative in iter_safe_repository_files(ROOT)[0]:
        relation = selected_relation_for_path(relative, mappings)
        if (
            relation is not None
            and relation.is_retained_by(modules)
            and is_protected_instruction_path(relative)
        ):
            decisions.append(
                {
                    "path": relative,
                    "decision": "TAKE",
                    "adoption_mode": "minimal-preservation",
                    "authorization_basis": "Fixture owner selects this exact path",
                    "authorized_scope": relative,
                    "reason": "Retained declaration lifecycle",
                }
            )
    marker: dict[str, Any] = {
        "template_sync": {
            "source_repo": "https://github.com/franklesniak/copilot-repo-template",
            "included_modules": sorted(modules),
            "protected_file_decisions": decisions,
            "protected_guide_contract_waivers": [
                {
                    "path": "AGENTS.md",
                    "contract_key": "agents-azure-devops-pr-review-protocol",
                    "target_module": "azure-devops-collaboration",
                    "reason": "Reviewed fixture protocol",
                    "authorization_basis": "Fixture owner retains protocol",
                }
            ],
        }
    }

    def adopt(target: Path, *, omit_guard: bool = False) -> subprocess.CompletedProcess[str]:
        arguments = [
            str(ROOT / ".template-sync/scripts/materialize_downstream_adoption.py"),
            "--template-root",
            str(ROOT),
            "--target-root",
            str(target),
            "--decisions-file",
            "decisions.yml",
        ]
        command = [sys.executable, "-B", *arguments]
        if omit_guard:
            # Change only the reviewed migration module in this child process.
            guard = "(path, anchor) not in failure_keys[content_root] or original_digest != digest"
            code = (
                "import sys, runpy\n"
                f"sys.path.insert(0, {str(ROOT / '.template-sync/scripts')!r})\n"
                "import instruction_profile_migration as migration\n"
                "source = migration.Path(migration.__file__).read_text(encoding='utf-8')\n"
                f"assert source.count({guard!r}) == 1\n"
                f"source = source.replace({guard!r}, 'False')\n"
                "exec(compile(source, migration.__file__, 'exec'), migration.__dict__)\n"
                f"sys.argv = {arguments!r}\n"
                "runpy.run_path(sys.argv[0], run_name='__main__')\n"
            )
            command = [sys.executable, "-B", "-c", code]
        return subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            timeout=60,
        )

    def snapshot(target: Path) -> dict[str, bytes]:
        files, skipped = iter_safe_repository_files(target, skipped_dirs=())
        assert not skipped
        return {relative: (target / relative).read_bytes() for relative in files}

    seed = tmp_path / "seed"
    write(seed, "decisions.yml", yaml.safe_dump(marker))
    initial = adopt(seed)
    assert initial.returncode == 0, initial.stdout + initial.stderr
    validated = run(seed)
    assert validated.returncode == 0, validated.stdout + validated.stderr
    for case in (
        "take-changed",
        "skip-changed",
        "take-same",
        "take-absent",
        "override-absent",
        "skip-stale",
    ):
        target = tmp_path / case
        shutil.copytree(seed, target)
        selection = yaml.safe_load(yaml.safe_dump(marker))
        del selection["template_sync"]["protected_guide_contract_waivers"]
        profile_path = target / ".github/instruction-profile.yml"
        document = yaml.safe_load(profile_path.read_text(encoding="utf-8"))
        relative = "docs/PR_REVIEW_PROMPTS.md" if case == "override-absent" else "AGENTS.md"
        if case in ("take-changed", "skip-changed", "skip-stale"):
            text = "<!-- Reviewed local annotation. -->\n" + (target / relative).read_text(
                encoding="utf-8"
            )
            write(target, relative, text)
            if case != "skip-stale":
                for declaration in document["exceptions"]:
                    if declaration["path"] == relative:
                        declaration["content_sha256"] = hashlib.sha256(text.encode()).hexdigest()
        if "absent" in case:
            (target / relative).unlink()
            document["exceptions"] = [
                item for item in document["exceptions"] if item["path"] != relative
            ]
            document["exceptions"].append(
                {
                    "path": relative,
                    "anchor": "file:absent",
                    "content_sha256": "absent",
                    "reason": "Explicit fixture absence",
                    "authorization_basis": "Fixture owner declaration",
                }
            )
        if case.startswith("skip"):
            next(
                item
                for item in selection["template_sync"]["protected_file_decisions"]
                if item["path"] == relative
            )["decision"] = "SKIP"
        if case == "override-absent":
            selection["template_sync"]["local_overrides"] = [
                {
                    "path": relative,
                    "default_decision": "TAKE",
                    "reason": "Reviewed staged documentation",
                }
            ]
        write(target, ".github/instruction-profile.yml", yaml.safe_dump(document, sort_keys=False))
        write(target, "decisions.yml", yaml.safe_dump(selection))
        before_validation = run(target)
        assert before_validation.returncode == (
            1 if case == "skip-stale" else 0
        ), before_validation.stderr
        before = snapshot(target)
        result = adopt(target)
        assert not (target / ".template-sync").exists()
        if case in ("skip-changed", "take-same"):
            assert result.returncode == 0, result.stdout + result.stderr
            retained = yaml.safe_load(profile_path.read_text(encoding="utf-8"))
            assert retained["exceptions"] == document["exceptions"]
            after_validation = run(target)
            assert after_validation.returncode == 0, (
                after_validation.stdout + after_validation.stderr
            )
            before_repeat = snapshot(target)
            repeated = adopt(target)
            assert repeated.returncode == 0, repeated.stdout + repeated.stderr
            assert snapshot(target) == before_repeat
        else:
            assert result.returncode == 1, result.stdout + result.stderr
            assert (
                f"Existing standalone exception conflicts with selected content: {relative}:"
                in result.stderr
            )
            assert snapshot(target) == before
            if case == "take-changed":
                mutant = adopt(target, omit_guard=True)
                assert mutant.returncode == 0, mutant.stdout + mutant.stderr
                assert snapshot(target) != before
                after_mutation = run(target)
                assert after_mutation.returncode == 1, after_mutation.stderr
                with pytest.raises(AssertionError):
                    assert mutant.returncode == 1
                with pytest.raises(AssertionError):
                    assert snapshot(target) == before


def test_excluded_scope_retires_prior_declaration_before_digest_check(tmp_path: Path) -> None:
    """Known module retirement does not require renewing an obsolete content hash."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    for root in (stage, target):
        catalog_path = root / ".github/instruction-contracts.yml"
        catalog = yaml.safe_load(catalog_path.read_text(encoding="utf-8"))
        catalog["instruction_contracts"][0]["requires_modules"] = ["python"]
        write(root, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
    declaration = {
        "path": "AGENTS.md",
        "anchor": "Agents MUST validate.",
        "content_sha256": "0" * 64,
        "reason": "Old declaration for removed scope",
        "authorization_basis": "Fixture owner",
    }
    document["exceptions"] = [declaration]
    write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
    migrated = run_migration_schema_control(stage, target)
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    destination = stage / ".github/instruction-profile.yml"
    generated = yaml.safe_load(destination.read_text(encoding="utf-8"))
    assert generated["exceptions"] == []
    assert generated["source_decisions"]["retired_exceptions"] == [declaration]
    validated = run(stage)
    assert validated.returncode == 0, validated.stdout + validated.stderr


def enforcement_adoption_decisions(
    modules: set[str], *, skipped_path: str | None = None
) -> dict[str, Any]:
    """Build public decisions for an actual template adoption fixture."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from template_sync_materialization_helpers import (
        is_protected_instruction_path,
        iter_safe_repository_files,
        parse_manifest_mappings,
        selected_relation_for_path,
    )

    manifest = yaml.safe_load((ROOT / ".template-sync/manifest.yml").read_text(encoding="utf-8"))
    _, mappings = parse_manifest_mappings(manifest)
    decisions = []
    for relative in iter_safe_repository_files(ROOT)[0]:
        relation = selected_relation_for_path(relative, mappings)
        if (
            relation
            and relation.is_retained_by(modules)
            and is_protected_instruction_path(relative)
        ):
            decisions.append(
                {
                    "path": relative,
                    "decision": "SKIP" if relative == skipped_path else "TAKE",
                    "adoption_mode": "minimal-preservation",
                    "authorization_basis": "Fixture owner selects this exact path",
                    "authorized_scope": relative,
                    "reason": "Enforcement input lifecycle",
                }
            )
    marker: dict[str, Any] = {
        "included_modules": sorted(modules),
        "source_repo": "https://github.com/franklesniak/copilot-repo-template",
        "protected_file_decisions": decisions,
        "protected_guide_contract_waivers": [
            {
                "path": "AGENTS.md",
                "contract_key": "agents-azure-devops-pr-review-protocol",
                "target_module": "azure-devops-collaboration",
                "reason": "Reviewed fixture protocol",
                "authorization_basis": "Fixture owner retains protocol",
            }
        ],
    }
    if skipped_path is not None and not is_protected_instruction_path(skipped_path):
        marker["local_overrides"] = [
            {
                "path": skipped_path,
                "default_decision": "SKIP",
                "reason": "Fixture owner retains local input",
            }
        ]
    return {"template_sync": marker}


def run_enforcement_adoption(
    target: Path, *, omit_check: str | None = None
) -> subprocess.CompletedProcess[str]:
    """Run adoption, optionally removing one preflight only inside a child process."""
    arguments = [
        str(ROOT / ".template-sync/scripts/materialize_downstream_adoption.py"),
        "--template-root",
        str(ROOT),
        "--target-root",
        str(target),
        "--decisions-file",
        "decisions.yml",
    ]
    command = [sys.executable, "-B", *arguments]
    if omit_check is not None:
        guards = {
            "inputs": "    validate_selected_enforcement_inputs(staging_root, target_root, marker)\n",
            "applicability": "    validate_skipped_profile_applicability(target_root, marker)\n",
        }
        guard = guards[omit_check]
        program = (
            "import sys, runpy\n"
            f"sys.path.insert(0, {str(ROOT / '.template-sync/scripts')!r})\n"
            "import instruction_profile_migration as migration\n"
            "source = migration.Path(migration.__file__).read_text(encoding='utf-8')\n"
            f"assert source.count({guard!r}) == 1\n"
            f"source = source.replace({guard!r}, '')\n"
            "exec(compile(source, migration.__file__, 'exec'), migration.__dict__)\n"
            f"sys.argv = {arguments!r}\n"
            "runpy.run_path(sys.argv[0], run_name='__main__')\n"
        )
        command = [sys.executable, "-B", "-c", program]
    return subprocess.run(command, capture_output=True, text=True, check=False, timeout=90)


def snapshot_enforcement_target(target: Path) -> dict[str, bytes]:
    """Capture every regular fixture file to detect writes before rejection."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    from template_sync_materialization_helpers import iter_safe_repository_files

    paths, skipped = iter_safe_repository_files(target, skipped_dirs=())
    assert not skipped
    return {relative: (target / relative).read_bytes() for relative in paths}


@pytest.mark.parametrize(
    ("relative", "sync"),
    [
        (".github/instruction-profile.yml", False),
        ("schemas/instruction-profile.schema.json", False),
        (".github/scripts/validate_instruction_profile.py", False),
        (".github/scripts/instruction_contract_core.py", False),
        (".github/scripts/instruction_contract_support.py", False),
        (".github/instruction-contracts.yml", False),
        ("schemas/instruction-contracts.schema.json", False),
        (".github/instruction-profile.yml", True),
    ],
)
def test_skipped_missing_enforcement_inputs_reject_before_adoption_writes(
    tmp_path: Path, relative: str, sync: bool
) -> None:
    """A retained hook cannot be installed successfully with a missing selected input."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement", "agent-codex"}
    if sync:
        modules.add("template-sync-support")
    marker = enforcement_adoption_decisions(modules, skipped_path=relative)
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    before = snapshot_enforcement_target(tmp_path)
    result = run_enforcement_adoption(tmp_path)
    assert result.returncode == 1, result.stdout + result.stderr
    assert (
        f"Selected enforcement input is missing or not a regular file: {relative}" in result.stderr
    )
    assert snapshot_enforcement_target(tmp_path) == before


@pytest.mark.parametrize("sync", [False, True])
def test_preserved_profile_directory_rejects_before_adoption_writes(
    tmp_path: Path, sync: bool
) -> None:
    """A wrong-kind profile fails before unrelated files are installed."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement", "agent-codex"}
    if sync:
        modules.add("template-sync-support")
    relative = ".github/instruction-profile.yml"
    marker = enforcement_adoption_decisions(modules, skipped_path=relative)
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    (tmp_path / relative).mkdir(parents=True)
    before = snapshot_enforcement_target(tmp_path)
    result = run_enforcement_adoption(tmp_path)
    assert result.returncode == 1, result.stdout + result.stderr
    assert (
        f"Selected enforcement input is missing or not a regular file: {relative}" in result.stderr
    )
    assert snapshot_enforcement_target(tmp_path) == before


@pytest.mark.parametrize(
    ("case", "sync", "field"),
    [
        ("wrong-mode", False, "mode"),
        ("missing-module", False, "modules"),
        ("extra-module", False, "modules"),
        ("wrong-mode", True, "mode"),
        ("upstream-context", True, "context"),
    ],
)
def test_preserved_profile_applicability_rejects_before_adoption_writes(
    tmp_path: Path, case: str, sync: bool, field: str
) -> None:
    """Schema-valid local applicability cannot silently override selected modules or mode."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement", "agent-codex"}
    if sync:
        modules.add("template-sync-support")
    relative = ".github/instruction-profile.yml"
    marker = enforcement_adoption_decisions(modules, skipped_path=relative)
    local: dict[str, Any] = {
        "version": 1,
        "mode": "standalone",
        "modules": sorted(modules - {"template-sync-support"}),
        "exceptions": [],
    }
    if (case == "wrong-mode" and not sync) or case == "upstream-context":
        local = {
            "version": 1,
            "mode": "marker",
            "context": "upstream-template" if sync else "downstream",
        }
    if case == "missing-module":
        local["modules"].remove("agent-codex")
    if case == "extra-module":
        local["modules"].append("agent-claude")
    write(tmp_path, relative, yaml.safe_dump(local))
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    before = snapshot_enforcement_target(tmp_path)
    result = run_enforcement_adoption(tmp_path)
    assert result.returncode == 1, result.stdout + result.stderr
    assert f"Preserved {relative} {field} conflicts" in result.stderr
    assert snapshot_enforcement_target(tmp_path) == before


@pytest.mark.parametrize("sync", [False, True])
def test_valid_skipped_profile_keeps_local_bytes_and_repeated_adoption(
    tmp_path: Path, sync: bool
) -> None:
    """Owner comments, exceptions and module ordering survive repeated materialization."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement", "agent-codex"}
    if sync:
        modules.add("template-sync-support")
    marker = enforcement_adoption_decisions(modules)
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    initial = run_enforcement_adoption(tmp_path)
    assert initial.returncode == 0, initial.stdout + initial.stderr
    validated = run(tmp_path)
    assert validated.returncode == 0, validated.stdout + validated.stderr
    relative = ".github/instruction-profile.yml"
    local = yaml.safe_load((tmp_path / relative).read_text(encoding="utf-8"))
    if not sync:
        local["modules"].reverse()
    preserved = "# Owner comment must remain.\n" + yaml.safe_dump(local, sort_keys=False)
    write(tmp_path, relative, preserved)
    marker = enforcement_adoption_decisions(modules, skipped_path=relative)
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    for _ in range(2):
        adopted = run_enforcement_adoption(tmp_path)
        assert adopted.returncode == 0, adopted.stdout + adopted.stderr
        assert (tmp_path / relative).read_bytes() == preserved.encode("utf-8")
        validated = run(tmp_path)
        assert validated.returncode == 0, validated.stdout + validated.stderr


def test_policy_only_materialization_does_not_require_profile(tmp_path: Path) -> None:
    """Optional enforcement remains optional in the actual materializer."""
    marker = enforcement_adoption_decisions({"baseline", "agent-instructions", "agent-codex"})
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    result = run_enforcement_adoption(tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    assert not (tmp_path / ".github/instruction-profile.yml").exists()
    assert not (tmp_path / ".github/scripts/validate_instruction_profile.py").exists()


@pytest.mark.parametrize(
    ("omitted", "case"),
    [
        ("inputs", "catalog"),
        ("inputs", "entrypoint"),
        ("applicability", "mode"),
        ("applicability", "modules"),
    ],
)
def test_actual_adoption_preflights_detect_independent_guard_removal(
    tmp_path: Path, omitted: str, case: str
) -> None:
    """Fixed failure oracles catch removed guards, including silent coverage loss."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement", "agent-codex"}
    relative = {
        "catalog": ".github/instruction-contracts.yml",
        "entrypoint": ".github/scripts/validate_instruction_profile.py",
        "mode": ".github/instruction-profile.yml",
        "modules": ".github/instruction-profile.yml",
    }[case]
    marker = enforcement_adoption_decisions(modules, skipped_path=relative)
    if case == "mode":
        local: dict[str, Any] = {"version": 1, "mode": "marker", "context": "downstream"}
        write(tmp_path, relative, yaml.safe_dump(local))
    elif case == "modules":
        local = {
            "version": 1,
            "mode": "standalone",
            "modules": sorted(modules - {"agent-codex"}),
            "exceptions": [],
        }
        write(tmp_path, relative, yaml.safe_dump(local))
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    before = snapshot_enforcement_target(tmp_path)
    rejected = run_enforcement_adoption(tmp_path)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert snapshot_enforcement_target(tmp_path) == before
    expected_diagnostic = (
        f"Selected enforcement input is missing or not a regular file: {relative}"
        if omitted == "inputs"
        else f"Preserved .github/instruction-profile.yml {case} conflicts"
    )
    assert expected_diagnostic in rejected.stderr
    mutant = run_enforcement_adoption(tmp_path, omit_check=omitted)
    if case != "entrypoint":
        # New independent backstops still reject; the removed early guard loses
        # its exact diagnostic, rather than creating a native false success.
        assert mutant.returncode == 1, mutant.stdout + mutant.stderr
        assert expected_diagnostic not in mutant.stderr
        backstop_diagnostic = {
            "catalog": "ERROR: Unable to read .github/instruction-contracts.yml: FileNotFoundError:",
            "mode": "ERROR: Selected preserved instruction profile mode conflicts with standalone migration.",
            "modules": (
                "ERROR: Retained instruction catalog conflicts with selected content: AGENTS.md: "
                "stale:agents-azure-devops-pr-review-protocol:heading:## Azure DevOps PR Review Protocol."
            ),
        }[case]
        assert backstop_diagnostic in mutant.stderr
        assert "Traceback" not in mutant.stderr
        assert snapshot_enforcement_target(tmp_path) == before
        with pytest.raises(AssertionError):
            assert expected_diagnostic in mutant.stderr
        return
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    with pytest.raises(AssertionError):
        assert mutant.returncode == 1
    with pytest.raises(AssertionError):
        assert snapshot_enforcement_target(tmp_path) == before
    if case == "entrypoint":
        assert not (tmp_path / relative).exists()
    else:
        validated = run(tmp_path)
        expected = 0 if case == "modules" else 2 if case == "mode" else 1
        assert validated.returncode == expected, validated.stdout + validated.stderr


def run_schema_reference_control(
    root: Path, *, unsafe_default: bool = False
) -> subprocess.CompletedProcess[str]:
    """Run the deployed CLI behind a public retrieval stub, never an external server."""
    program = (
        "import io, json, runpy, sys, urllib.request\n"
        "calls = []\n"
        "def retrieve(request, *args, **kwargs):\n"
        "    calls.append(request.full_url)\n"
        "    return io.BytesIO(b'{}')\n"
        "urllib.request.urlopen = retrieve\n"
        "sys.dont_write_bytecode = True\n"
        f"if {unsafe_default!r}:\n"
        "    import jsonschema\n"
        "    from referencing import Registry, Resource\n"
        "    from referencing.jsonschema import DRAFT202012\n"
        "    original_validator = jsonschema.Draft202012Validator\n"
        "    def retrieve_schema(uri):\n"
        "        with urllib.request.urlopen(urllib.request.Request(uri)) as response:\n"
        "            return Resource(contents=json.load(response), specification=DRAFT202012)\n"
        "    def dependency_default(schema, *args, **kwargs):\n"
        "        if 'registry' not in kwargs:\n"
        "            kwargs['registry'] = Registry(retrieve=retrieve_schema)\n"
        "        return original_validator(schema, *args, **kwargs)\n"
        "    dependency_default.check_schema = original_validator.check_schema\n"
        "    jsonschema.Draft202012Validator = dependency_default\n"
        "root = sys.argv[1]\n"
        "sys.path.insert(0, root + '/.github/scripts')\n"
        "sys.argv = [root + '/.github/scripts/validate_instruction_profile.py', '--repo-root', root]\n"
        "try:\n"
        "    runpy.run_path(sys.argv[0], run_name='__main__')\n"
        "finally:\n"
        "    print('REFERENCE_RETRIEVALS=' + json.dumps(calls))\n"
    )
    return subprocess.run(
        [sys.executable, "-c", program, str(root)],
        check=False,
        capture_output=True,
        text=True,
        timeout=30,
    )


@pytest.mark.parametrize("slot", ["profile", "contracts"])
@pytest.mark.parametrize("keyword", ["$ref", "$dynamicRef"])
@pytest.mark.parametrize(
    "reference", ["https://schema.invalid/private", "file:///private.json", "other.json"]
)
def test_instruction_schema_external_references_fail_without_retrieval(
    tmp_path: Path, slot: str, keyword: str, reference: str
) -> None:
    """Both standalone schema ingresses reject HTTP, file and relative references."""
    profile(tmp_path, ["agent-instructions", "instruction-enforcement", "github-actions"])
    path = tmp_path / f"schemas/instruction-{slot}.schema.json"
    schema = json.loads(path.read_text(encoding="utf-8"))
    schema["allOf"] = [{keyword: reference}]
    write(tmp_path, path.relative_to(tmp_path).as_posix(), json.dumps(schema))
    result = run_schema_reference_control(tmp_path)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Only local instruction schema references are supported." in result.stderr
    assert "REFERENCE_RETRIEVALS=[]" in result.stdout
    assert "Traceback" not in result.stderr
    assert reference not in result.stderr
    assert not (tmp_path / ".template-sync").exists()
    assert not (tmp_path / ".pre-commit-config.yaml").exists()
    assert not (tmp_path / "pyproject.toml").exists()


@pytest.mark.parametrize(
    ("case", "expected"),
    [
        ("fragment", 0),
        ("anchor", 0),
        ("dynamic-anchor", 0),
        ("annotations", 0),
        ("boolean", 0),
        ("missing-fragment", 1),
        ("invalid-reference-type", 1),
        ("indirect-external", 1),
        ("unused-external", 1),
    ],
)
def test_instruction_schema_reference_boundaries(tmp_path: Path, case: str, expected: int) -> None:
    """Preserve local reuse and annotation data while rejecting unresolved schemas."""
    profile(tmp_path)
    path = tmp_path / "schemas/instruction-profile.schema.json"
    schema = json.loads(path.read_text(encoding="utf-8"))
    if case == "fragment":
        schema["$defs"]["local"] = {"type": "object"}
        schema["allOf"] = [{"$ref": "#/$defs/local"}]
    elif case in {"anchor", "dynamic-anchor"}:
        anchor = "$dynamicAnchor" if case == "dynamic-anchor" else "$anchor"
        schema["$defs"]["local"] = {anchor: "local", "type": "object"}
        schema["allOf"] = [{"$dynamicRef" if case == "dynamic-anchor" else "$ref": "#local"}]
    elif case == "annotations":
        schema["examples"] = [{"$ref": 7, "$dynamicRef": "https://schema.invalid/data"}]
        schema["default"] = {"$ref": "file:///annotation.json"}
    elif case == "boolean":
        schema["allOf"] = [True]
    elif case == "missing-fragment":
        schema["allOf"] = [{"$ref": "#/$defs/missing"}]
    elif case == "invalid-reference-type":
        schema["allOf"] = [{"$ref": 7}]
    elif case == "indirect-external":
        schema["examples"] = [{"$ref": "https://schema.invalid/indirect"}]
        schema["allOf"] = [{"$ref": "#/examples/0"}]
    else:
        schema["$defs"]["unused"] = {"$ref": "https://schema.invalid/unused"}
    write(tmp_path, path.relative_to(tmp_path).as_posix(), json.dumps(schema))
    result = run_schema_reference_control(tmp_path)
    assert result.returncode == expected, result.stdout + result.stderr
    assert "REFERENCE_RETRIEVALS=[]" in result.stdout
    assert "Traceback" not in result.stderr
    if case in {"missing-fragment", "indirect-external"}:
        assert "Unable to resolve an instruction validation schema reference." in result.stderr
    elif case == "invalid-reference-type":
        assert "Invalid instruction validation schema." in result.stderr


@pytest.mark.parametrize("guard", ["preflight", "registry"])
def test_instruction_schema_reference_guards_have_independent_native_oracles(
    tmp_path: Path, guard: str
) -> None:
    """Removing either guard causes a fixed rejected deployment to pass natively."""
    profile(tmp_path)
    path = tmp_path / "schemas/instruction-profile.schema.json"
    schema = json.loads(path.read_text(encoding="utf-8"))
    if guard == "preflight":
        schema["$defs"]["unused"] = {"$ref": "https://schema.invalid/unused"}
    else:
        schema["examples"] = [{"$ref": "https://schema.invalid/indirect"}]
        schema["allOf"] = [{"$ref": "#/examples/0"}]
    write(tmp_path, path.relative_to(tmp_path).as_posix(), json.dumps(schema))
    rejected = run_schema_reference_control(tmp_path, unsafe_default=guard == "registry")
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "REFERENCE_RETRIEVALS=[]" in rejected.stdout
    source_path = tmp_path / ".github/scripts/instruction_contract_support.py"
    source = source_path.read_text(encoding="utf-8")
    anchor = (
        "        validate_local_schema_references(schema, referencing_schema.DRAFT202012)\n"
        if guard == "preflight"
        else ", registry=referencing_module.Registry()"
    )
    assert source.count(anchor) == 1
    source_path.write_text(source.replace(anchor, ""), encoding="utf-8")
    mutant = run_schema_reference_control(tmp_path, unsafe_default=guard == "registry")
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    assert "Traceback" not in mutant.stderr
    expected_calls = "[]" if guard == "preflight" else '["https://schema.invalid/indirect"]'
    assert "REFERENCE_RETRIEVALS=" + expected_calls in mutant.stdout


@pytest.mark.parametrize("nested", [False, True])
@pytest.mark.parametrize("known", [False, True])
def test_instruction_schema_reference_scan_preserves_dialect_semantics(
    tmp_path: Path, nested: bool, known: bool
) -> None:
    """Root remains explicit 2020-12; nested known dialects retain existing evolution."""
    profile(tmp_path)
    path = tmp_path / "schemas/instruction-profile.schema.json"
    schema = json.loads(path.read_text(encoding="utf-8"))
    dialect = (
        "http://json-schema.org/draft-07/schema#" if known else "https://schema.invalid/dialect"
    )
    declaration = {"$schema": dialect, "dependentRequired": {"mode": ["privateMissing"]}}
    if nested:
        schema["allOf"] = [declaration]
    else:
        schema.update(declaration)
    write(tmp_path, path.relative_to(tmp_path).as_posix(), json.dumps(schema))
    result = run_schema_reference_control(tmp_path)
    assert result.returncode == (0 if nested and known else 1), result.stdout + result.stderr
    assert "REFERENCE_RETRIEVALS=[]" in result.stdout
    assert "Traceback" not in result.stderr


def run_catalog_selection_adoption(
    target: Path, *, control: str | None = None
) -> subprocess.CompletedProcess[str]:
    """Exercise the public CLI with independent child-only guard and input controls."""
    program = (
        "import sys, yaml\n"
        f"sys.path.insert(0, {str(ROOT / '.template-sync/scripts')!r})\n"
        "import instruction_profile_migration as migration\n"
        "source = migration.Path(migration.__file__).read_text(encoding='utf-8')\n"
    )
    mutations = {
        "claude-state": (
            "        validate_selected_claude_state(staging_root, target_root, marker, reports)\n",
            "",
        ),
        "claude-selection": (
            "if item.path not in removed and selected_root(item.path) == content_root",
            "if item.path not in removed",
        ),
        "catalog": (
            (
                "        catalog_root, _ = selected_content_root(\n"
                '            ".github/instruction-contracts.yml", staging_root, target_root, marker\n'
                "        )\n"
            ),
            "        catalog_root = staging_root\n",
        ),
        "preflight": (
            (
                "            validate_retained_catalog_selection(\n"
                "                staging_root, target_root, marker, document, reports, path_applicability\n"
                "            )\n"
            ),
            "            pass\n",
        ),
        "waiver": (
            (
                "                declaration_applies(\n"
                "                    waiver, contracts, section_obligations, modules, reference_obligations\n"
            ),
            (
                "                taken and declaration_applies(\n"
                "                    waiver, contracts, section_obligations, modules, reference_obligations\n"
            ),
        ),
        "mode": (
            (
                '    if effective_profile["mode"] != "standalone":\n'
                "        raise TemplateSyncMaterializationError(\n"
                '            "Selected preserved instruction profile mode conflicts with standalone migration. "\n'
                '            "Review the profile or supply an explicit protected selection."\n'
                "        )\n"
            ),
            "",
        ),
        "skipped-profile": (
            (
                " or any(\n"
                '            item["path"] == PROFILE_PATH and item["decision"] == "SKIP"\n'
                '            for item in marker.get("protected_file_decisions", [])\n'
                "        )"
            ),
            "",
        ),
        "removal-scope": (
            (
                "            if not path_applicability.get(path, False):\n"
                "                continue\n"
            ),
            "",
        ),
        "removal-presence": (
            (
                "            if removal_target.exists() or removal_target.is_symlink():\n"
                "                raise TemplateSyncMaterializationError(\n"
                '                    f"Selected protected removal is not complete: {path}. "\n'
                '                    "Complete the reviewed local removal before standalone migration."\n'
                "                )\n"
            ),
            "",
        ),
        "removal": (
            (
                "        removed_paths = {\n"
                '            item["path"]\n'
                '            for item in marker.get("protected_file_decisions", [])\n'
                '            if item["decision"] == "REMOVE-LOCAL"\n'
                "        }\n"
            ),
            "        removed_paths: set[str] = set()\n",
        ),
        "diagnostic": (
            "        raise TemplateSyncMaterializationError(str(error)) from error\n",
            "        raise\n",
        ),
    }
    if control in mutations:
        original, replacement = mutations[control]
        program += (
            f"assert source.count({original!r}) == 1\n"
            f"source = source.replace({original!r}, {replacement!r})\n"
        )
    program += (
        "exec(compile(source, migration.__file__, 'exec'), migration.__dict__)\n"
        "import materialize_downstream_adoption as materializer\n"
    )
    if control in {"semantic", "diagnostic"}:
        program += (
            "original_writer = materializer.write_staged_candidate\n"
            "def prepared(**kwargs):\n"
            "    result = original_writer(**kwargs)\n"
            "    path = kwargs['staging_root'] / '.github/instruction-contracts.yml'\n"
            "    catalog = yaml.safe_load(path.read_text(encoding='utf-8'))\n"
            "    contract = next(c for c in catalog['instruction_contracts'] if c['path'] == 'AGENTS.md')\n"
            "    extra = dict(contract['required_sections'][0])\n"
            "    extra['required_paragraphs'] = ['Distinct schema-valid paragraph']\n"
            "    contract['required_sections'].append(extra)\n"
            "    path.write_text(yaml.safe_dump(catalog), encoding='utf-8')\n"
            "    return result\n"
            "materializer.write_staged_candidate = prepared\n"
        )
    program += (
        f"raise SystemExit(materializer.main(['--template-root', {str(ROOT)!r}, "
        f"'--target-root', {str(target)!r}, '--decisions-file', 'decisions.yml']))\n"
    )
    return subprocess.run(
        [sys.executable, "-B", "-c", program],
        check=False,
        capture_output=True,
        text=True,
        timeout=90,
    )


@pytest.mark.parametrize("case", ["incompatible", "compatible", "waived", "take", "skip-profile"])
def test_selected_local_catalog_controls_native_adoption(tmp_path: Path, case: str) -> None:
    """The installed catalog and effective profile control actual selected guide bytes."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement", "agent-codex"}
    marker = enforcement_adoption_decisions(
        modules, skipped_path=None if case == "take" else ".github/instruction-contracts.yml"
    )
    catalog = yaml.safe_load(
        (ROOT / ".github/instruction-contracts.yml").read_text(encoding="utf-8")
    )
    contract = next(
        item for item in catalog["instruction_contracts"] if item["path"] == "AGENTS.md"
    )
    anchor = (
        "Canonical Instructions" if case == "compatible" else "Local reviewed clause is absent."
    )
    contract.setdefault("required_phrases", []).append(anchor)
    catalog_text = yaml.safe_dump(catalog)
    write(tmp_path, ".github/instruction-contracts.yml", catalog_text)
    if case in {"waived", "skip-profile"}:
        marker["template_sync"]["instruction_contract_waivers"] = [
            {
                "path": "AGENTS.md",
                "anchor": anchor,
                "reason": "Fixture owner reviewed local catalog requirement",
                "authorization_basis": "Explicit private fixture declaration",
            }
        ]
    if case == "skip-profile":
        for decision in marker["template_sync"]["protected_file_decisions"]:
            if decision["path"] == ".github/instruction-profile.yml":
                decision["decision"] = "SKIP"
        write(
            tmp_path,
            ".github/instruction-profile.yml",
            yaml.safe_dump(
                {
                    "version": 1,
                    "mode": "standalone",
                    "modules": sorted(modules),
                    "exceptions": [],
                }
            ),
        )
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    before = snapshot_enforcement_target(tmp_path)
    adopted = run_catalog_selection_adoption(tmp_path)
    if case in {"incompatible", "skip-profile"}:
        assert adopted.returncode == 1, adopted.stdout + adopted.stderr
        assert "Retained instruction catalog conflicts with selected content" in adopted.stderr
        assert anchor in adopted.stderr
        assert snapshot_enforcement_target(tmp_path) == before
        return
    assert adopted.returncode == 0, adopted.stdout + adopted.stderr
    checked = run(tmp_path)
    assert checked.returncode == 0, checked.stdout + checked.stderr
    assert not (tmp_path / ".template-sync").exists()
    if case != "take":
        assert (tmp_path / ".github/instruction-contracts.yml").read_text(
            encoding="utf-8"
        ) == catalog_text
    profile_bytes = (tmp_path / ".github/instruction-profile.yml").read_bytes()
    repeated = run_catalog_selection_adoption(tmp_path)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert (tmp_path / ".github/instruction-profile.yml").read_bytes() == profile_bytes


@pytest.mark.parametrize("guard", ["catalog", "preflight"])
def test_retained_catalog_guards_have_independent_native_oracles(
    tmp_path: Path, guard: str
) -> None:
    """Removing either guard restores adoption success with a failing installed hook."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement", "agent-codex"}
    marker = enforcement_adoption_decisions(
        modules, skipped_path=".github/instruction-contracts.yml"
    )
    catalog = yaml.safe_load(
        (ROOT / ".github/instruction-contracts.yml").read_text(encoding="utf-8")
    )
    contract = next(
        item for item in catalog["instruction_contracts"] if item["path"] == "AGENTS.md"
    )
    contract.setdefault("required_phrases", []).append("Fixed unwaived local obligation.")
    write(tmp_path, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    before = snapshot_enforcement_target(tmp_path)
    rejected = run_catalog_selection_adoption(tmp_path)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "Retained instruction catalog conflicts" in rejected.stderr
    assert snapshot_enforcement_target(tmp_path) == before
    mutant = run_catalog_selection_adoption(tmp_path, control=guard)
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    assert snapshot_enforcement_target(tmp_path) != before
    invalid = run(tmp_path)
    assert invalid.returncode == 1, invalid.stdout + invalid.stderr
    assert "Fixed unwaived local obligation." in invalid.stdout
    with pytest.raises(AssertionError):
        assert mutant.returncode == 1


@pytest.mark.parametrize("active", [False, True])
def test_direct_skip_waivers_follow_current_selected_failures(tmp_path: Path, active: bool) -> None:
    """Satisfied SKIP waivers reject; exact missing clauses remain valid declarations."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    document = profile(stage)
    profile(target)
    if active:
        write(target, "AGENTS.md", "Agents MUST preserve authority.\n")
    marker = {
        "template_sync": {
            "included_modules": document["modules"],
            "protected_file_decisions": [
                {
                    "path": path,
                    "decision": decision,
                    "adoption_mode": "minimal-preservation",
                    "authorization_basis": "Fixture owner",
                    "authorized_scope": path,
                    "reason": "Reviewed selected input",
                }
                for path, decision in (
                    ("AGENTS.md", "SKIP"),
                    (".github/instruction-contracts.yml", "TAKE"),
                    (".github/instruction-profile.yml", "TAKE"),
                )
            ],
            "instruction_contract_waivers": [
                {
                    "path": "AGENTS.md",
                    "anchor": "Agents MUST validate.",
                    "reason": "Reviewed exact selected clause",
                    "authorization_basis": "Fixture owner",
                }
            ],
        }
    }
    before = (stage / ".github/instruction-profile.yml").read_bytes()
    result = run_migration_schema_control(stage, target, marker_document=marker)
    if not active:
        assert result.returncode == 1, result.stdout + result.stderr
        assert "Instruction waiver conflicts with selected preserved content" in result.stderr
        assert (stage / ".github/instruction-profile.yml").read_bytes() == before
        return
    assert result.returncode == 0, result.stdout + result.stderr
    write(stage, "AGENTS.md", (target / "AGENTS.md").read_text(encoding="utf-8"))
    checked = run(stage)
    assert checked.returncode == 0, checked.stdout + checked.stderr
    assert not (stage / ".template-sync").exists()


def test_direct_skip_waiver_guard_catches_native_false_success(tmp_path: Path) -> None:
    """A satisfied known paragraph cannot become an active standalone exception."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement", "agent-codex"}
    marker = enforcement_adoption_decisions(modules)
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    initial = run_enforcement_adoption(tmp_path)
    assert initial.returncode == 0, initial.stdout + initial.stderr
    for decision in marker["template_sync"]["protected_file_decisions"]:
        if decision["path"] == "AGENTS.md":
            decision["decision"] = "SKIP"
    marker["template_sync"]["instruction_contract_waivers"] = [
        {
            "path": "AGENTS.md",
            "anchor": "section:## Execution:paragraph:2eea7e4977d47083984c652d3bfcbccce8d48317ca6c14354736ec3c672e9894",
            "reason": "Reviewed fixed paragraph",
            "authorization_basis": "Fixture owner",
        }
    ]
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    before = snapshot_enforcement_target(tmp_path)
    rejected = run_catalog_selection_adoption(tmp_path)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "Instruction waiver conflicts with selected preserved content" in rejected.stderr
    assert snapshot_enforcement_target(tmp_path) == before
    mutant = run_catalog_selection_adoption(tmp_path, control="waiver")
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    invalid = run(tmp_path)
    assert invalid.returncode == 1, invalid.stdout + invalid.stderr
    assert "Exception does not match a current failure and exact content" in invalid.stderr
    with pytest.raises(AssertionError):
        assert mutant.returncode == 1


def test_migration_core_errors_preserve_native_diagnostic_boundary(tmp_path: Path) -> None:
    """Semantic core failures remain nonzero without leaking an uncaught traceback."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement", "agent-codex"}
    write(tmp_path, "decisions.yml", yaml.safe_dump(enforcement_adoption_decisions(modules)))
    before = snapshot_enforcement_target(tmp_path)
    rejected = run_catalog_selection_adoption(tmp_path, control="semantic")
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "ERROR: Duplicate required section: ## Execution" in rejected.stderr
    assert "Traceback" not in rejected.stderr
    assert snapshot_enforcement_target(tmp_path) == before
    mutant = run_catalog_selection_adoption(tmp_path, control="diagnostic")
    assert mutant.returncode == 1, mutant.stdout + mutant.stderr
    assert "Duplicate required section: ## Execution" in mutant.stderr
    assert "Traceback" in mutant.stderr
    assert snapshot_enforcement_target(tmp_path) == before
    with pytest.raises(AssertionError):
        assert "Traceback" not in mutant.stderr


def test_unresolved_preserved_marker_profile_fails_without_traceback(tmp_path: Path) -> None:
    """An unresolved marker profile cannot be treated as an authorized standalone candidate."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement", "agent-codex"}
    marker = enforcement_adoption_decisions(
        modules, skipped_path=".github/instruction-contracts.yml"
    )
    marker["template_sync"]["protected_file_decisions"] = [
        item
        for item in marker["template_sync"]["protected_file_decisions"]
        if item["path"] != ".github/instruction-profile.yml"
    ]
    write(
        tmp_path,
        ".github/instruction-contracts.yml",
        (ROOT / ".github/instruction-contracts.yml").read_text(encoding="utf-8"),
    )
    write(
        tmp_path,
        ".github/instruction-profile.yml",
        yaml.safe_dump(
            {
                "version": 1,
                "mode": "marker",
                "context": "downstream",
            }
        ),
    )
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    before = snapshot_enforcement_target(tmp_path)
    rejected = run_catalog_selection_adoption(tmp_path)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "ERROR: Selected preserved instruction profile mode conflicts" in rejected.stderr
    assert "Traceback" not in rejected.stderr
    assert snapshot_enforcement_target(tmp_path) == before
    mutant = run_catalog_selection_adoption(tmp_path, control="mode")
    assert mutant.returncode == 1, mutant.stdout + mutant.stderr
    assert "KeyError: 'exceptions'" in mutant.stderr
    with pytest.raises(AssertionError):
        assert "Traceback" not in mutant.stderr


@pytest.mark.parametrize("existing", [False, True])
def test_selected_removal_preserves_original_absence_declaration(
    tmp_path: Path, existing: bool
) -> None:
    """Authorized final absence preserves the original declaration without renewing a hash."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement", "agent-codex"}
    marker = enforcement_adoption_decisions(
        modules, skipped_path=".github/instruction-contracts.yml"
    )
    for item in marker["template_sync"]["protected_file_decisions"]:
        if item["path"] == "AGENTS.md":
            item["decision"] = "REMOVE-LOCAL"
    declaration = {
        "path": "AGENTS.md",
        "anchor": "file:absent",
        "content_sha256": "absent",
        "reason": "Original absence rationale",
        "authorization_basis": "Original fixture owner",
    }
    write(
        tmp_path,
        ".github/instruction-profile.yml",
        yaml.safe_dump(
            {
                "version": 1,
                "mode": "standalone",
                "modules": sorted(modules),
                "exceptions": [declaration],
            }
        ),
    )
    write(
        tmp_path,
        ".github/instruction-contracts.yml",
        yaml.safe_dump(
            {
                "instruction_contracts": [
                    {
                        "path": "AGENTS.md",
                        "requires_modules": ["agent-instructions", "agent-codex"],
                        "required_phrases": ["Reviewed absence fixture."],
                    }
                ],
            }
        ),
    )
    if existing:
        write(tmp_path, "AGENTS.md", "Content selected for explicit removal.\n")
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    before = snapshot_enforcement_target(tmp_path)
    if existing:
        rejected = run_catalog_selection_adoption(tmp_path)
        assert rejected.returncode == 1, rejected.stdout + rejected.stderr
        assert "Selected protected removal is not complete: AGENTS.md" in rejected.stderr
        assert snapshot_enforcement_target(tmp_path) == before
        mutant = run_catalog_selection_adoption(tmp_path, control="removal-presence")
        assert mutant.returncode == 0, mutant.stdout + mutant.stderr
        assert (tmp_path / "AGENTS.md").read_bytes() == before["AGENTS.md"]
        invalid = run(tmp_path)
        assert invalid.returncode == 1, invalid.stdout + invalid.stderr
        assert "Exception does not match a current failure and exact content" in invalid.stderr
        with pytest.raises(AssertionError):
            assert mutant.returncode == 1
        return
    mutant = run_catalog_selection_adoption(tmp_path, control="removal")
    assert mutant.returncode == 1, mutant.stdout + mutant.stderr
    assert "Existing standalone exception conflicts" in mutant.stderr
    assert snapshot_enforcement_target(tmp_path) == before
    adopted = run_catalog_selection_adoption(tmp_path)
    assert adopted.returncode == 0, adopted.stdout + adopted.stderr
    assert not (tmp_path / "AGENTS.md").exists()
    generated = yaml.safe_load(
        (tmp_path / ".github/instruction-profile.yml").read_text(encoding="utf-8")
    )
    assert generated["exceptions"] == [declaration]
    checked = run(tmp_path)
    assert checked.returncode == 0, checked.stdout + checked.stderr
    assert not (tmp_path / ".template-sync").exists()
    repeated = run_catalog_selection_adoption(tmp_path)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert yaml.safe_load(
        (tmp_path / ".github/instruction-profile.yml").read_text(encoding="utf-8")
    )["exceptions"] == [declaration]


@pytest.mark.parametrize("declaration_kind", ["wrong-absence-digest", "nonabsence"])
def test_selected_removal_does_not_renew_invalid_declarations(
    tmp_path: Path, declaration_kind: str
) -> None:
    """Removal does not make invalid hashes or old content-specific anchors valid."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    write(target, "AGENTS.md", "Agents MUST preserve authority.\n")
    declaration = {
        "path": "AGENTS.md",
        "anchor": "file:absent",
        "content_sha256": "0" * 64,
        "reason": "Unchanged reviewed record",
        "authorization_basis": "Fixture owner",
    }
    if declaration_kind == "nonabsence":
        declaration["anchor"] = "Agents MUST validate."
        declaration["content_sha256"] = hashlib.sha256(
            (target / "AGENTS.md").read_bytes()
        ).hexdigest()
    document["exceptions"] = [declaration]
    write(target, ".github/instruction-profile.yml", yaml.safe_dump(document))
    marker = {
        "template_sync": {
            "included_modules": document["modules"],
            "protected_file_decisions": [
                {"path": "AGENTS.md", "decision": "REMOVE-LOCAL"},
            ],
        }
    }
    # The owner has completed removal; old content-specific authority remains invalid.
    (target / "AGENTS.md").unlink()
    before = (stage / ".github/instruction-profile.yml").read_bytes()
    rejected = run_migration_schema_control(stage, target, marker_document=marker)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "Existing standalone exception conflicts" in rejected.stderr
    assert (stage / ".github/instruction-profile.yml").read_bytes() == before
    assert not (target / "AGENTS.md").exists()


@pytest.mark.parametrize("stale", [False, True])
def test_skipped_profile_with_taken_catalog_uses_effective_exceptions(
    tmp_path: Path, stale: bool
) -> None:
    """A discarded generated profile cannot retire the retained profile's exception."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement"}
    marker = enforcement_adoption_decisions(modules, skipped_path=".github/instruction-profile.yml")
    # No agent-codex means the new catalog excludes AGENTS; its prior exception is stale.
    exceptions = (
        [
            {
                "path": "AGENTS.md",
                "anchor": "file:absent",
                "content_sha256": "absent",
                "reason": "Earlier catalog clause",
                "authorization_basis": "Fixture owner",
            }
        ]
        if stale
        else []
    )
    write(
        tmp_path,
        ".github/instruction-profile.yml",
        yaml.safe_dump(
            {
                "version": 1,
                "mode": "standalone",
                "modules": sorted(modules),
                "exceptions": exceptions,
            }
        ),
    )
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    before = snapshot_enforcement_target(tmp_path)
    result = run_catalog_selection_adoption(tmp_path)
    if not stale:
        assert result.returncode == 0, result.stdout + result.stderr
        validated = run(tmp_path)
        assert validated.returncode == 0, validated.stdout + validated.stderr
        return
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Retained instruction catalog conflicts with selected profile exception" in result.stderr
    assert snapshot_enforcement_target(tmp_path) == before
    mutant = run_catalog_selection_adoption(tmp_path, control="skipped-profile")
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    invalid = run(tmp_path)
    assert invalid.returncode == 1, invalid.stdout + invalid.stderr
    assert "Exception does not match a current failure and exact content" in invalid.stderr
    with pytest.raises(AssertionError):
        assert mutant.returncode == 1


@pytest.mark.parametrize("explicit", ["SKIP", "REMOVE-LOCAL", "TAKE"])
def test_candidate_installation_fixture_preserves_explicit_decisions(explicit: str) -> None:
    """Synthetic defaults never override an explicit profile or catalog selection."""
    marker = {
        "template_sync": {
            "included_modules": ["baseline"],
            "protected_file_decisions": [
                {"path": path, "decision": explicit, "reason": "Exact original record"}
                for path in (
                    ".github/instruction-profile.yml",
                    ".github/instruction-contracts.yml",
                    "AGENTS.md",
                )
            ],
        }
    }
    before = json.dumps(marker)
    selected = candidate_installation_marker(marker)
    assert selected == marker
    assert selected is not marker
    assert json.dumps(marker) == before


@pytest.mark.parametrize(
    "path",
    [
        ".github/instructions/local-owner-note.md",
        ".github/instructions/yaml.instructions.md",
    ],
)
def test_unrelated_or_inactive_removal_keeps_existing_reconciliation(
    tmp_path: Path, path: str
) -> None:
    """Only applicable enforcement absence requires completed local removal."""
    modules = {"baseline", "agent-instructions", "instruction-enforcement", "agent-codex"}
    marker = enforcement_adoption_decisions(modules)
    decision = {
        "path": path,
        "decision": "REMOVE-LOCAL",
        "authorized_scope": path,
        "authorization_basis": "Fixture owner",
        "reason": "Separate recorded local cleanup",
        "adoption_mode": "minimal-preservation",
    }
    marker["template_sync"]["protected_file_decisions"].append(decision)
    local_text = "Local content outside currently applicable instruction contracts.\n"
    write(tmp_path, path, local_text)
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    before = snapshot_enforcement_target(tmp_path)
    mutant = run_catalog_selection_adoption(tmp_path, control="removal-scope")
    assert mutant.returncode == 1, mutant.stdout + mutant.stderr
    assert f"Selected protected removal is not complete: {path}" in mutant.stderr
    assert snapshot_enforcement_target(tmp_path) == before
    with pytest.raises(AssertionError):
        assert mutant.returncode == 0
    adopted = run_catalog_selection_adoption(tmp_path)
    assert adopted.returncode == 0, adopted.stdout + adopted.stderr
    assert (tmp_path / path).read_bytes() == before[path]
    document = yaml.safe_load(
        (tmp_path / ".github/instruction-profile.yml").read_text(encoding="utf-8")
    )
    assert decision in document["source_decisions"]["protected_file_decisions"]
    assert all(item["path"] != path for item in document["exceptions"])
    validated = run(tmp_path)
    assert validated.returncode == 0, validated.stdout + validated.stderr
    assert not (tmp_path / ".template-sync").exists()


def selected_claude_adoption_fixture(target: Path) -> dict[str, Any]:
    """Install valid standalone Claude content before testing a selected local change."""
    modules = {
        "baseline",
        "agent-instructions",
        "instruction-enforcement",
        "agent-claude",
        "azure-devops-collaboration",
    }
    marker = enforcement_adoption_decisions(modules)
    write(target, "decisions.yml", yaml.safe_dump(marker))
    adopted = run_enforcement_adoption(target)
    assert adopted.returncode == 0, adopted.stdout + adopted.stderr
    validated = run(target)
    assert validated.returncode == 0, validated.stdout + validated.stderr
    assert not (target / ".template-sync").exists()
    for decision in marker["template_sync"]["protected_file_decisions"]:
        if decision["path"] == "CLAUDE.md":
            decision["decision"] = "SKIP"
    return marker


@pytest.mark.parametrize("selection", ["take", "catalog-skip", "profile-skip"])
@pytest.mark.parametrize("kind", ["import", "root-memory", "nested-memory"])
def test_selected_claude_state_rejects_before_native_adoption(
    tmp_path: Path, selection: str, kind: str
) -> None:
    """Every standalone output rejects non-exceptable selected content before writes."""
    marker = selected_claude_adoption_fixture(tmp_path)
    if kind == "import":
        path = tmp_path / "CLAUDE.md"
        write(tmp_path, "CLAUDE.md", "@private-policy.md\n\n" + path.read_text(encoding="utf-8"))
        expected = "Active Claude imports:"
    else:
        memory = "sub/CLAUDE.local.md" if kind == "nested-memory" else "CLAUDE.local.md"
        write(tmp_path, memory, "Private fixture instruction.\n")
        for command in (
            ["git", "init", "--quiet"],
            ["git", "-c", "core.hooksPath=NUL", "add", "--force", "--", memory],
        ):
            initialized = subprocess.run(
                command, cwd=tmp_path, capture_output=True, text=True, check=False
            )
            assert initialized.returncode == 0, initialized.stdout + initialized.stderr
        expected = "Tracked Claude local memory:"
    if selection != "take":
        selected_path = (
            ".github/instruction-contracts.yml"
            if selection == "catalog-skip"
            else ".github/instruction-profile.yml"
        )
        for decision in marker["template_sync"]["protected_file_decisions"]:
            if decision["path"] == selected_path:
                decision["decision"] = "SKIP"
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    before = snapshot_enforcement_target(tmp_path)
    rejected = run_catalog_selection_adoption(tmp_path)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "Selected non-exceptable Claude instruction content" in rejected.stderr
    assert "Traceback" not in rejected.stderr
    assert snapshot_enforcement_target(tmp_path) == before
    mutant = run_catalog_selection_adoption(tmp_path, control="claude-state")
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    if selection == "profile-skip":
        assert snapshot_enforcement_target(tmp_path) == before
    else:
        assert snapshot_enforcement_target(tmp_path) != before
    invalid = run(tmp_path)
    assert invalid.returncode == 1, invalid.stdout + invalid.stderr
    assert expected in invalid.stdout
    assert "Missing required anchors:" not in invalid.stdout
    assert "Stale protected-guide sections" not in invalid.stdout
    assert not (tmp_path / ".template-sync").exists()
    with pytest.raises(AssertionError):
        assert mutant.returncode == 1


@pytest.mark.parametrize("kind", ["clean", "discarded", "inline-code", "fenced-code", "untracked"])
def test_selected_claude_state_preserves_valid_native_adoption(tmp_path: Path, kind: str) -> None:
    """Selected clean bytes and literal/untracked examples preserve standalone no-op behavior."""
    marker = selected_claude_adoption_fixture(tmp_path)
    if kind == "untracked":
        write(tmp_path, "CLAUDE.local.md", "Private untracked local instructions.\n")
        initialized = subprocess.run(
            ["git", "init", "--quiet"], cwd=tmp_path, capture_output=True, text=True, check=False
        )
        assert initialized.returncode == 0, initialized.stdout + initialized.stderr
    elif kind != "clean":
        prefix = {
            "discarded": "@private-policy.md\n\n",
            "inline-code": "`@private-policy.md`\n\n",
            "fenced-code": "```text\n@private-policy.md\n```\n\n",
        }[kind]
        path = tmp_path / "CLAUDE.md"
        write(tmp_path, "CLAUDE.md", prefix + path.read_text(encoding="utf-8"))
        if kind == "discarded":
            for decision in marker["template_sync"]["protected_file_decisions"]:
                if decision["path"] == "CLAUDE.md":
                    decision["decision"] = "TAKE"
    write(tmp_path, "decisions.yml", yaml.safe_dump(marker))
    if kind == "discarded":
        before = snapshot_enforcement_target(tmp_path)
        mutant = run_catalog_selection_adoption(tmp_path, control="claude-selection")
        assert mutant.returncode == 1, mutant.stdout + mutant.stderr
        assert "Selected non-exceptable Claude instruction content" in mutant.stderr
        assert snapshot_enforcement_target(tmp_path) == before
        with pytest.raises(AssertionError):
            assert mutant.returncode == 0
    adopted = run_catalog_selection_adoption(tmp_path)
    assert adopted.returncode == 0, adopted.stdout + adopted.stderr
    validated = run(tmp_path)
    assert validated.returncode == 0, validated.stdout + validated.stderr
    assert not (tmp_path / ".template-sync").exists()
    profile_bytes = (tmp_path / ".github/instruction-profile.yml").read_bytes()
    repeated = run_catalog_selection_adoption(tmp_path)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert (tmp_path / ".github/instruction-profile.yml").read_bytes() == profile_bytes
    repeated_validation = run(tmp_path)
    assert repeated_validation.returncode == 0, (
        repeated_validation.stdout + repeated_validation.stderr
    )


@pytest.mark.parametrize("selected", [False, True])
def test_staged_claude_import_respects_catalog_module_selection(
    tmp_path: Path, selected: bool
) -> None:
    """The staged-content check reuses applicable core reports, including excluded contracts."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    modules = ["baseline", "agent-instructions", "instruction-enforcement"]
    if selected:
        modules.append("agent-claude")
    for root in (stage, target):
        profile(root, modules)
        write(
            root,
            ".github/instruction-contracts.yml",
            yaml.safe_dump(
                {
                    "instruction_contracts": [
                        {
                            "path": "CLAUDE.md",
                            "requires_modules": ["agent-instructions", "agent-claude"],
                            "required_phrases": ["Retained Claude rule."],
                        }
                    ]
                }
            ),
        )
        write(root, "CLAUDE.md", "Retained Claude rule.\n")
    write(stage, "CLAUDE.md", "@private-policy.md\n\nRetained Claude rule.\n")
    marker = candidate_installation_marker(
        {
            "template_sync": {
                "included_modules": modules,
                "protected_file_decisions": [
                    {
                        "path": "CLAUDE.md",
                        "decision": "TAKE",
                        "adoption_mode": "minimal-preservation",
                        "authorization_basis": "Fixture owner selects staged instructions",
                        "authorized_scope": "CLAUDE.md",
                        "reason": "Exact candidate selection",
                    }
                ],
            }
        }
    )
    before = (stage / ".github/instruction-profile.yml").read_bytes()
    rendered = run_migration_schema_control(stage, target, marker_document=marker)
    assert rendered.returncode == (1 if selected else 0), rendered.stdout + rendered.stderr
    if selected:
        assert "Selected non-exceptable Claude instruction content" in rendered.stderr
        assert (stage / ".github/instruction-profile.yml").read_bytes() == before
    else:
        valid = run(stage)
        assert valid.returncode == 0, valid.stdout + valid.stderr
