"""Native schema-output boundaries for protected profile migration."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path
from typing import Any

import jsonschema
import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest
from tests.test_instruction_profile import (
    ROOT,
    candidate_installation_marker,
    profile,
    run,
    write,
)

pytestmark = pytest.mark.upstream_template_only

OUTPUT_GUARD = """    core.support.validate_schema(
        document,
        load_reviewed_schema("schemas/instruction-profile.schema.json"),
        destination,
        staging_root,
    )
"""


def run_output_migration(
    stage: Path,
    target: Path,
    marker: dict[str, Any],
    *,
    remove_guard: bool = False,
    remove_byte_guard: bool = False,
) -> subprocess.CompletedProcess[str]:
    """Run the actual renderer, preserving native exits and installed schema trust."""
    source_path = ROOT / ".template-sync/scripts/instruction_profile_migration.py"
    source = source_path.read_text(encoding="utf-8")
    assert source.count(OUTPUT_GUARD) == 1
    if remove_guard:
        source = source.replace(OUTPUT_GUARD, "")
    if remove_byte_guard:
        guard = '    if len(rendered.encode("utf-8")) > core.MAXIMUM_INPUT_BYTES:\n'
        assert source.count(guard) == 1
        source = source.replace(guard, "    if False:  # Remove only the serialized-size gate.\n")
    marker_path = target.parent / "marker-input.json"
    marker_path.write_text(json.dumps(marker), encoding="utf-8")
    program = (
        "import json, sys\n"
        "from pathlib import Path\n"
        f"sys.path.insert(0, {str(source_path.parent)!r})\n"
        f"namespace = {{'__file__': {str(source_path)!r}, '__name__': 'profile_output_control'}}\n"
        f"exec(compile({source!r}, {str(source_path)!r}, 'exec'), namespace)\n"
        "try:\n"
        "    namespace['render_instruction_profile'](\n"
        "        staging_root=Path(sys.argv[1]), target_root=Path(sys.argv[2]),\n"
        "        marker_document=json.loads(Path(sys.argv[3]).read_text(encoding='utf-8')))\n"
        "except namespace['TemplateSyncMaterializationError'] as error:\n"
        "    print(str(error), file=sys.stderr)\n"
        "    raise SystemExit(1)\n"
    )
    return subprocess.run(
        [sys.executable, "-c", program, str(stage), str(target), str(marker_path)],
        check=False,
        capture_output=True,
        text=True,
    )


def output_fixture(root: Path, count: int, reason_length: int) -> tuple[Path, Path, dict[str, Any]]:
    """Build real missing obligations with marker-schema-valid waiver declarations."""
    stage, target = root / "stage", root / "target"
    profile(stage)
    document = profile(target)
    phrases = [f"Required obligation number {number}." for number in range(count)]
    for directory in (stage, target):
        write(directory, "AGENTS.md", "Retained fixture file.\n")
        write(
            directory,
            ".github/instruction-contracts.yml",
            yaml.safe_dump(
                {
                    "instruction_contracts": [
                        {
                            "path": "AGENTS.md",
                            "requires_modules": ["agent-instructions"],
                            "required_phrases": phrases,
                        }
                    ]
                }
            ),
        )
    waivers = [
        {
            "path": "AGENTS.md",
            "anchor": phrase,
            "reason": "r" * reason_length,
            "authorization_basis": "Explicit test fixture",
        }
        for phrase in phrases
    ]
    marker_schema = json.loads(
        (ROOT / "schemas/template-sync-marker.schema.json").read_text(encoding="utf-8")
    )
    jsonschema.Draft202012Validator(
        {
            "$defs": marker_schema["$defs"],
            "type": "array",
            "items": {"$ref": "#/$defs/instructionContractWaiver"},
        }
    ).validate(waivers)
    return (
        stage,
        target,
        candidate_installation_marker(
            {
                "template_sync": {
                    "included_modules": document["modules"],
                    "instruction_contract_waivers": waivers,
                }
            }
        ),
    )


@pytest.mark.parametrize(
    ("count", "reason_length", "accepted"),
    [(256, 1, True), (257, 1, False), (1, 4096, True), (1, 4097, False)],
)
def test_migration_checks_complete_output_before_writing(
    tmp_path: Path, count: int, reason_length: int, accepted: bool
) -> None:
    """Fixed schema boundaries reject invalid candidates without changing either profile."""
    stage, target, marker = output_fixture(tmp_path, count, reason_length)
    destination = stage / ".github/instruction-profile.yml"
    previous = target / ".github/instruction-profile.yml"
    staged_before, target_before = destination.read_bytes(), previous.read_bytes()
    result = run_output_migration(stage, target, marker)
    assert result.returncode == (0 if accepted else 1), result.stdout + result.stderr
    schema = json.loads(
        (ROOT / "schemas/instruction-profile.schema.json").read_text(encoding="utf-8")
    )
    validator = jsonschema.Draft202012Validator(schema)
    if accepted:
        rendered = yaml.safe_load(destination.read_text(encoding="utf-8"))
        validator.validate(rendered)
        assert len(rendered["exceptions"]) == count
    else:
        assert "Schema validation failed for .github/instruction-profile.yml" in result.stderr
        assert destination.read_bytes() == staged_before
        mutant = run_output_migration(stage, target, marker, remove_guard=True)
        assert mutant.returncode == 0, mutant.stdout + mutant.stderr
        rendered = yaml.safe_load(destination.read_text(encoding="utf-8"))
        assert not validator.is_valid(rendered)
        assert len(rendered["exceptions"]) == count
    assert previous.read_bytes() == target_before


def test_schema_output_gate_preserves_valid_marker_mode(tmp_path: Path) -> None:
    """The same output gate accepts the renderer's explicit marker context."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    profile(stage)
    document = profile(target)
    marker = {
        "template_sync": {"included_modules": [*document["modules"], "template-sync-support"]}
    }
    result = run_output_migration(stage, target, marker)
    assert result.returncode == 0, result.stdout + result.stderr
    rendered = yaml.safe_load(
        (stage / ".github/instruction-profile.yml").read_text(encoding="utf-8")
    )
    assert rendered == {"version": 1, "mode": "marker", "context": "downstream"}


@pytest.mark.parametrize(("count", "accepted"), [(120, True), (128, False)])
def test_schema_valid_profile_output_obeys_existing_reader_byte_limit(
    tmp_path: Path, count: int, accepted: bool
) -> None:
    """Preserved declaration history cannot make successful output unreadable."""
    stage, target, marker = output_fixture(tmp_path, count, 4096)
    destination = stage / ".github/instruction-profile.yml"
    previous = target / ".github/instruction-profile.yml"
    staged_before, target_before = destination.read_bytes(), previous.read_bytes()
    result = run_output_migration(stage, target, marker)
    assert result.returncode == (0 if accepted else 1), result.stdout + result.stderr
    if not accepted:
        assert "Rendered .github/instruction-profile.yml exceeds" in result.stderr
        assert destination.read_bytes() == staged_before
        mutant = run_output_migration(stage, target, marker, remove_byte_guard=True)
        assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    raw = destination.read_bytes()
    assert (len(raw) <= 1_048_576) is accepted
    rendered = yaml.safe_load(raw)
    schema = json.loads(
        (ROOT / "schemas/instruction-profile.schema.json").read_text(encoding="utf-8")
    )
    jsonschema.Draft202012Validator(schema).validate(rendered)
    assert len(rendered["exceptions"]) == count
    assert len(rendered["source_decisions"]["instruction_contract_waivers"]) == count
    deployed = run(stage)
    assert deployed.returncode == (0 if accepted else 1), deployed.stdout + deployed.stderr
    if not accepted:
        assert "exceeds the 1048576-byte input limit" in deployed.stderr
    assert previous.read_bytes() == target_before
