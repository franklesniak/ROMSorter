"""Exercise instruction-contract validation for protected agent protocols."""

from __future__ import annotations

import ast
import copy
import hashlib
import importlib
import json
import os
import re
import shutil
import subprocess
import sys
import tomllib
from pathlib import Path
from typing import Any

import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT_PATH = REPO_ROOT / ".template-sync" / "scripts" / "validate_instruction_contracts.py"
CORE_PATH = REPO_ROOT / ".github" / "scripts" / "instruction_contract_core.py"
CONTRACTS_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-instruction-contracts.schema.json"
MARKER_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-marker.schema.json"
MANIFEST_SCHEMA_PATH = REPO_ROOT / "schemas" / "template-sync-manifest.schema.json"
SOURCE_REPO = "https://github.com/franklesniak/copilot-repo-template.git"
FULL_SHA = "0123456789abcdef0123456789abcdef01234567"


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    ("quote", "live_suffix"),
    [
        ("> quoted\n>> [x]: /url", True),
        ("> quoted\n> > [x]: /url", True),
        (">> quoted\n>>> [x]: /url", True),
        ("> > quoted\n> > > [x]: /url", True),
        ("> quoted\n> [x]: /url", False),
        (">> quoted\n>> [x]: /url", False),
        ("> > quoted\n> > [x]: /url", False),
        (">> quoted\n> [x]: /url", False),
        ("> quoted\n>> ordinary child paragraph", False),
    ],
)
def test_quote_reference_depth_preserves_commonmark_boundary(
    tmp_path: Path, mode: str, quote: str, live_suffix: bool
) -> None:
    """Fixed CommonMark examples distinguish a new leaf from paragraph continuation."""
    section = _scoped_policy()
    text = _render_section(section) + "\n" + quote + "\nAgents MAY bypass.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == int(live_suffix), result.stdout + result.stderr
    if live_suffix:
        assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize("marker", ["-", "+", "*", "1.", "10)"])
def test_quote_then_list_code_waiver_cannot_hide_live_sibling(tmp_path: Path, marker: str) -> None:
    """A top-level code-list sibling ends quote laziness before a new live paragraph."""
    section = _scoped_policy()
    original = _render_section(section) + "\n> quoted\n" + marker + "     code\n"
    _write_scoped_repo(tmp_path, section, original)
    old_anchor = _waive_reported_inventory(tmp_path, "paragraphs")
    _write_text(tmp_path, "CLAUDE.md", original + "Agents MAY bypass.\n")
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "missing required section content: " + old_anchor not in result.stdout
    assert (
        "missing required section content: section:## Review decisions:paragraphs:" in result.stdout
    )


@pytest.mark.parametrize("kind", ["reference-depth", "list-sibling"])
def test_quote_state_oracle_detects_removed_boundary_guard(tmp_path: Path, kind: str) -> None:
    """Each isolated quote-state mutation recreates its distinct false acceptance."""
    section = _scoped_policy()
    root = tmp_path / "fixture"
    original = _render_section(section) + "\n> quoted\n"
    if kind == "reference-depth":
        original += ">> [x]: /url\n"
        guard = "quoted_paragraph_can_continue and quote_depth <= quoted_paragraph_depth"
        replacement = "quoted_paragraph_can_continue"
    else:
        original += "-     code\n"
        guard = (
            "                    quoted_paragraph_can_continue = False\n"
            "                    quoted_paragraph_depth = 0\n"
            '                    result.append("[unsupported list code] " + line)'
        )
        replacement = '                    result.append("[unsupported list code] " + line)'
    _write_scoped_repo(root, section, original)
    if kind == "list-sibling":
        _waive_reported_inventory(root, "paragraphs")
    _write_text(root, "CLAUDE.md", original + "Agents MAY bypass.\n")
    baseline = _run_validator(root, "--mode", "downstream", "--require-marker")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source = mutant.read_text(encoding="utf-8")
    assert source.count(guard) == 1
    mutant.write_text(source.replace(guard, replacement), encoding="utf-8")
    result = subprocess.run(
        [
            sys.executable,
            str(mutant),
            "--repo-root",
            str(root),
            "--mode",
            "downstream",
            "--require-marker",
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    if kind == "list-sibling":
        assert "passed with waivers" in result.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("prefix", ["", "- "])
@pytest.mark.parametrize(
    ("expected", "observed"),
    [
        ("See [policy](docs/policy.md).", "See [policy]<!--x-->(docs/policy.md)."),
        ("See ![diagram](image.png).", "See ![diagram]<!--x-->(image.png)."),
        ("See <https://example.com>.", "See <https:<!--x-->//example.com>."),
        ("Use **strong** text.", "Use *<!--x-->*strong** text."),
        ("Use &amp; safely.", "Use &amp<!--x-->; safely."),
        ("Use &#38; safely.", "Use &#3<!--x-->8; safely."),
        ("See [policy][rules].", "See [policy]<!--x-->[rules]."),
        ("See [policy][].", "See [policy]<!--x-->[]."),
    ],
)
def test_inline_comments_cannot_synthesize_markdown(
    tmp_path: Path, mode: str, prefix: str, expected: str, observed: str
) -> None:
    """A source token boundary cannot be deleted to satisfy expected Markdown."""
    section: dict[str, Any] = {
        "heading": "## Rules",
        "next_heading": None,
        "required_paragraphs": [prefix + expected],
    }
    _write_scoped_repo(tmp_path, section, section["heading"] + "\n\n" + prefix + observed)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr

    assert "section:## Rules:paragraphs:" in result.stdout


def test_inline_comment_source_retention_and_normalized_identity() -> None:
    """Source tokens survive code precedence and keep the existing normalization contract."""
    program = r"""
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)
for clause in [
    "See [policy]<!--x-->(docs/policy.md).",
    "- See [policy]<!--x-->(docs/policy.md).",
    "Use `<!--x-->` literally.",
    "Use `code`<!-- note --> safely.",
]:
    section = validator.RequiredSection("## Rules", (clause,), ())
    assert validator.section_failures("## Rules\n\n" + clause, (section,)) == [], clause
section = validator.RequiredSection("## Rules", ("Act.",), ())
def failures(comment):
    return validator.section_failures("## Rules\n\nAct." + comment, (section,))
assert failures("<!--x-->") != failures("<!--y-->")
assert failures("<!-- x -->") == failures("<!--  x -->") == failures("<!--\tx -->")
# Even explicitly catalogued comments cannot make table rows supported.
table = validator.RequiredTable(("State", "Action"), (("Pending<!--x-->", "Wait"),))
section = validator.RequiredSection("## Rules", (), (table,))
text = "## Rules\n\n| State | Action |\n| --- | --- |\n| Pending<!--x--> | Wait |"
assert validator.section_failures(text, (section,))
print("Literal comments, code precedence, table restriction and normalized identities passed.")
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_inline_token_oracle_detects_removed_comment_retention(tmp_path: Path) -> None:
    """Removing retention alone falsely recreates a required link and is detected."""
    section: dict[str, Any] = {
        "heading": "## Rules",
        "next_heading": None,
        "required_paragraphs": ["See [policy](docs/policy.md)."],
    }
    root = tmp_path / "fixture"
    _write_scoped_repo(root, section, "## Rules\n\nSee [policy]<!--x-->(docs/policy.md).")
    baseline = _run_validator(root, "--mode", "downstream")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    assert "section:## Rules:paragraphs:" in baseline.stdout
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source = mutant.read_text(encoding="utf-8")
    guard = "visible.append(line[column : end + 3])"
    assert source.count(guard) == 1
    mutant.write_text(source.replace(guard, "pass"), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(mutant), "--repo-root", str(root), "--mode", "downstream"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_large_section_validation_uses_bounded_line_work() -> None:
    """Many short valid sections must not rescan the complete line sequence."""
    program = r"""
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)

class CountedLines(list):
    work = 0
    def __iter__(self):
        for value in super().__iter__():
            self.work += 1
            yield value
    def __getitem__(self, key):
        value = super().__getitem__(key)
        self.work += len(value) if isinstance(key, slice) else 1
        return value

count = 400
lines = CountedLines()
sections = []
for number in range(count):
    heading = f"## Section {number}"
    following = f"## Section {number + 1}" if number + 1 < count else None
    lines.extend([heading, "", "Act.", ""])
    sections.append(validator.RequiredSection(heading, ("Act.",), (), following))
# This public scanner seam isolates repeated document traversal from lexing.
validator.operative_markdown_lines = lambda text: lines
assert validator.section_failures("fixture", tuple(sections)) == []
assert lines.work <= 12 * len(lines), (lines.work, len(lines))
print("Bounded document work:", lines.work)
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_successor_catalog_and_optional_chains_avoid_repeated_walks() -> None:
    """Shared acyclic suffixes are traversed once without a wall-clock assertion."""
    program = r"""
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)

class CountedHeading(str):
    hashes = 0
    def __hash__(self):
        type(self).hashes += 1
        return super().__hash__()

count = 400
headings = [CountedHeading(f"## Section {number}") for number in range(count)]
raw = {"required_sections": [
    {"heading": value, "next_heading": headings[number + 1] if number + 1 < count else None,
     "required_paragraphs": ["Act."], "requires_modules": ["optional"]}
    for number, value in enumerate(headings)
]}
sections = validator.parse_required_sections(raw)
assert CountedHeading.hashes <= 30 * count, CountedHeading.hashes
mapping = {section.heading: section for section in sections}
CountedHeading.hashes = 0
cache = {}
for section in sections:
    boundary = validator.applicable_section_boundary(section, mapping, set(), set(), cache)
    assert boundary.next_heading is None
assert CountedHeading.hashes <= 20 * count, CountedHeading.hashes
# A present optional boundary must still stop the same traversal.
cache = {}
boundary = validator.applicable_section_boundary(sections[0], mapping, {headings[2]}, set(), cache)
assert boundary.next_heading == headings[2]
print("Bounded successor work and present optional boundary passed.")
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("field", ["heading", "next_heading"])
@pytest.mark.parametrize("suffix", [" ", "\t", "\n", "\r"])
def test_catalog_heading_whitespace_fails_loading(
    tmp_path: Path, mode: str, field: str, suffix: str
) -> None:
    """Impossible expectations fail as catalog errors, not document drift."""
    section = _scoped_policy()
    section["next_heading"] = "## End"
    text = _render_section(section) + "\n## End\n"
    section[field] += suffix
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "instruction-contracts.yml" in result.stderr
    assert ":boundary:" not in result.stdout


def test_catalog_heading_schema_and_semantic_grammar_agree() -> None:
    """Both consumers reject full-value defects and retain literal supported spelling."""
    program = r"""
import json, sys
from pathlib import Path
from jsonschema import Draft202012Validator
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)
schema = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))
heading_schema = schema["$defs"]["policyHeading"]
assert heading_schema["pattern"] == validator.POLICY_HEADING_PATTERN.pattern
cases = [
    ("# X", True), ("###### X", True), ("## Two  words", True),
    ("## X\tY", True), ("## \u00c9tat \u4e2d\u6587\u00a0", True),
    ("##", False), ("####### X", False), ("## ", False), ("## \t", False),
    ("## X ", False), ("## X\t", False), ("## X\n", False), ("## X\r", False),
    ("## X\nY", False), ("## X\rY", False),
]
for field in ("heading", "next_heading"):
    for value, expected in cases:
        section = {"heading": "## Rules", "next_heading": "## End",
                   "required_paragraphs": ["Act."]}
        section[field] = value
        document = {"instruction_contracts": [
            {"path": "AGENTS.md", "requires_modules": ["agent-instructions"],
             "required_sections": [section]}]}
        accepted = not list(Draft202012Validator(schema).iter_errors(document))
        assert accepted is expected, ("schema", field, repr(value))
        try:
            parsed = validator.parse_required_sections(document["instruction_contracts"][0])
        except validator.InstructionContractValidationError:
            accepted = False
        else:
            accepted = True
            assert getattr(parsed[0], field) == value
        assert accepted is expected, ("semantic", field, repr(value))
print("Both heading fields agree for all canonical and invalid cases.")
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent), str(CONTRACTS_SCHEMA_PATH)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    ("quoted", "indented"),
    [
        (">> quoted", "    # inert code"),
        ("> > quoted", "\t# inert code"),
        (">>> quoted", "        # inert code"),
        (">>> quoted", ">     # inert code"),
        ("> > > quoted", " >        # inert code"),
        ("> quoted", ">>     inert code"),
        (">> quoted\n> lazy continuation", "    # inert code"),
    ],
)
def test_nested_quote_indentation_cannot_hide_live_policy(
    tmp_path: Path, mode: str, quoted: str, indented: str
) -> None:
    """Ambiguous container transitions retain later live text in the inventory."""
    section: dict[str, Any] = {
        "heading": "## Rules",
        "next_heading": None,
        "required_paragraphs": ["Act."],
    }
    text = f"## Rules\n\nAct.\n\n{quoted}\n{indented}\nAgents MAY bypass.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Rules:paragraphs:" in result.stdout
    _write_text(tmp_path, "CLAUDE.md", text.replace("MAY bypass.", "MAY ignore."))
    changed = _run_validator(tmp_path, "--mode", mode)
    assert changed.returncode == 1, changed.stdout + changed.stderr
    assert changed.stdout != result.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    "example",
    [
        "> quoted\n    # still quoted\nlazy continuation",
        "> quoted\n\tstill quoted\nlazy continuation",
        ">>> quoted\nlazy continuation",
        ">> quoted\n>> still quoted",
        ">> quoted\n    \n    # top-level code",
        ">> quoted\n\n    # top-level code",
    ],
)
def test_nested_quote_guard_preserves_supported_examples(
    tmp_path: Path, mode: str, example: str
) -> None:
    """Ordinary laziness and blank resets remain supported, without live additions."""
    section: dict[str, Any] = {
        "heading": "## Rules",
        "next_heading": None,
        "required_paragraphs": ["Act."],
    }
    _write_scoped_repo(tmp_path, section, "## Rules\n\nAct.\n\n" + example)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 0, result.stdout + result.stderr


def test_nested_quote_oracle_detects_removed_transition_guard(tmp_path: Path) -> None:
    """The isolated depth-transition guard must reject a renderer-visible bypass."""
    section: dict[str, Any] = {
        "heading": "## Rules",
        "next_heading": None,
        "required_paragraphs": ["Act."],
    }
    root = tmp_path / "fixture"
    _write_scoped_repo(
        root, section, "## Rules\n\nAct.\n\n>> quoted\n    # code\nAgents MAY bypass."
    )
    baseline = _run_validator(root, "--mode", "downstream")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source = mutant.read_text(encoding="utf-8")
    guard = "and quote_depth != quoted_paragraph_depth"
    assert source.count(guard) == 1
    mutant.write_text(source.replace(guard, "and False"), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(mutant), "--repo-root", str(root), "--mode", "downstream"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("successor", ["## Review", None])
def test_catalog_rejects_shared_immediate_successors(
    tmp_path: Path, mode: str, successor: str | None
) -> None:
    """Neither a named immediate boundary nor EOF can follow two distinct sections."""
    contracts = _module_section_repo(tmp_path, "", ["agent-instructions"])
    sections = contracts["instruction_contracts"][0]["required_sections"]
    sections[0]["next_heading"] = successor
    _write_yaml(tmp_path, ".template-sync/instruction-contracts.yml", contracts)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Shared section successor" in result.stderr
    assert ":boundary:" not in result.stdout


@pytest.mark.parametrize("successor", ["## End", None])
def test_shared_successor_oracle_detects_removed_loader_guard(
    tmp_path: Path, successor: str | None
) -> None:
    """Removing just shared-successor rejection makes the invalid catalog load again."""
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source = mutant.read_text(encoding="utf-8")
    guard = "if next_heading in successor_owners:"
    assert source.count(guard) == 1
    mutant.write_text(source.replace(guard, "if False:"), encoding="utf-8")
    program = """
import json, sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)
successor = json.loads(sys.argv[2])
raw = {"required_sections": [
    {"heading": heading, "next_heading": successor, "required_paragraphs": ["Act."]}
    for heading in ["## A", "## B"]
]}
validator.parse_required_sections(raw)
print("Catalog accepted")
"""
    for directory, expected in [(SCRIPT_PATH.parent, 1), (mutant_dir, 0)]:
        result = subprocess.run(
            [sys.executable, "-c", program, str(directory), json.dumps(successor)],
            check=False,
            capture_output=True,
            text=True,
        )
        assert result.returncode == expected, result.stdout + result.stderr
        if expected:
            assert "Shared section successor" in result.stderr
        else:
            assert result.stdout.strip() == "Catalog accepted"


def _canonical_host_contract_fixture(
    tmp_path: Path, modules: list[str]
) -> tuple[dict[str, Any], str]:
    """Use the actual canonical contract while retaining an independent literal deletion."""
    catalog = yaml.safe_load(
        (REPO_ROOT / ".template-sync/instruction-contracts.yml").read_text(encoding="utf-8")
    )
    contract = next(
        item
        for item in catalog["instruction_contracts"]
        if item["path"] == ".github/copilot-instructions.md"
    )
    contracts = {"instruction_contracts": [contract]}
    _write_common_contract_repo(tmp_path, contracts)
    _write_yaml(tmp_path, ".template-sync/marker.yml", _marker(modules))
    text = (REPO_ROOT / ".github/copilot-instructions.md").read_text(encoding="utf-8")
    _write_text(tmp_path, ".github/copilot-instructions.md", text)
    return contracts, text


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    "change", ["host-paragraph", "azure-section", "azure-heading", "azure-clause"]
)
def test_actual_canonical_host_protocol_rejects_loss(
    tmp_path: Path, mode: str, change: str
) -> None:
    """The original host role needs complete body ownership, not only a boundary heading."""
    _contracts_data, text = _canonical_host_contract_fixture(
        tmp_path, ["agent-instructions", "azure-devops-collaboration"]
    )
    intact = _run_validator(tmp_path, "--mode", mode)
    assert intact.returncode == 0, intact.stdout + intact.stderr
    azure = "### Azure DevOps Services with Azure Repos"
    following = "## Linting and Validation Configurations"
    if change == "azure-section":
        text = text[: text.index(azure)] + text[text.index(following) :]
    elif change == "azure-heading":
        text = text.replace(azure, "", 1)
    elif change == "azure-clause":
        text = text.replace(
            "- The repository must be a Git repository in Azure Repos; TFVC is not supported.",
            "",
            1,
        )
    else:
        text = text.replace(
            "agents MUST NOT rename, weaken, or replace the GitHub protocol",
            "agents MAY replace the GitHub protocol",
            1,
        )
    _write_text(tmp_path, ".github/copilot-instructions.md", text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:" in result.stdout


@pytest.mark.upstream_template_only
def test_actual_canonical_host_protocol_preserves_excluded_host(tmp_path: Path) -> None:
    """Excluded Azure sections may be absent while the primary-host relationship remains."""
    _contracts_data, text = _canonical_host_contract_fixture(tmp_path, ["agent-instructions"])
    azure = "### Azure DevOps Services with Azure Repos"
    following = "## Linting and Validation Configurations"
    text = text[: text.index(azure)] + text[text.index(following) :]
    _write_text(tmp_path, ".github/copilot-instructions.md", text)
    accepted = _run_validator(tmp_path, "--mode", "downstream")
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    text = text.replace(
        "agents MUST NOT rename, weaken, or replace the GitHub protocol",
        "agents MAY replace the GitHub protocol",
        1,
    )
    _write_text(tmp_path, ".github/copilot-instructions.md", text)
    rejected = _run_validator(tmp_path, "--mode", "downstream")
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("section", ["host", "azure"])
def test_canonical_host_clause_oracle_detects_removed_catalog_ownership(
    tmp_path: Path, section: str
) -> None:
    """Deleting one catalog owner must resurrect its independently deleted-clause defect."""
    contracts, text = _canonical_host_contract_fixture(
        tmp_path, ["agent-instructions", "azure-devops-collaboration"]
    )
    if section == "host":
        heading = "## Host-Specific PR Review Protocols"
        text = text.replace(
            "agents MUST NOT rename, weaken, or replace the GitHub protocol",
            "agents MAY replace the GitHub protocol",
            1,
        )
    else:
        heading = "### Azure DevOps Services with Azure Repos"
        text = text.replace(
            "- The repository must be a Git repository in Azure Repos; TFVC is not supported.",
            "",
            1,
        )
    _write_text(tmp_path, ".github/copilot-instructions.md", text)
    baseline = _run_validator(tmp_path, "--mode", "downstream")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    sections = contracts["instruction_contracts"][0]["required_sections"]
    assert sum(item["heading"] == heading for item in sections) == 1
    sections[:] = [item for item in sections if item["heading"] != heading]
    _write_yaml(tmp_path, ".template-sync/instruction-contracts.yml", contracts)
    mutant = _run_validator(tmp_path, "--mode", "downstream")
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr


BLOCK_ORDER_TABLE = "| State | Action |\n| --- | --- |\n| Pending | Wait |"


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("change", ["intact", "cross-paragraph", "swap-tables"])
def test_multiple_tables_preserve_interleaved_order(tmp_path: Path, mode: str, change: str) -> None:
    """Distinct tables retain both their own order and their relation to prose."""
    second = BLOCK_ORDER_TABLE.replace("Pending", "Complete").replace("Wait", "Stop")
    section = {
        "heading": "## Rules",
        "next_heading": None,
        "required_paragraphs": ["Before.", "Between.", "After."],
        "required_tables": [
            {"headers": ["State", "Action"], "rows": [["Pending", "Wait"]]},
            {"headers": ["State", "Action"], "rows": [["Complete", "Stop"]]},
        ],
        "required_block_order": ["paragraph", "table", "paragraph", "table", "paragraph"],
    }
    parts = ["Before.", BLOCK_ORDER_TABLE, "Between.", second, "After."]
    if change == "cross-paragraph":
        parts = ["Before.", "Between.", BLOCK_ORDER_TABLE, second, "After."]
    elif change == "swap-tables":
        parts = ["Before.", second, "Between.", BLOCK_ORDER_TABLE, "After."]
    _write_scoped_repo(tmp_path, section, "## Rules\n\n" + "\n\n".join(parts))
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == (0 if change == "intact" else 1), result.stdout + result.stderr
    if change != "intact":
        assert ":blocks:" in result.stdout
    if change == "swap-tables":
        assert ":tables:" in result.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("boundary", ["heading", "list", "table"])
def test_hard_break_guard_preserves_block_final_padding(
    tmp_path: Path, mode: str, boundary: str
) -> None:
    """Spaces before a new block do not turn a final line into a hard break."""
    section: dict[str, Any] = {
        "heading": "## Rules",
        "next_heading": None,
        "required_paragraphs": ["Act."],
    }
    if boundary == "heading":
        section["next_heading"] = "## Next"
        following = "## Next\n\nUncontracted text."
    elif boundary == "list":
        section["required_paragraphs"].append("- Continue.")
        following = "- Continue."
    else:
        section["required_tables"] = [
            {"headers": ["State", "Action"], "rows": [["Pending", "Wait"]]}
        ]
        following = BLOCK_ORDER_TABLE
    _write_scoped_repo(tmp_path, section, "## Rules\n\nAct.  \n" + following)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 0, result.stdout + result.stderr


def _interleaved_policy() -> dict[str, Any]:
    """Use independent before/table/after source in ordering tests."""
    return {
        "heading": "## Rules",
        "next_heading": None,
        "required_paragraphs": ["Before.", "After."],
        "required_tables": [{"headers": ["State", "Action"], "rows": [["Pending", "Wait"]]}],
        "required_block_order": ["paragraph", "table", "paragraph"],
    }


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("position", ["first", "middle", "last"])
def test_mixed_policy_blocks_keep_table_placement(tmp_path: Path, mode: str, position: str) -> None:
    """Moving an unchanged table across prose requires an ordering waiver."""
    parts = {
        "first": [BLOCK_ORDER_TABLE, "Before.", "After."],
        "middle": ["Before.", BLOCK_ORDER_TABLE, "After."],
        "last": ["Before.", "After.", BLOCK_ORDER_TABLE],
    }
    _write_scoped_repo(
        tmp_path, _interleaved_policy(), "## Rules\n\n" + "\n\n".join(parts[position])
    )
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == (0 if position == "middle" else 1), result.stdout + result.stderr
    if position != "middle":
        assert "section:## Rules:blocks:" in result.stdout
        assert "section:## Rules:tables:" not in result.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    "order",
    [
        [],
        ["paragraph", "table"],
        ["paragraph", "table", "paragraph", "table"],
        ["paragraph", "unknown", "paragraph"],
    ],
)
def test_catalog_rejects_invalid_block_order(tmp_path: Path, mode: str, order: list[str]) -> None:
    """A sequence cannot omit, duplicate or invent inventory kinds."""
    section = _interleaved_policy()
    section["required_block_order"] = order
    _write_scoped_repo(tmp_path, section, "## Rules\n\nBefore.")
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    if order and all(kind in {"paragraph", "table"} for kind in order):
        assert "Block order must match paragraph and table counts" in result.stderr
    else:
        assert "schema" in result.stderr.lower()


def test_default_block_order_is_enforced_and_missing_clause_stays_individual(
    tmp_path: Path,
) -> None:
    """Omission means prose then tables; missing clauses keep separate failures."""
    section = _interleaved_policy()
    section.pop("required_block_order")
    _write_scoped_repo(tmp_path, section, "## Rules\n\nBefore.\n\nAfter.\n\n" + BLOCK_ORDER_TABLE)
    accepted = _run_validator(tmp_path, "--mode", "downstream")
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    _write_text(tmp_path, "CLAUDE.md", "## Rules\n\n" + BLOCK_ORDER_TABLE + "\n\nBefore.\n\nAfter.")
    rejected = _run_validator(tmp_path, "--mode", "downstream")
    assert rejected.returncode == 1 and ":blocks:" in rejected.stdout
    _write_scoped_repo(
        tmp_path, _interleaved_policy(), "## Rules\n\n" + BLOCK_ORDER_TABLE + "\n\nAfter."
    )
    missing = _run_validator(tmp_path, "--mode", "downstream")
    assert missing.returncode == 1 and ":paragraph:" in missing.stdout
    assert ":blocks:" not in missing.stdout


@pytest.mark.parametrize(
    "table",
    [BLOCK_ORDER_TABLE.replace("Wait", "Ignore"), BLOCK_ORDER_TABLE.replace("---", "bad", 1)],
)
def test_table_waiver_identity_includes_cross_type_placement(tmp_path: Path, table: str) -> None:
    """Changed or malformed table waivers cannot survive relocation."""
    identities = []
    for parts in [["Before.", table, "After."], ["Before.", "After.", table]]:
        _write_scoped_repo(tmp_path, _interleaved_policy(), "## Rules\n\n" + "\n\n".join(parts))
        result = _run_validator(tmp_path, "--mode", "downstream")
        assert result.returncode == 1, result.stdout + result.stderr
        identities.append(re.findall(r"section:## Rules:tables:[0-9a-f]+", result.stdout))
    assert identities[0] and identities[1] and identities[0] != identities[1]


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
@pytest.mark.parametrize("spaces", ["  ", "   "])
def test_space_hard_breaks_are_not_soft_wrapping(
    tmp_path: Path, mode: str, ending: str, spaces: str
) -> None:
    """Hard breaks keep content-bound failures in both CLI modes."""
    section = {
        "heading": "## Rules",
        "next_heading": None,
        "required_paragraphs": ["Act. Continue. Finish."],
    }
    text = ("## Rules\n\nAct." + spaces + "\nContinue. Finish.").replace("\n", ending)
    _write_scoped_repo(tmp_path, section, text)
    # Avoid Windows text translation turning an explicit CRLF into CR-CRLF.
    (tmp_path / "CLAUDE.md").write_bytes(text.encode("utf-8"))
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    first = re.findall(r"section:## Rules:hard-break:[0-9a-f]+", result.stdout)
    assert first
    changed_text = ("## Rules\n\nAct. Continue." + spaces + "\nFinish.").replace("\n", ending)
    (tmp_path / "CLAUDE.md").write_bytes(changed_text.encode("utf-8"))
    changed = _run_validator(tmp_path, "--mode", mode)
    assert changed.returncode == 1, changed.stdout + changed.stderr
    assert re.findall(r"section:## Rules:hard-break:[0-9a-f]+", changed.stdout) != first


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    ("clause", "source"),
    [
        ("Act. Continue.", "Act.\nContinue."),
        ("Act. Continue.", "Act. \nContinue."),
        ("Act. Continue.", "Act.\t\nContinue."),
        ("Act. Continue.", "Act. Continue.   "),
        ("Use \x60literal example\x60.", "Use \x60literal  \nexample\x60."),
        ("Act.", "Act.  \n\n> quoted  \n> example"),
        ("Act.", "Act.  \n\n~~~\ncode  \nexample\n~~~"),
        ("Act.", "Act.  \n\n<!-- comment  \nexample -->"),
    ],
)
def test_hard_break_guard_preserves_soft_and_inert_controls(
    tmp_path: Path, mode: str, clause: str, source: str
) -> None:
    """Wrapping, block-final padding and examples remain supported."""
    section = {"heading": "## Rules", "next_heading": None, "required_paragraphs": [clause]}
    _write_scoped_repo(tmp_path, section, "## Rules\n\n" + source)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("case", ["blocks", "hard-break"])
def test_order_and_hard_break_oracles_detect_removed_guards(tmp_path: Path, case: str) -> None:
    """Removing one guard restores its independently constructed false clean."""
    root = tmp_path / "fixture"
    if case == "blocks":
        section = _interleaved_policy()
        text = "## Rules\n\nBefore.\n\nAfter.\n\n" + BLOCK_ORDER_TABLE
        guard, replacement = "if observed_ranks != sorted(observed_ranks):", "if False:"
    else:
        section = {
            "heading": "## Rules",
            "next_heading": None,
            "required_paragraphs": ["Act. Continue."],
        }
        text = "## Rules\n\nAct.  \nContinue."
        guard, replacement = (
            'line.endswith("  ")',
            "False",
        )
    _write_scoped_repo(root, section, text)
    baseline = _run_validator(root, "--mode", "downstream")
    assert baseline.returncode == 1 and f":{case}:" in baseline.stdout
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source = mutant.read_text(encoding="utf-8")
    assert source.count(guard) == 1
    mutant.write_text(source.replace(guard, replacement), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(mutant), "--repo-root", str(root), "--mode", "downstream"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
def test_paragraph_waivers_cannot_authorize_block_relocation(tmp_path: Path, mode: str) -> None:
    """An accepted altered clause stays bound to its original side of a table."""
    section = _interleaved_policy()
    positioned = "## Rules\n\nBefore.\n\n" + BLOCK_ORDER_TABLE + "\n\nAltered."
    moved = "## Rules\n\nBefore.\n\nAltered.\n\n" + BLOCK_ORDER_TABLE
    _write_scoped_repo(tmp_path, section, positioned)
    baseline = _run_validator(tmp_path, "--mode", mode)
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    anchors = re.findall(r"section:## Rules:paragraphs?:[0-9a-f]{64}", baseline.stdout)
    assert len(anchors) == 2
    if mode == "downstream":
        _write_yaml(
            tmp_path,
            ".template-sync/marker.yml",
            _marker(
                ["agent-instructions"],
                waivers=[
                    {
                        "path": "CLAUDE.md",
                        "anchor": anchor,
                        "reason": "Fixture owner permits Altered. after the table.",
                        "authorization_basis": "Explicit fixture grant for this exact content and placement.",
                    }
                    for anchor in anchors
                ],
            ),
        )
        accepted = _run_validator(tmp_path, "--mode", mode)
        assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    _write_text(tmp_path, "CLAUDE.md", moved)
    rejected = _run_validator(tmp_path, "--mode", mode)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    changed = re.findall(r"section:## Rules:paragraphs:[0-9a-f]{64}", rejected.stdout)
    assert changed and changed[0] not in anchors


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("thematic", ["***", "___"])
@pytest.mark.parametrize("separation", ["\n", "\n\n"])
@pytest.mark.parametrize("padding", ["before", "after", "both"])
def test_thematic_boundaries_preserve_harmless_padding(
    tmp_path: Path, mode: str, thematic: str, separation: str, padding: str
) -> None:
    """The admitted thematic source representation keeps block-final spaces harmless."""
    clauses = [f"Act. {thematic} After."] if separation == "\n" else ["Act.", thematic, "After."]
    section = {"heading": "## Rules", "next_heading": None, "required_paragraphs": clauses}
    parts = [
        "Act." + ("  " if padding in {"before", "both"} else ""),
        thematic + ("  " if padding in {"after", "both"} else ""),
        "After.",
    ]
    _write_scoped_repo(tmp_path, section, "## Rules\n\n" + separation.join(parts))
    accepted = _run_validator(tmp_path, "--mode", mode)
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    # A real break later in the same admitted source remains independently rejected.
    clauses[-1] += " Continue."
    text = "## Rules\n\n" + separation.join(parts) + "  \nContinue."
    _write_scoped_repo(tmp_path, section, text)
    rejected = _run_validator(tmp_path, "--mode", mode)
    assert rejected.returncode == 1 and ":hard-break:" in rejected.stdout


@pytest.mark.parametrize("case", ["paragraph-identity", "thematic-boundary", "block-count"])
def test_inventory_and_boundary_oracles_detect_removed_guards(tmp_path: Path, case: str) -> None:
    """Independent properties fail when their specific guard is removed."""
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source = mutant.read_text(encoding="utf-8")
    if case == "paragraph-identity":
        guard = "identity = policy_inventory_digest(\n                serialized_policy_blocks(expected_blocks), observed_tables\n            )"
        replacement = "identity = policy_inventory_digest(expected_paragraphs, paragraphs)"
    elif case == "thematic-boundary":
        guard = "and not is_policy_thematic_break(following)"
        replacement = "and True"
    else:
        guard = 'if "required_block_order" in raw and ('
        replacement = "if False and ("
    assert source.count(guard) == 1
    mutant.write_text(source.replace(guard, replacement), encoding="utf-8")
    program = r"""
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as v
v = sys.modules.get("instruction_contract_core", v)
case = sys.argv[2]
table = v.RequiredTable(("State", "Action"), (("Pending", "Wait"),))
section = v.RequiredSection("## Rules", ("Before.", "After."), (table,), None, (), ("paragraph", "table", "paragraph"))
literal = "| State | Action |\n| --- | --- |\n| Pending | Wait |"
if case == "paragraph-identity":
    positioned = v.section_failures("## Rules\n\nBefore.\n\n" + literal + "\n\nAltered.", (section,))
    moved = v.section_failures("## Rules\n\nBefore.\n\nAltered.\n\n" + literal, (section,))
    assert positioned != moved, (positioned, moved)
elif case == "thematic-boundary":
    section = v.RequiredSection("## Rules", ("Act. *** After.",), ())
    assert v.section_failures("## Rules\n\nAct.  \n***\nAfter.", (section,)) == []
else:
    raw = {"required_sections": [{"heading": "## Rules", "next_heading": None,
        "required_paragraphs": ["Before.", "After."],
        "required_tables": [{"headers": ["State", "Action"], "rows": [["Pending", "Wait"]]}],
        "required_block_order": ["paragraph", "table"]}]}
    try:
        sections = v.parse_required_sections(raw)
    except v.InstructionContractValidationError as exc:
        assert "Block order must match paragraph and table counts" in str(exc)
    else:
        failures = v.section_failures("## Rules\n\nBefore.\n\nAfter.\n\n" + literal, sections)
        assert False, ("Invalid catalog accepted", failures)
"""
    for directory, expected in [(SCRIPT_PATH.parent, 0), (mutant_dir, 1)]:
        result = subprocess.run(
            [sys.executable, "-c", program, str(directory), case],
            check=False,
            capture_output=True,
            text=True,
        )
        assert result.returncode == expected, result.stdout + result.stderr
        if expected:
            assert "AssertionError" in result.stderr


RETAINED_AZURE_HOSTS = [
    "CLAUDE.md",
    "GEMINI.md",
    ".hermes.md",
    ".cursor/rules/repository-instructions.mdc",
]


def _actual_azure_host_fixture(
    tmp_path: Path, path: str, modules: list[str]
) -> tuple[dict[str, Any], str]:
    """Copy an actual host contract and independent source text."""
    catalog = yaml.safe_load(
        (REPO_ROOT / ".template-sync/instruction-contracts.yml").read_text(encoding="utf-8")
    )
    contract = next(item for item in catalog["instruction_contracts"] if item["path"] == path)
    contracts = {"instruction_contracts": [contract]}
    _write_common_contract_repo(tmp_path, contracts)
    _write_yaml(tmp_path, ".template-sync/marker.yml", _marker(modules))
    text = (REPO_ROOT / path).read_text(encoding="utf-8")
    _write_text(tmp_path, path, text)
    return contracts, text


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("path", RETAINED_AZURE_HOSTS)
@pytest.mark.parametrize("change", ["section", "clause", "boundary"])
def test_actual_retained_azure_hosts_reject_protocol_loss(
    tmp_path: Path, mode: str, path: str, change: str
) -> None:
    """Every retained host must preserve its complete Azure protocol."""
    _contracts_data, text = _actual_azure_host_fixture(
        tmp_path, path, ["agent-instructions", "azure-devops-collaboration"]
    )
    intact = _run_validator(tmp_path, "--mode", mode)
    assert intact.returncode == 0, intact.stdout + intact.stderr
    heading = "## Azure DevOps PR Review Protocol"
    if change == "section":
        start = text.index(heading)
        end = text.find("\n## ", start + len(heading))
        text = text[:start] + (text[end:] if end != -1 else "")
    elif change == "clause":
        assert "does not satisfy required-reviewer policies" in text
        text = text.replace(
            "does not satisfy required-reviewer policies", "satisfies required-reviewer policies", 1
        )
    else:
        text = text.replace(heading, heading + "\n\n### Uncontracted boundary", 1)
    _write_text(tmp_path, path, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:" in result.stdout


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("path", RETAINED_AZURE_HOSTS)
def test_actual_azure_host_exclusion_and_owner_removal_oracle(tmp_path: Path, path: str) -> None:
    """Exclusion permits absence; removing ownership restores the clause defect."""
    contracts, text = _actual_azure_host_fixture(tmp_path, path, ["agent-instructions"])
    heading = "## Azure DevOps PR Review Protocol"
    start = text.index(heading)
    end = text.find("\n## ", start + len(heading))
    absent = text[:start] + (text[end:] if end != -1 else "")
    _write_text(tmp_path, path, absent)
    excluded = _run_validator(tmp_path, "--mode", "downstream")
    assert excluded.returncode == 0, excluded.stdout + excluded.stderr
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["agent-instructions", "azure-devops-collaboration"]),
    )
    changed = text.replace(
        "does not satisfy required-reviewer policies", "satisfies required-reviewer policies", 1
    )
    assert changed != text
    _write_text(tmp_path, path, changed)
    baseline = _run_validator(tmp_path, "--mode", "downstream")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    contract = contracts["instruction_contracts"][0]
    sections = contract["required_sections"]
    assert sum(item["heading"] == heading for item in sections) == 1
    sections[:] = [item for item in sections if item["heading"] != heading]
    if not sections:
        del contract["required_sections"]
        contract["required_headings"] = ["## Protected Instruction Files"]
    _write_yaml(tmp_path, ".template-sync/instruction-contracts.yml", contracts)
    mutant = _run_validator(tmp_path, "--mode", "downstream")
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr


@pytest.mark.upstream_template_only
def test_retained_azure_contracts_preserve_four_stale_decisions() -> None:
    """Retained enforcement cannot add an excluded-host policy decision."""
    catalog = yaml.safe_load(
        (REPO_ROOT / ".template-sync/instruction-contracts.yml").read_text(encoding="utf-8")
    )
    paths = {
        item["path"]
        for item in catalog["protected_guide_section_obligations"]
        if "azure-devops-collaboration" in item["target_modules"]
        and "## Azure DevOps PR Review Protocol" in item.get("stale_headings", [])
    }
    assert paths == {"AGENTS.md", "CLAUDE.md", "GEMINI.md", ".hermes.md"}


def _write_text(repo_root: Path, relative_path: str, text: str) -> None:
    """Write text below a fixture repository root."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _write_yaml(repo_root: Path, relative_path: str, data: dict[str, Any]) -> None:
    """Write YAML below a fixture repository root."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data, sort_keys=False), encoding="utf-8")


def _copy_schemas(repo_root: Path) -> None:
    """Copy the real validator schemas into a fixture repository."""
    schemas_dir = repo_root / "schemas"
    schemas_dir.mkdir(parents=True, exist_ok=True)
    for source_path in (CONTRACTS_SCHEMA_PATH, MARKER_SCHEMA_PATH, MANIFEST_SCHEMA_PATH):
        shutil.copyfile(source_path, schemas_dir / source_path.name)


def _manifest() -> dict[str, Any]:
    """Build a small schema-valid manifest fixture."""
    return {
        "template_manifest": {
            "version": 2,
            "modules": [
                {"name": "agent-copilot", "description": "Explicit retained agent."},
                {"name": "agent-codex", "description": "Explicit retained agent."},
                {"name": "agent-claude", "description": "Explicit retained agent."},
                {"name": "agent-cursor", "description": "Explicit retained agent."},
                {"name": "agent-gemini", "description": "Explicit retained agent."},
                {"name": "agent-hermes", "description": "Explicit retained agent."},
                {
                    "name": "agent-instructions",
                    "description": "Agent instruction files.",
                },
                {
                    "name": "template-sync-support",
                    "description": "Template sync support files.",
                },
                {
                    "name": "azure-devops-collaboration",
                    "description": "Azure DevOps collaboration files.",
                },
                {
                    "name": "schema",
                    "description": "Schema files.",
                },
            ],
            "path_mappings": [
                {
                    "pattern": "CLAUDE.md",
                    "requires_all": ["agent-instructions"],
                },
                {
                    "pattern": ".template-sync/instruction-contracts.yml",
                    "requires_all": ["template-sync-support"],
                },
            ],
            "filtering": {
                "default_semantics": "AND",
                "requires_any_semantics": "OR",
                "path_matching": "most_specific_match_wins",
                "same_specificity_action": "union_modules",
                "unmapped_action": "surface_for_owner",
            },
            "notes": {
                "downstream_retention": "Downstream repositories keep marker data for syncs.",
            },
        }
    }


def _contracts(
    *,
    required_headings: list[str] | None = None,
    required_phrases: list[str] | None = None,
) -> dict[str, Any]:
    """Build a small instruction-contract fixture."""
    contract: dict[str, Any] = {
        "path": "CLAUDE.md",
        "requires_modules": ["agent-instructions"],
    }
    if required_headings is not None:
        contract["required_headings"] = required_headings
    if required_phrases is not None:
        contract["required_phrases"] = required_phrases
    return {"instruction_contracts": [contract]}


def _host_specific_contracts() -> dict[str, Any]:
    """Build contract fixtures for default GitHub and optional Azure DevOps protocols."""
    return {
        "instruction_contracts": [
            {
                "path": "CLAUDE.md",
                "requires_modules": ["agent-instructions"],
                "required_headings": ["## Handling Code Review Comments"],
            },
            {
                "path": "GEMINI.md",
                "requires_modules": [
                    "agent-instructions",
                    "azure-devops-collaboration",
                ],
                "required_headings": ["## Azure DevOps PR Review Protocol"],
            },
        ]
    }


def _marker(
    included_modules: list[str],
    *,
    protected_decisions: list[dict[str, str]] | None = None,
    waivers: list[dict[str, str]] | None = None,
    protected_guide_waivers: list[dict[str, str]] | None = None,
) -> dict[str, Any]:
    """Build a small schema-valid marker fixture."""
    # Legacy all-agent fixtures now state that selection explicitly.
    if "agent-instructions" in included_modules:
        included_modules = list(
            dict.fromkeys(
                [
                    *included_modules,
                    "agent-copilot",
                    "agent-codex",
                    "agent-claude",
                    "agent-cursor",
                    "agent-gemini",
                    "agent-hermes",
                ]
            )
        )
    template_sync: dict[str, Any] = {
        "source_repo": SOURCE_REPO,
        "last_reviewed_template_commit": FULL_SHA,
        "included_modules": included_modules,
    }
    if protected_decisions is not None:
        template_sync["protected_file_decisions"] = protected_decisions
    if waivers is not None:
        template_sync["instruction_contract_waivers"] = waivers
    if protected_guide_waivers is not None:
        template_sync["protected_guide_contract_waivers"] = protected_guide_waivers
    return {"template_sync": template_sync}


def _write_common_contract_repo(repo_root: Path, contracts: dict[str, Any]) -> None:
    """Write schemas, manifest, and instruction contracts to a fixture repository."""
    _copy_schemas(repo_root)
    _write_yaml(repo_root, ".template-sync/manifest.yml", _manifest())
    _write_yaml(repo_root, ".template-sync/instruction-contracts.yml", contracts)


def _run_validator(repo_root: Path, *extra_args: str) -> subprocess.CompletedProcess[str]:
    """Run the instruction-contract validator against a fixture repository."""
    return subprocess.run(
        [
            sys.executable,
            str(SCRIPT_PATH),
            "--repo-root",
            str(repo_root),
            *extra_args,
        ],
        check=False,
        capture_output=True,
        text=True,
    )


def _copy_real_instruction_contract_surface(repo_root: Path) -> None:
    """Copy the actual protected-contract inputs into an isolated repository."""
    catalog_path = REPO_ROOT / ".template-sync/instruction-contracts.yml"
    catalog = yaml.safe_load(catalog_path.read_text(encoding="utf-8"))
    assert isinstance(catalog, dict)
    contracts = catalog.get("instruction_contracts")
    assert isinstance(contracts, list)

    _copy_schemas(repo_root)
    for relative_path in (
        ".template-sync/manifest.yml",
        ".template-sync/instruction-contracts.yml",
    ):
        destination = repo_root / relative_path
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(REPO_ROOT / relative_path, destination)
    for contract in contracts:
        assert isinstance(contract, dict)
        contract_path = contract.get("path")
        assert isinstance(contract_path, str)
        destination = repo_root / contract_path
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(REPO_ROOT / contract_path, destination)


def _section_entries(output: str, heading: str) -> set[str]:
    """Return bullet entries rendered under a named output section."""
    entries: set[str] = set()
    in_section = False
    for line in output.splitlines():
        if line and not line.startswith(" ") and line.endswith(":"):
            in_section = line == f"{heading}:"
            continue
        if in_section and line.startswith("  - "):
            entries.add(line.removeprefix("  - ").strip())
    return entries


def test_mode_is_required(tmp_path: Path) -> None:
    """The validator must not fall back to implicit mode detection."""
    result = _run_validator(tmp_path)

    assert result.returncode == 2
    assert "--mode" in result.stderr


@pytest.mark.upstream_template_only
def test_intact_upstream_claude_contract_passes() -> None:
    """The committed upstream Claude protocol satisfies the default contract."""
    result = _run_validator(REPO_ROOT, "--mode", "upstream-template")

    assert result.returncode == 0, result.stderr
    assert "Instruction-contract validation passed." in result.stdout
    assert "Contracts checked: 10" in result.stdout


FOLLOW_UP_GOVERNANCE_MUTATIONS = (
    pytest.param(
        "### Safe PR-head placement",
        "Do not use a stale tracking ref as proof.",
        "A stale tracking ref is sufficient proof.",
        id="r01-fresh-authenticated-head",
    ),
    pytest.param(
        "### Safe PR-head placement",
        "Agents MUST inspect the entire outgoing commit range and every changed path",
        "Agents MAY inspect only the latest outgoing commit and selected changed paths",
        id="r01-entire-outgoing-range",
    ),
    pytest.param(
        "### Safe PR-head placement",
        "agents MUST use an explicit non-force source-to-destination refspec",
        "agents MAY use an implicit or force source-to-destination refspec",
        id="r01-explicit-non-force-refspec",
    ),
    pytest.param(
        "### Safe PR-head placement",
        "do not treat a series of partial file writes as equivalent tested placement",
        "treat a series of partial file writes as equivalent tested placement",
        id="r01-api-tested-tree-equivalence",
    ),
    pytest.param(
        "### Safe PR-head placement",
        "A failed, uncertain, or mismatched readback leaves placement incomplete",
        "A failed, uncertain, or mismatched readback establishes placement success",
        id="r01-authenticated-readback",
    ),
    pytest.param(
        "### Ownership and delegation",
        "Agents MUST preserve unrelated user and agent work.",
        "Agents MAY overwrite unrelated user and agent work.",
        id="r02-preserve-unrelated-work",
    ),
    pytest.param(
        "### Ownership and delegation",
        "Before mutation or delegation, record the repository, worktree, branch, head and tree identities, allowed scope and paths, applicable findings, permitted public actions, and authority limits.",
        "Before mutation or delegation, record only the branch name.",
        id="r02-exact-task-input-and-scope",
    ),
    pytest.param(
        "### Ownership and delegation",
        "Agents MUST prevent overlapping writers to a file, worktree, index, branch ref, or remote object.",
        "Agents MAY allow overlapping writers to a file, worktree, index, branch ref, or remote object.",
        id="r02-exclusive-writers",
    ),
    pytest.param(
        "### Ownership and delegation",
        "Workers MUST stay within the assigned authority and MUST NOT create unbounded descendants.",
        "Workers MAY expand their authority and create unbounded descendants.",
        id="r02-bounded-worker-authority",
    ),
    pytest.param(
        "### Ownership and delegation",
        "The integration owner MUST verify worker claims against actual files, diffs, and native validation results before integration.",
        "The integration owner MAY trust worker summaries without inspecting native evidence.",
        id="r02-parent-native-verification",
    ),
    pytest.param(
        "### Ownership and delegation",
        "A self-review MUST NOT be described as an independent review.",
        "A self-review MAY be described as an independent review.",
        id="r02-independent-review-truth",
    ),
    pytest.param(
        "### Continuity and recovery",
        "Analysis, an option selection, or a next-step preview is not completion while authorized work remains.",
        "Analysis or a next-step preview completes the task while authorized work remains.",
        id="r03-continue-through-validation",
    ),
    pytest.param(
        "### Continuity and recovery",
        "agents MUST keep one compact task-private, untracked state index",
        "agents MAY rely on an unrecorded conversation summary",
        id="r03-task-private-index",
    ),
    pytest.param(
        "### Continuity and recovery",
        "Do not reconstruct requirements, authority, results, or pending operations from memory or a summary.",
        "Reconstruct requirements, authority, results, and pending operations from memory or a summary.",
        id="r03-full-input-recovery",
    ),
    pytest.param(
        "### Continuity and recovery",
        "reconcile uncertain remote operations before retrying",
        "retry uncertain remote operations before reconciliation",
        id="r03-uncertain-operation-reconciliation",
    ),
    pytest.param(
        "### Continuity and recovery",
        "Reuse passing results only when their relevant inputs are unchanged and repository policy permits it.",
        "Reuse passing results after relevant inputs change.",
        id="r03-input-bound-evidence-reuse",
    ),
    pytest.param(
        "### Continuity and recovery",
        "The record is evidence, not authority",
        "The record grants authority",
        id="r03-record-is-not-authority",
    ),
    pytest.param(
        "### Finding inventory and decisions",
        "A standalone `@codex review` or a command-only `@copilot` comment, with harmless surrounding whitespace, is request evidence rather than a finding",
        "Every bot command is a finding and surrounding whitespace changes its classification",
        id="r04-command-only-request-evidence",
    ),
    pytest.param(
        "### Finding inventory and decisions",
        "If a comment also contains substantive feedback, inventory that feedback regardless of its prefix.",
        "Discard substantive feedback when a comment begins with a bot command.",
        id="r04-command-with-substantive-feedback",
    ),
    pytest.param(
        "### Finding inventory and decisions",
        "agents MUST enumerate all-state review threads and their comments, including resolved, unresolved, and outdated threads.",
        "agents MAY enumerate only current unresolved review threads.",
        id="r04-complete-all-state-inventory",
    ),
    pytest.param(
        "### Finding inventory and decisions",
        "Agents MUST NOT filter inventory membership by REST `commit_id == current head`",
        "Agents MAY filter inventory membership by REST `commit_id == current head`",
        id="r04-no-mutable-head-filter",
    ),
    pytest.param(
        "### Finding inventory and decisions",
        "Keep current and original commit identities as provenance, not membership filters.",
        "Use current and original commit identities as membership filters.",
        id="r04-commit-identities-are-provenance",
    ),
    pytest.param(
        "### Finding inventory and decisions",
        "agents MUST perform a bounded search for the same root cause in relevant helpers and callers, copies of the same policy or configuration, and retained platform or module variants",
        "agents MAY inspect only the reported line and skip related variants",
        id="r05-bounded-related-defect-search",
    ),
    pytest.param(
        "### Finding inventory and decisions",
        "A materially different concern needs its own finding and decision",
        "Group a materially different concern into the existing decision",
        id="r05-separate-material-concerns",
    ),
    pytest.param(
        "### Finding inventory and decisions",
        "discovery does not expand task or protected-content authority",
        "discovery expands task and protected-content authority",
        id="r05-search-does-not-expand-authority",
    ),
    pytest.param(
        "### Finding inventory and decisions",
        "tests MUST include positive and negative controls and a targeted assertion-removal or failure-injection case with an expected result independent of the production predicate",
        "tests MAY repeat the production predicate without negative or mutation controls",
        id="r05-independent-guard-oracle",
    ),
    pytest.param(
        "### Finding inventory and decisions",
        "Do not require mutation tests for every prose, formatting, or cosmetic edit.",
        "Require mutation tests for every prose, formatting, and cosmetic edit.",
        id="r05-proportionate-mutation-scope",
    ),
)


GOVERNANCE_REFINEMENT_MUTATIONS = (
    pytest.param(
        ".github/instructions/docs.instructions.md",
        "### Synchronizing `Last Updated` and `Version` on Content Changes",
        "Internal topic, work-in-progress, and iteration commits are not separate published transitions.",
        "Each internal commit is a separate published transition.",
        "paragraph",
        id="published-finalization-not-internal-commits",
    ),
    pytest.param(
        ".github/instructions/docs.instructions.md",
        "### Synchronizing `Last Updated` and `Version` on Content Changes",
        "`Last Updated` in the final document MUST be the current UTC date in `YYYY-MM-DD` form at the finalization point.",
        "Last Updated MAY use any internal commit date.",
        "paragraph",
        id="finalization-date",
    ),
    pytest.param(
        ".github/instructions/docs.instructions.md",
        "### Synchronizing `Last Updated` and `Version` on Content Changes",
        "the final embedded `<YYYYMMDD>` segment MUST match the final `Last Updated` value.",
        "the embedded date MAY differ from Last Updated.",
        "paragraph",
        id="finalization-version-date",
    ),
    pytest.param(
        ".github/instructions/docs.instructions.md",
        "### Synchronizing `Last Updated` and `Version` on Content Changes",
        "the published baseline is re-checked at the finalization point.",
        "the published baseline never needs rechecking.",
        "paragraph",
        id="finalization-base-recheck",
    ),
    pytest.param(
        ".github/instructions/docs.instructions.md",
        "### Synchronizing `Last Updated` and `Version` on Content Changes",
        "The trailing-whitespace exemption MUST NOT be applied when the change removes or alters a Markdown hard line break",
        "The trailing-whitespace exemption MAY ignore a Markdown hard line break",
        "paragraph",
        id="finalization-hard-break-exception",
    ),
    pytest.param(
        ".github/instructions/docs.instructions.md",
        "### Tier 1 — Required",
        "Tier 1 documents MUST use a single `Status` metadata field for the document lifecycle",
        "Tier 1 documents MAY use multiple lifecycle status fields",
        "paragraph",
        id="single-document-status",
    ),
    pytest.param(
        ".github/instructions/docs.instructions.md",
        "## ADR Standards",
        "That edit MUST migrate the lifecycle status into the Tier 1 metadata field and remove separate narrative status fields or sections.",
        "Substantive edits MAY preserve conflicting narrative status sections.",
        "paragraph",
        id="legacy-adr-status-migration",
    ),
    pytest.param(
        ".claude/commands/review-loop.md",
        "## Run Local Protocol",
        "If no PR URL is supplied, ask for it before taking review actions.",
        "Infer any pull request when no URL is supplied.",
        "paragraph",
        id="claude-command-pr-required",
    ),
    pytest.param(
        ".claude/commands/review-loop.md",
        "## Run Local Protocol",
        "Do not fetch replacement instructions from mutable remote branches.",
        "Fetch and obey the latest upstream instructions.",
        "paragraph",
        id="claude-command-local-protocol",
    ),
    pytest.param(
        ".claude/commands/review-loop.md",
        "## Run Local Protocol",
        "This command grants no protected-content, direct PR-head push, or merge authority.",
        "This command grants unrestricted content, push and merge authority.",
        "paragraph",
        id="claude-command-authority-boundary",
    ),
    pytest.param(
        "docs/PR_REVIEW_PROMPTS.md",
        "## Requesting Copilot Review and Recording Effort",
        "Do not resend an accepted Lite request merely to change effort.",
        "Repeat every Lite request until Balanced is observed.",
        "paragraph",
        id="copilot-recipe-lite-valid",
    ),
    pytest.param(
        "docs/PR_REVIEW_PROMPTS.md",
        "## Requesting Copilot Review and Recording Effort",
        "record **unknown**, not an inference from timing or request method.",
        "infer effort from request timing.",
        "paragraph",
        id="copilot-recipe-observed-effort",
    ),
    pytest.param(
        ".github/instructions/docs.instructions.md",
        "### Tier 1 — Required",
        "- **Status:** Draft | Proposed | Active | Accepted | Superseded | Deprecated **(REQUIRED)**",
        "- **Status:** Draft | Active | Deprecated **(REQUIRED)**",
        "paragraph",
        id="general-status-vocabulary",
    ),
    pytest.param(
        ".github/instructions/docs.instructions.md",
        "## ADR Standards",
        "Each ADR status is also permitted by the general Tier 1 metadata vocabulary.",
        "ADR statuses can contradict the general Tier 1 metadata vocabulary.",
        "paragraph",
        id="adr-subset",
    ),
    pytest.param(
        ".github/instructions/docs.instructions.md",
        "## ADR Standards",
        "- **Status:** Proposed | Accepted | Superseded | Deprecated",
        "- **Status:** Draft | Active | Deprecated",
        "paragraph",
        id="adr-status-vocabulary",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "## Shared Review Governance",
        "applies to every real review finding and to material non-review bugs, design questions, investigation or test findings, and implementation choices",
        "applies only to real review findings",
        "paragraph",
        id="material-non-review-scope",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "## Shared Review Governance",
        "Apply PR inventory, native posting, reviewer attribution, replies, thread or body closure, and paired-review duties only to actual PR review findings.",
        "Apply all PR duties to every non-review finding.",
        "paragraph",
        id="pr-duties-only-for-pr-findings",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "## Shared Review Governance",
        "Do not create a PR solely to process a non-review finding.",
        "Create a PR for every non-review finding.",
        "paragraph",
        id="no-pr-created-for-non-review-finding",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "## Shared Review Governance",
        "For non-review work only, agents MAY reuse a complete existing decision record for mechanical implementation of the same finding when all relevant inputs remain unchanged and implementation authority is already granted.",
        "Agents MAY reuse any incomplete decision for any finding.",
        "paragraph",
        id="mechanical-reuse-guards",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "## Shared Review Governance",
        "A new material choice or changed relevant input requires a new evaluation.",
        "A new material choice can reuse the old evaluation.",
        "paragraph",
        id="changed-input-needs-new-evaluation",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "## Shared Review Governance",
        "This exception MUST NOT waive the mandatory analysis of a real review finding.",
        "This exception MAY waive analysis of a real review finding.",
        "paragraph",
        id="review-analysis-not-waived",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Finding inventory and decisions",
        "For a GitHub body-only finding, post its prompt as a standalone PR comment that records its synthetic finding key, source review identity, reviewed commit, and location when available.",
        "Post every GitHub finding's prompt in the same native review thread.",
        "paragraph",
        id="body-only-guide-prompt-routing",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Finding inventory and decisions",
        "Record the prompt comment's native identity and keep the secondary guide action pending until it has an attributable disposition.",
        "Mark the secondary guide action complete when the prompt is posted.",
        "paragraph",
        id="secondary-guide-pending-disposition",
    ),
    pytest.param(
        "AGENTS.md",
        "#### Protected-file authorization terms",
        "For a GitHub body-only finding, use a standalone PR comment that records its synthetic finding key, source review identity, reviewed commit, and location when available.",
        "For a GitHub body-only finding, use the same native review thread.",
        "paragraph",
        id="codex-body-only-guide-prompt-routing",
    ),
    pytest.param(
        "CLAUDE.md",
        "### Protected-file authorization terms",
        "For a GitHub body-only finding, use a standalone PR comment that records its synthetic finding key, source review identity, reviewed commit, and location when available.",
        "For a GitHub body-only finding, use the same native review thread.",
        "paragraph",
        id="claude-body-only-guide-prompt-routing",
    ),
    pytest.param(
        "AGENTS.md",
        "## Execution",
        "Preserve its narrow mechanical-reuse conditions and its distinction between general decisions and PR-specific duties.",
        "Apply PR-specific duties to every general decision.",
        "paragraph",
        id="codex-entry-link",
    ),
    pytest.param(
        "CLAUDE.md",
        "## Execution",
        "Preserve its narrow mechanical-reuse conditions and its distinction between general decisions and PR-specific duties.",
        "Apply PR-specific duties to every general decision.",
        "paragraph",
        id="claude-entry-link",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Selected-action writing rule",
        "does not claim full conformance with the external specification",
        "certifies full conformance with the external specification",
        "paragraph",
        id="no-full-conformance-claim",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Selected-action writing rule",
        "No external specification or general prose-linting engine is required to apply this local rule.",
        "An external specification and general prose-linting engine are mandatory.",
        "paragraph",
        id="local-rule-no-general-engine",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Selected-action writing rule",
        "- Limit each instruction sentence to 20 words and each description sentence to 25 words.",
        "- Use any sentence length.",
        "paragraph",
        id="sentence-limits",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Selected-action writing rule",
        "Name the affected files and intended behavior. Explain the decisive rationale, tradeoffs, and any lost guarantee or coverage.",
        "Describe the change without affected files, tradeoffs, or lost coverage.",
        "paragraph",
        id="action-content",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Selected-action writing rule",
        "Distinguish pending tests from observed passes.",
        "Describe pending tests as observed passes.",
        "paragraph",
        id="test-result-truth",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "## Agent Execution",
        "agents MUST verify the intended resulting state through authenticated evidence before claiming success",
        "agents MAY claim success from an unverified request",
        "paragraph",
        id="verify-remote-result",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "## Agent Execution",
        "Reuse reliable returned final-state evidence when it establishes the result; otherwise read the affected object.",
        "Always trust an acknowledgment without resulting-state evidence.",
        "paragraph",
        id="reliable-returned-state",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "## Agent Execution",
        "A delivery acknowledgment alone is insufficient. Reconcile an uncertain, failed, or mismatched result before retrying.",
        "A delivery acknowledgment proves success and permits an immediate retry.",
        "paragraph",
        id="uncertain-result-before-retry",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "## Agent Execution",
        "Do not require separate per-command receipts or local-edit readbacks.",
        "Require a remote receipt for every local edit.",
        "paragraph",
        id="proportionate-verification",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Protected authority and deferral",
        "the originating finding/PR/review link",
        "no origin record",
        "tables",
        id="deferral-origin",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Protected authority and deferral",
        "an explicit condition for resuming work",
        "no condition for resuming work",
        "tables",
        id="resume-condition",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Ownership and delegation",
        "the requested worker model and reasoning effort, or inherited defaults",
        "only the worker name",
        "paragraph",
        id="requested-settings-or-defaults",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Ownership and delegation",
        "Record effective settings only when reliable runtime evidence identifies them; otherwise state that they are unavailable.",
        "Treat requested settings as proven effective settings.",
        "paragraph",
        id="effective-settings-evidence",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Ownership and delegation",
        "Require applicable analysis, first-edit, validation, and public-mutation checkpoints.",
        "Require only a final checkpoint.",
        "paragraph",
        id="applicable-checkpoints",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Ownership and delegation",
        "A read-only assignment has no first-edit checkpoint.",
        "Every read-only assignment has a fictional first edit.",
        "paragraph",
        id="read-only-checkpoint-applicability",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Ownership and delegation",
        "Public mutations remain with the integration owner unless an explicitly delegated bounded operation falls within existing authority.",
        "Every worker may perform unbounded public mutations.",
        "paragraph",
        id="public-mutation-owner",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Ownership and delegation",
        "These requirements do not mandate a model-selection API or invent an unavailable capability.",
        "These requirements mandate an unavailable model-selection API.",
        "paragraph",
        id="no-invented-settings-capability",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Continuity and recovery",
        "meaningful phase, finding, blocker, and handoff boundaries",
        "every command boundary",
        "paragraph",
        id="meaningful-update-boundaries",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Continuity and recovery",
        "State the current result, remaining uncertainty, and next useful action.",
        "State only that work continues.",
        "paragraph",
        id="useful-update-content",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Continuity and recovery",
        "Follow the active runtime's communication timing requirements.",
        "Ignore the active runtime's communication timing requirements.",
        "paragraph",
        id="runtime-timing",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Continuity and recovery",
        "Normal quiet reasoning is not proof of a hang, but this distinction MUST NOT justify indefinite silence.",
        "Quiet reasoning always proves a hang or permits indefinite silence.",
        "paragraph",
        id="quiet-reasoning-boundary",
    ),
    pytest.param(
        ".github/copilot-instructions.md",
        "### Continuity and recovery",
        "do not add telemetry, a polling framework, or mandatory per-command narration",
        "add telemetry and narrate every command",
        "paragraph",
        id="no-telemetry-or-command-narration",
    ),
    pytest.param(
        "AGENTS.md",
        "## Execution",
        "For runtimes that follow [OpenAI's documented instruction discovery]",
        "Every runtime always follows undocumented Codex loading behavior.",
        "paragraph",
        id="runtime-qualified-discovery",
    ),
    pytest.param(
        "AGENTS.md",
        "## Execution",
        "Global guidance precedes project guidance from the root through the launch working directory.",
        "Only the nearest instruction file applies.",
        "paragraph",
        id="startup-chain",
    ),
    pytest.param(
        "AGENTS.md",
        "## Execution",
        "Codex MUST perform bounded discovery along the relevant file paths and read applicable instructions.",
        "Codex MAY skip deeper applicable instructions.",
        "paragraph",
        id="bounded-deeper-discovery",
    ),
    pytest.param(
        "AGENTS.md",
        "## Execution",
        "If active guidance is stale, use the runtime's supported refresh; the documented CLI procedure is to restart in the target directory.",
        "Continue with stale guidance without refreshing.",
        "paragraph",
        id="stale-session-refresh",
    ),
    pytest.param(
        "AGENTS.md",
        "## Execution",
        "Preserve the shared disk-recovery requirements and exact-file rereads after compaction or instruction changes.",
        "Reconstruct instructions from memory after compaction.",
        "paragraph",
        id="preserve-disk-recovery",
    ),
    pytest.param(
        "AGENTS.md",
        "## Execution",
        "This guidance does not require changing configuration or apply Codex loading semantics to other agents.",
        "Change configuration and apply Codex loading semantics to every agent.",
        "paragraph",
        id="codex-only-no-config-change",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "A workflow MAY use `actions/setup-node` with `node-version-file` instead of `node-version` only when all conditions below hold. This optional exception does not change the direct release-line default, other setup actions, or Azure Pipelines selector rules.",
        "A workflow MAY use any version file instead of a direct selector.",
        "paragraph",
        id="node-version-file-is-bounded-exception",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "- The file MUST be tracked, repository-relative, and read from the reviewed revision. External, generated, or untracked version sources do not qualify.",
        "- The file MAY be external, generated, or untracked.",
        "paragraph",
        id="version-file-reviewed-provenance",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "- The referenced action revision MUST document support for the format. The inspected setup-node v7 format set is `.nvmrc`, `.node-version`, `.tool-versions`, and `package.json`. Do not infer support for a later format from newer action documentation.",
        "- Any format documented by a later action release is accepted.",
        "paragraph",
        id="format-support-bound-to-action-revision",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "- The selected value MUST be one exact stable `major.minor.patch` version. Ranges, wildcards, aliases, release channels, prereleases, and build metadata do not qualify.",
        "- The selected value MAY be a range, alias, prerelease, or floating channel.",
        "paragraph",
        id="exact-stable-version",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "- For `.nvmrc` and `.node-version`, use only the exact version. For `.tool-versions`, use one unambiguous `node` or `nodejs` entry with that exact version.",
        "- Plain and tool version files MAY contain ambiguous selectors.",
        "paragraph",
        id="literal-version-file-formats",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "- For `package.json`, account for the action's precedence: `volta.node`, then the first `devEngines.runtime` entry with a case-insensitive `node` name and a version, then `engines.node`, then recursive `volta.extends`. Higher-precedence fields MUST be absent or select the same exact version as the declared canonical field. Multiple Node runtime entries MUST agree. Any inherited file MUST also be tracked, reviewed, repository-contained, and cycle-free.",
        "- For `package.json`, read any convenient field and ignore conflicting or inherited values.",
        "paragraph",
        id="package-json-precedence-and-agreement",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "- The setup step MUST NOT also supply `node-version`. The action gives that input priority, which would make the file non-authoritative.",
        "- The setup step MAY also supply a competing `node-version` input.",
        "paragraph",
        id="no-competing-direct-input",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "- Before dependency installation, build, lint, test, or other Node-dependent project work, a later step MUST read the same canonical field and compare the installed version with it. The job MUST fail unless `process.versions.node` equals the expected version exactly. Disable optional automatic package-manager caching when it would perform dependent work before this check.",
        "- Version verification MAY occur after Node-dependent project work and need not fail on a mismatch.",
        "paragraph",
        id="exact-verification-before-dependent-work",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "See the [setup-node version-file documentation](https://github.com/actions/setup-node/blob/820762786026740c76f36085b0efc47a31fe5020/docs/advanced-usage.md#node-version-file) and [its selected-field parser](https://github.com/actions/setup-node/blob/820762786026740c76f36085b0efc47a31fe5020/src/util.ts#L11-L73). Exact pins require deliberate patch maintenance; they do not constitute a transitive dependency lock.",
        "An exact runtime pin is a complete transitive dependency lock and needs no maintenance.",
        "paragraph",
        id="exact-pin-is-not-dependency-lock",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "Compliant example: this tracked `package.json` declares only `engines.node` as the canonical Node.js field. The version is illustrative, not a runtime-currency recommendation.",
        "Compliant example: use any unreviewed field and treat this value as a currency recommendation.",
        "paragraph",
        id="example-selected-field-and-currency-limit",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "The following step sequence uses Bash for the verification step. It reads the same JSON field, rejects a non-exact value, and stops before `npm ci` on mismatch.",
        "The verification step may read another field or run after `npm ci`.",
        "paragraph",
        id="example-verifies-selected-field-first",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "Non-compliant cases include a file containing `24`, `24.x`, `>=24`, `lts/*`, or `24.18.0-rc.1`; both setup inputs; a check of another field; a missing equality check; or verification after `npm ci`. Correct the source or verification before dependent work.",
        "Release lines, prereleases, competing inputs, and late verification are compliant.",
        "paragraph",
        id="noncompliant-boundaries",
    ),
    pytest.param(
        ".github/instructions/yaml.instructions.md",
        "### Exact Node.js version-file exception",
        "Instruction contracts and focused example tests protect this guidance. They do not validate every downstream workflow or prove agent compliance. When a repository retains a toolchain inventory scanner, its selected-file parsing MUST agree with the action. The inventory does not replace checks of tracked provenance, exactness, or verification ordering. This rule does not require retaining an optional scanner or its module.",
        "The optional inventory proves every downstream workflow compliant and is mandatory in every profile.",
        "paragraph",
        id="contract-and-scanner-scope",
    ),
)


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    ("variant", "required_scope"),
    (
        (
            "Secondary Guide Prompt Only",
            "Do not implement that secondary guide change in this task, even if earlier authority would allow it.",
        ),
        (
            "Authorized Secondary Guide Changes",
            "Implement a selected fix or secondary style-guide change only when specific current-task authority covers its content.",
        ),
        (
            "No Secondary Guide Proposal",
            "Assess guide impact, but do not prepare or implement a separate secondary guide proposal.",
        ),
    ),
)
def test_review_prompt_variants_keep_copyable_scope(variant: str, required_scope: str) -> None:
    """Each copied prompt retains its own authority boundary and guide-output scope."""
    document = (REPO_ROOT / "docs/PR_REVIEW_PROMPTS.md").read_text(encoding="utf-8")
    heading = f"### Evaluate, Decide, and Implement — {variant}\n"
    section = document.split(heading, 1)[1].split("\n### ", 1)[0]
    prompts = re.findall(r"^```markdown\n(.*?)\n```$", section, re.MULTILINE | re.DOTALL)
    assert len(prompts) == 1
    # Inspect the copyable block, not explanatory prose that a user may omit.
    prompt = " ".join(prompts[0].split())
    assert "Shared Review Governance and Protected Instruction Files" in prompt
    assert ".github/copilot-instructions.md" in prompt
    assert "fresh weighted rubric, displayed scores, pre-edit evaluation" in prompt
    assert "This prompt grants no protected-content, branch-placement or merge authority." in prompt
    assert required_scope in prompt


@pytest.mark.upstream_template_only
def test_finalization_examples_and_command_limits_preserve_portable_policy() -> None:
    """Check published-transition examples and wrapper limits, not calendar correctness."""
    guide = (REPO_ROOT / ".github/instructions/docs.instructions.md").read_text(encoding="utf-8")
    section = guide.split("### Synchronizing `Last Updated` and `Version` on Content Changes\n", 1)[
        1
    ].split("### Non-Normative Historical Artifacts\n", 1)[0]
    for example in (
        "`1.6.20260502.0` → this change `1.6.20260502.1`",
        "change finalized 2026-05-03 UTC → `1.6.20260503.0`",
        "one published transition to `1.6.20260502.5`, not `.7`",
        "two or more trailing spaces, or a trailing backslash",
    ):
        assert example in section
    command = (REPO_ROOT / ".claude/commands/review-loop.md").read_text(encoding="utf-8")
    assert "**$ARGUMENTS**" in command
    assert "../../CLAUDE.md" in command
    assert "../../.github/copilot-instructions.md" in command
    assert "80" not in command
    assert "8 review" not in command
    assert "raw.githubusercontent.com" not in command


@pytest.mark.upstream_template_only
def test_issue_evaluation_prompt_preserves_nonmutating_copyable_contract() -> None:
    """The copied issue prompt preserves evidence and has no implicit action grant."""
    source = (REPO_ROOT / "docs/ISSUE_EVALUATION_PROMPT.md").read_text(encoding="utf-8")
    blocks = re.findall(r"^```markdown\n(.*?)\n```$", source, re.MULTILINE | re.DOTALL)
    assert len(blocks) == 1
    prompt = " ".join(blocks[0].split())
    assert "Do not create or edit issues, files, branches, comments or settings" in prompt
    assert "do not implement the proposed change" in prompt
    assert "Preserve the reported problem, useful evidence and intended outcome." in prompt
    assert "Distinguish verified facts, inferences, proposed remedies and open questions." in prompt
    assert "acceptance conditions and validation" in prompt
    assert "A prose link alone does not establish a native dependency." in prompt
    assert "This prompt grants no protected-content, branch-placement or merge authority." in prompt
    assert "STYLE_GUIDE.md" not in source
    assert "STYLE_GUIDE_RATIONALE.md" not in source


@pytest.mark.upstream_template_only
def test_push_scope_examples_preserve_explicit_ref_intent() -> None:
    """The two real guide examples express the documented GitHub event shapes."""
    guide = (REPO_ROOT / ".github/instructions/yaml.instructions.md").read_text(encoding="utf-8")
    section = guide.split("## GitHub Actions Push Ref and Path Scope\n", 1)[1].split(
        "## GitHub Actions Setup Version Pins\n", 1
    )[0]
    examples = re.findall(r"```yaml\n(.*?)```", section, re.DOTALL)
    assert len(examples) == 2
    # BaseLoader keeps GitHub's "on" key literal instead of YAML 1.1 Boolean coercion.
    branch_only, tag_enabled = (yaml.load(item, Loader=yaml.BaseLoader) for item in examples)
    assert branch_only["on"]["push"] == {
        "branches": ["**"],
        "paths": ["docs/**"],
    }
    assert tag_enabled["on"]["push"] == {
        "branches": ["**"],
        "tags": ["v*"],
        "paths": ["docs/**"],
    }
    assert "regardless of changed paths" in section


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    ("clause", "replacement"),
    [
        ("authors MUST define its branch and tag intent explicitly", ""),
        (
            "a tag-enabled workflow MUST NOT rely on those filters to select tag events",
            "a tag-enabled workflow MAY rely on those filters to select tag events",
        ),
        (
            "For a branch-only workflow, define `branches` or `branches-ignore` and omit tag filters.",
            "Path filters alone establish branch-only behavior.",
        ),
        ("This rule applies to GitHub Actions push events only.", ""),
    ],
)
def test_push_scope_contract_rejects_missing_or_weakened_rule(
    tmp_path: Path, clause: str, replacement: str
) -> None:
    """A changed trigger obligation fails through the production contract CLI."""
    _copy_real_instruction_contract_surface(tmp_path)
    baseline = _run_validator(tmp_path, "--mode", "upstream-template")
    assert baseline.returncode == 0, baseline.stdout + baseline.stderr
    path = tmp_path / ".github/instructions/yaml.instructions.md"
    text = path.read_text(encoding="utf-8")
    assert text.count(clause) == 1
    path.write_text(text.replace(clause, replacement, 1), encoding="utf-8")
    result = _run_validator(tmp_path, "--mode", "upstream-template")
    assert result.returncode == 1, result.stdout + result.stderr
    assert ".github/instructions/yaml.instructions.md" in result.stdout
    assert "section:## GitHub Actions Push Ref and Path Scope:paragraph:" in result.stdout


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    ("path", "clause", "replacement"),
    (
        (
            ".github/copilot-instructions.md",
            "grants MUST remain valid across a verified resume of the same task, repository, PR, scope, and action class",
            "grants expire after every interruption",
        ),
        (
            ".github/copilot-instructions.md",
            "Protected-content, branch-placement, and merge authority remain separate; a resume creates none of them.",
            "A resume grants all publication authority.",
        ),
        (
            "AGENTS.md",
            "for this specific PR in the current task, including a verified resume under the shared continuity rule",
            "for any PR in any task",
        ),
        (
            "CLAUDE.md",
            "This is an intentional platform policy difference, not a transferable grant.",
            "This platform grant applies to all agents.",
        ),
    ),
)
def test_verified_resume_authority_rejects_scope_drift(
    tmp_path: Path, path: str, clause: str, replacement: str
) -> None:
    """Real contracts detect lost continuity and expanded cross-task/platform grants."""
    _copy_real_instruction_contract_surface(tmp_path)
    baseline = _run_validator(tmp_path, "--mode", "upstream-template")
    assert baseline.returncode == 0, baseline.stdout + baseline.stderr
    policy = tmp_path / path
    text = policy.read_text(encoding="utf-8")
    assert text.count(clause) == 1
    policy.write_text(text.replace(clause, replacement, 1), encoding="utf-8")
    result = _run_validator(tmp_path, "--mode", "upstream-template")
    assert result.returncode == 1, result.stdout + result.stderr
    assert path in result.stdout
    assert ":paragraph:" in result.stdout


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    "clause",
    (
        "it is descriptive, not a second authoritative pin.",
        "but does not create vulnerability alerts for SHA-pinned actions.",
    ),
)
def test_action_pin_contract_rejects_lost_authority_or_alert_limit(
    tmp_path: Path, clause: str
) -> None:
    """The actual policy cannot lose its pin authority or truthful alert boundary."""
    _copy_real_instruction_contract_surface(tmp_path)
    baseline = _run_validator(tmp_path, "--mode", "upstream-template")
    assert baseline.returncode == 0, baseline.stdout + baseline.stderr
    path = tmp_path / ".github/copilot-instructions.md"
    source = path.read_text(encoding="utf-8")
    assert source.count(clause) == 1
    path.write_text(source.replace(clause, "", 1), encoding="utf-8")
    result = _run_validator(tmp_path, "--mode", "upstream-template")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:### Immutable action pins and release comments:paragraph:" in result.stdout


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(("heading", "clause", "weakened"), FOLLOW_UP_GOVERNANCE_MUTATIONS)
@pytest.mark.parametrize("mutation", ["delete", "weaken"])
def test_follow_up_governance_clauses_reject_deletion_and_weakening(
    tmp_path: Path,
    heading: str,
    clause: str,
    weakened: str,
    mutation: str,
) -> None:
    """Removed or weakened governance clauses fail through the real contract validator."""
    _copy_real_instruction_contract_surface(tmp_path)
    baseline = _run_validator(tmp_path, "--mode", "upstream-template")
    assert baseline.returncode == 0, baseline.stdout + baseline.stderr

    policy_path = tmp_path / ".github/copilot-instructions.md"
    policy = policy_path.read_text(encoding="utf-8")
    assert policy.count(clause) == 1
    replacement = "" if mutation == "delete" else weakened
    policy_path.write_text(policy.replace(clause, replacement, 1), encoding="utf-8")

    result = _run_validator(tmp_path, "--mode", "upstream-template")

    assert result.returncode == 1, result.stdout + result.stderr
    assert ".github/copilot-instructions.md" in result.stdout
    assert f"section:{heading}:paragraph:" in result.stdout


@pytest.mark.upstream_template_only
def test_document_status_vocabulary_contains_the_adr_subset() -> None:
    """The real guide and ADR examples use one consistent status vocabulary."""
    guide = (REPO_ROOT / ".github/instructions/docs.instructions.md").read_text(encoding="utf-8")
    tier_one = guide.split("### Tier 1 — Required\n", 1)[1].split("\n### Tier 2 — Not Required", 1)[
        0
    ]
    adr_rules = guide.split("## ADR Standards\n", 1)[1].split(
        "\n## Requirements Documentation Standards", 1
    )[0]

    general_match = re.search(
        r"^- \*\*Status:\*\* (?P<values>.+?) \*\*\(REQUIRED\)\*\*$",
        tier_one,
        re.MULTILINE,
    )
    adr_match = re.search(
        r"^  - \*\*Status:\*\* (?P<values>.+)$",
        adr_rules,
        re.MULTILINE,
    )
    assert general_match is not None
    assert adr_match is not None
    general = {item.strip() for item in general_match.group("values").split("|")}
    adr = {item.strip() for item in adr_match.group("values").split("|")}
    assert general == {
        "Draft",
        "Proposed",
        "Active",
        "Accepted",
        "Superseded",
        "Deprecated",
    }
    assert adr == {"Proposed", "Accepted", "Superseded", "Deprecated"}
    assert adr < general

    guide_status = re.search(r"^- \*\*Status:\*\* (?P<value>\w+)$", guide, re.MULTILINE)
    assert guide_status is not None
    assert guide_status.group("value") == "Active"
    assert guide_status.group("value") in general
    for status in sorted(adr):
        example = f"# ADR\n\n## Metadata\n\n- **Status:** {status}\n"
        match = re.search(r"^- \*\*Status:\*\* (?P<value>\w+)$", example, re.MULTILINE)
        assert match is not None
        assert match.group("value") in adr
        assert match.group("value") in general


@pytest.mark.upstream_template_only
def test_document_status_oracle_rejects_the_old_contradiction() -> None:
    """The previous general list cannot contain the retained ADR lifecycle."""
    guide = (REPO_ROOT / ".github/instructions/docs.instructions.md").read_text(encoding="utf-8")
    old_general = (
        "- **Status:** Draft | Proposed | Active | Accepted | Superseded | Deprecated "
        "**(REQUIRED)**"
    )
    contradictory = guide.replace(
        old_general,
        "- **Status:** Draft | Active | Deprecated **(REQUIRED)**",
        1,
    )
    assert contradictory != guide
    tier_one = contradictory.split("### Tier 1 — Required\n", 1)[1].split(
        "\n### Tier 2 — Not Required", 1
    )[0]
    adr_rules = contradictory.split("## ADR Standards\n", 1)[1].split(
        "\n## Requirements Documentation Standards", 1
    )[0]
    general_match = re.search(
        r"^- \*\*Status:\*\* (?P<values>.+?) \*\*\(REQUIRED\)\*\*$",
        tier_one,
        re.MULTILINE,
    )
    adr_match = re.search(
        r"^  - \*\*Status:\*\* (?P<values>.+)$",
        adr_rules,
        re.MULTILINE,
    )
    assert general_match is not None
    assert adr_match is not None
    general = {item.strip() for item in general_match.group("values").split("|")}
    adr = {item.strip() for item in adr_match.group("values").split("|")}
    assert not adr.issubset(general)
    assert adr - general == {"Proposed", "Accepted", "Superseded"}


@pytest.mark.upstream_template_only
def test_codex_discovery_guidance_remains_codex_specific() -> None:
    """Codex discovery and recovery stay in AGENTS without changing Claude loading rules."""
    agents = (REPO_ROOT / "AGENTS.md").read_text(encoding="utf-8")
    claude = (REPO_ROOT / "CLAUDE.md").read_text(encoding="utf-8")
    agents_execution = agents.split("## Execution\n", 1)[1].split(
        "\n## Protected Instruction Files", 1
    )[0]
    claude_execution = claude.split("## Execution\n", 1)[1].split(
        "\n## Protected Instruction Files", 1
    )[0]
    for clause in (
        "OpenAI's documented instruction discovery",
        "AGENTS.override.md",
        "bounded discovery along the relevant file paths",
        "restart in the target directory",
        "exact-file rereads after compaction or instruction changes",
    ):
        assert clause in agents_execution
        assert clause not in claude_execution
    assert "does not require changing configuration" in agents_execution
    assert "apply Codex loading semantics to other agents" in agents_execution


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    ("path", "heading", "clause", "weakened", "anchor_kind"),
    GOVERNANCE_REFINEMENT_MUTATIONS,
)
@pytest.mark.parametrize("mutation", ["delete", "weaken"])
def test_governance_refinements_reject_removed_or_narrowed_obligations(
    tmp_path: Path,
    path: str,
    heading: str,
    clause: str,
    weakened: str,
    anchor_kind: str,
    mutation: str,
) -> None:
    """The real catalog rejects removal or weakening of each expanded decision, execution, and version-file rule."""
    _copy_real_instruction_contract_surface(tmp_path)
    baseline = _run_validator(tmp_path, "--mode", "upstream-template")
    assert baseline.returncode == 0, baseline.stdout + baseline.stderr

    policy_path = tmp_path / path
    policy = policy_path.read_text(encoding="utf-8")
    assert policy.count(clause) == 1, (path, clause)
    replacement = "" if mutation == "delete" else weakened
    policy_path.write_text(policy.replace(clause, replacement, 1), encoding="utf-8")

    result = _run_validator(tmp_path, "--mode", "upstream-template")
    assert result.returncode == 1, result.stdout + result.stderr
    assert path in result.stdout
    assert f"section:{heading}:{anchor_kind}:" in result.stdout


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    ("included_modules", "include_guide", "mutate_clause", "expected_exit"),
    (
        pytest.param(
            ["agent-instructions", "yaml"],
            True,
            False,
            0,
            id="yaml-retained",
        ),
        pytest.param(
            ["agent-instructions", "yaml"],
            True,
            True,
            1,
            id="yaml-retained-weakened",
        ),
        pytest.param(
            ["agent-instructions"],
            False,
            False,
            0,
            id="yaml-excluded",
        ),
    ),
)
def test_node_version_file_contract_follows_yaml_module_selection(
    tmp_path: Path,
    included_modules: list[str],
    include_guide: bool,
    mutate_clause: bool,
    expected_exit: int,
) -> None:
    """The real version-file contract is enforced only when the YAML guide is retained."""
    catalog = yaml.safe_load(
        (REPO_ROOT / ".template-sync/instruction-contracts.yml").read_text(encoding="utf-8")
    )
    yaml_contract = next(
        contract
        for contract in catalog["instruction_contracts"]
        if contract["path"] == ".github/instructions/yaml.instructions.md"
    )
    node_section = next(
        section
        for section in yaml_contract["required_sections"]
        if section["heading"] == "### Exact Node.js version-file exception"
    )
    assert node_section["next_heading"] == "## GitHub Actions Documentation Comment URLs"

    _copy_schemas(tmp_path)
    (tmp_path / ".template-sync").mkdir(parents=True, exist_ok=True)
    shutil.copyfile(
        REPO_ROOT / ".template-sync/manifest.yml",
        tmp_path / ".template-sync/manifest.yml",
    )
    _write_yaml(
        tmp_path,
        ".template-sync/instruction-contracts.yml",
        {"instruction_contracts": [yaml_contract]},
    )
    _write_yaml(tmp_path, ".template-sync/marker.yml", _marker(included_modules))

    if include_guide:
        guide_path = tmp_path / ".github/instructions/yaml.instructions.md"
        guide_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(REPO_ROOT / ".github/instructions/yaml.instructions.md", guide_path)
        if mutate_clause:
            clause = (
                "- The setup step MUST NOT also supply `node-version`. The action gives that "
                "input priority, which would make the file non-authoritative."
            )
            guide = guide_path.read_text(encoding="utf-8")
            assert guide.count(clause) == 1
            guide_path.write_text(
                guide.replace(
                    clause,
                    "- The setup step MAY also supply a competing `node-version` input.",
                    1,
                ),
                encoding="utf-8",
            )

    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == expected_exit, result.stdout + result.stderr
    if included_modules == ["agent-instructions"]:
        assert "Instruction-contract validation passed." in result.stdout
        skipped = _section_entries(
            result.stdout, "Contracts skipped by downstream module selection"
        )
        assert any(
            item.startswith(".github/instructions/yaml.instructions.md ") for item in skipped
        )
    elif mutate_clause:
        assert ".github/instructions/yaml.instructions.md" in result.stdout
        assert "section:### Exact Node.js version-file exception:paragraph:" in result.stdout
    else:
        assert "Instruction-contract validation passed." in result.stdout


@pytest.mark.parametrize(
    "missing_heading",
    [
        "## Handling Code Review Comments",
        "## Automated Review Loop",
    ],
)
def test_upstream_missing_required_heading_fails(
    tmp_path: Path,
    missing_heading: str,
) -> None:
    """Required Claude protocol headings are enforced in upstream-template mode."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=[missing_heading]),
    )
    _write_text(
        tmp_path,
        "CLAUDE.md",
        "# Agent Instructions\n\n## Different Heading\n\nProtected-file authorization checkpoint\n",
    )

    result = _run_validator(tmp_path, "--mode", "upstream-template")

    assert result.returncode == 1
    assert "Instruction-contract validation failed." in result.stdout
    assert f"CLAUDE.md: missing required heading: {missing_heading}" in result.stdout


def test_upstream_missing_required_phrase_fails(tmp_path: Path) -> None:
    """Required protocol phrases are reported with exact missing text."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_phrases=["Protected-file authorization checkpoint"]),
    )
    _write_text(tmp_path, "CLAUDE.md", "# Agent Instructions\n")

    result = _run_validator(tmp_path, "--mode", "upstream-template")

    assert result.returncode == 1
    assert (
        "CLAUDE.md: missing required phrase: Protected-file authorization checkpoint"
        in result.stdout
    )


def test_downstream_missing_marker_skips_by_default_and_fails_when_required(
    tmp_path: Path,
) -> None:
    """Downstream mode preserves the marker validator's require-marker semantics."""
    default_result = _run_validator(tmp_path, "--mode", "downstream")
    required_result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")

    assert default_result.returncode == 0, default_result.stderr
    assert "No marker found at .template-sync/marker.yml" in default_result.stdout
    assert "instruction-contract validation skipped" in default_result.stdout
    assert required_result.returncode == 1
    assert "Marker is required but was not found" in required_result.stderr


def test_upstream_mode_with_marker_present_warns_without_failing(tmp_path: Path) -> None:
    """A present marker is a non-blocking warning in upstream-template mode."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=["## Handling Code Review Comments"]),
    )
    _write_yaml(tmp_path, ".template-sync/marker.yml", {"not_template_sync": True})
    _write_text(tmp_path, "CLAUDE.md", "# Agent Instructions\n\n## Handling Code Review Comments\n")

    result = _run_validator(tmp_path, "--mode", "upstream-template")

    assert result.returncode == 0, result.stderr
    assert "WARNING: --mode upstream-template was invoked while .template-sync/marker.yml" in (
        result.stdout
    )
    assert "Instruction-contract validation passed." in result.stdout


def test_valid_downstream_waiver_is_reported_loudly(tmp_path: Path) -> None:
    """A valid marker waiver can pass validation but is not ordinary success."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=["## Handling Code Review Comments"]),
    )
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            waivers=[
                {
                    "path": "CLAUDE.md",
                    "anchor": "## Handling Code Review Comments",
                    "reason": "Downstream owner uses a different review protocol.",
                    "authorization_basis": "Owner authorized this waiver on 2026-05-27.",
                }
            ],
        ),
    )
    _write_text(tmp_path, "CLAUDE.md", "# Agent Instructions\n")

    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")

    assert result.returncode == 0, result.stderr
    assert "Instruction-contract validation passed with waivers." in result.stdout
    assert "Instruction contract waivers applied:" in result.stdout
    assert "CLAUDE.md: ## Handling Code Review Comments" in result.stdout
    assert "Owner authorized this waiver on 2026-05-27." in result.stdout


def test_required_heading_inside_indented_code_block_is_not_satisfied(
    tmp_path: Path,
) -> None:
    """A heading inside an indented code block (4+ leading spaces) must not satisfy."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=["## Handling Code Review Comments"]),
    )
    _write_text(
        tmp_path,
        "CLAUDE.md",
        "# Agent Instructions\n\nSee example:\n\n    ## Handling Code Review Comments\n",
    )

    result = _run_validator(tmp_path, "--mode", "upstream-template")

    assert result.returncode == 1
    assert "CLAUDE.md: missing required heading: ## Handling Code Review Comments" in result.stdout


def test_required_heading_inside_tab_indented_line_is_not_satisfied(
    tmp_path: Path,
) -> None:
    """A heading with a leading tab is treated as indented code per CommonMark."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=["## Handling Code Review Comments"]),
    )
    _write_text(
        tmp_path,
        "CLAUDE.md",
        "# Agent Instructions\n\n\t## Handling Code Review Comments\n",
    )

    result = _run_validator(tmp_path, "--mode", "upstream-template")

    assert result.returncode == 1
    assert "CLAUDE.md: missing required heading: ## Handling Code Review Comments" in result.stdout


def test_required_heading_with_three_leading_spaces_is_satisfied(tmp_path: Path) -> None:
    """CommonMark allows up to 3 leading spaces for an ATX heading."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=["## Handling Code Review Comments"]),
    )
    _write_text(
        tmp_path,
        "CLAUDE.md",
        "# Agent Instructions\n\n   ## Handling Code Review Comments\n",
    )

    result = _run_validator(tmp_path, "--mode", "upstream-template")

    assert result.returncode == 0, result.stderr
    assert "Instruction-contract validation passed." in result.stdout


def test_required_heading_inside_fenced_code_block_is_not_satisfied(tmp_path: Path) -> None:
    """A heading nested inside a fenced code block must not satisfy the contract."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=["## Handling Code Review Comments"]),
    )
    _write_text(
        tmp_path,
        "CLAUDE.md",
        "# Agent Instructions\n\n```markdown\n## Handling Code Review Comments\n```\n",
    )

    result = _run_validator(tmp_path, "--mode", "upstream-template")

    assert result.returncode == 1
    assert "CLAUDE.md: missing required heading: ## Handling Code Review Comments" in result.stdout


def test_blockquote_fence_closure_allows_later_required_heading(
    tmp_path: Path,
) -> None:
    """Instruction contracts use the shared GFM containing-block fence boundary."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=["## Handling Code Review Comments"]),
    )
    _write_text(
        tmp_path,
        "CLAUDE.md",
        (
            "# Agent Instructions\n\n"
            "> ```\n"
            "> ## Handling Code Review Comments\n"
            "\n"
            "## Handling Code Review Comments\n"
        ),
    )

    result = _run_validator(tmp_path, "--mode", "upstream-template")

    assert result.returncode == 0, result.stderr
    assert "Instruction-contract validation passed." in result.stdout


def test_required_phrase_inside_fenced_code_block_is_not_satisfied(tmp_path: Path) -> None:
    """A phrase nested inside a fenced code block must not satisfy the contract."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_phrases=["Protected-file authorization checkpoint"]),
    )
    _write_text(
        tmp_path,
        "CLAUDE.md",
        "# Agent Instructions\n\n~~~text\nProtected-file authorization checkpoint\n~~~\n",
    )

    result = _run_validator(tmp_path, "--mode", "upstream-template")

    assert result.returncode == 1
    assert (
        "CLAUDE.md: missing required phrase: Protected-file authorization checkpoint"
        in result.stdout
    )


def test_upstream_mode_skip_if_marker_present_exits_zero_without_validating(
    tmp_path: Path,
) -> None:
    """--skip-if-marker-present makes upstream-template mode a no-op downstream."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=["## Handling Code Review Comments"]),
    )
    _write_yaml(tmp_path, ".template-sync/marker.yml", _marker(["agent-instructions"]))
    _write_text(tmp_path, "CLAUDE.md", "# Agent Instructions\n")

    result = _run_validator(
        tmp_path,
        "--mode",
        "upstream-template",
        "--skip-if-marker-present",
    )

    assert result.returncode == 0, result.stderr
    assert "--mode upstream-template skipped" in result.stdout
    assert ".template-sync/marker.yml" in result.stdout
    assert "Instruction-contract validation failed." not in result.stdout


def test_upstream_mode_skip_if_marker_present_runs_when_marker_absent(
    tmp_path: Path,
) -> None:
    """--skip-if-marker-present has no effect when the marker is absent."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=["## Handling Code Review Comments"]),
    )
    _write_text(tmp_path, "CLAUDE.md", "# Agent Instructions\n")

    result = _run_validator(
        tmp_path,
        "--mode",
        "upstream-template",
        "--skip-if-marker-present",
    )

    assert result.returncode == 1
    assert "CLAUDE.md: missing required heading: ## Handling Code Review Comments" in result.stdout


def test_downstream_skips_azure_devops_contract_when_module_excluded(
    tmp_path: Path,
) -> None:
    """Azure-specific contracts are not mandatory for GitHub-only adopters."""
    _write_common_contract_repo(tmp_path, _host_specific_contracts())
    _write_yaml(tmp_path, ".template-sync/marker.yml", _marker(["agent-instructions"]))
    _write_text(tmp_path, "CLAUDE.md", "# Agent Instructions\n\n## Handling Code Review Comments\n")

    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")

    assert result.returncode == 0, result.stderr
    assert "Instruction-contract validation passed." in result.stdout
    # CLAUDE.md is checked while the Azure-only GEMINI.md contract is skipped
    # because azure-devops-collaboration is not retained. Assert on these stable
    # signals rather than the exact skipped-contract line, whose module ordering
    # and phrasing are incidental formatting details.
    assert "Contracts checked: 1" in result.stdout
    skipped_entries = _section_entries(
        result.stdout, "Contracts skipped by downstream module selection"
    )
    assert len(skipped_entries) == 1
    (skipped_entry,) = skipped_entries
    assert skipped_entry.startswith("GEMINI.md ")
    assert "azure-devops-collaboration" in skipped_entry


def test_downstream_checks_azure_devops_contract_when_module_retained(
    tmp_path: Path,
) -> None:
    """Azure-specific contracts are enforced only when their Azure module is retained."""
    _write_common_contract_repo(tmp_path, _host_specific_contracts())
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["agent-instructions", "azure-devops-collaboration"]),
    )
    _write_text(tmp_path, "CLAUDE.md", "# Agent Instructions\n\n## Handling Code Review Comments\n")
    _write_text(tmp_path, "GEMINI.md", "# Agent Instructions\n")

    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")

    assert result.returncode == 1
    assert "Instruction-contract validation failed." in result.stdout
    assert (
        "GEMINI.md: missing required heading: ## Azure DevOps PR Review Protocol" in result.stdout
    )


def test_downstream_stale_protected_guide_section_fails_when_module_excluded(
    tmp_path: Path,
) -> None:
    """Protected-guide sections for excluded modules need owner review or a waiver."""
    contracts = _contracts(required_headings=["## Handling Code Review Comments"])
    contracts["protected_guide_section_obligations"] = [
        {
            "key": "agents-azure-devops-pr-review-protocol",
            "path": "AGENTS.md",
            "target_modules": ["azure-devops-collaboration"],
            "stale_headings": ["## Azure DevOps PR Review Protocol"],
        }
    ]
    _write_common_contract_repo(tmp_path, contracts)
    _write_yaml(tmp_path, ".template-sync/marker.yml", _marker(["agent-instructions"]))
    _write_text(tmp_path, "CLAUDE.md", "## Handling Code Review Comments\n")
    _write_text(tmp_path, "AGENTS.md", "## Azure DevOps PR Review Protocol\n")

    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")

    assert result.returncode == 1
    assert "Stale protected-guide sections requiring owner review:" in result.stdout
    assert "AGENTS.md: agents-azure-devops-pr-review-protocol: stale heading" in result.stdout


def test_downstream_stale_protected_guide_section_waiver_passes_loudly(
    tmp_path: Path,
) -> None:
    """A protected-guide section waiver passes validation but remains visible."""
    contracts = _contracts(required_headings=["## Handling Code Review Comments"])
    contracts["protected_guide_section_obligations"] = [
        {
            "key": "agents-azure-devops-pr-review-protocol",
            "path": "AGENTS.md",
            "target_modules": ["azure-devops-collaboration"],
            "stale_headings": ["## Azure DevOps PR Review Protocol"],
        }
    ]
    _write_common_contract_repo(tmp_path, contracts)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            protected_guide_waivers=[
                {
                    "path": "AGENTS.md",
                    "contract_key": "agents-azure-devops-pr-review-protocol",
                    "target_module": "azure-devops-collaboration",
                    "reason": "GitHub-only fixture retains the protected Azure protocol.",
                    "authorization_basis": "Owner authorized this protected-guide waiver.",
                }
            ],
        ),
    )
    _write_text(tmp_path, "CLAUDE.md", "## Handling Code Review Comments\n")
    _write_text(tmp_path, "AGENTS.md", "## Azure DevOps PR Review Protocol\n")

    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")

    assert result.returncode == 0, result.stderr
    assert "Instruction-contract validation passed with waivers." in result.stdout
    assert "Protected guide contract waivers applied:" in result.stdout
    assert "AGENTS.md: agents-azure-devops-pr-review-protocol" in result.stdout


def test_downstream_duplicate_waiver_pairs_fail(tmp_path: Path) -> None:
    """Duplicate (path, anchor) waivers fail fast instead of silently de-duplicating."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=["## Handling Code Review Comments"]),
    )
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            waivers=[
                {
                    "path": "CLAUDE.md",
                    "anchor": "## Handling Code Review Comments",
                    "reason": "First waiver.",
                    "authorization_basis": "Owner authorized this waiver on 2026-05-27.",
                },
                {
                    "path": "CLAUDE.md",
                    "anchor": "## Handling Code Review Comments",
                    "reason": "Conflicting second waiver for the same anchor.",
                    "authorization_basis": "Owner re-authorized on 2026-05-27.",
                },
            ],
        ),
    )
    _write_text(tmp_path, "CLAUDE.md", "# Agent Instructions\n")

    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")

    assert result.returncode == 1
    assert (
        "Duplicate template_sync.instruction_contract_waivers (path, anchor) pair(s):"
        in result.stderr
    )
    assert "(CLAUDE.md, ## Handling Code Review Comments)" in result.stderr


def test_file_absent_without_authorized_remove_local_fails(tmp_path: Path) -> None:
    """A retained contract file cannot disappear without protected-file authorization."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=["## Handling Code Review Comments"]),
    )
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["agent-instructions"]),
    )

    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")

    assert result.returncode == 1
    assert "Required instruction files absent without authorized removal:" in result.stdout
    assert "CLAUDE.md" in result.stdout


def test_file_absent_with_authorized_remove_local_is_visible_skip(tmp_path: Path) -> None:
    """An authorized protected-file removal skips anchors visibly."""
    _write_common_contract_repo(
        tmp_path,
        _contracts(required_headings=["## Handling Code Review Comments"]),
    )
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            protected_decisions=[
                {
                    "path": "CLAUDE.md",
                    "decision": "REMOVE-LOCAL",
                    "authorization_basis": "Owner explicitly authorized removing CLAUDE.md.",
                    "authorized_scope": "CLAUDE.md only.",
                    "reason": "Claude agent is not used by this downstream repository.",
                }
            ],
        ),
    )

    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")

    assert result.returncode == 0, result.stderr
    assert "Authorized removals skipped:" in result.stdout
    assert "CLAUDE.md" in result.stdout
    assert "Owner explicitly authorized removing CLAUDE.md." in result.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("marker", ["0.", "2.", "2)", "0002.", "999999999.", "  2."])
@pytest.mark.parametrize("fence", ["```", "~~~"])
def test_noninterrupting_ordered_fence_keeps_live_clause(
    tmp_path: Path, mode: str, marker: str, fence: str
) -> None:
    """A new non-1 ordered item cannot turn paragraph continuation into code."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": ["Agents MUST review changes."],
    }
    text = _render_section(section) + marker + " " + fence + "\n   Agents MAY bypass.\n   " + fence
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize(
    ("clause", "separator", "marker"),
    [
        ("Agents MUST review changes.", "\n", "2."),
        ("Agents MUST review changes.", "", "1."),
        ("Agents MUST review changes.", "", "0001)"),
        ("1. Agents MUST review changes.", "", "2."),
        ("- Agents MUST review changes.", "", "2)"),
    ],
)
def test_real_list_fences_remain_inert(
    tmp_path: Path, clause: str, separator: str, marker: str
) -> None:
    """Blank-separated lists, interrupting starts, and real siblings keep code inert."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [clause],
    }
    margin = " " * (len(marker) + 1)
    text = (
        _render_section(section)
        + separator
        + marker
        + " ```\n"
        + margin
        + "Agents MAY bypass.\n"
        + margin
        + "```\n"
    )
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("fence", ["```", "~~~", "`````"])
@pytest.mark.parametrize("trailing", ["\t", " \t", "\t \t"])
@pytest.mark.parametrize("successor", [None, "## Following section"])
def test_tabbed_closing_fence_exposes_later_policy(
    tmp_path: Path, mode: str, fence: str, trailing: str, successor: str | None
) -> None:
    """Valid closing-fence whitespace cannot conceal the following live clause."""
    section = _scoped_policy()
    text = _render_section(section) + "\n" + fence + "\nexample\n" + fence + trailing
    text += "\nAgents MAY bypass.\n"
    section["next_heading"] = successor
    if successor is not None:
        text += successor + "\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    ("paragraphs", "diagnostic"),
    [
        ([" \t\n"], "Empty normalized"),
        (["\u00a0\u2003"], "Empty normalized"),
        (["Rule  text", "Rule text"], "Duplicate normalized"),
        (["Rule\ntext", "Rule\ttext"], "Duplicate normalized"),
    ],
)
def test_catalog_rejects_unsatisfiable_normalized_paragraphs(
    tmp_path: Path, mode: str, paragraphs: list[str], diagnostic: str
) -> None:
    """Structurally valid but unsatisfiable expectations fail during semantic loading."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": paragraphs,
    }
    _write_scoped_repo(tmp_path, section, "## Review decisions\n\nRule text.\n")
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert diagnostic + " contract paragraph" in result.stderr
    assert "missing required section content" not in result.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
def test_catalog_accepts_unique_wrapped_paragraphs(tmp_path: Path, mode: str) -> None:
    """Normalization preserves author wrapping without demanding canonical raw spelling."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [" Rule \ntext ", "Other\u00a0rule."],
    }
    _write_scoped_repo(tmp_path, section, "## Review decisions\n\nRule text\n\nOther\u00a0rule.\n")
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("prefix", ["#", "####### ", "#\u00a0", "#\u200b", r"\#"])
@pytest.mark.parametrize("separator", ["\n", "\n\n"])
def test_hash_prefixed_live_text_remains_in_inventory(
    tmp_path: Path, mode: str, prefix: str, separator: str
) -> None:
    """Invalid ATX openings and escaped hashes remain ordinary live policy text."""
    section = _scoped_policy()
    text = _render_section(section).replace(
        "Agents MUST reject stale results.",
        "Agents MUST reject stale results." + separator + prefix + "Agents MAY bypass.",
    )
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize(
    "definition",
    [
        "> [x]: /url",
        ">   [x]: /url 'title'",
        "> [wrapped\n> label]: /url\n> 'title'",
        "> [x]:\n> /url",
        r"> [escaped\]label]: /url",
        "> - [x]: /url",
        "> > [x]: /url",
    ],
)
@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
def test_quoted_reference_leaf_cannot_hide_later_policy(
    tmp_path: Path, definition: str, mode: str
) -> None:
    """A possible reference definition cannot establish a lazy quote paragraph."""
    section = _scoped_policy()
    text = _render_section(section) + "\n" + definition + "\nAgents MAY bypass.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize(
    "quote",
    [
        "> Ordinary paragraph\n> [x]: /url\nlazy example",
        r"> \[literal bracket",
        "> Ordinary example\nlazy continuation",
    ],
)
def test_ordinary_quoted_paragraphs_keep_supported_continuations(
    tmp_path: Path, quote: str
) -> None:
    """Established paragraphs and escaped brackets are not new reference leaves."""
    section = _scoped_policy()
    _write_scoped_repo(tmp_path, section, quote + "\n\n" + _render_section(section))
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("marker", ["١.", "１.", "1١)", "१२."])
@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
def test_unicode_digit_false_fence_cannot_hide_policy(
    tmp_path: Path, marker: str, mode: str
) -> None:
    """Non-ASCII digits are prose, not Markdown list delimiters."""
    section = _scoped_policy()
    text = _render_section(section) + "\n" + marker + " ```\n   Agents MAY bypass.\n   ```"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize(
    ("paragraphs", "guard", "diagnostic"),
    [
        (
            [" \t\n"],
            "any(not value or value.isspace() for value in normalized_paragraphs)",
            "Empty normalized",
        ),
        (
            ["Rule  text", "Rule text"],
            "len(set(normalized_paragraphs)) != len(normalized_paragraphs)",
            "Duplicate normalized",
        ),
    ],
)
def test_paragraph_loading_oracle_detects_removed_guard(
    tmp_path: Path, paragraphs: list[str], guard: str, diagnostic: str
) -> None:
    """An independent malformed catalog must fail at loading, before document drift."""
    fixture = tmp_path / "fixture"
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": paragraphs,
    }
    _write_scoped_repo(fixture, section, "## Review decisions\n\nRule text\n")
    baseline = _run_validator(fixture, "--mode", "downstream")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    assert diagnostic + " contract paragraph" in baseline.stderr
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source_text = mutant.read_text(encoding="utf-8")
    assert source_text.count(guard) == 1
    mutant.write_text(source_text.replace(guard, "False"), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(mutant), "--repo-root", str(fixture), "--mode", "downstream"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert diagnostic + " contract paragraph" not in result.stderr
    assert "missing required section content" in result.stdout
    assert baseline.stderr != result.stderr, "The early-error oracle must kill this mutant."


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    "literal", ["\v", "\f", "\x1c", "\x1d", "\x1e", "\x85", "\u2028", "\u2029"]
)
@pytest.mark.parametrize("fence", ["```", "~~~"])
def test_nonphysical_separator_cannot_expose_a_required_clause(
    tmp_path: Path, mode: str, literal: str, fence: str
) -> None:
    """A non-CR/LF suffix leaves the sole obligation inside literal fenced code."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": ["Agents MUST reject stale results."],
    }
    text = section["heading"] + "\n\n" + fence + "\nexample\n" + fence + literal
    text += "\n" + section["required_paragraphs"][0] + "\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraph:" in result.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    "literal",
    ["\v", "\f", "\x1c", "\x1d", "\x1e", "\x85", "\u2028", "\u2029", "\u00a0", "\u2003", "\x1f"],
)
@pytest.mark.parametrize("position", ["prefix", "suffix"])
def test_non_ascii_whitespace_cannot_supply_required_headings(
    tmp_path: Path, mode: str, literal: str, position: str
) -> None:
    """Neither loose nor section headings discard literal Unicode/control characters."""
    section = _scoped_policy()
    heading = section["heading"]
    fake = literal + heading if position == "prefix" else heading + literal
    text = _render_section(section).replace(heading, fake, 1)
    _write_scoped_repo(tmp_path, section, text)
    contracts = _contracts(required_headings=[heading])
    contracts["instruction_contracts"][0]["required_sections"] = [section]
    _write_yaml(tmp_path, ".template-sync/instruction-contracts.yml", contracts)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "missing required heading: ## Review decisions" in result.stdout
    assert "missing required section content: section:## Review decisions" in result.stdout


@pytest.mark.parametrize("kind", ["splitlines", "strip"])
def test_physical_line_and_heading_oracle_detects_removed_guard(tmp_path: Path, kind: str) -> None:
    """Restoring Python's broader semantics falsely accepts independently invalid input."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": ["Agents MUST reject stale results."],
    }
    if kind == "splitlines":
        text = section["heading"] + "\n\n```\nexample\n```\v\n"
        text += section["required_paragraphs"][0]
        original, replacement = "lines = markdown_lines(text)", "lines = text.splitlines()"
    else:
        text = "\u00a0" + _render_section(section)
        original = 'observed = "".join(visible).strip(" \\t")'
        replacement = 'observed = "".join(visible).strip()'
    fixture = tmp_path / "fixture"
    _write_scoped_repo(fixture, section, text)
    baseline = _run_validator(fixture, "--mode", "downstream")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source_text = mutant.read_text(encoding="utf-8")
    assert source_text.count(original) == 1
    mutant.write_text(source_text.replace(original, replacement), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(mutant), "--repo-root", str(fixture), "--mode", "downstream"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_non_ascii_blank_lines_preserve_quote_and_code_span_state() -> None:
    """Literal whitespace cannot reset unsupported quotes or inline span boundaries."""
    program = r"""
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)
for literal in ("\v", "\f", "\x1c", "\x1d", "\x1e", "\x85", "\u2028", "\u2029", "\u00a0", "\u2003", "\x1f"):
    lines = ["`start", literal, "end`"]
    assert validator.policy_code_span_ends(lines)[(0, 0)] == (2, 4)
    observed = validator.operative_markdown_lines("> <unsupported>\n" + literal + "\nAgents MUST act.")
    assert observed[-1] == "[unsupported quoted policy] Agents MUST act."
for blank in ("", " ", "\t", " \t"):
    assert validator.policy_code_span_ends(["`start", blank, "end`"]) == {}
    observed = validator.operative_markdown_lines("> <unsupported>\n" + blank + "\nAgents MUST act.")
    assert observed[-1] == "Agents MUST act."
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
def test_physical_lines_preserve_literal_boundaries_and_ascii_headings(
    tmp_path: Path, mode: str
) -> None:
    """Unicode inside a fence stays inert; supported ASCII headings remain operative."""
    section = _scoped_policy()
    for indent in range(4):
        text = _render_section(section).replace(
            section["heading"], " " * indent + section["heading"] + " \t", 1
        )
        for literal in ("\v", "\f", "\x1c", "\x1d", "\x1e", "\x85", "\u2028", "\u2029"):
            text += "\n~~~\nexample\n~~~" + literal + "\n## Inert example heading\n~~~\n"
        _write_scoped_repo(tmp_path, section, text)
        result = _run_validator(tmp_path, "--mode", mode)
        assert result.returncode == 0, result.stdout + result.stderr


def _scoped_policy() -> dict[str, Any]:
    """Provide an independent security/failure-truth oracle for parser tests."""
    return {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [
            "Agents MUST reject stale results.",
            "Retry delivery after 120 seconds, at most twice.",
        ],
        "required_tables": [
            {
                "headers": ["State", "Action", "Gate"],
                "rows": [
                    ["Failed", "Retry only failed service, at most three attempts", "Not clean"],
                    ["Stale", "Count toward ten complete pending observations", "Unknown"],
                    ["One clean service", "Continue the other service", "Incomplete"],
                    ["Both clean", "Reconcile all findings and CI", "Review complete"],
                ],
            }
        ],
    }


def _render_section(section: dict[str, Any]) -> str:
    """Render fixture input independently of the production Markdown parser."""
    paragraphs = list(section.get("required_paragraphs", []))
    tables: list[str] = []
    for table in section.get("required_tables", []):
        lines = ["| " + " | ".join(table["headers"]) + " |"]
        lines.append("| " + " | ".join("---" for _ in table["headers"]) + " |")
        lines.extend("| " + " | ".join(row) + " |" for row in table["rows"])
        tables.append("\n".join(lines))
    order = section.get(
        "required_block_order", ["paragraph"] * len(paragraphs) + ["table"] * len(tables)
    )
    blocks = [section["heading"]]
    for kind in order:
        blocks.append(paragraphs.pop(0) if kind == "paragraph" else tables.pop(0))
    if section.get("next_heading") is not None:
        blocks.append(section["next_heading"])
    return "\n\n".join(blocks) + "\n"


def _write_scoped_repo(tmp_path: Path, section: dict[str, Any], text: str) -> None:
    """Use a portable retained-agent fixture, never optional upstream files."""
    contracts = _contracts()
    contracts["instruction_contracts"][0]["required_sections"] = [section]
    _write_common_contract_repo(tmp_path, contracts)
    _write_text(tmp_path, "CLAUDE.md", text)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(sorted({"agent-instructions", *section.get("requires_modules", [])})),
    )


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
def test_scoped_policy_cli_accepts_complete_live_content(tmp_path: Path, mode: str) -> None:
    """Full clauses and decision rows are accepted in both supported modes."""
    section = _scoped_policy()
    _write_scoped_repo(tmp_path, section, _render_section(section))
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize(
    ("original", "replacement"),
    [
        ("MUST reject", "MAY reject"),
        ("120 seconds", "121 seconds"),
        ("at most twice", "without a bound"),
        ("at most three attempts", "at most four attempts"),
        ("Not clean", "Clean"),
        ("Count toward ten complete pending observations", "Reset the wait counter"),
        ("Continue the other service", "Finish the pair"),
        ("Reconcile all findings and CI", "Ignore earlier findings"),
        ("| --- | --- | --- |", "| not a delimiter | --- | --- |"),
        ("| Failed |", "| Failed | extra |"),
    ],
)
def test_scoped_policy_cli_rejects_semantic_mutations(
    tmp_path: Path, original: str, replacement: str
) -> None:
    """Independent expected failures guard retry, attribution, and completion truth."""
    section = _scoped_policy()
    _write_scoped_repo(tmp_path, section, _render_section(section).replace(original, replacement))
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "missing required section content" in result.stdout


@pytest.mark.parametrize("decoy", ["fence", "quote", "lazy-quote", "indent", "comment", "sibling"])
def test_scoped_policy_cli_rejects_inert_or_relocated_decoys(tmp_path: Path, decoy: str) -> None:
    """A complete-looking example cannot replace the operative section."""
    section = _scoped_policy()
    content = _render_section(section)
    if decoy == "fence":
        content = "```markdown\n" + content + "```\n"
    elif decoy == "quote":
        content = "\n".join("> " + line for line in content.splitlines())
    elif decoy == "lazy-quote":
        content = (
            section["heading"]
            + "\n\n> Example only\n"
            + "\n".join(section["required_paragraphs"])
            + "\n"
        )
    elif decoy == "indent":
        content = "\n".join("    " + line for line in content.splitlines())
    elif decoy == "comment":
        content = "<!--\n" + content + "-->\n"
    else:
        content = section["heading"] + "\n\n## Unrelated\n" + content.split("\n", 1)[1]
    _write_scoped_repo(tmp_path, section, content)
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1, result.stdout + result.stderr


@pytest.mark.parametrize(
    "mutation", ["heading", "table", "row", "extra", "order", "paragraph-order"]
)
def test_scoped_policy_cli_rejects_duplicates_and_reordering(tmp_path: Path, mutation: str) -> None:
    """Ambiguous sections, extra decisions, and changed execution order fail closed."""
    section = _scoped_policy()
    content = _render_section(section)
    changed = copy.deepcopy(section)
    if mutation == "heading":
        content += content
    elif mutation == "table":
        changed["required_tables"] *= 2
        content = _render_section(changed)
    elif mutation == "row":
        changed["required_tables"][0]["rows"] *= 2
        content = _render_section(changed)
    elif mutation == "extra":
        changed["required_tables"][0]["rows"].append(["Expired", "Declare success", "Clean"])
        content = _render_section(changed)
    elif mutation == "order":
        changed["required_tables"][0]["rows"].reverse()
        content = _render_section(changed)
    else:
        changed["required_paragraphs"].reverse()
        content = _render_section(changed)
    _write_scoped_repo(tmp_path, section, content)
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1, result.stdout + result.stderr


def _catalog_sections() -> list[Any]:
    """Collect catalog data retained by template-sync support in downstream trees."""
    catalog = yaml.safe_load(
        (REPO_ROOT / ".template-sync/instruction-contracts.yml").read_text(encoding="utf-8")
    )
    return [
        pytest.param(
            section,
            sorted(set(contract["requires_modules"]) | set(section.get("requires_modules", []))),
            id=f"{contract['path']}:{section['heading']}",
        )
        for contract in catalog["instruction_contracts"]
        for section in contract.get("required_sections", [])
    ]


@pytest.mark.parametrize(("section", "modules"), _catalog_sections())
def test_every_catalog_invariant_has_downstream_cli_mutation_coverage(
    tmp_path: Path, section: dict[str, Any], modules: list[str]
) -> None:
    """Every retained contract clause/row is enforced without optional source files.

    The fixture is contract data; the expected native failure is independent of
    the production predicates. Disabling either validation loop makes this test
    fail. This suite deliberately remains enabled under downstream selection.
    """
    content = _render_section(section)
    _write_scoped_repo(tmp_path, section, content)
    shutil.copyfile(
        REPO_ROOT / ".template-sync/manifest.yml", tmp_path / ".template-sync/manifest.yml"
    )
    _write_yaml(tmp_path, ".template-sync/marker.yml", _marker(modules))
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 0, result.stdout + result.stderr
    mutations: list[str] = []
    for paragraph in section.get("required_paragraphs", []):
        mutations.append(content.replace(paragraph, "Removed obligation.", 1))
    for table in section.get("required_tables", []):
        for row in table["rows"]:
            original = "| " + " | ".join(row) + " |"
            mutations.append(
                content.replace(original, original.replace(row[-1], "Changed gate"), 1)
            )
    assert mutations, "Every scoped contract must enforce actual content."
    for index, mutated in enumerate(mutations):
        _write_text(tmp_path, "CLAUDE.md", mutated)
        result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
        assert result.returncode == 1, f"mutation {index}: {result.stdout}{result.stderr}"
        assert "missing required section content" in result.stdout


def test_scoped_waiver_is_specific_and_reported(tmp_path: Path) -> None:
    """An explicit clause waiver cannot silently waive the whole policy."""
    section = _scoped_policy()
    text = _render_section(section).replace("Agents MUST reject stale results.", "")
    _write_scoped_repo(tmp_path, section, text)
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            waivers=[
                {
                    "path": "CLAUDE.md",
                    "anchor": "section:## Review decisions:paragraph:"
                    + hashlib.sha256(b"Agents MUST reject stale results.").hexdigest(),
                    "reason": "Owner selected a different local review policy.",
                    "authorization_basis": "Explicit fixture owner authorization for this clause only.",
                }
            ],
        ),
    )
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 0, result.stdout + result.stderr
    assert "passed with waivers" in result.stdout
    _write_text(tmp_path, "CLAUDE.md", text.replace("Not clean", "Clean"))
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1
    assert "section:## Review decisions:tables" in result.stdout


def test_malformed_table_waiver_cannot_hide_a_weakened_clause(tmp_path: Path) -> None:
    """A waived table failure leaves independently violated paragraphs failing."""
    section = _scoped_policy()
    text = _render_section(section).replace("MUST reject", "MAY reject")
    text = text.replace("| --- | --- | --- |", "| broken | --- | --- |")
    _write_scoped_repo(tmp_path, section, text)
    initial = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert initial.returncode == 1
    anchor = re.search(r"section:## Review decisions:tables:[0-9a-f]{64}", initial.stdout)
    assert anchor is not None
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            waivers=[
                {
                    "path": "CLAUDE.md",
                    "anchor": anchor.group(),
                    "reason": "Fixture owner permits a different table format.",
                    "authorization_basis": "Explicit authorization for this table inventory only.",
                }
            ],
        ),
    )
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1, result.stdout + result.stderr
    assert (
        "missing required section content: section:## Review decisions:paragraph:" in result.stdout
    )
    assert "Instruction contract waivers applied:" in result.stdout


def test_section_heading_waiver_cannot_hide_its_content(tmp_path: Path) -> None:
    """A missing-heading exception does not authorize removal of the whole policy."""
    _write_scoped_repo(tmp_path, _scoped_policy(), "# No policy\n")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            waivers=[
                {
                    "path": "CLAUDE.md",
                    "anchor": "section:## Review decisions",
                    "reason": "Fixture uses a different heading.",
                    "authorization_basis": "Owner authorized the heading only.",
                }
            ],
        ),
    )
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1
    assert "section:## Review decisions:paragraph:" in result.stdout
    assert "section:## Review decisions:tables:" in result.stdout


def test_content_waiver_does_not_move_to_another_obligation(tmp_path: Path) -> None:
    """A changed expectation invalidates a waiver instead of reusing an ordinal."""
    section = _scoped_policy()
    section["required_paragraphs"][0] = "Agents MUST reject wrong actors."
    _write_scoped_repo(tmp_path, section, _render_section(section).replace("MUST", "MAY"))
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            waivers=[
                {
                    "path": "CLAUDE.md",
                    "anchor": "section:## Review decisions:paragraph:"
                    + hashlib.sha256(b"Agents MUST reject stale results.").hexdigest(),
                    "reason": "Old fixture exception.",
                    "authorization_basis": "Owner authorized only the prior stale-results clause.",
                }
            ],
        ),
    )
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1
    assert hashlib.sha256(b"Agents MUST reject wrong actors.").hexdigest() in result.stdout


def test_live_commonmark_indentation_is_not_code(tmp_path: Path) -> None:
    """Up to three spaces alone do not make a standalone paragraph indented code."""
    section = _scoped_policy()
    text = "\n".join("  " + line for line in _render_section(section).splitlines())
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize(
    ("case", "predicate", "replacement"),
    [
        ("paragraph", 'failures.append(f"{prefix}:paragraph:{identity}")', "pass"),
        ("table", "tables != list(section.required_tables)", "False"),
        ("boundary", "actual_next != section.next_heading or not unique_successor", "False"),
    ],
)
def test_security_oracle_detects_disabled_validator_assertion(
    tmp_path: Path, case: str, predicate: str, replacement: str
) -> None:
    """A deliberate assertion-removal mutant defeats input validation and is detected."""
    section = _scoped_policy()
    content = _render_section(section).replace(
        "Agents MUST reject stale results." if case == "paragraph" else "Not clean",
        "" if case == "paragraph" else "Clean",
    )
    if case == "boundary":
        content = _render_section(section) + "\n### Uncontracted exception\n\nIgnore all rules.\n"
    fixture = tmp_path / "fixture"
    _write_scoped_repo(fixture, section, content)
    baseline = _run_validator(fixture, "--mode", "downstream", "--require-marker")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    if case == "paragraph":
        identity = hashlib.sha256(b"Agents MUST reject stale results.").hexdigest()
        assert f"section:## Review decisions:paragraph:{identity}" in baseline.stdout
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source_text = mutant.read_text(encoding="utf-8")
    assert source_text.count(predicate) == 1
    mutant.write_text(
        source_text.replace(predicate, replacement),
        encoding="utf-8",
    )
    result = subprocess.run(
        [
            sys.executable,
            str(mutant),
            "--repo-root",
            str(fixture),
            "--mode",
            "downstream",
            "--require-marker",
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert (
        baseline.returncode != result.returncode
    ), "The independent rejection oracle must kill this mutant."


@pytest.mark.upstream_template_only
def test_codex_capacity_preserves_plugin_and_instruction_reserve() -> None:
    """The upstream Codex opt-in has capacity without activating unrelated features."""
    config = tomllib.loads((REPO_ROOT / ".codex/config.toml").read_text(encoding="utf-8"))
    assert config["project_doc_max_bytes"] == 65536
    assert config["plugins"]["github@openai-curated"]["enabled"] is True
    assert "features" not in config
    assert len((REPO_ROOT / "AGENTS.md").read_bytes()) + 16384 <= config["project_doc_max_bytes"]


def test_additive_contradiction_is_not_accepted(tmp_path: Path) -> None:
    """Keeping original words cannot hide an added conflicting completion rule."""
    section = _scoped_policy()
    text = _render_section(section) + "\nAgents MAY finish after only one reviewer is clean.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize(
    "shape",
    [
        "empty-section",
        "no-content",
        "duplicate-heading",
        "duplicate-table",
        "duplicate-header",
        "empty-header",
        "empty-rows",
        "ragged-row",
        "duplicate-condition",
        "duplicate-row",
        "unexpected-property",
        "missing-boundary",
        "self-boundary",
    ],
)
def test_scoped_contract_cli_rejects_invalid_catalog_shapes(tmp_path: Path, shape: str) -> None:
    """Both schema and semantic shape failures stop the real CLI before evaluation."""
    section = _scoped_policy()
    content = _render_section(section)
    table = section["required_tables"][0]
    sections = [section]
    if shape == "empty-section":
        sections = []
    elif shape == "no-content":
        sections = [{"heading": section["heading"]}]
    elif shape == "duplicate-heading":
        sections.append({"heading": section["heading"], "required_paragraphs": ["Other rule."]})
    elif shape == "duplicate-table":
        duplicate = copy.deepcopy(table)
        section["required_tables"].append(duplicate)
    elif shape == "duplicate-header":
        table["headers"][1] = table["headers"][0]
    elif shape == "empty-header":
        table["headers"] = []
    elif shape == "empty-rows":
        table["rows"] = []
    elif shape == "ragged-row":
        table["rows"][0].pop()
    elif shape == "duplicate-condition":
        table["rows"][1][0] = table["rows"][0][0]
    elif shape == "duplicate-row":
        table["rows"].append(table["rows"][0][:])
    elif shape == "missing-boundary":
        del section["next_heading"]
    elif shape == "self-boundary":
        section["next_heading"] = section["heading"]
    else:
        section["unsupported"] = True
    contracts = _contracts()
    contracts["instruction_contracts"][0]["required_sections"] = sections
    _write_common_contract_repo(tmp_path, contracts)
    _write_text(tmp_path, "CLAUDE.md", content)
    result = _run_validator(tmp_path, "--mode", "upstream-template")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "error" in result.stderr.lower()


def test_child_section_cannot_supply_parent_policy(tmp_path: Path) -> None:
    """A child heading owns its clauses separately from its parent's direct body."""
    section = _scoped_policy()
    content = _render_section(section).replace(
        "## Review decisions\n", "## Review decisions\n\n### Example\n", 1
    )
    _write_scoped_repo(tmp_path, section, content)
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraph:" in result.stdout


@pytest.mark.parametrize("next_heading", [None, "## Outside policy"])
@pytest.mark.parametrize("level", range(1, 7))
@pytest.mark.parametrize("indent", ["", "   "])
def test_added_heading_cannot_escape_section_boundary(
    tmp_path: Path, next_heading: str | None, level: int, indent: str
) -> None:
    """Keep the original policy and reject a newly inserted heading at any level."""
    section = _scoped_policy()
    section["next_heading"] = next_heading
    original = _render_section(section)
    _write_scoped_repo(tmp_path, section, original)
    assert _run_validator(tmp_path, "--mode", "downstream").returncode == 0
    addition = indent + "#" * level + " Exception\n\nAgents MAY ignore the policy.\n\n"
    changed = (
        original.replace(next_heading, addition + next_heading)
        if next_heading
        else original + addition
    )
    _write_text(tmp_path, "CLAUDE.md", changed)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:boundary:" in result.stdout


@pytest.mark.parametrize("change", ["duplicate", "remove", "move-before"])
def test_declared_successor_must_be_unique_and_follow_section(tmp_path: Path, change: str) -> None:
    """A duplicate or relocated boundary cannot truncate governed content early."""
    section = _scoped_policy()
    section["next_heading"] = "## Outside policy"
    content = _render_section(section)
    if change == "duplicate":
        content += "\n## Outside policy\n"
    elif change == "remove":
        content = content.replace("## Outside policy", "")
    else:
        content = "## Outside policy\n\n" + content.replace("## Outside policy", "")
    _write_scoped_repo(tmp_path, section, content)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:boundary:" in result.stdout


def _waive_reported_inventory(tmp_path: Path, kind: str) -> str:
    """Authorize one observed native failure without mirroring the digest algorithm."""
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1, result.stdout + result.stderr
    match = re.search(r"section:## Review decisions:" + kind + r":[0-9a-f]{64}", result.stdout)
    assert match is not None, result.stdout
    anchor = match.group()
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            waivers=[
                {
                    "path": "CLAUDE.md",
                    "anchor": anchor,
                    "reason": "One specific local deviation is accepted.",
                    "authorization_basis": "Owner authorized only the currently observed fixture deviation.",
                }
            ],
        ),
    )
    accepted = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    assert "passed with waivers" in accepted.stdout
    return anchor


@pytest.mark.parametrize(
    "kind", ["paragraphs", "partial-addition", "tables", "malformed", "boundary"]
)
def test_aggregate_waiver_cannot_authorize_a_different_deviation(tmp_path: Path, kind: str) -> None:
    """A previously authorized local difference cannot silently expand or change."""
    section = _scoped_policy()
    original = _render_section(section)
    if kind in {"paragraphs", "partial-addition"}:
        first = original + "\nAccepted local A.\n\nAccepted local B.\n"
        second = first.replace(
            "Accepted local A.", "Changed local rule." if kind == "paragraphs" else ""
        )
        anchor_kind = "paragraphs"
    elif kind == "boundary":
        first = original + "\n### Exception\n\nAccepted local A.\n"
        second = first.replace("Accepted local A.", "A different unauthorized exception.")
        anchor_kind = "boundary"
    else:
        first = original.replace("Not clean", "Accepted gate")
        if kind == "malformed":
            first = first.replace("| --- | --- | --- |", "| malformed | --- | --- |")
        second = first.replace("Accepted gate", "Different gate")
        anchor_kind = "tables"
    _write_scoped_repo(tmp_path, section, first)
    old_anchor = _waive_reported_inventory(tmp_path, anchor_kind)
    _write_text(tmp_path, "CLAUDE.md", second)
    rejected = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "missing required section content: " + old_anchor not in rejected.stdout
    assert (
        "missing required section content: section:## Review decisions:" + anchor_kind
        in rejected.stdout
    )
    _write_text(tmp_path, "CLAUDE.md", original)
    restored = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert restored.returncode == 0, restored.stdout + restored.stderr
    assert "passed with waivers" not in restored.stdout


@pytest.mark.parametrize(
    "example",
    [
        "```html\n<!--\n```\n\n",
        "~~~html\n<!-- literal -->\n~~~\n\n",
        "<!--\n```markdown\n-->\n\n",
        "> <!-- example only\n\n",
        "    <!-- example only\n\n",
    ],
)
def test_code_or_comment_example_cannot_corrupt_following_policy(
    tmp_path: Path, example: str
) -> None:
    """Literal examples stay inert and a fence in a real comment stays commented."""
    section = _scoped_policy()
    _write_scoped_repo(tmp_path, section, example + _render_section(section))
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize(
    "literal", ["`<!--`", "`<!-- example -->`", "``literal ` <!-- -->``", "\\<!--"]
)
def test_inline_code_and_escaped_comment_delimiters_remain_literal(
    tmp_path: Path, literal: str
) -> None:
    """A comment delimiter in a matched code span or after an escape is visible text."""
    section = _scoped_policy()
    section["required_paragraphs"][0] = "Agents MUST preserve " + literal + " in examples."
    _write_scoped_repo(tmp_path, section, _render_section(section))
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


def test_wrapped_inline_code_span_preserves_comment_literal(tmp_path: Path) -> None:
    """Inline code can wrap within one paragraph without turning into an HTML comment."""
    section = _scoped_policy()
    section["required_paragraphs"][
        0
    ] = "Agents MUST preserve `literal continued <!-- example -->` text."
    text = _render_section(section).replace("literal continued", "literal\ncontinued")
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


def test_unmatched_backtick_does_not_hide_a_live_comment(tmp_path: Path) -> None:
    """Unmatched literal backticks cannot make commented obligations operative."""
    section = _scoped_policy()
    clause = section["required_paragraphs"][0]
    text = _render_section(section).replace(clause, "Unmatched ` <!-- " + clause + " -->", 1)
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraph:" in result.stdout


@pytest.mark.parametrize(
    "example",
    [
        "```markdown\n### Example heading\n```",
        "> ### Example heading",
        "    ### Example heading",
        "<!--\n### Example heading\n-->",
    ],
)
def test_inert_headings_do_not_change_a_declared_boundary(tmp_path: Path, example: str) -> None:
    """Only live headings can interrupt a section's declared successor relation."""
    section = _scoped_policy()
    section["next_heading"] = "## Outside policy"
    text = _render_section(section).replace("## Outside policy", example + "\n\n## Outside policy")
    text += "\nThis unrelated section can have its own uncontracted content.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


def test_indented_pseudoheading_cannot_end_a_lazy_quote(tmp_path: Path) -> None:
    """Four-space pseudo-headings cannot promote quoted clauses into live policy."""
    section = _scoped_policy()
    clause = section["required_paragraphs"][0]
    text = _render_section(section).replace(
        clause, "> Example paragraph\n    ### Inert pseudo-heading\n" + clause, 1
    )
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraph:" in result.stdout


@pytest.mark.parametrize("indent", range(4))
def test_live_heading_ends_a_lazy_quote(tmp_path: Path, indent: int) -> None:
    """Zero-to-three-space ATX headings can start a live policy after a quote."""
    section = _scoped_policy()
    text = "> Example paragraph\n" + " " * indent + _render_section(section)
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize(
    "quote_end",
    [
        "> note\n- Extra rule.",
        "> note\n2. Extra rule.",
        "> note\n123456789) Extra rule.",
        "> note\n-",
        "> note\n2.",
        "> note\n***",
        "> note\n---",
        "> note\n```\nliteral example\n```",
        "> note\n~~~\nliteral example\n~~~",
        "> note\n<!-- example -->",
        "> ### Example heading",
        "> > ### Nested heading",
        "> - ### List heading",
        "> ```\n> literal example\n> ```",
        "> - ```\n>   literal example\n>   ```",
        "> ~~~\n> literal example\n> ~~~",
        ">     indented example",
        "> text\n> ===",
        "> text\n> ---",
        "> text\n>",
        "> ***",
    ],
)
def test_live_policy_after_quote_blocks_cannot_disappear(tmp_path: Path, quote_end: str) -> None:
    """An outside extra clause must fail even when no blank ends the preceding quote."""
    section = _scoped_policy()
    text = _render_section(section) + "\n" + quote_end + "\nAgents MAY bypass review.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize(
    "quoted_paragraph",
    [
        "> Example paragraph",
        "> - Example paragraph",
        "> 2. Example paragraph",
        "> > Example paragraph",
        "> Example paragraph\n>     indented continuation",
        "> Example paragraph\n    - indented continuation",
        "> Example paragraph\n\t### indented continuation",
        "> ===",
    ],
)
def test_true_lazy_quote_cannot_supply_required_clause(
    tmp_path: Path, quoted_paragraph: str
) -> None:
    """Actual paragraph continuations remain inert, including nested/list paragraphs."""
    section = _scoped_policy()
    clause = section["required_paragraphs"][0]
    text = _render_section(section).replace(clause, quoted_paragraph + "\n" + clause, 1)
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraph:" in result.stdout


@pytest.mark.parametrize(
    "example",
    [
        "> note\n```\nliteral example\n```",
        "> note\n~~~\nliteral example\n~~~",
        "> ### Example heading",
        "> ```\n> literal example\n> ```",
    ],
)
def test_known_quote_block_before_live_clause_is_accepted(tmp_path: Path, example: str) -> None:
    """Harmless supported quote/fence leaves do not hide the following required clause."""
    section = _scoped_policy()
    clause = section["required_paragraphs"][0]
    text = _render_section(section).replace(clause, example + "\n" + clause, 1)
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize(
    "example",
    [
        "> <custom>",
        "> <!-- example -->",
        "> | example | table |\n> | --- | --- |\n> | text | text |",
        "> - > mixed nested container",
        "> paragraph\n<custom>",
    ],
)
def test_ambiguous_quote_syntax_fails_closed(tmp_path: Path, example: str) -> None:
    """Unsupported regions neither hide added policy nor supply expected live clauses."""
    section = _scoped_policy()
    original = _render_section(section)
    _write_scoped_repo(tmp_path, section, original + "\n" + example + "\nAn extra rule.\n")
    added = _run_validator(tmp_path, "--mode", "downstream")
    assert added.returncode == 1, added.stdout + added.stderr
    clause = section["required_paragraphs"][0]
    _write_text(tmp_path, "CLAUDE.md", original.replace(clause, example + "\n" + clause, 1))
    quoted = _run_validator(tmp_path, "--mode", "downstream")
    assert quoted.returncode == 1, quoted.stdout + quoted.stderr
    assert "section:## Review decisions:paragraph:" in quoted.stdout


@pytest.mark.parametrize("position", ["header", "data"])
@pytest.mark.parametrize(
    "cell",
    [
        " State",
        "State ",
        "A  B",
        "   ",
        "A\tB",
        "A\nB",
        "State\n",
        "A\u00a0B",
        "A\u001cB",
        "A\u0085B",
        "A\ufeffB",
        "",
    ],
)
def test_noncanonical_policy_cell_is_a_catalog_error(
    tmp_path: Path, position: str, cell: str
) -> None:
    """Invalid catalog spelling fails schema and semantic loading before drift checks."""
    section = _scoped_policy()
    content = _render_section(section)
    table = section["required_tables"][0]
    target = table["headers"] if position == "header" else table["rows"][0]
    target[0] = cell
    _write_scoped_repo(tmp_path, section, content)

    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "error" in result.stderr.lower()
    assert "missing required section content:" not in result.stdout

    result = subprocess.run(
        [
            sys.executable,
            "-c",
            (
                "import sys, pathlib, yaml; sys.path.insert(0, sys.argv[1]); "
                "import validate_instruction_contracts as validator; "
                "document = yaml.safe_load(pathlib.Path(sys.argv[2]).read_text(encoding='utf-8')); "
                "validator.parse_required_sections(document['instruction_contracts'][0])"
            ),
            str(SCRIPT_PATH.parent),
            str(tmp_path / ".template-sync/instruction-contracts.yml"),
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode != 0
    assert "Noncanonical contract table cell: ## Review decisions" in result.stderr


@pytest.mark.parametrize("cell", ["State", "State value", "État", "状态", "Gate: `pending`"])
def test_canonical_policy_cells_remain_usable(tmp_path: Path, cell: str) -> None:
    """Single spaces, Unicode text, punctuation, and markup retain exact meaning."""
    section = _scoped_policy()
    section["required_tables"][0]["headers"][0] = cell
    section["required_tables"][0]["rows"][0][0] = cell
    _write_scoped_repo(tmp_path, section, _render_section(section))
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 0, result.stdout + result.stderr


def test_unmatched_code_runs_have_bounded_line_visits() -> None:
    """Distinct unmatched widths cannot cause repeated scans of paragraph tails."""
    program = """
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)

class BoundedLines(list):
    reads = 0
    def __getitem__(self, key):
        self.reads += len(range(*key.indices(len(self)))) if isinstance(key, slice) else 1
        assert self.reads <= 4 * len(self), "Scanner repeatedly revisited paragraph lines"
        return super().__getitem__(key)
    def __iter__(self):
        for index in range(len(self)):
            yield self[index]

source = ["literal " + chr(96) * width for width in range(1, 1001)]
lines = BoundedLines(source)

assert validator.policy_code_span_ends(lines) == {}
assert lines.reads > 0
observed = validator.operative_markdown_lines("\\n".join(source))
assert observed == source
print("Bounded scan preserved every unmatched delimiter.")
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent)],
        check=False,
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert "Bounded scan preserved every unmatched delimiter." in result.stdout


@pytest.mark.parametrize(
    "paragraph",
    [
        "A literal \\`` span <!-- marker` remains visible.",
        "A literal \\``` span <!-- marker`` remains visible.",
        "The ``outer ` inner <!-- marker`` span remains visible.",
        "The `first` and ``second`` spans remain visible.",
    ],
)
def test_indexed_inline_spans_preserve_literal_comments(tmp_path: Path, paragraph: str) -> None:
    """Escaped opener prefixes and mixed delimiter widths retain comment precedence."""
    section = _scoped_policy()
    section["required_paragraphs"][0] = paragraph
    _write_scoped_repo(tmp_path, section, _render_section(section))
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize(
    "parent",
    [
        "- Agents MUST review changes.",
        "+ Agents MUST review changes.",
        "1. Agents MUST review changes.",
        "10) Agents MUST review changes.",
        "-    Agents MUST review changes.",
        "-\tAgents MUST review changes.",
    ],
)
@pytest.mark.parametrize("gap", ["\n", "\n\n"])
def test_nested_list_policy_cannot_disappear(tmp_path: Path, parent: str, gap: str) -> None:
    """A nested exception remains visible across tight and loose list boundaries."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [" ".join(parent.split())],
    }
    text = section["heading"] + "\n\n" + parent + gap + "     - Agents MAY bypass review.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize(
    "nested",
    [
        "    - Agents MAY bypass review.",
        "\tAgents MAY bypass review.",
        "    ### Hidden exception\n    Agents MAY bypass review.",
        "    > Agents MAY bypass review.",
        "    | Exception | Allowed |",
        "    <!-->\n    Agents MAY bypass review.",
        "    ```\n    Agents MAY bypass review.",
        "  <!-- hidden -->\n    Agents MAY bypass review.",
        "lazy continuation\n\n    Agents MAY bypass review.",
    ],
)
def test_nested_list_leaves_fail_closed(tmp_path: Path, nested: str) -> None:
    """Nested structural syntax cannot turn the rest of an item into inert content."""
    parent = "- Agents MUST review changes."
    expected = parent + " lazy continuation" if nested.startswith("lazy") else parent
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [expected],
    }
    _write_scoped_repo(tmp_path, section, section["heading"] + "\n\n" + parent + "\n" + nested)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize("sibling", ["-", "- <!-- hidden -->", "2.", "2. <!-- hidden -->"])
def test_empty_list_sibling_keeps_nested_content_live(tmp_path: Path, sibling: str) -> None:
    """An empty or commented sibling still owns a following indented paragraph."""
    parent = "- Agents MUST review changes."
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [parent, sibling.split()[0]],
    }
    text = section["heading"] + "\n\n" + parent + "\n\n" + sibling
    _write_scoped_repo(tmp_path, section, text + "\n    Agents MAY bypass review.\n")
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize(
    "exit_block", ["<!-- list ends -->", "## Outside policy", "A direct paragraph.\n"]
)
def test_proven_list_exit_keeps_standalone_code_inert(tmp_path: Path, exit_block: str) -> None:
    """A dedented new block ends the list before a standalone indented example."""
    parent = "- Agents MUST review changes."
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [parent],
    }
    if exit_block.startswith("#"):
        section["next_heading"] = exit_block
    elif exit_block.startswith("A"):
        section["required_paragraphs"].append(exit_block.strip())
    text = section["heading"] + "\n\n" + parent + "\n\n" + exit_block
    _write_scoped_repo(tmp_path, section, text + "\n\n    Agents MAY bypass review.\n")
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("indent", ["    ", "\t", "  \t"])
@pytest.mark.parametrize("gap", ["\n", "\n\n", " <!--\nhidden\n-->\n", "\n<!-- hidden -->\n"])
def test_indented_paragraph_context_controls_visibility(
    tmp_path: Path, indent: str, gap: str
) -> None:
    """Only a live paragraph continuation can make otherwise indented text operative."""
    parent = "Agents MUST review changes."
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [parent],
    }
    text = section["heading"] + "\n\n" + parent + gap + indent + "Agents MAY bypass review.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    expected = 1 if gap in {"\n", " <!--\nhidden\n-->\n"} else 0
    assert result.returncode == expected, result.stdout + result.stderr
    if expected:
        assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize("comment", ["<!-->", "<!--->"])
@pytest.mark.parametrize("prefix", ["", " ", "  ", "   ", "<!-- ordinary -->"])
@pytest.mark.parametrize("suffix", ["\n", "\n-->\n", " <!--> <!--->\n"])
def test_short_comment_cannot_hide_later_policy(
    tmp_path: Path, comment: str, prefix: str, suffix: str
) -> None:
    """An overlapping close ends the comment before a later live extra clause."""
    section = _scoped_policy()
    text = (
        _render_section(section) + "\n" + prefix + comment + suffix + "Agents MAY bypass review.\n"
    )
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize("comment", ["<!-->", "<!--->", "<!-- ordinary -->", "<!--\nordinary\n-->"])
def test_complete_comments_remain_inert(tmp_path: Path, comment: str) -> None:
    """Standalone short and ordinary comments remain inert examples."""
    section = _scoped_policy()
    text = comment + "\n\n" + _render_section(section)
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize(
    "kind",
    [
        "list",
        "paragraph",
        "comment",
        "inline-comment",
        "gfm-comment",
        "ordered",
        "tab-fence",
        "hash",
        "reference",
        "unicode-list",
        "comment-space",
        "quoted-padding",
        "quoted-tab",
    ],
)
def test_security_oracle_detects_removed_scanner_guard(tmp_path: Path, kind: str) -> None:
    """Independent live-clause fixtures detect removal of each scanner safeguard."""
    clause = "- Agents MUST review changes." if kind == "list" else "Agents MUST review changes."
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [clause],
    }
    extra = (
        "\n<!-->\nAgents MAY bypass review."
        if kind == "comment"
        else "\n    Agents MAY bypass review."
    )
    if kind == "inline-comment":
        extra = " <!--\nAgents MAY bypass review."
    elif kind == "gfm-comment":
        extra = " <!-- Agents MAY -- bypass review. -->"
    elif kind == "ordered":
        extra = "\n2. ```\n   Agents MAY bypass review.\n   ```"
    elif kind == "tab-fence":
        extra = "\n\n```\nexample\n```\t\nAgents MAY bypass review."
    elif kind == "hash":
        extra = "\n\n#Agents MAY bypass review."
    elif kind == "reference":
        extra = "\n\n> [x]: /url\nAgents MAY bypass review."
    elif kind == "unicode-list":
        extra = "\n\n\u0661. ```\n   Agents MAY bypass review.\n   ```"
    elif kind == "comment-space":
        extra = ""
    elif kind == "quoted-padding":
        extra = "\n\n> -     code\nAgents MAY bypass review."
    elif kind == "quoted-tab":
        extra = "\n\n> > - \tcode\nAgents MAY bypass review."
    fixture = tmp_path / "fixture"
    text = section["heading"] + "\n\n" + clause + extra
    if kind == "comment-space":
        text = text.replace("Agents MUST", "Agents<!--x-->MUST")
    _write_scoped_repo(fixture, section, text)
    baseline = _run_validator(fixture, "--mode", "downstream")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    assert "section:## Review decisions:paragraphs:" in baseline.stdout
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    target = (
        mutant_dir / "instruction_contract_support.py"
        if kind in {"tab-fence", "unicode-list"}
        else mutant
    )
    source_text = target.read_text(encoding="utf-8")
    changes = {
        "list": (
            'result.append("[unsupported nested policy] " + line)',
            'result.append("")',
        ),
        "paragraph": (
            '"[unsupported indented policy] " + line if paragraph_can_continue else ""',
            '""',
        ),
        "comment": ('line.find("-->", column + 2)', 'line.find("-->", column + 4)'),
        "inline-comment": ("end == -1 and not block_comment", "False"),
        "gfm-comment": ("not block_comment and ambiguous_inline", "False"),
        "ordered": ("and int(ordered.group(1)) != 1", "and False"),
        "tab-fence": (
            'stripped[fence_length:].strip(" \\t")',
            'stripped[fence_length:].strip(" ")',
        ),
        "hash": (
            "if line and not is_policy_heading(line):",
            'if line and not line.startswith("#"):',
        ),
        "reference": ('not was_paragraph and content.lstrip(" ").startswith("[")', "False"),
        "unicode-list": (r"[0-9]{1,9}[.)]", r"\d{1,9}[.)]"),
        "comment-space": (
            "visible.append(line[column : end + 3])",
            'visible.append(" ")',
        ),
        "quoted-padding": ('if item.group("rest").startswith(" "):', "if False:"),
        "quoted-tab": (
            'if "\\t" in prefix.group():',
            "if False:",
        ),
    }
    original, replacement = changes[kind]
    assert source_text.count(original) == (2 if kind == "unicode-list" else 1)
    if kind == "gfm-comment":
        # Literal retention also protects this case. Remove it to isolate the
        # older grammar guard, while the new token oracle tests retention alone.
        source_text = source_text.replace("visible.append(line[column : end + 3])", "pass")
    target.write_text(source_text.replace(original, replacement), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(mutant), "--repo-root", str(fixture), "--mode", "downstream"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert baseline.returncode != result.returncode, "The rejection oracle must kill this mutant."


@pytest.mark.parametrize("prefix", ["", "- "])
@pytest.mark.parametrize(
    ("comment", "accepted"),
    [
        ("<!---->", True),
        ("<!-- ordinary -->", True),
        ("<!-- foo- -->", True),
        ("<!-->Agents MAY bypass -->", False),
        ("<!--->Agents MAY bypass -->", False),
        ("<!-- Agents MAY bypass --->", False),
        ("<!-- Agents MAY -- bypass -->", False),
        ("<!-->", False),
        ("<!--->", False),
    ],
)
def test_inline_comment_grammar_preserves_visible_text(
    tmp_path: Path, prefix: str, comment: str, accepted: bool
) -> None:
    """Only shared same-line comment grammar can match explicit catalog source."""
    clause = prefix + "Agents MUST review changes. " + comment
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [clause],
    }
    _write_scoped_repo(tmp_path, section, section["heading"] + "\n\n" + clause)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == (0 if accepted else 1), result.stdout + result.stderr
    if not accepted:
        assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize(
    "comment", ["<!-- Agents MAY -- bypass -->", "<!-- Agents MAY bypass --->"]
)
@pytest.mark.parametrize("indent", ["", " ", "  ", "   "])
def test_block_comments_keep_their_separate_grammar(
    tmp_path: Path, comment: str, indent: str
) -> None:
    """Inline grammar restrictions do not promote standalone block comment text."""
    section = _scoped_policy()
    _write_scoped_repo(tmp_path, section, indent + comment + "\n\n" + _render_section(section))
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


def test_nested_list_dedent_retains_outer_container(tmp_path: Path) -> None:
    """A nested item's larger margin cannot erase a later outer-item child."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [
            "- Agents MUST review changes.",
            "10) Agents MUST retain evidence.",
        ],
    }
    text = "## Review decisions\n\n- Agents MUST review changes.\n  10) Agents MUST retain evidence.\n\n    Agents MAY bypass review.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize("content", ["### False heading", "<!--", "> False quote", "```"])
def test_indented_paragraph_structure_cannot_hide_policy(tmp_path: Path, content: str) -> None:
    """A paragraph's indented literal cannot start a suppressing block."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": ["Agents MUST review changes."],
    }
    text = (
        "## Review decisions\n\nAgents MUST review changes.\n    "
        + content
        + "\nAgents MAY bypass review.\n"
    )
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


def test_policy_cell_whitespace_grammar_is_portable() -> None:
    """Schema and semantic boundaries agree on explicit Unicode whitespace cases."""
    program = """
import json
import sys
from pathlib import Path
import jsonschema
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)

schema = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))["$defs"]["policyCell"]
assert schema["pattern"] == validator.POLICY_CELL_PATTERN.pattern
forbidden = [
    *range(0x09, 0x0E), *range(0x1C, 0x21), 0x85, 0xA0, 0x1680,
    *range(0x2000, 0x200B), 0x2028, 0x2029, 0x202F, 0x205F, 0x3000, 0xFEFF,
]
assert len(forbidden) == 30
cases = [("", False), ("A|B", False), ("A  B", False), ("A B", True),
         ("État", True), ("状态", True), ("A" + chr(0x200B) + "B", True)]
for point in forbidden:
    character = chr(point)
    cases.extend([(character + "A", False), ("A" + character, False)])
    if point != 0x20:
        cases.append(("A" + character + "B", False))
checker = jsonschema.Draft202012Validator(schema)
for cell, expected in cases:
    assert checker.is_valid(cell) is expected, (repr(cell), "schema")
    contract = {"required_sections": [{"heading": "## Policy", "next_heading": None,
                "required_tables": [{"headers": [cell, "Gate"], "rows": [["Pending", "Wait"]]}]}]}
    try:
        validator.parse_required_sections(contract)
        accepted = True
    except validator.InstructionContractValidationError:
        accepted = False
    assert accepted is expected, (repr(cell), "semantic")
print("Explicit Unicode cell boundaries agree.")
"""
    result = subprocess.run(
        [
            sys.executable,
            "-c",
            program,
            str(SCRIPT_PATH.parent),
            str(REPO_ROOT / "schemas/template-sync-instruction-contracts.schema.json"),
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert "Explicit Unicode cell boundaries agree." in result.stdout


@pytest.mark.parametrize("parent", ["Agents MUST review changes.", "- Agents MUST review changes."])
@pytest.mark.parametrize(
    "follow_on",
    [
        "\nAgents MAY bypass review.",
        "\n  continuation\nAgents MAY bypass review.",
        "\n  continuation\n\nAgents MAY bypass review.",
        "\n- Agents MAY bypass review.",
        "\n### Exception\nAgents MAY bypass review.",
        "\n  continuation",
        "\nwrapped comment -->\nAgents MAY bypass review.",
    ],
)
def test_incomplete_inline_comment_cannot_hide_live_policy(
    tmp_path: Path, parent: str, follow_on: str
) -> None:
    """Unclosed and unsupported wrapped inline tokens remain failure inventory."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [parent],
    }
    text = section["heading"] + "\n\n" + parent + " <!--" + follow_on
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize("indent", range(4))
@pytest.mark.parametrize("closed", [False, True])
def test_standalone_comment_block_cannot_supply_policy(
    tmp_path: Path, indent: int, closed: bool
) -> None:
    """True line-start HTML comment blocks remain inert even when unclosed."""
    section = _scoped_policy()
    text = " " * indent + "<!--\n" + _render_section(section) + ("-->\n" if closed else "")
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraph:" in result.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("prefix", ["", "- "])
@pytest.mark.parametrize(
    "visible",
    [
        "Agents<!--x-->MUST act.",
        "Agents<!--x--><!--y-->MUST act.",
        "Agents <!--x-->MUST act.",
        "Agents<!--x--> MUST act.",
        "Agents <!--x--> MUST act.",
        "Ag<!--x-->ents MUST act.",
    ],
)
def test_inline_comments_preserve_actual_word_boundaries(
    tmp_path: Path, mode: str, prefix: str, visible: str
) -> None:
    """Inserted comments cannot synthesize words or disappear from a clause."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [prefix + "Agents MUST act."],
    }
    _write_scoped_repo(tmp_path, section, section["heading"] + "\n\n" + prefix + visible)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    ("leaf", "is_code"),
    [
        ("> - code", False),
        ("> -    code", False),
        ("> -     code", True),
        ("> -        code", True),
        ("> 1.     code", True),
        ("> 10)     code", True),
        ("> > +     code", True),
        ("> -\tcode", None),
        ("> - \tcode", None),
        ("> 1. \tcode", None),
        ("> > 1. \tcode", None),
        ("> > - \tcode", None),
        (">\t- code", None),
        ("> \t> - code", None),
        ("> > > 10)\tcode", None),
        (" > -\tcode", None),
        ("  > 1.\tcode", None),
        ("   > > + \tcode", None),
        ("> plain\ttext", False),
        ("> - plain\ttext", False),
        ("> 1. plain\ttext", False),
        ("> prior\n> -     code", True),
    ],
)
def test_quoted_list_padding_preserves_live_and_lazy_policy(
    tmp_path: Path, mode: str, leaf: str, is_code: bool | None
) -> None:
    """Known paragraphs can continue lazily; ambiguous structural tabs fail closed."""
    clause = "Agents MUST act."
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [clause],
    }
    for label, text, expected in (
        (
            "required",
            section["heading"] + "\n\n" + leaf + "\n" + clause,
            0 if is_code is True else 1,
        ),
        (
            "extra",
            _render_section(section) + "\n" + leaf + "\nAgents MAY bypass.",
            0 if is_code is False else 1,
        ),
    ):
        root = tmp_path / label
        _write_scoped_repo(root, section, text)
        result = _run_validator(root, "--mode", mode)
        assert result.returncode == expected, result.stdout + result.stderr
        if expected:
            assert "section:## Review decisions:paragraph" in result.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("change", ["none", "omit", "reorder", "alter", "duplicate"])
def test_distinct_same_header_tables_keep_complete_ordered_identity(
    tmp_path: Path, mode: str, change: str
) -> None:
    """Repeated column names across tables do not weaken full inventory checks."""
    section = _scoped_policy()
    second = copy.deepcopy(section["required_tables"][0])
    second["rows"] = [["Owner decision", "Keep PR open", "Not merged"]]
    section["required_tables"].append(second)
    document = copy.deepcopy(section)
    if change == "omit":
        document["required_tables"].pop()
    elif change == "reorder":
        document["required_tables"].reverse()
    elif change == "alter":
        document["required_tables"][1]["rows"][0][-1] = "Merged"
    elif change == "duplicate":
        document["required_tables"].append(copy.deepcopy(second))
    _write_scoped_repo(tmp_path, section, _render_section(document))
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == (0 if change == "none" else 1), result.stdout + result.stderr
    if change != "none":
        assert "section:## Review decisions:tables:" in result.stdout


def test_semantic_loader_rejects_complete_duplicate_tables() -> None:
    """Direct semantic loading retains the duplicate guard beyond schema checks."""
    program = """
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)
table = {"headers": ["State", "Action"], "rows": [["Pending", "Wait"]]}
raw = {"required_sections": [{"heading": "## Rules", "next_heading": None,
                             "required_tables": [table, table]}]}
try:
    validator.parse_required_sections(raw)
except validator.InstructionContractValidationError as error:
    assert "Duplicate contract table" in str(error), str(error)
else:
    raise AssertionError("Complete duplicate tables must fail semantic loading.")
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    "kind",
    ["heading", "list", "ordered", "table-delimiter", "block-tail", "wrapped-block-tail"],
)
def test_comment_elision_cannot_create_policy_structure(
    tmp_path: Path, mode: str, kind: str
) -> None:
    """Inline text removal cannot create a source heading, list, or table."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": ["Agents MUST act."],
    }
    if kind == "list":
        section["required_paragraphs"] = ["- Agents MUST act."]
    elif kind == "ordered":
        section["required_paragraphs"] = ["1. Agents MUST act."]
    elif kind == "table-delimiter":
        section["required_tables"] = [
            {"headers": ["State", "Action"], "rows": [["Pending", "Wait"]]}
        ]
    text = _render_section(section)
    if kind == "heading":
        text = text.replace("## Review", "##<!--x--> Review")
    elif kind == "list":
        text = text.replace("- Agents", "-<!--x--> Agents")
    elif kind == "ordered":
        text = text.replace("1. Agents", "1.<!--x--> Agents")
    elif kind == "table-delimiter":
        text = text.replace("| --- |", "| -<!--x-->-- |", 1)
    elif kind == "block-tail":
        text = text.replace("Agents MUST", "<!--x-->Agents MUST")
    else:
        text = text.replace("Agents MUST", "<!--\nx\n-->Agents MUST")
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:" in result.stdout


@pytest.mark.parametrize("kind", ["heading", "list", "table", "block-tail", "wrapped-tail"])
def test_comment_structure_oracle_detects_removed_guard(tmp_path: Path, kind: str) -> None:
    """Structural guards remain independently effective without token retention."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": ["Agents MUST act."],
    }
    if kind == "list":
        section["required_paragraphs"] = ["- Agents MUST act."]
    elif kind == "table":
        section["required_tables"] = [
            {"headers": ["State", "Action"], "rows": [["Pending", "Wait"]]}
        ]
    text = _render_section(section)
    if kind == "heading":
        text = text.replace("## Review", "##<!--x--> Review")
    elif kind == "list":
        text = text.replace("- Agents", "-<!--x--> Agents")
    elif kind == "table":
        text = text.replace("| --- |", "| -<!--x-->-- |", 1)
    elif kind == "block-tail":
        text = text.replace("Agents MUST", "<!--x-->Agents MUST")
    else:
        text = text.replace("Agents MUST", "<!--\nx\n-->Agents MUST")
    root = tmp_path / "fixture"
    _write_scoped_repo(root, section, text)
    baseline = _run_validator(root, "--mode", "downstream")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    assert "section:## Review decisions:" in baseline.stdout

    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source_text = mutant.read_text(encoding="utf-8")
    if kind == "block-tail":
        original = 'if block_comment and end != -1 and line[end + 3 :].strip(" \\t"):'
        replacement = "if False:"
    elif kind == "wrapped-tail":
        original = 'if line[column:].strip(" \\t"):'
        replacement = "if False:"
    else:
        original = "if elided_comment and ("
        replacement = "if False and ("
        source_text = source_text.replace("visible.append(line[column : end + 3])", "pass")
    assert source_text.count(original) == 1
    mutant.write_text(source_text.replace(original, replacement), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(mutant), "--repo-root", str(root), "--mode", "downstream"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def _module_section_repo(tmp_path: Path, text: str, modules: list[str]) -> dict[str, Any]:
    """Write an independent three-section fixture with one optional host protocol."""
    sections = [
        {
            "heading": "## Plugin",
            "next_heading": "## Azure",
            "required_paragraphs": ["Use the plugin."],
        },
        {
            "heading": "## Azure",
            "next_heading": "## Review",
            "requires_modules": ["azure-devops-collaboration"],
            "required_paragraphs": ["Keep Azure authentication secure."],
        },
        {
            "heading": "## Review",
            "next_heading": None,
            "required_paragraphs": ["Review all findings."],
        },
    ]
    contracts = _contracts()
    contracts["instruction_contracts"][0]["required_sections"] = sections
    contracts["protected_guide_section_obligations"] = [
        {
            "key": "optional-azure",
            "path": "CLAUDE.md",
            "target_modules": ["azure-devops-collaboration"],
            "stale_headings": ["## Azure"],
        }
    ]
    _write_common_contract_repo(tmp_path, contracts)
    _write_text(tmp_path, "CLAUDE.md", text)
    _write_yaml(tmp_path, ".template-sync/marker.yml", _marker(modules))
    return contracts


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("retained", [False, True])
@pytest.mark.parametrize("present", [False, True])
def test_module_section_applicability_cli(
    tmp_path: Path, mode: str, retained: bool, present: bool
) -> None:
    """Upstream enforces all protocols; downstream checks modules and stale retention."""
    text = "## Plugin\n\nUse the plugin.\n\n"
    if present:
        text += "## Azure\n\nKeep Azure authentication secure.\n\n"
    text += "## Review\n\nReview all findings.\n"
    modules = ["agent-instructions"] + (["azure-devops-collaboration"] if retained else [])
    _module_section_repo(tmp_path, text, modules)
    result = _run_validator(tmp_path, "--mode", mode)
    expected = (
        (0 if present else 1) if mode == "upstream-template" or retained else (1 if present else 0)
    )
    assert result.returncode == expected, result.stdout + result.stderr
    if mode == "downstream" and not retained and present:
        assert "CLAUDE.md: optional-azure: stale heading: ## Azure" in result.stdout
        _write_yaml(
            tmp_path,
            ".template-sync/marker.yml",
            _marker(
                modules,
                protected_guide_waivers=[
                    {
                        "path": "CLAUDE.md",
                        "contract_key": "optional-azure",
                        "target_module": "azure-devops-collaboration",
                        "reason": "The fixture owner retains the optional protocol.",
                        "authorization_basis": "Explicit fixture authorization for this stale section.",
                    }
                ],
            ),
        )
        waived = _run_validator(tmp_path, "--mode", mode)
        assert waived.returncode == 0, waived.stdout + waived.stderr
        assert "passed with waivers" in waived.stdout
        # The excluded host body is outside applicable section contracts. Its
        # intentional retention is governed by the separate stale-section waiver.
        _write_text(
            tmp_path,
            "CLAUDE.md",
            text.replace(
                "Keep Azure authentication secure.",
                "Local inactive host note.\n\n### Local inactive subsection",
            ),
        )
        inactive = _run_validator(tmp_path, "--mode", mode)
        assert inactive.returncode == 0, inactive.stdout + inactive.stderr


@pytest.mark.parametrize(
    "replacement",
    [
        "### Hidden exception\n\nSkip review.\n\n",
        "## Azure\n\n## Azure\n\n",
        "## Unrelated\n\n",
    ],
)
def test_module_section_traversal_rejects_unexpected_boundaries(
    tmp_path: Path, replacement: str
) -> None:
    """Excluding Azure cannot authorize arbitrary, duplicate, or inserted boundaries."""
    text = "## Plugin\n\nUse the plugin.\n\n" + replacement + "## Review\n\nReview all findings.\n"
    _module_section_repo(tmp_path, text, ["agent-instructions"])
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Plugin:boundary:" in result.stdout


@pytest.mark.parametrize("end", ["contracted", "uncontracted", "eof"])
def test_module_section_traversal_follows_multiple_absent_successors(
    tmp_path: Path, end: str
) -> None:
    """Traversal skips declared excluded sections and stops at the explicit scope end."""
    text = "## Plugin\n\nUse the plugin.\n"
    if end != "eof":
        text += "\n## Review\n\nReview all findings.\n"
    contracts = _module_section_repo(tmp_path, text, ["agent-instructions"])
    sections = contracts["instruction_contracts"][0]["required_sections"]
    sections[1]["next_heading"] = "## Second optional"
    sections.insert(
        2,
        {
            "heading": "## Second optional",
            "next_heading": None if end == "eof" else "## Review",
            "requires_modules": ["schema"],
            "required_paragraphs": ["Keep the second protocol."],
        },
    )
    if end != "contracted":
        sections.pop()
    _write_yaml(tmp_path, ".template-sync/instruction-contracts.yml", contracts)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr
    _write_text(
        tmp_path, "CLAUDE.md", text.replace("Use the plugin.", "Use the plugin.\n\n## Intruder")
    )
    rejected = _run_validator(tmp_path, "--mode", "downstream")
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "section:## Plugin:boundary:" in rejected.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("value", [[], ["schema", "schema"], [3], "schema", ["unknown-module"]])
def test_section_module_catalog_rejects_invalid_requirements(
    tmp_path: Path, mode: str, value: Any
) -> None:
    """Malformed or unknown section modules are catalog errors, not skipped checks."""
    contracts = _module_section_repo(tmp_path, "", ["agent-instructions"])
    contracts["instruction_contracts"][0]["required_sections"][1]["requires_modules"] = value
    _write_yaml(tmp_path, ".template-sync/instruction-contracts.yml", contracts)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "requires_modules" in result.stderr or "unknown manifest module" in result.stderr


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
def test_section_module_catalog_rejects_successor_cycles(tmp_path: Path, mode: str) -> None:
    """An excluded successor cycle fails loading and cannot hang traversal."""
    contracts = _module_section_repo(tmp_path, "", ["agent-instructions"])
    contracts["instruction_contracts"][0]["required_sections"][1]["next_heading"] = "## Plugin"
    _write_yaml(tmp_path, ".template-sync/instruction-contracts.yml", contracts)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Cyclic section successor chain" in result.stderr


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    ("heading", "successor"),
    [
        ("## GitHub Plugin Usage", "## Azure DevOps PR Review Protocol"),
        ("## Azure DevOps PR Review Protocol", "## PR Review Workflow (Codex-adapted)"),
    ],
)
def test_actual_codex_protocol_removal_is_rejected(
    tmp_path: Path, mode: str, heading: str, successor: str
) -> None:
    """Real catalog coverage cannot disappear with a data-driven fixture expectation."""
    catalog = yaml.safe_load(
        (REPO_ROOT / ".template-sync/instruction-contracts.yml").read_text(encoding="utf-8")
    )
    contract = next(
        item for item in catalog["instruction_contracts"] if item["path"] == "AGENTS.md"
    )
    _write_common_contract_repo(tmp_path, {"instruction_contracts": [contract]})
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["agent-instructions", "azure-devops-collaboration"]),
    )
    text = (REPO_ROOT / "AGENTS.md").read_text(encoding="utf-8")
    _write_text(tmp_path, "AGENTS.md", text)
    intact = _run_validator(tmp_path, "--mode", mode)
    assert intact.returncode == 0, intact.stdout + intact.stderr
    start = text.index(heading + "\n")
    end = text.index(successor + "\n", start)
    _write_text(tmp_path, "AGENTS.md", text[:start] + text[end:])
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert f"section:{heading}" in result.stdout


@pytest.mark.parametrize("mutation", ["skip-applicable", "disable-traversal"])
def test_section_module_security_mutants_are_detected(tmp_path: Path, mutation: str) -> None:
    """Independent native expectations kill disabled applicability and traversal checks."""
    fixture = tmp_path / "fixture"
    if mutation == "skip-applicable":
        section = {
            "heading": "## Azure",
            "next_heading": None,
            "requires_modules": ["azure-devops-collaboration"],
            "required_paragraphs": ["Keep Azure authentication secure."],
        }
        _write_scoped_repo(fixture, section, "")
        predicate = "if not section_applies(section, included_modules):"
        replacement = "if section.requires_modules:"
        expected, mutant_expected = 1, 0
    else:
        _module_section_repo(
            fixture,
            "## Plugin\n\nUse the plugin.\n\n## Review\n\nReview all findings.\n",
            ["agent-instructions"],
        )
        predicate = "while successor in sections and successor not in live_lines:"
        replacement = "while False:"
        expected, mutant_expected = 0, 1
    baseline = _run_validator(fixture, "--mode", "downstream")
    assert baseline.returncode == expected, baseline.stdout + baseline.stderr
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source_text = mutant.read_text(encoding="utf-8")
    assert source_text.count(predicate) == 1
    mutant.write_text(source_text.replace(predicate, replacement), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(mutant), "--repo-root", str(fixture), "--mode", "downstream"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == mutant_expected, result.stdout + result.stderr
    assert result.returncode != baseline.returncode


def test_module_section_boundary_waiver_tracks_effective_successor(tmp_path: Path) -> None:
    """A waiver for excluded Azure cannot authorize a missing required Azure boundary."""
    text = "## Plugin\n\nUse the plugin.\n\n## Intruder\n\n## Review\n\nReview all findings.\n"
    _module_section_repo(tmp_path, text, ["agent-instructions"])
    anchors: list[str] = []
    for modules in (
        ["agent-instructions"],
        ["agent-instructions", "schema"],
        ["agent-instructions", "azure-devops-collaboration"],
    ):
        _write_yaml(tmp_path, ".template-sync/marker.yml", _marker(modules))
        result = _run_validator(tmp_path, "--mode", "downstream")
        assert result.returncode == 1, result.stdout + result.stderr
        match = re.search(r"section:## Plugin:boundary:[0-9a-f]{64}", result.stdout)
        assert match is not None, result.stdout
        anchors.append(match.group())
    assert (
        anchors[0] == anchors[1]
    ), "An unrelated module cannot change the same boundary deviation."
    assert (
        anchors[0] != anchors[2]
    ), "Required Azure changes the expected boundary and waiver identity."


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    ("original", "replacement"),
    [
        ("| Review state |", "| Review  state |"),
        ("| Review state |", "| Review\tstate |"),
        ("| Review state |", "| \x1cReview state |"),
        ("Not clean", "Not  clean"),
        ("Not clean", "Not\tclean"),
        ("Not clean", "Not\x1cclean"),
        ("Not clean", "Not\x85clean"),
        ("Not clean", "Not\u00a0clean"),
        ("Not clean", "Not\u2028clean"),
        ("Not clean", "Not\ufeffclean"),
        ("| --- |", "| \x1c---\x1c |"),
        ("| --- |", "| \u00a0---\u00a0 |"),
        ("| --- |", "| --\t- |"),
        ("| --- |", "| -- - |"),
    ],
)
def test_observed_table_cells_reject_noncanonical_whitespace(
    tmp_path: Path, mode: str, original: str, replacement: str
) -> None:
    """Both CLIs retain invalid header, data and delimiter cell spelling."""
    section = _scoped_policy()
    section["required_tables"][0]["headers"][0] = "Review state"
    text = _render_section(section).replace(original, replacement, 1)
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:tables:" in result.stdout


@pytest.mark.parametrize("padding", ["", " ", "   ", "\t", " \t"])
def test_observed_table_ascii_padding_preserves_canonical_cells(
    tmp_path: Path, padding: str
) -> None:
    """Only cell-edge syntax padding changes; Unicode words and alignment stay usable."""
    section = _scoped_policy()
    table = section["required_tables"][0]
    table["headers"][0] = "État 中文"
    text = _render_section(section)
    lines = []
    for line in text.split("\n"):
        if line.startswith("|"):
            cells = [cell.strip(" ") for cell in line[1:-1].split("|")]
            cells = [":---:" if cell == "---" else cell for cell in cells]
            line = "|" + "|".join(padding + cell + padding for cell in cells) + "|"
        lines.append(line)
    _write_scoped_repo(tmp_path, section, "\n".join(lines))
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


def test_observed_whitespace_character_matrix() -> None:
    """Literal oracles cover every explicit cell separator and paragraph character."""
    program = r"""
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)
points = [
    *range(0x09, 0x0E), *range(0x1C, 0x21), 0x85, 0xA0, 0x1680,
    *range(0x2000, 0x200B), 0x2028, 0x2029, 0x202F, 0x205F, 0x3000, 0xFEFF,
]
assert len(points) == 30
for point in points:
    character = chr(point)
    variants = [character + "A B", "A" + character + "B", "A B" + character]
    for location in ("header", "data"):
        for position, value in enumerate(variants):
            lines = ["| A B | Gate |", "| --- | --- |", "| A B | Not clean |"]
            row = 0 if location == "header" else 2
            lines[row] = lines[row].replace("A B", value)
            expected = point in (0x09, 0x20) if position != 1 else point == 0x20
            try:
                validator.parse_policy_body(lines)
                accepted = True
            except validator.InstructionContractValidationError:
                accepted = False
            assert accepted is expected, (point, location, position, accepted)
    # Delimiter edges have their own structural check after cell validation.
    lines = ["| A | B |", "|" + character + "---" + character + "| --- |", "| C | D |"]
    try:
        validator.parse_policy_body(lines)
        accepted = True
    except validator.InstructionContractValidationError:
        accepted = False
    assert accepted is (point in (0x09, 0x20)), (point, "delimiter")
section = validator.RequiredSection("## Policy", ("Agents MUST act.",), ())
for point in points:
    if point in (0x09, 0x0A, 0x0D, 0x20):
        continue
    character = chr(point)
    variants = [character + "Agents MUST act.", "Agents" + character + "MUST act.",
                "Agents MUST act." + character]
    for value in variants:
        assert validator.section_failures("## Policy\n\n" + value + "\n", (section,)), point
        literal = validator.RequiredSection("## Policy", (value,), ())
        assert validator.section_failures("## Policy\n\n" + value + "\n", (literal,)) == [], point
print("Explicit observed-cell and literal-paragraph character matrices passed.")
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("character", ["\x1c", "\x85", "\v", "\u00a0"])
@pytest.mark.parametrize("position", ["leading", "internal", "trailing"])
def test_paragraph_whitespace_drift_remains_literal(
    tmp_path: Path, mode: str, character: str, position: str
) -> None:
    """Replacing or adding a non-ASCII wrapping character cannot satisfy a clause."""
    section = _scoped_policy()
    original = section["required_paragraphs"][0]
    value = (
        character + original
        if position == "leading"
        else (
            original + character
            if position == "trailing"
            else original.replace("Agents MUST", "Agents" + character + "MUST")
        )
    )
    _write_scoped_repo(tmp_path, section, _render_section(section).replace(original, value))
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraph:" in result.stdout
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("wrapping", ["  ", "\t", "\n", "\r\n", "\r"])
def test_paragraph_ascii_wrapping_preserves_literal_unicode(
    tmp_path: Path, mode: str, wrapping: str
) -> None:
    """ASCII wrapping folds while distinct Unicode-bearing catalog clauses remain distinct."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": ["Agents MUST act.", "Agents\u00a0MUST act.", "État 中文."],
    }
    text = _render_section(section).replace("Agents MUST", "Agents" + wrapping + "MUST")
    _write_scoped_repo(tmp_path, section, text)
    # Preserve literal CR/LF combinations without host text-mode translation.
    (tmp_path / "CLAUDE.md").write_bytes(text.encode("utf-8"))
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 0, result.stdout + result.stderr
    (tmp_path / "CLAUDE.md").write_bytes(
        text.replace("Agents\u00a0MUST", "Agents MUST").encode("utf-8")
    )
    changed = _run_validator(tmp_path, "--mode", mode)
    assert changed.returncode == 1, changed.stdout + changed.stderr
    assert "section:## Review decisions:paragraph:" in changed.stdout


@pytest.mark.parametrize("kind", ["table", "paragraph", "boundary"])
def test_whitespace_waiver_cannot_authorize_another_deviation(tmp_path: Path, kind: str) -> None:
    """An authorized control-character deviation cannot waive another or canonical content."""
    section = _scoped_policy()
    if kind == "boundary":
        section["next_heading"] = "## End"
    canonical = _render_section(section)
    if kind == "table":
        first = canonical.replace("Not clean", "Not\x1cclean")
    elif kind == "paragraph":
        first = canonical.replace("Agents MUST", "Agents\x1cMUST")
    else:
        first = canonical.replace("## End", "## Extra\n\nLocal\x1ctext.\n\n## End")
    second = first.replace("\x1c", "\x85")
    _write_scoped_repo(tmp_path, section, first)
    initial = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert initial.returncode == 1, initial.stdout + initial.stderr
    anchors = set(
        re.findall(
            r"section:## Review decisions:(?:paragraph|paragraphs|tables|boundary):[0-9a-f]{64}",
            initial.stdout,
        )
    )
    assert anchors
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(
            ["agent-instructions"],
            waivers=[
                {
                    "path": "CLAUDE.md",
                    "anchor": anchor,
                    "reason": "Fixture owner accepts only this exact observed deviation.",
                    "authorization_basis": "Explicit fixture authorization for the captured text.",
                }
                for anchor in sorted(anchors)
            ],
        ),
    )
    same = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert same.returncode == 0, same.stdout + same.stderr
    assert "passed with waivers" in same.stdout
    _write_text(tmp_path, "CLAUDE.md", second)
    changed = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert changed.returncode == 1, changed.stdout + changed.stderr
    _write_text(tmp_path, "CLAUDE.md", canonical)
    restored = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert restored.returncode == 0, restored.stdout + restored.stderr
    assert "Instruction contract waivers applied:" not in restored.stdout


@pytest.mark.parametrize("kind", ["table", "paragraph"])
def test_whitespace_matching_oracle_detects_broad_normalization(tmp_path: Path, kind: str) -> None:
    """Restoring broad folding falsely accepts an independently invalid literal fixture."""
    section = _scoped_policy()
    if kind == "table":
        text = _render_section(section).replace("Not clean", "Not\x1cclean")
        original = 'cell.strip(" \\t")'
        replacement = '" ".join(cell.split())'
    else:
        text = _render_section(section).replace("Agents MUST", "Agents\x1cMUST")
        original = 'return re.sub(r"[ \\t\\r\\n]+", " ", text).strip(" ")'
        replacement = 'return " ".join(text.split())'
    fixture = tmp_path / "fixture"
    _write_scoped_repo(fixture, section, text)
    baseline = _run_validator(fixture, "--mode", "downstream")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source_text = mutant.read_text(encoding="utf-8")
    assert source_text.count(original) == 1
    mutant.write_text(source_text.replace(original, replacement), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(mutant), "--repo-root", str(fixture), "--mode", "downstream"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("kind", ["malformed-table", "boundary"])
def test_whitespace_identity_oracle_detects_normalized_hashes(tmp_path: Path, kind: str) -> None:
    """Distinct observed controls need distinct anchors even after another defect is waived."""
    program = r"""
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)
kind = sys.argv[2]
if kind == "malformed-table":
    section = validator.RequiredSection("## Policy", (), (
        validator.RequiredTable(("State", "Gate"), (("Failed", "Not clean"),)),
    ))
    text = "## Policy\n\n| State | Gate |\n| broken | --- |\n| Failed | Not\x1cclean |\n"
    prefix = "section:## Policy:tables:"
else:
    section = validator.RequiredSection("## Policy", ("Agents MUST act.",), (), "## End")
    text = "## Policy\n\nAgents MUST act.\n\n## Extra\n\nLocal\x1ctext.\n\n## End\n"
    prefix = "section:## Policy:boundary:"
first = next(a for a in validator.section_failures(text, (section,)) if a.startswith(prefix))
second = next(a for a in validator.section_failures(text.replace("\x1c", "\x85"), (section,))
              if a.startswith(prefix))
print("distinct" if first != second else "collision")
"""
    baseline = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent), kind],
        check=False,
        capture_output=True,
        text=True,
    )
    assert baseline.returncode == 0, baseline.stdout + baseline.stderr
    assert baseline.stdout.strip() == "distinct"
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source_text = mutant.read_text(encoding="utf-8")
    if kind == "malformed-table":
        # Fold source before it fans out into both malformed rows and the full
        # body; mutating only one copy leaves the other protection effective.
        original = "            malformed = True\n            tables = []\n"
        replacement = original + '            body = [" ".join(line.split()) for line in body]\n'
    else:
        original = "[normalize_policy_paragraph(line) for line in lines[start:end] if line]"
        replacement = '[" ".join(line.split()) for line in lines[start:end] if line]'
    assert source_text.count(original) == 1
    mutant.write_text(source_text.replace(original, replacement), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, "-c", program, str(mutant_dir), kind],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert result.stdout.strip() == "collision"


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    ("opening", "closing", "blank_terminated"),
    [
        ("<script>", "</script>", False),
        ("   <StYlE\ttype='text/css'>", "</PRE>", False),
        ("<pre", "</style>", False),
        ("<?instruction", "?>", False),
        ("<!DOCTYPE", ">", False),
        ("<![CDATA[", "]]>", False),
        ("<!--", "-->", False),
        ("<div>", "", True),
        ("</DETAILS>", "", True),
        ("<custom data-x='>' disabled>", "", True),
        ("</custom >", "", True),
    ],
)
def test_raw_html_blocks_cannot_supply_contracted_policy(
    tmp_path: Path, mode: str, opening: str, closing: str, blank_terminated: bool
) -> None:
    """Each HTML family hides headings or clauses until its real block boundary."""
    section = _scoped_policy()
    original = _render_section(section)
    cases = {
        "heading": opening + "\n" + original,
        "clause": original.replace(
            section["required_paragraphs"][0],
            opening + "\n" + section["required_paragraphs"][0] + "\n" + closing,
            1,
        ),
    }
    if not blank_terminated:
        cases["blank-does-not-close"] = opening + "\n\n" + original
    for label, text in cases.items():
        root = tmp_path / label
        _write_scoped_repo(root, section, text)
        result = _run_validator(root, "--mode", mode)
        assert result.returncode == 1, result.stdout + result.stderr
        assert "section:## Review decisions:paragraph:" in result.stdout

    # A real terminator restores subsequent policy, including type-1 cross-tag closes.
    root = tmp_path / "closed"
    _write_scoped_repo(root, section, opening + "\nexample\n" + closing + "\n\n" + original)
    result = _run_validator(root, "--mode", mode)
    assert result.returncode == 0, result.stdout + result.stderr

    # Closing the block must not swallow a later unauthorized live addition.
    text = original + "\n" + opening + "\nexample\n" + closing + "\n\nAgents MAY bypass review.\n"
    _write_text(root, "CLAUDE.md", text)
    result = _run_validator(root, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraphs:" in result.stdout


@pytest.mark.parametrize(
    "prefix",
    [
        "Example paragraph\n",
        "Unmatched `\n",
        "> Example paragraph\n",
        "> <ambiguous>\n",
        "- Example paragraph\n",
        "- Example paragraph\n\n",
    ],
)
def test_interrupting_html_after_other_blocks_hides_policy(tmp_path: Path, prefix: str) -> None:
    """Top-level HTML ends prior containers and outranks potential inline spans."""
    section = _scoped_policy()
    text = prefix + "<script>\n`\n\n" + _render_section(section)
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraph:" in result.stdout


@pytest.mark.parametrize(
    "prefix",
    [
        "<script></script> trailing raw HTML\n",
        "<?instruction?> trailing raw HTML\n",
        "<!DOCTYPE html> trailing raw HTML\n",
        "<![CDATA[example]]> trailing raw HTML\n",
        "```html\n<script>\n```\n",
        "    <script>\n",
        "> <script>\n",
        "\\<script>\n",
        "`<script>`\n",
        "<!-- <script> -->\n",
        "<div>\n<script>\n\n",
    ],
)
def test_inert_html_examples_preserve_following_live_policy(tmp_path: Path, prefix: str) -> None:
    """Literal examples and completed blocks cannot open a later phantom HTML block."""
    section = _scoped_policy()
    _write_scoped_repo(tmp_path, section, prefix + "\n" + _render_section(section))
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("tag", ["<custom>", "<custom data-x='a b'>", "</custom>", "<script/>"])
def test_standalone_tag_cannot_interrupt_a_live_paragraph(tmp_path: Path, tag: str) -> None:
    """Only block-interrupting HTML families may hide a paragraph continuation."""
    clause = "Agents MUST act. " + tag + " Agents MAY bypass."
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [clause],
    }
    text = section["heading"] + "\n\nAgents MUST act.\n" + tag + "\nAgents MAY bypass.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr
    section["required_paragraphs"] = ["Agents MUST act."]
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 1, result.stdout + result.stderr


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize("marker", ["-", "+", "*", "1.", "10)", "999999999."])
@pytest.mark.parametrize("padding", ["     ", "        ", "\t\t"])
def test_overpadded_list_content_cannot_supply_policy(
    tmp_path: Path, mode: str, marker: str, padding: str
) -> None:
    """Five-plus visual columns make first item content code, never a live clause."""
    clause = marker + " Agents MUST act."
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [clause],
    }
    text = section["heading"] + "\n\n" + marker + padding + "Agents MUST act.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", mode)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "section:## Review decisions:paragraph:" in result.stdout


@pytest.mark.parametrize(
    ("prefix", "accepted"),
    [
        ("- ", True),
        ("-  ", True),
        ("-   ", True),
        ("-    ", True),
        ("-\t", True),
        ("- \t", True),
        (" -\t", True),
        ("  -\t", True),
        ("   -\t", True),
        ("   - \t", True),
        (" -\t\t", False),
        ("  -\t\t", False),
        ("   -\t ", False),
        ("   -\t\t", False),
    ],
)
def test_list_padding_uses_visual_columns(tmp_path: Path, prefix: str, accepted: bool) -> None:
    """Marker indentation and tabs use independent fixed rendering expectations."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": ["- Agents MUST act."],
    }
    text = section["heading"] + "\n\n" + prefix + "Agents MUST act.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == (0 if accepted else 1), result.stdout + result.stderr


@pytest.mark.parametrize("padding", [" ", "    ", "        ", "\t\t"])
def test_empty_list_items_keep_their_existing_contract(tmp_path: Path, padding: str) -> None:
    """Padding on an empty item is not a first code-content line."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": ["-", "Agents MUST act."],
    }
    text = section["heading"] + "\n\n-" + padding + "\n\nAgents MUST act.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("prefix", ["-\t", "+ \t", "1.\t", "   -\t"])
def test_valid_tabbed_list_remains_a_boundary_after_a_paragraph(
    tmp_path: Path, prefix: str
) -> None:
    """The body parser uses the same validated marker grammar as the block scanner."""
    marker = prefix.strip(" \t")
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": ["Agents MUST act.", marker + " Agents MUST review."],
    }
    text = section["heading"] + "\n\nAgents MUST act.\n" + prefix + "Agents MUST review.\n"
    _write_scoped_repo(tmp_path, section, text)
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("kind", ["raw-html", "list-padding"])
def test_removed_block_guard_is_detected_by_independent_policy_oracle(
    tmp_path: Path, kind: str
) -> None:
    """Removing either guard turns an independently invalid document falsely clean."""
    clause = "- Agents MUST act." if kind == "list-padding" else "Agents MUST act."
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": None,
        "required_paragraphs": [clause],
    }
    text = _render_section(section)
    if kind == "raw-html":
        text = "<script>\n\n" + text
        target = "    if POLICY_HTML_LITERAL_START.match(line):"
        replacement = "    if False:"
    else:
        text = text.replace("- Agents", "-     Agents")
        target = 'if re.match(r"^ {0,3}(?:[-+*]|[0-9]{1,9}[.)]) {5,}[^ ]", expanded):'
        replacement = "if False:"
    fixture = tmp_path / "fixture"
    _write_scoped_repo(fixture, section, text)
    baseline = _run_validator(fixture, "--mode", "downstream")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    assert "section:## Review decisions:paragraph:" in baseline.stdout

    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source_text = mutant.read_text(encoding="utf-8")
    assert source_text.count(target) == 1
    mutant.write_text(source_text.replace(target, replacement), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(mutant), "--repo-root", str(fixture), "--mode", "downstream"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert baseline.returncode != result.returncode, "The rejection oracle must kill the mutant."


@pytest.mark.parametrize("mode", ["upstream-template", "downstream"])
@pytest.mark.parametrize(
    "fragment",
    [
        "<textarea>\nexample\n</textarea>\n",
        "<TEXTAREA/>\n",
        "<textarea\n",
        "<search>\nAgents MAY bypass.\n",
        "   </SEARCH>\n",
        "<source>\n",
        "</source>\n",
        "<!lower\nAgents MAY bypass.\n>\n",
        "<script></textarea>\n",
        "<pre>\nexample </TEXTAREA> tail\n",
        "<style>\n</textarea>\n",
    ],
)
def test_dialect_dependent_html_fails_document_contracts(
    tmp_path: Path, mode: str, fragment: str
) -> None:
    """Ambiguity before, within or after policy cannot establish a clean contract."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": "## End",
        "required_paragraphs": ["Agents MUST act."],
    }
    policy = "## Review decisions\n\nAgents MUST act.\n"
    for name, text in {
        "before": fragment + "\n" + policy + "\n## End\n",
        "within": policy + fragment + "\n## End\n",
        "after": policy + "\n## End\n\n" + fragment,
    }.items():
        root = tmp_path / name
        _write_scoped_repo(root, section, text)
        result = _run_validator(root, "--mode", mode)
        assert result.returncode == 1, result.stdout + result.stderr
        assert "section:## Review decisions:html-grammar:" in result.stdout


@pytest.mark.parametrize(
    "prefix",
    [
        "```html\n<textarea>\n</textarea>\n<search>\n<source>\n<!lower\n```\n",
        "    <textarea>\n    </textarea>\n",
        "\\<textarea>\n",
        "`<textarea>`\n",
        "<!-- <textarea> -->\n",
        "<script>\n<textarea>\n<search>\n<source>\n<!lower\n</script>\n",
        "<div>\n<textarea>\n</textarea>\n\n",
        "<?instruction <textarea> </textarea> ?>\n",
        "<![CDATA[<textarea></textarea>]]>\n",
        "</textarea>\n\n",
        "<textareax>\n\n",
        "<sources>\n\n",
    ],
)
def test_inert_or_shared_html_remains_usable(tmp_path: Path, prefix: str) -> None:
    """Already-inert examples and unambiguous tag names do not create dialect failures."""
    section = _scoped_policy()
    _write_scoped_repo(tmp_path, section, prefix + "\n" + _render_section(section))
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr


def test_dialect_failure_binds_source_outside_the_section(tmp_path: Path) -> None:
    """A waiver cannot hide changes below an ambiguous block or the section end."""
    section: dict[str, Any] = {
        "heading": "## Review decisions",
        "next_heading": "## End",
        "required_paragraphs": ["Agents MUST act."],
    }
    prefix = "<script></textarea>\n\n## Review decisions\n\nAgents MUST act.\n\n## End\n\n"
    anchors: list[str] = []
    for tail in ["First outside text.\n", "Changed outside text.\n"]:
        _write_scoped_repo(tmp_path, section, prefix + tail)
        result = _run_validator(tmp_path, "--mode", "downstream")
        assert result.returncode == 1, result.stdout + result.stderr
        matches = re.findall(
            r"section:## Review decisions:html-grammar:[0-9a-f]{64}", result.stdout
        )
        assert matches, result.stdout + result.stderr
        anchors.append(matches[0])
    assert anchors[0] != anchors[1]


def test_removed_dialect_guard_is_detected_by_independent_oracle(tmp_path: Path) -> None:
    """GFM keeps the entire required section inside the unclosed script block."""
    section = _scoped_policy()
    fixture = tmp_path / "fixture"
    _write_scoped_repo(fixture, section, "<script></textarea>\n\n" + _render_section(section))
    baseline = _run_validator(fixture, "--mode", "downstream")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    assert "section:## Review decisions:html-grammar:" in baseline.stdout

    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source_text = mutant.read_text(encoding="utf-8")
    target = "        if ambiguous_html:"
    assert source_text.count(target) == 1
    mutant.write_text(source_text.replace(target, "        if False:"), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(mutant), "--repo-root", str(fixture), "--mode", "downstream"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert baseline.returncode != result.returncode, "The independent rejection oracle must fail."


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    "mutation", ["encoding-weakened", "encoding-fenced", "live-noncanonical", "schematic-ban"]
)
def test_yaml_encoding_and_live_placeholder_contract_rejects_policy_drift(
    tmp_path: Path, mutation: str
) -> None:
    """New YAML policies require operative clauses, including their permitted schematic scope."""
    _copy_real_instruction_contract_surface(tmp_path)
    path = tmp_path / ".github/instructions/yaml.instructions.md"
    text = path.read_text(encoding="utf-8")
    assert _run_validator(tmp_path, "--mode", "upstream-template").returncode == 0
    if mutation == "encoding-weakened":
        changed = text.replace("YAML files **MUST** use UTF-8", "YAML files **MAY** use UTF-8", 1)
    elif mutation == "encoding-fenced":
        heading = "## Encoding\n\n"
        start = text.index(heading) + len(heading)
        end = text.index("\n\n## Formatting Rules", start)
        changed = text[:start] + "```text\n" + text[start:end] + "\n```" + text[end:]
    elif mutation == "live-noncanonical":
        changed = text.replace(
            "URLs **MUST** use literal `OWNER/REPO`",
            "URLs **MUST** use literal `<owner>/<repo>`",
            1,
        )
    else:
        changed = text.replace("notation **MAY** appear", "notation **MUST NOT** appear", 1)
    assert changed != text
    path.write_text(changed, encoding="utf-8")
    result = _run_validator(tmp_path, "--mode", "upstream-template")
    assert result.returncode == 1, result.stdout + result.stderr
    assert ".github/instructions/yaml.instructions.md" in result.stdout


def _write_claude_import_fixture(repo_root: Path, text: str) -> None:
    """Write a minimal cataloged Claude contract for import-scanner tests."""
    _write_common_contract_repo(repo_root, _contracts(required_headings=["# Fixture"]))
    _write_text(repo_root, "CLAUDE.md", text)


def _initialize_git_repository(repo_root: Path) -> None:
    """Initialize an isolated Git worktree without requiring commit identity."""
    result = subprocess.run(
        ["git", "init", "--quiet", str(repo_root)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def _git_add(repo_root: Path, relative_path: str) -> None:
    """Force one fixture path into the isolated index."""
    result = subprocess.run(
        ["git", "-C", str(repo_root), "add", "-f", "--", relative_path],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def _copy_validator_sources(destination: Path) -> Path:
    """Copy the validator module surface for an independent mutant run."""
    destination.mkdir()
    _copy_mutation_surface(destination)
    return destination / SCRIPT_PATH.name


def _run_validator_with_environment(
    repo_root: Path, environment: dict[str, str]
) -> subprocess.CompletedProcess[str]:
    """Run upstream validation with an explicit inherited environment."""
    return subprocess.run(
        [
            sys.executable,
            str(SCRIPT_PATH),
            "--repo-root",
            str(repo_root),
            "--mode",
            "upstream-template",
        ],
        check=False,
        capture_output=True,
        text=True,
        env=environment,
    )


@pytest.mark.parametrize(
    ("text", "targets"),
    [
        pytest.param(
            "# Fixture\n\n@relative.md @/absolute.md @~/personal.md @../parent.md @LICENSE",
            {"relative.md", "/absolute.md", "~/personal.md", "../parent.md", "LICENSE"},
            id="active-target-forms",
        ),
        pytest.param(
            "# Fixture\n\n>@quoted.md\n\n- > @nested.md",
            {"quoted.md", "nested.md"},
            id="blockquote-and-list",
        ),
        pytest.param(
            "# Fixture\n\nText <!-- @inline.md --> remains prose.",
            {"inline.md"},
            id="inline-comment-is-not-exempt",
        ),
        pytest.param(
            "# Fixture\n\nUnmatched ` delimiter leaves @visible.md live.",
            {"visible.md"},
            id="unmatched-code-run",
        ),
        pytest.param(
            "# Fixture\n\nUse `@inline.md` and ``@double.md``.\n"
            "A `multiline\n@multi.md` span.\n\n"
            "```text\n@fenced.md\n```\n~~~text\n@tilde.md\n~~~\n"
            "> ```text\n> @quoted-fence.md\n> ```\n"
            "- ~~~text\n  @list-fence.md\n  ~~~\n",
            set(),
            id="code-spans-and-fences",
        ),
        pytest.param(
            "# Fixture\n\n<!--\n```\n@hidden.md\n```\n--> @tail.md",
            {"tail.md"},
            id="block-comment-state-and-live-tail",
        ),
        pytest.param(
            "# Fixture\n\n> <!--\n> @hidden.md\n> -->\n\n- <!-- @also-hidden.md -->",
            set(),
            id="block-comments-in-containers",
        ),
        pytest.param(
            "# Fixture\n\n@claude start review loop\n@codex review\n"
            "mail@example.com and (@not-boundary.md)",
            set(),
            id="commands-email-and-nonboundary",
        ),
        pytest.param(
            "# Fixture\n\n@claude-policy.md @codex.md",
            {"claude-policy.md", "codex.md"},
            id="bot-prefixes-are-imports",
        ),
    ],
)
def test_active_claude_import_scanner_literals_and_boundaries(
    tmp_path: Path, text: str, targets: set[str]
) -> None:
    """Only documented literal contexts and exact bot commands are exempt."""
    _write_claude_import_fixture(tmp_path, text)
    result = _run_validator(tmp_path, "--mode", "upstream-template")
    assert result.returncode == (1 if targets else 0), result.stdout + result.stderr
    entries = _section_entries(result.stdout, "Active Claude imports")
    observed = {entry.rsplit("@", 1)[-1] for entry in entries}
    assert observed == targets


@pytest.mark.parametrize(
    ("body", "targets"),
    [
        pytest.param(
            "> > `start\n> > @README`\n",
            set(),
            id="nested-quote-multiline-literal",
        ),
        pytest.param(
            "- > `start\n  > @README`\n",
            set(),
            id="list-quote-multiline-literal",
        ),
        pytest.param(
            "> - `start\n>   @README`\n",
            set(),
            id="quote-list-multiline-literal",
        ),
        pytest.param(
            "`start\n    @README`\n",
            set(),
            id="indented-paragraph-continuation-literal",
        ),
        pytest.param(
            "- - `start\n    @README`\n",
            set(),
            id="nested-list-multiline-literal",
        ),
        pytest.param(
            "> > `start\n> @README\n> > end`\n",
            set(),
            id="partial-lazy-nested-quote-literal",
        ),
        pytest.param(
            "100. `start\n     @README`\n",
            set(),
            id="ordered-list-multiline-literal",
        ),
        pytest.param(
            "- `unmatched\n- @README\n- closing`\n",
            {"README"},
            id="new-list-items-break-unmatched-span",
        ),
        pytest.param(
            "> - `unmatched\n> - @README\n> - closing`\n",
            {"README"},
            id="new-quoted-list-items-break-unmatched-span",
        ),
        pytest.param(
            "`unmatched\n> @README\nclosing`\n",
            {"README"},
            id="new-quote-breaks-unmatched-span",
        ),
        pytest.param(
            "> `unmatched\n@README\n> closing`\n",
            set(),
            id="lazy-quote-continuation-remains-literal",
        ),
        pytest.param(
            "`open\n0. @README\nclose`\n",
            set(),
            id="ordered-zero-cannot-interrupt-root-paragraph",
        ),
        pytest.param(
            "`open\n2. @README\nclose`\n",
            set(),
            id="ordered-two-cannot-interrupt-root-paragraph",
        ),
        pytest.param(
            "`open\n003. @README\nclose`\n",
            set(),
            id="ordered-leading-zero-cannot-interrupt-root-paragraph",
        ),
        pytest.param(
            "`open\n2) @README\nclose`\n",
            set(),
            id="ordered-parenthesis-cannot-interrupt-root-paragraph",
        ),
        pytest.param(
            "> `open\n> 2. @README\n> close`\n",
            set(),
            id="ordered-two-cannot-interrupt-quoted-paragraph",
        ),
        pytest.param(
            "`open\n2. # @README\nclose`\n",
            set(),
            id="heading-after-ordered-two-remains-paragraph-text",
        ),
        pytest.param(
            "`open\n2. > @README\nclose`\n",
            set(),
            id="quote-after-ordered-two-remains-paragraph-text",
        ),
        pytest.param(
            "> `open\n> 2. # @README\n> close`\n",
            set(),
            id="quoted-heading-after-ordered-two-remains-paragraph-text",
        ),
        pytest.param(
            "> `open\n> 2. > @README\n> close`\n",
            set(),
            id="quoted-quote-after-ordered-two-remains-paragraph-text",
        ),
        pytest.param(
            "`open\n1. @README\nclose`\n",
            {"README"},
            id="ordered-one-interrupts-root-paragraph",
        ),
        pytest.param(
            "`open\n- @README\nclose`\n",
            {"README"},
            id="bullet-interrupts-root-paragraph",
        ),
        pytest.param(
            "1. `open\n2. @README\n3. close`\n",
            {"README"},
            id="ordered-sibling-items-remain-boundaries",
        ),
        pytest.param(
            "1. `open\n2) @README\n3. close`\n",
            {"README"},
            id="ordered-delimiter-change-remains-boundary",
        ),
        pytest.param(
            "- `open\n2. @README\nclose`\n",
            {"README"},
            id="bullet-list-exit-before-ordered-two-remains-boundary",
        ),
        pytest.param(
            "1. `open\n   2. @README\n   close`\n",
            set(),
            id="ordered-two-under-active-list-remains-literal",
        ),
        pytest.param(
            "> `open\n2. @README\nclose`\n",
            {"README"},
            id="ordered-two-after-quote-exit-remains-boundary",
        ),
        pytest.param(
            "- > `open\n> @README\nclose`\n",
            {"README"},
            id="quote-after-list-exit-remains-boundary",
        ),
        pytest.param(
            "- > `open\n@README\n> close`\n",
            {"README"},
            id="later-quote-after-list-exit-remains-boundary",
        ),
        pytest.param(
            "- > > `open\n> @README\n> close`\n",
            {"README"},
            id="nested-quote-after-list-exit-remains-boundary",
        ),
        pytest.param(
            "- - > `open\n  > @README\n  > close`\n",
            {"README"},
            id="quote-after-nested-list-exit-remains-boundary",
        ),
        pytest.param(
            "- > `open\n  > @README\n  > close`\n",
            set(),
            id="properly-indented-list-quote-remains-literal",
        ),
        pytest.param(
            "- > `open\n@README\nclose`\n",
            set(),
            id="fully-lazy-list-quote-remains-literal",
        ),
        pytest.param(
            "> - `open\n> @README\n> close`\n",
            set(),
            id="quote-list-lazy-continuation-remains-literal",
        ),
        pytest.param(
            "- > > `open\n  > @README\n  > > close`\n",
            set(),
            id="partial-nested-quote-with-list-indent-remains-literal",
        ),
    ],
)
def test_active_import_code_spans_respect_markdown_paragraph_containers(
    tmp_path: Path, body: str, targets: set[str]
) -> None:
    """Code spans stay within one explicit or lazy container paragraph."""
    _write_claude_import_fixture(tmp_path, "# Fixture\n\n" + body)
    result = _run_validator(tmp_path, "--mode", "upstream-template")
    assert result.returncode == (1 if targets else 0), result.stdout + result.stderr
    entries = _section_entries(result.stdout, "Active Claude imports")
    observed = {entry.rsplit("@", 1)[-1] for entry in entries}
    assert observed == targets


@pytest.mark.parametrize(
    "mutation",
    [
        "paragraph-key",
        "container-continuation",
        "ordered-interruption",
        "ordered-after-list-exit",
        "quote-after-list-exit",
        "split-list-laziness",
    ],
)
def test_active_import_container_span_mutants_are_detected(tmp_path: Path, mutation: str) -> None:
    """Independent outcomes kill paragraph and container-transition regressions."""
    if mutation == "paragraph-key":
        body = "- `unmatched\n- @README\n- closing`\n"
        expected_returncode = 0
        original = "key = (line.paragraph_id, width)"
        replacement = "key = (0, width)"
    elif mutation == "container-continuation":
        body = "> > `start\n> > @README`\n"
        expected_returncode = 1
        original = "same_container = continuation is not None"
        replacement = "same_container = False"
    elif mutation == "ordered-interruption":
        body = "`open\n2. @README\nclose`\n"
        expected_returncode = 1
        original = "if marker[0].isdigit() and int(marker[:-1]) != 1:"
        replacement = "if False:"
    elif mutation == "ordered-after-list-exit":
        body = "- `open\n2. @README\nclose`\n"
        expected_returncode = 0
        original = "if relative_list is not None and not omitted_container_prefix:"
        replacement = "if relative_list is not None:"
    elif mutation == "quote-after-list-exit":
        body = "- > `open\n> @README\nclose`\n"
        expected_returncode = 0
        original = "if consumed and omitted_list_indentation:"
        replacement = "if False and omitted_list_indentation:"
    else:
        body = "- > `open\n@README\nclose`\n"
        expected_returncode = 1
        original = (
            "else:\n"
            "            omitted_container_prefix = True\n"
            "            omitted_list_indentation = True"
        )
        replacement = "else:\n            return None"
    _write_claude_import_fixture(tmp_path, "# Fixture\n\n" + body)
    baseline = _run_validator(tmp_path, "--mode", "upstream-template")
    assert baseline.returncode == 1 - expected_returncode, baseline.stdout + baseline.stderr

    mutant = _copy_validator_sources(tmp_path / f"mutant-{mutation}")
    source = mutant.read_text(encoding="utf-8")
    assert source.count(original) == 1
    mutant.write_text(source.replace(original, replacement), encoding="utf-8")
    result = subprocess.run(
        [
            sys.executable,
            str(mutant),
            "--repo-root",
            str(tmp_path),
            "--mode",
            "upstream-template",
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == expected_returncode, result.stdout + result.stderr


def test_active_import_container_normalization_is_bounded_and_fail_visible() -> None:
    """A pathological alternating prefix cannot hide its import or rescan its tail."""
    program = r"""
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)

text = "- > " * 25_000 + "@README\n"
imports = validator.active_claude_imports("CLAUDE.md", text)
assert [(item.line, item.target) for item in imports] == [(1, "README")]
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent)],
        check=False,
        capture_output=True,
        text=True,
        timeout=5,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_active_import_scan_covers_cataloged_nested_claude_file(tmp_path: Path) -> None:
    """A retained nested exact-basename Claude contract receives the same scan."""
    contracts = {
        "instruction_contracts": [
            {
                "path": "tools/CLAUDE.md",
                "requires_modules": ["agent-instructions"],
                "required_headings": ["# Nested"],
            }
        ]
    }
    _write_common_contract_repo(tmp_path, contracts)
    _write_text(tmp_path, "tools/CLAUDE.md", "# Nested\n\n@outside.md\n")
    result = _run_validator(tmp_path, "--mode", "upstream-template")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "tools/CLAUDE.md:3: @outside.md" in result.stdout


def test_active_import_scan_skips_inapplicable_downstream_contract(tmp_path: Path) -> None:
    """Module exclusion remains an explicit skip rather than reading the file."""
    contracts = {
        "instruction_contracts": [
            {
                "path": "CLAUDE.md",
                "requires_modules": ["azure-devops-collaboration"],
                "required_headings": ["# Fixture"],
            }
        ]
    }
    _write_common_contract_repo(tmp_path, contracts)
    _write_text(tmp_path, "CLAUDE.md", "# Fixture\n\n@outside.md\n")
    _write_yaml(tmp_path, ".template-sync/marker.yml", _marker(["agent-instructions"]))
    result = _run_validator(tmp_path, "--mode", "downstream")
    assert result.returncode == 0, result.stdout + result.stderr
    assert "Contracts skipped by downstream module selection:" in result.stdout
    assert "Active Claude imports:" not in result.stdout


@pytest.mark.parametrize(
    "relative_path",
    ["CLAUDE.local.md", "tools/CLAUDE.LOCAL.md", "more/claude.Local.MD"],
)
def test_tracked_claude_local_memory_rejects_index_paths_without_reading_worktree(
    tmp_path: Path, relative_path: str
) -> None:
    """Root, nested, mixed-case, and index-only local memory all fail."""
    _write_claude_import_fixture(tmp_path, "# Fixture\n")
    _initialize_git_repository(tmp_path)
    _write_text(tmp_path, relative_path, "personal content must not be read")
    _git_add(tmp_path, relative_path)
    (tmp_path / relative_path).unlink()
    result = _run_validator(tmp_path, "--mode", "upstream-template")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Tracked Claude local memory:" in result.stdout
    assert relative_path in result.stdout


def _fsmonitor_fixture(root: Path, *, tracked_memory: bool = False) -> Path:
    """Create a native hook whose only effect is a sentinel inside the isolated Git dir."""
    _initialize_git_repository(root)
    relative_path = "tools/CLAUDE.LOCAL.md" if tracked_memory else "ordinary.txt"
    _write_text(root, relative_path, "fixture\n")
    _git_add(root, relative_path)
    if tracked_memory:
        (root / relative_path).unlink()
    hook = root / ".git" / "test-fsmonitor"
    hook.write_text(
        "#!/bin/sh\nprintf invoked > .git/fsmonitor-invoked\nexit 1\n", encoding="utf-8"
    )
    hook.chmod(0o755)
    subprocess.run(
        ["git", "-C", str(root), "config", "core.fsmonitor", ".git/test-fsmonitor"],
        check=True,
        capture_output=True,
    )
    sentinel = root / ".git" / "fsmonitor-invoked"
    # A positive control proves the platform actually executes this hook fixture.
    subprocess.run(["git", "-C", str(root), "ls-files"], check=True, capture_output=True)
    assert sentinel.read_text(encoding="utf-8") == "invoked"
    sentinel.unlink()
    return sentinel


@pytest.mark.parametrize("tracked_memory", [False, True])
def test_git_inventory_disables_native_fsmonitor(
    tmp_path: Path, monkeypatch: Any, tracked_memory: bool
) -> None:
    """Neither a passing inventory nor an index-only violation may execute local hooks."""
    sentinel = _fsmonitor_fixture(tmp_path, tracked_memory=tracked_memory)
    monkeypatch.syspath_prepend(str(SCRIPT_PATH.parent))
    validator = importlib.import_module("validate_instruction_contracts")
    validator = sys.modules.get("instruction_contract_core", validator)
    version = subprocess.run(["git", "--version"], check=True, capture_output=True, text=True)
    match = re.match(r"git version ([0-9]+)\.([0-9]+)", version.stdout)
    if match is None or tuple(map(int, match.groups())) < (2, 36):
        pytest.skip("The modern Git override requires Git >=2.36")
    inventory, applicable = validator.tracked_claude_local_memory(tmp_path)
    assert applicable
    assert bool(inventory) == tracked_memory
    if tracked_memory:
        assert inventory[0].path == "tools/CLAUDE.LOCAL.md"
    assert not sentinel.exists()


@pytest.mark.parametrize("version", [b"git version 2.35.1", b"unparseable Git"])
@pytest.mark.parametrize("setting", [None, "", "true", "false", ".git/test-fsmonitor"])
def test_git_inventory_legacy_fsmonitor_gate(
    tmp_path: Path, monkeypatch: Any, version: bytes, setting: str | None
) -> None:
    """Unknown/old Git inspects configuration without entering the index when configured."""
    sentinel = _fsmonitor_fixture(tmp_path, tracked_memory=True)
    command = ["git", "-C", str(tmp_path), "config"]
    command += ["--unset-all", "core.fsmonitor"] if setting is None else ["core.fsmonitor", setting]
    subprocess.run(command, check=True, capture_output=True)
    monkeypatch.syspath_prepend(str(SCRIPT_PATH.parent))
    validator = importlib.import_module("validate_instruction_contracts")
    validator = sys.modules.get("instruction_contract_core", validator)
    real_runner = validator.run_bounded_git
    calls: list[list[str]] = []

    def runner(root: Path, args: list[str], **kwargs: Any) -> bytes:
        calls.append(args)
        return version if args == ["--version"] else real_runner(root, args, **kwargs)

    monkeypatch.setitem(
        validator.tracked_claude_local_memory.__globals__, "run_bounded_git", runner
    )
    if setting is None:
        inventory, applicable = validator.tracked_claude_local_memory(tmp_path)
        assert applicable and inventory[0].path == "tools/CLAUDE.LOCAL.md"
    else:
        with pytest.raises(
            validator.InstructionContractValidationError, match="requires Git detected"
        ):
            validator.tracked_claude_local_memory(tmp_path)
        assert not any("ls-files" in call or "rev-parse" in call for call in calls)
    assert not sentinel.exists()


@pytest.mark.parametrize("kind", ["modern-override", "legacy-rejection", "absent-config-exit"])
def test_git_fsmonitor_oracle_detects_removed_guard(tmp_path: Path, kind: str) -> None:
    """Isolated source mutants distinguish callback suppression from safe absence handling."""
    root = tmp_path / "fixture"
    root.mkdir()
    sentinel = _fsmonitor_fixture(root)
    if kind == "absent-config-exit":
        subprocess.run(
            ["git", "-C", str(root), "config", "--unset-all", "core.fsmonitor"],
            check=True,
            capture_output=True,
        )
    mutant_dir = tmp_path / "mutant"
    mutant_dir.mkdir()
    _copy_mutation_surface(mutant_dir)
    mutant = mutant_dir / SCRIPT_PATH.name
    source = mutant.read_text(encoding="utf-8")
    replacements = {
        "modern-override": ('return ["-c", "core.fsmonitor=false"]', "return []"),
        "legacy-rejection": ("if configured:\n", "if False:\n"),
        "absent-config-exit": ("allowed_returncodes=(0, 1)", "allowed_returncodes=(0,)"),
    }
    guard, replacement = replacements[kind]
    assert source.count(guard) == 1
    mutant.write_text(source.replace(guard, replacement), encoding="utf-8")
    program = r"""
import sys
from pathlib import Path
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)
real_runner = validator.run_bounded_git
def runner(root, args, **kwargs):
    if args == ["--version"]:
        return b"git version 2.36.0" if sys.argv[3] == "modern-override" else b"git version 2.35.1"
    return real_runner(root, args, **kwargs)
validator.run_bounded_git = runner
try:
    validator.tracked_claude_local_memory(Path(sys.argv[2]))
except validator.InstructionContractValidationError as error:
    print(error)
    raise SystemExit(1)
"""
    baseline = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent), str(root), kind],
        check=False,
        capture_output=True,
        text=True,
    )
    assert baseline.returncode == int(kind == "legacy-rejection"), baseline.stdout + baseline.stderr
    assert not sentinel.exists()
    result = subprocess.run(
        [sys.executable, "-c", program, str(mutant_dir), str(root), kind],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == int(kind == "absent-config-exit"), result.stdout + result.stderr
    assert sentinel.exists() == (kind != "absent-config-exit")


def test_untracked_and_near_miss_local_memory_remain_valid(tmp_path: Path) -> None:
    """Personal untracked memory and a tracked suffix near miss are not violations."""
    _write_claude_import_fixture(tmp_path, "# Fixture\n")
    _initialize_git_repository(tmp_path)
    _write_text(tmp_path, "CLAUDE.local.md", "untracked personal content")
    _write_text(tmp_path, "CLAUDE.local.md.bak", "tracked near miss")
    _git_add(tmp_path, "CLAUDE.local.md.bak")
    result = _run_validator(tmp_path, "--mode", "upstream-template")
    assert result.returncode == 0, result.stdout + result.stderr
    assert "Tracked Claude local memory:\n" not in result.stdout


def test_non_git_and_outer_repository_roots_report_inventory_not_applicable(
    tmp_path: Path,
) -> None:
    """An absent exact-root marker neither invokes nor borrows Git inventory."""
    outer = tmp_path / "outer"
    fixture = outer / "materialized"
    _write_claude_import_fixture(fixture, "# Fixture\n")
    _initialize_git_repository(outer)
    environment = os.environ.copy()
    environment["PATH"] = ""
    result = _run_validator_with_environment(fixture, environment)
    assert result.returncode == 0, result.stdout + result.stderr
    assert "Tracked Claude local memory: N/A" in result.stdout
    assert "Tracked Claude local memory:\n" not in result.stdout


def test_present_git_marker_requires_working_git_and_valid_metadata(tmp_path: Path) -> None:
    """Missing Git and corrupt exact-root metadata cannot become silent N/A."""
    valid = tmp_path / "valid"
    _write_claude_import_fixture(valid, "# Fixture\n")
    _initialize_git_repository(valid)
    environment = os.environ.copy()
    environment["PATH"] = ""
    missing_git = _run_validator_with_environment(valid, environment)
    assert missing_git.returncode == 1
    assert "Git is required" in missing_git.stderr
    assert "N/A" not in missing_git.stdout

    corrupt = tmp_path / "corrupt"
    _write_claude_import_fixture(corrupt, "# Fixture\n")
    (corrupt / ".git").mkdir()
    broken = _run_validator(corrupt, "--mode", "upstream-template")
    assert broken.returncode == 1
    assert "Git inventory failed" in broken.stderr
    assert "N/A" not in broken.stdout


def test_git_inventory_ignores_repository_config_and_pathspec_environment(
    tmp_path: Path,
) -> None:
    """Caller Git selectors cannot redirect or change the exact index query."""
    _write_claude_import_fixture(tmp_path, "# Fixture\n")
    _initialize_git_repository(tmp_path)
    relative_path = "tools/CLAUDE.LOCAL.md"
    _write_text(tmp_path, relative_path, "personal content must not be read")
    _git_add(tmp_path, relative_path)
    (tmp_path / relative_path).unlink()
    environment = os.environ.copy()
    environment.update(
        {
            "GIT_CONFIG": str(tmp_path / "bogus-config"),
            "GIT_CONFIG_COUNT": "1",
            "GIT_CONFIG_KEY_0": "core.bare",
            "GIT_CONFIG_VALUE_0": "true",
            "GIT_DIR": str(tmp_path / "missing-git-dir"),
            "GIT_INDEX_FILE": str(tmp_path / "missing-index"),
            "GIT_WORK_TREE": str(tmp_path / "missing-work-tree"),
            "GIT_LITERAL_PATHSPECS": "1",
            "GIT_GLOB_PATHSPECS": "1",
            "GIT_NOGLOB_PATHSPECS": "1",
            "GIT_ICASE_PATHSPECS": "1",
        }
    )
    result = _run_validator_with_environment(tmp_path, environment)
    assert result.returncode == 1, result.stdout + result.stderr
    assert relative_path in result.stdout


def test_git_output_limit_is_enforced_during_capture(tmp_path: Path) -> None:
    """The independent low-limit mutant proves Git output cannot grow unchecked."""
    _write_claude_import_fixture(tmp_path, "# Fixture\n")
    _initialize_git_repository(tmp_path)
    mutant = _copy_validator_sources(tmp_path / "mutant-output")
    source = mutant.read_text(encoding="utf-8")
    assert source.count("MAXIMUM_GIT_OUTPUT_BYTES = 1024 * 1024") == 1
    mutant.write_text(
        source.replace("MAXIMUM_GIT_OUTPUT_BYTES = 1024 * 1024", "MAXIMUM_GIT_OUTPUT_BYTES = 1"),
        encoding="utf-8",
    )
    result = subprocess.run(
        [
            sys.executable,
            str(mutant),
            "--repo-root",
            str(tmp_path),
            "--mode",
            "upstream-template",
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 1
    assert "one-mebibyte per-stream limit" in result.stderr


@pytest.mark.parametrize(
    ("kind", "message"),
    [
        ("reader-error", "Git inventory output could not be read"),
        ("held-open", "Git inventory streams did not close"),
    ],
)
def test_git_reader_failures_and_descendant_held_pipes_fail_boundedly(
    kind: str, message: str
) -> None:
    """Reader exceptions and inherited open pipes cannot admit partial output or hang."""
    program = r"""
import io
import sys
import threading
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)

kind = sys.argv[2]
closed = threading.Event()

class BrokenStream:
    def read(self, _size):
        error = OSError(5, "synthetic reader failure", "synthetic-private-reader")
        error.filename2 = "synthetic-private-reader-two"
        raise error
    def close(self):
        pass

class HeldStream:
    def read(self, _size):
        closed.wait()
        return b""
    def close(self):
        closed.set()

class FakeProcess:
    returncode = 0
    def __init__(self):
        self.stdout = BrokenStream() if kind == "reader-error" else HeldStream()
        self.stderr = io.BytesIO()
    def wait(self, timeout=None):
        return 0
    def kill(self):
        closed.set()

validator.run_bounded_git.__globals__["GIT_TIMEOUT_SECONDS"] = 0.05
validator.subprocess.Popen = lambda *args, **kwargs: FakeProcess()
try:
    validator.run_bounded_git(validator.Path.cwd(), ["status"])
except validator.InstructionContractValidationError as error:
    if kind == "reader-error":
        assert isinstance(error.__cause__, OSError)
    print(error)
else:
    raise AssertionError("Synthetic Git stream failure was accepted.")
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent), kind],
        check=False,
        capture_output=True,
        text=True,
        timeout=5,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert message in result.stdout
    assert "synthetic-private-reader" not in result.stdout + result.stderr


def test_real_descendant_held_git_pipe_cannot_defeat_deadline() -> None:
    """A real inherited write handle must not block the validator's caller."""
    program = r"""
import subprocess
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)

real_popen = subprocess.Popen
git_parent = r'''
import subprocess
import sys
subprocess.Popen([sys.executable, "-c", "import time; time.sleep(2)"])
'''

def spawn_git_parent(*args, **kwargs):
    return real_popen([sys.executable, "-c", git_parent], **kwargs)

validator.run_bounded_git.__globals__["GIT_TIMEOUT_SECONDS"] = 0.5
validator.subprocess.Popen = spawn_git_parent
try:
    validator.run_bounded_git(validator.Path.cwd(), ["status"])
except validator.InstructionContractValidationError as error:
    print(error)
else:
    raise AssertionError("A descendant-held Git pipe was accepted.")
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent)],
        check=False,
        capture_output=True,
        text=True,
        timeout=1.5,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert "Git inventory streams did not close" in result.stdout


@pytest.mark.parametrize(
    ("kind", "context"),
    [
        ("spawn", "Git inventory could not start"),
        ("marker", "Cannot inspect the repository Git marker"),
        ("canonical", "Cannot canonicalize the Git top level"),
    ],
)
def test_git_os_error_paths_exclude_exception_filenames(kind: str, context: str) -> None:
    """Spawn, marker, and canonicalization failures retain safe native causes."""
    program = r"""
import stat
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)

kind = sys.argv[2]
private_one = "synthetic-private-directory/CLAUDE.local.md"
private_two = "synthetic-private-directory-two/CLAUDE.local.md"
native_error = PermissionError(13, "Access denied", private_one)
native_error.filename2 = private_two

def raise_native_error(*_args, **_kwargs):
    raise native_error

root = validator.Path.cwd()
if kind == "spawn":
    validator.subprocess.Popen = raise_native_error
    operation = lambda: validator.run_bounded_git(root, ["status"])
elif kind == "marker":
    validator.Path.lstat = raise_native_error
    operation = lambda: validator.tracked_claude_local_memory(root)
else:
    class MarkerStatus:
        st_mode = stat.S_IFDIR
    validator.Path.lstat = lambda _path: MarkerStatus()
    validator.run_bounded_git = lambda _root, args, **_kwargs: (
        b"git version 2.36.0" if args == ["--version"] else str(root).encode("utf-8")
    )
    validator.Path.resolve = raise_native_error
    operation = lambda: validator.tracked_claude_local_memory(root)

try:
    operation()
except validator.InstructionContractValidationError as error:
    diagnostic = str(error)
    assert error.__cause__ is native_error
    assert "PermissionError: Access denied" in diagnostic
    assert private_one not in diagnostic
    assert private_two not in diagnostic
    print(diagnostic)
else:
    raise AssertionError("Synthetic Git OSError was accepted.")
"""
    result = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent), kind],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert context in result.stdout


def test_os_error_diagnostic_fallback_and_raw_string_mutant(tmp_path: Path) -> None:
    """Missing strerror stays useful and a raw exception rendering leaks filenames."""
    privacy_program = r"""
import sys
sys.path.insert(0, sys.argv[1])
import validate_instruction_contracts as validator
validator = sys.modules.get("instruction_contract_core", validator)

private_one = "synthetic-private-directory/CLAUDE.local.md"
private_two = "synthetic-private-directory-two/CLAUDE.local.md"
native_error = PermissionError(13, "Access denied", private_one)
native_error.filename2 = private_two
diagnostic = validator.os_error_diagnostic(native_error)
if private_one in diagnostic or private_two in diagnostic:
    raise SystemExit(9)
assert diagnostic == "PermissionError: Access denied"
"""
    program = (
        privacy_program
        + '\nassert validator.os_error_diagnostic(OSError("opaque")) == "OSError: I/O error"\n'
    )
    baseline = subprocess.run(
        [sys.executable, "-c", program, str(SCRIPT_PATH.parent)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert baseline.returncode == 0, baseline.stdout + baseline.stderr

    mutant = _copy_validator_sources(tmp_path / "mutant-raw-os-error")
    source = mutant.read_text(encoding="utf-8")
    original = "return f\"{type(error).__name__}: {error.strerror or 'I/O error'}\""
    assert source.count(original) == 1
    mutant.write_text(source.replace(original, "return str(error)"), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, "-c", privacy_program, str(mutant.parent)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 9, result.stdout + result.stderr


@pytest.mark.parametrize("mutation", ["scanner-call", "blockquote-prefix"])
def test_active_import_mutants_are_detected(tmp_path: Path, mutation: str) -> None:
    """Independent CLI outcomes kill scan-removal and container-prefix mutants."""
    _write_claude_import_fixture(tmp_path, "# Fixture\n\n>@outside.md\n")
    baseline = _run_validator(tmp_path, "--mode", "upstream-template")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    mutant = _copy_validator_sources(tmp_path / f"mutant-{mutation}")
    source = mutant.read_text(encoding="utf-8")
    if mutation == "scanner-call":
        original = "active_imports.extend(active_claude_imports(contract.path, text))"
        replacement = "pass"
    else:
        original = "normalized, content = claude_container_content(line)"
        replacement = "normalized, content = line, line"
    assert source.count(original) == 1
    mutant.write_text(source.replace(original, replacement), encoding="utf-8")
    result = subprocess.run(
        [
            sys.executable,
            str(mutant),
            "--repo-root",
            str(tmp_path),
            "--mode",
            "upstream-template",
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("mutation", ["failure-gate", "case-sensitive-basename"])
def test_tracked_local_memory_mutants_are_detected(tmp_path: Path, mutation: str) -> None:
    """Independent mixed-case index fixtures kill gate and exact-case mutants."""
    _write_claude_import_fixture(tmp_path, "# Fixture\n")
    _initialize_git_repository(tmp_path)
    relative_path = "tools/CLAUDE.LOCAL.md"
    _write_text(tmp_path, relative_path, "personal content")
    _git_add(tmp_path, relative_path)
    (tmp_path / relative_path).unlink()
    baseline = _run_validator(tmp_path, "--mode", "upstream-template")
    assert baseline.returncode == 1, baseline.stdout + baseline.stderr
    mutant = _copy_validator_sources(tmp_path / f"mutant-{mutation}")
    source = mutant.read_text(encoding="utf-8")
    if mutation == "failure-gate":
        original = "or self.tracked_claude_local_memory"
        replacement = "or False"
    else:
        original = 'parts[-1].casefold() == "claude.local.md"'
        replacement = 'parts[-1] == "CLAUDE.local.md"'
    assert source.count(original) == 1
    mutant.write_text(source.replace(original, replacement), encoding="utf-8")
    result = subprocess.run(
        [
            sys.executable,
            str(mutant),
            "--repo-root",
            str(tmp_path),
            "--mode",
            "upstream-template",
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def _copy_mutation_surface(destination: Path) -> None:
    """Compose adapter and shared engine only for existing independent mutants.

    Mutations still act on the actual production engine text. The standalone
    suite separately exercises the uncomposed deployment and import boundary.
    """
    for directory in (SCRIPT_PATH.parent, CORE_PATH.parent):
        for source in directory.glob("*.py"):
            shutil.copyfile(source, destination / source.name)
    adapter = SCRIPT_PATH.read_text(encoding="utf-8")
    lines = adapter.splitlines(keepends=True)
    for node in reversed(ast.parse(adapter).body):
        if isinstance(node, ast.ImportFrom) and node.module in {
            "instruction_contract_core",
            "__future__",
        }:
            del lines[node.lineno - 1 : node.end_lineno]
    composed = CORE_PATH.read_text(encoding="utf-8") + "\n" + "".join(lines)
    (destination / SCRIPT_PATH.name).write_text(composed, encoding="utf-8")
