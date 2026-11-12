"""Native standalone reference checks, exact exceptions, and independent controls."""

from __future__ import annotations

import hashlib
import subprocess
import sys
from pathlib import Path
from typing import Any

import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest
from tests.test_instruction_profile import (
    ROOT,
    profile,
    run,
    run_migration_schema_control,
    write,
)
from tests.test_instruction_profile import (
    test_standalone_runs_after_sync_owned_files_are_physically_deleted as deploy_without_sync,
)

pytestmark = pytest.mark.upstream_template_only

KINDS = ("prose-reference", "absolute-url", "markdown-relative-link")
TOKENS = {
    "prose-reference": "docs/azure-devops-support.md",
    "absolute-url": "https://example.invalid/azure-guide",
    "markdown-relative-link": "docs/azure-devops-support.md#setup",
}


def reference_fixture(root: Path, kind: str) -> tuple[dict[str, Any], str]:
    """Declare one excluded-module obligation in a minimal deployed runtime."""
    document = profile(root)
    path = root / ".github/instruction-contracts.yml"
    catalog = yaml.safe_load(path.read_text(encoding="utf-8"))
    obligation: dict[str, Any] = {
        "key": "azure-reference",
        "path": "AGENTS.md",
        "reference_kind": kind,
        "target_modules": ["azure-devops-platform"],
        "target_path": "docs/azure-devops-support.md",
    }
    if kind != "markdown-relative-link":
        obligation["tokens"] = [TOKENS[kind]]
    catalog["protected_guide_reference_obligations"] = [obligation]
    write(root, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
    prose = TOKENS[kind]
    if kind == "markdown-relative-link":
        prose = f"[Azure guide]({prose})"
    return document, prose


def append_reference(root: Path, prose: str) -> None:
    """Add fixture prose without changing the required anchor controls."""
    path = root / "AGENTS.md"
    write(root, "AGENTS.md", path.read_text(encoding="utf-8") + f"See {prose}.\n")


def require_reference_failure(root: Path) -> None:
    """Keep the native failure oracle independent of the validator implementation."""
    result = run(root)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Stale protected-guide references" in result.stdout
    assert "azure-reference" in result.stdout


@pytest.mark.parametrize("kind", KINDS)
def test_standalone_rejects_each_declared_reference_kind(tmp_path: Path, kind: str) -> None:
    """A clean positive control passes before excluded target prose is restored."""
    _, prose = reference_fixture(tmp_path, kind)
    clean = run(tmp_path)
    assert clean.returncode == 0, clean.stdout + clean.stderr
    append_reference(tmp_path, prose)
    require_reference_failure(tmp_path)
    assert not (tmp_path / ".template-sync").exists()


@pytest.mark.parametrize("kind", KINDS)
@pytest.mark.parametrize("control", ["fence", "retained-target", "excluded-source"])
def test_standalone_reference_applicability_controls(
    tmp_path: Path, kind: str, control: str
) -> None:
    """Examples and module exclusions retain their existing scoped semantics."""
    document, prose = reference_fixture(tmp_path, kind)
    if control == "fence":
        append_reference(tmp_path, f"\n```markdown\n{prose}\n```\n")
    else:
        append_reference(tmp_path, prose)
    if control == "retained-target":
        document["modules"].append("azure-devops-platform")
        write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
    if control == "excluded-source":
        path = tmp_path / ".github/instruction-contracts.yml"
        catalog = yaml.safe_load(path.read_text(encoding="utf-8"))
        catalog["instruction_contracts"][0]["requires_modules"] = ["agent-codex"]
        write(tmp_path, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
        (tmp_path / "AGENTS.md").unlink()
    result = run(tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("kind", KINDS)
@pytest.mark.parametrize("change", ["none", "content", "path", "key", "kind", "target"])
def test_reference_exception_is_exact(tmp_path: Path, kind: str, change: str) -> None:
    """Only the declared path, reference identity and unchanged content can pass."""
    document, prose = reference_fixture(tmp_path, kind)
    append_reference(tmp_path, prose)
    require_reference_failure(tmp_path)
    digest = hashlib.sha256((tmp_path / "AGENTS.md").read_bytes()).hexdigest()
    declaration = {
        "path": "AGENTS.md",
        "anchor": f"reference:azure-reference:{kind}:{TOKENS[kind]}",
        "content_sha256": digest,
        "reason": "Fixture retains this exact local reference.",
        "authorization_basis": "Explicit fixture decision.",
    }
    if change == "content":
        append_reference(tmp_path, "a changed local policy")
    elif change == "path":
        declaration["path"] = "OTHER.md"
    elif change == "key":
        declaration["anchor"] = declaration["anchor"].replace("azure-reference", "other-key")
    elif change == "kind":
        other = "absolute-url" if kind != "absolute-url" else "prose-reference"
        declaration["anchor"] = declaration["anchor"].replace(kind, other)
    elif change == "target":
        declaration["anchor"] += "-different"
    document["exceptions"] = [declaration]
    write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
    result = run(tmp_path)
    assert result.returncode == (0 if change == "none" else 1), result.stdout + result.stderr
    if change == "none":
        assert "Applied local exception: AGENTS.md: reference:" in result.stdout
        # The exception cannot hide a separate mandatory anchor failure.
        write(tmp_path, "AGENTS.md", f"See {prose}.\n")
        declaration["content_sha256"] = hashlib.sha256(
            (tmp_path / "AGENTS.md").read_bytes()
        ).hexdigest()
        write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
        unrelated = run(tmp_path)
        assert unrelated.returncode == 1, unrelated.stdout + unrelated.stderr
        assert "missing required phrase" in unrelated.stdout
    else:
        assert "Exception does not match a current failure and exact content" in result.stderr


@pytest.mark.parametrize("kind", KINDS)
def test_reference_oracle_detects_removed_adapter_guard(tmp_path: Path, kind: str) -> None:
    """Removing the real reference call defeats the unchanged native failure oracle."""
    _, prose = reference_fixture(tmp_path, kind)
    append_reference(tmp_path, prose)
    require_reference_failure(tmp_path)
    path = tmp_path / ".github/scripts/validate_instruction_profile.py"
    source = path.read_text(encoding="utf-8")
    guard = (
        "        protected_guide_reference_obligations=core.parse_protected_guide_reference_obligations(\n"
        "            catalog, known_modules\n"
        "        ),\n"
    )
    assert source.count(guard) == 1
    write(tmp_path, path.relative_to(tmp_path).as_posix(), source.replace(guard, ""))
    mutant = run(tmp_path)
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    with pytest.raises(AssertionError):
        require_reference_failure(tmp_path)


@pytest.mark.parametrize("change", ["missing", "oversized", "unsafe", "invalid-utf8"])
def test_reference_only_input_fails_closed(tmp_path: Path, change: str) -> None:
    """Reference-only paths receive bounded and contained reads independently of anchors."""
    reference_fixture(tmp_path, "prose-reference")
    path = tmp_path / ".github/instruction-contracts.yml"
    catalog = yaml.safe_load(path.read_text(encoding="utf-8"))
    catalog["protected_guide_reference_obligations"][0]["path"] = "REFERENCE.md"
    if change == "unsafe":
        catalog["protected_guide_reference_obligations"][0]["path"] = "../outside.md"
    write(tmp_path, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
    if change == "oversized":
        write(tmp_path, "REFERENCE.md", "x" * (1024 * 1024 + 1))
    elif change == "invalid-utf8":
        (tmp_path / "REFERENCE.md").write_bytes(b"\xff")
    result = run(tmp_path)
    assert result.returncode == 1, result.stdout + result.stderr
    if change == "missing":
        assert "Required instruction files absent" in result.stdout
        assert "REFERENCE.md" in result.stdout
    else:
        assert "ERROR:" in result.stderr


def test_azure_reference_fails_after_all_sync_files_are_deleted(tmp_path: Path) -> None:
    """The real selected catalog checks stale Azure prose with physical sync removal."""
    manifest = yaml.safe_load((ROOT / ".template-sync/manifest.yml").read_text(encoding="utf-8"))
    all_modules = {item["name"] for item in manifest["template_manifest"]["modules"]}
    selected = {"agent-instructions", "instruction-enforcement", "baseline"}
    deploy_without_sync(tmp_path, all_modules - selected)
    path = tmp_path / ".github/copilot-instructions.md"
    write(
        tmp_path,
        path.relative_to(tmp_path).as_posix(),
        path.read_text(encoding="utf-8")
        + "\nSee docs/azure-devops-support.md for Azure guidance.\n",
    )
    result = run(tmp_path)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "copilot-instructions-azure-devops-support-guide-path" in result.stdout
    assert "Stale protected-guide references" in result.stdout
    assert not (tmp_path / ".template-sync").exists()


MULTILINE_LINK_CASES = (
    '[Guide](docs/azure-devops-support.md\n "Azure guide")',
    '[Guide](docs/azure-devops-support.md "Azure\nguide")',
    "[Azure\nguide](docs/azure-devops-support.md)",
    '[Guide](\n docs/azure-devops-support.md\n "Azure guide"\n)',
    '[Guide](<docs/azure-devops-support.md>\n "Azure guide")',
    '[Guide]:\n docs/azure-devops-support.md\n "Azure guide"\n\n[Guide]',
    "[Azure\nguide]: docs/azure-devops-support.md\n\n[Azure guide]",
    "[Guide](docs/azure-devops-support.md\n 'Azure\nguide')",
    "[Guide](docs/azure-devops-support.md\n (Azure\nguide))",
    '[Guide](docs/azure-devops-support.md\n "Azure \\"guide\\" title")',
    '> [Guide](docs/azure-devops-support.md\n> "Azure guide")',
    '- [Guide](docs/azure-devops-support.md\n  "Azure guide")',
)
NON_LINK_CASES = (
    '```markdown\n[Guide](docs/azure-devops-support.md\n "Azure guide")\n```',
    '[Guide](docs/azure-devops-support.md\n```text\nexample\n```\n "Azure guide")',
    '[Guide](docs/azure-devops-support.md "Azure\n\nguide")',
    "[Azure\n\nguide](docs/azure-devops-support.md)",
    "[Guide](<docs/azure-devops-\nsupport.md>)",
    '[Guide](docs/azure-devops-support.md "unfinished)',
    "[Guide](docs/azure-devops-support.md \"wrong title')",
    '[Guide](docs/azure-devops-support.md\n# Boundary\n "Azure guide")',
    '- [Guide](docs/azure-devops-support.md\n- "Azure guide")',
    '[Guide](docs/azure-devops-support.md\n> "Azure guide")',
)


@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
@pytest.mark.parametrize("body", MULTILINE_LINK_CASES)
def test_multiline_links_preserve_exact_target_and_physical_line(body: str, ending: str) -> None:
    """Fixed expected tuples independently cover the supported multiline grammar."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    text = "```text\nexample\n```\n\n" + body + "\n"
    assert core.markdown_link_targets_from_text(text.replace("\n", ending)) == (
        (5, "docs/azure-devops-support.md"),
    )


@pytest.mark.parametrize("body", NON_LINK_CASES)
def test_multiline_link_boundaries_do_not_synthesize_targets(body: str) -> None:
    """Removed fences, blank lines and distinct blocks cannot create a live link."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(body) == ()


@pytest.mark.parametrize("body", MULTILINE_LINK_CASES[:7])
def test_deployed_multiline_reference_fails_without_sync_and_exact_exception_passes(
    tmp_path: Path, body: str
) -> None:
    """Use the deployed CLI, then preserve existing exact declaration semantics."""
    document, _ = reference_fixture(tmp_path, "markdown-relative-link")
    # The helper adds prose before this body; definitions require a new block.
    append_reference(tmp_path, "\n\n" + body)
    require_reference_failure(tmp_path)
    assert not (tmp_path / ".template-sync").exists()
    declaration = {
        "path": "AGENTS.md",
        "anchor": "reference:azure-reference:markdown-relative-link:docs/azure-devops-support.md",
        "content_sha256": hashlib.sha256((tmp_path / "AGENTS.md").read_bytes()).hexdigest(),
        "reason": "Exact multiline fixture declaration",
        "authorization_basis": "Fixture owner approval",
    }
    document["exceptions"] = [declaration]
    write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
    accepted = run(tmp_path)
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    assert "Applied local exception" in accepted.stdout
    append_reference(tmp_path, "changed content")
    rejected = run(tmp_path)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "Exception does not match" in rejected.stderr


def test_multiline_link_parser_mutations_have_independent_native_oracles(tmp_path: Path) -> None:
    """Restore the previous parser and remove each boundary separately."""
    reference_fixture(tmp_path, "markdown-relative-link")
    body = MULTILINE_LINK_CASES[0]
    append_reference(tmp_path, body)
    require_reference_failure(tmp_path)
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    start = source.index("def markdown_link_targets_from_text(")
    stop = source.index("\ndef normalize_markdown_target(", start)
    old = r"""def markdown_link_targets_from_text(text, *, fence_context=MARKDOWN_FENCE_CONTEXT):
    targets = []
    inline = re.compile(r'(?<!!)\[[^\]\n]+\]\((?P<target><[^>\n]+>|[^)\s\n]+)(?:\s+[^)\n]*)?\)')
    reference = re.compile(r'^ {0,3}\[[^\]\n]+\]:\s+(?P<target><[^>\n]+>|\S+)')
    for number, line in lines_outside_markdown_fences(text, fence_context=fence_context):
        for match in inline.finditer(line):
            targets.append((number, normalize_markdown_target(match.group("target"))))
        match = reference.match(line)
        if match is not None:
            targets.append((number, normalize_markdown_target(match.group("target"))))
    return tuple(targets)

"""
    path.write_text(source[:start] + old + source[stop:], encoding="utf-8", newline="\n")
    missed = run(tmp_path)
    assert missed.returncode == 0, missed.stdout + missed.stderr
    for guard, replacement, fragment in (
        ("line_number != previous_line + 1", "False", NON_LINK_CASES[1]),
        (r'if not content.strip(" \t"):', "if False:", NON_LINK_CASES[2]),
    ):
        boundary_start = source.index("def markdown_link_spans(")
        boundary_stop = source.index("def markdown_delimiter_pairs(", boundary_start)
        boundaries = source[boundary_start:boundary_stop]
        assert boundaries.count(guard) == 1
        write(
            tmp_path,
            "AGENTS.md",
            "Agents MUST validate.\nAgents MUST preserve authority.\n" + fragment,
        )
        path.write_text(source, encoding="utf-8", newline="\n")
        positive = run(tmp_path)
        assert positive.returncode == 0, positive.stdout + positive.stderr
        mutant = (
            source[:boundary_start]
            + boundaries.replace(guard, replacement)
            + source[boundary_stop:]
        )
        path.write_text(mutant, encoding="utf-8", newline="\n")
        synthetic = run(tmp_path)
        assert synthetic.returncode == 1, synthetic.stdout + synthetic.stderr
        assert "Stale protected-guide references" in synthetic.stdout


def test_multiline_reference_migration_keeps_exact_identity_and_noop(tmp_path: Path) -> None:
    """Newly recognized links use the existing exact marker waiver translation."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    for root in (stage, target):
        document, _ = reference_fixture(root, "markdown-relative-link")
        append_reference(root, MULTILINE_LINK_CASES[0])
    waiver = {
        "path": "AGENTS.md",
        "contract_key": "azure-reference",
        "target_path": "docs/azure-devops-support.md",
        "reason": "Exact retained link",
        "authorization_basis": "Fixture owner approval",
    }
    marker = {
        "template_sync": {
            "included_modules": document["modules"],
            "protected_guide_contract_waivers": [waiver],
        }
    }
    migrated = run_migration_schema_control(stage, target, marker_document=marker)
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    accepted = run(stage)
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    profile_path = stage / ".github/instruction-profile.yml"
    before = profile_path.read_bytes()
    generated = yaml.safe_load(before)
    assert generated["exceptions"][0]["anchor"] == (
        "reference:azure-reference:markdown-relative-link:docs/azure-devops-support.md"
    )
    write(target, ".github/instruction-profile.yml", before.decode())
    repeated = run_migration_schema_control(stage, target, marker_document=marker)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert profile_path.read_bytes() == before


def test_multiline_link_parser_near_limit_malformed_input_finishes(tmp_path: Path) -> None:
    """Repeated unmatched delimiters must not rescan a near-limit suffix quadratically."""
    for name, text in {
        "labels": "[" * 1_040_000,
        "destinations": "[x](<" * 170_000 + "tail",
        "parentheses": "[x](" * 200_000 + "tail",
        "titles": '[x](target "' + "x" * 1_039_000,
        "reference-prefixes": "prefix " + "[x]:a " * 170_000,
    }.items():
        candidate = tmp_path / (name + ".md")
        candidate.write_text(text, encoding="utf-8")
        program = (
            "import sys; from pathlib import Path; "
            f"sys.path.insert(0, {str(ROOT / '.github/scripts')!r}); "
            "import instruction_contract_core as c; "
            "assert c.markdown_link_targets_from_text(Path(sys.argv[1]).read_text()) == ()"
        )
        result = subprocess.run(
            [sys.executable, "-c", program, str(candidate)],
            capture_output=True,
            text=True,
            check=False,
            timeout=15,
        )
        assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("[x]([x]([x](target)))", ((1, "[x]([x](target))"),)),
        ('[x](first "[nested](not-a-target)") [y](second)', ((1, "first"), (1, "second"))),
        ('[x]: first "[nested](not-a-target)"\n\n[y](second)', ((1, "first"), (3, "second"))),
        (
            '[x](<docs/azure%2Ddevops-support.md#section>\n "title")',
            ((1, "docs/azure%2Ddevops-support.md#section"),),
        ),
        (r"[x](docs/a\(b\).md)", ((1, r"docs/a\(b\).md"),)),
        ("[x](a(b(c)).md) [y](other.md)", ((1, "a(b(c)).md"), (1, "other.md"))),
    ],
)
def test_multiline_link_destinations_have_independent_fixed_oracles(
    text: str, expected: tuple[tuple[int, str], ...]
) -> None:
    """Accepted destinations and titles cannot be reinterpreted as nested links."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(text) == expected


def restore_raw_link_candidate_scanner(source: str) -> str:
    """Restore pre-code-context bracket enumeration in an isolated mutant only."""
    candidate_start = source.index("def markdown_link_candidates(")
    candidate_stop = source.index("\ndef markdown_span_link_targets(", candidate_start)
    raw_candidates = """def markdown_link_candidates(text, starts, whitespace, reference_labels=None, definition_starts=None):
    labels, closes = markdown_delimiter_pairs(text)
    candidates = []
    for opening, closing in labels:
        component = markdown_link_component(
            text, opening, closing, closes, whitespace,
            starts[bisect_right(starts, opening) - 1])
        if component is not None:
            candidates.append((opening, *component))
    return candidates

"""
    return source[:candidate_start] + raw_candidates + source[candidate_stop:]


def test_multiline_link_output_is_bounded_and_guard_mutant_is_caught(tmp_path: Path) -> None:
    """Near-limit accepted nesting emits one target; removing consumption breaks it."""
    reference_fixture(tmp_path, "markdown-relative-link")
    script_root = tmp_path / ".github/scripts"
    program = (
        "import sys; "
        f"sys.path.insert(0, {str(script_root)!r}); "
        "import instruction_contract_core as c; "
        "n=int(sys.argv[1]); text='[x]('*n+'target'+')'*n; "
        "targets=c.markdown_link_targets_from_text(text); "
        "assert targets == ((1, text[4:-1]),), (len(targets), len(text)); "
        "assert sum(len(target) for _,target in targets) <= len(text)"
    )
    result = subprocess.run(
        [sys.executable, "-c", program, "200000"],
        capture_output=True,
        text=True,
        check=False,
        timeout=15,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    path = script_root / "instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    guard = "        if opening < consumed_until:\n            continue\n"
    assert source.count(guard) == 1
    # The forward scan independently excludes accepted destinations. Restore
    # raw enumeration before removing the final guard to reproduce old overlap.
    restored = restore_raw_link_candidate_scanner(source)
    path.write_text(restored.replace(guard, ""), encoding="utf-8", newline="\n")
    mutant = subprocess.run(
        [sys.executable, "-c", program, "20"],
        capture_output=True,
        text=True,
        check=False,
        timeout=15,
    )
    assert mutant.returncode == 1, mutant.stdout + mutant.stderr
    assert "AssertionError: (20, 106)" in mutant.stderr


@pytest.mark.parametrize(
    "prefix",
    ["", " ", "  ", "   ", "prefix ", " " * 100000],
    ids=["zero", "one", "two", "three", "inline", "long"],
)
def test_multiline_reference_definition_prefix_remains_bounded(prefix: str) -> None:
    """Only zero through three literal spaces qualify as definition indentation."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    expected = ((1, "target.md"),) if len(prefix) <= 3 else ()
    assert core.markdown_link_targets_from_text(prefix + "[x]:\n target.md") == expected


BANG_PARITY_CASES = (
    (0, False),
    (1, True),
    (2, False),
    (3, True),
    (4, False),
    (5, True),
    (6, False),
)


CODE_SPAN_LINK_CASES = (
    ("[a `]`](target.md)", ((1, "target.md"),)),
    ("[a `[`](target.md)", ((1, "target.md"),)),
    ("[a `[]`](target.md)", ((1, "target.md"),)),
    ("[a ``]``](target.md)", ((1, "target.md"),)),
    ("[a ``]`[``](target.md)", ((1, "target.md"),)),
    (r"[a `\]`](target.md)", ((1, "target.md"),)),
    (r"[a `]\`](target.md)", ((1, "target.md"),)),
    (r"[a \``]`](target.md)", ((1, "target.md"),)),
    ("[a `](target.md)", ((1, "target.md"),)),
    ("[a ``]`](target.md)", ()),
    ("[a `]\nx`](target.md)", ((1, "target.md"),)),
    ("[a\n `[`](target.md)", ((1, "target.md"),)),
    ("[a `) [`](target.md)", ((1, "target.md"),)),
    ("[a `[z](hidden.md)`](target.md)", ((1, "target.md"),)),
    ("`[a](hidden.md)` [b](target.md)", ((1, "target.md"),)),
    ("``[a](hidden.md)``", ()),
    ("[a `](hidden.md) and `tail", ()),
    (r"[a \`]`](hidden.md)", ()),
    ("![a `]`](hidden.md)", ()),
    (r"\![a `]`](target.md)", ((1, "target.md"),)),
    (r"\\![a `]`](hidden.md)", ()),
    ("[a `]`](first.md) [b](second.md)", ((1, "first.md"), (1, "second.md"))),
    ("[a](a`b.md) [b](second.md) and `tail", ((1, "a`b.md"), (1, "second.md"))),
    ('[a](first.md "`title") [b](second.md) and `tail', ((1, "first.md"), (1, "second.md"))),
    ("[a](a`(b).md) and `z`", ((1, "a`(b).md"),)),
    ("[a `x`]: target.md", ((1, "target.md"),)),
    ("[a `x]: target.md", ((1, "target.md"),)),
    ("[a `]`]: hidden.md", ()),
    ("`\n[a]: hidden.md\n`", ()),
    ("[a `]\n\nx`](hidden.md)", ()),
    ("```markdown\n[a `]`](hidden.md)\n```", ()),
    ("[a `]\n```text\nx\n```\n`](hidden.md)", ()),
)


@pytest.mark.parametrize(("body", "expected"), CODE_SPAN_LINK_CASES)
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
def test_code_span_link_contexts_have_fixed_targets_and_lines(
    body: str, expected: tuple[tuple[int, str], ...], ending: str
) -> None:
    """CommonMark-derived expectations preserve raw targets and outer label lines."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(body.replace("\n", ending)) == expected


@pytest.mark.parametrize("width", [1, 2, 7, 257])
def test_code_span_delimiter_width_is_exact(width: int) -> None:
    """Different-width inner runs cannot terminate the enclosing code span."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    delimiter = "`" * width
    inner = "`" * (width + 1)
    text = f"[a {delimiter}] {inner} [ {delimiter}](target.md)"
    assert core.markdown_link_targets_from_text(text) == ((1, "target.md"),)


@pytest.mark.parametrize(
    ("body", "expected_exit"),
    [
        ("[a `]`](docs/azure-devops-support.md)", 1),
        ("[a\n `[`](docs/azure-devops-support.md)", 1),
        ("[a ``]`[``](docs/azure-devops-support.md)", 1),
        (r"\![a `]`](docs/azure-devops-support.md)", 1),
        ("`[a](docs/azure-devops-support.md)`", 0),
        ("![a `]`](docs/azure-devops-support.md)", 0),
    ],
)
def test_deployed_code_span_link_has_native_outcome_without_sync(
    tmp_path: Path, body: str, expected_exit: int
) -> None:
    """Only live links fail in the physically deployed standalone runtime."""
    reference_fixture(tmp_path, "markdown-relative-link")
    append_reference(tmp_path, body)
    assert not (tmp_path / ".template-sync").exists()
    result = run(tmp_path)
    assert result.returncode == expected_exit, result.stdout + result.stderr
    assert ("Stale protected-guide references" in result.stdout) is (expected_exit == 1)
    if expected_exit == 1:
        assert "AGENTS.md:3:" in result.stdout


def test_code_span_link_exception_and_retained_module_remain_scoped(tmp_path: Path) -> None:
    """Code in a label does not change exact exception or applicability rules."""
    body = "[a `]`](docs/azure-devops-support.md)"
    test_deployed_multiline_reference_fails_without_sync_and_exact_exception_passes(
        tmp_path / "excepted", body
    )
    retained = tmp_path / "retained"
    document, _ = reference_fixture(retained, "markdown-relative-link")
    document["modules"].append("azure-devops-platform")
    write(retained, ".github/instruction-profile.yml", yaml.safe_dump(document))
    append_reference(retained, body)
    result = run(retained)
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("mutation", ["remove-code-context", "restore-raw-scanner"])
def test_code_span_skip_removal_has_independent_native_oracles(
    tmp_path: Path, mutation: str
) -> None:
    """Removing code context restores both a missed live link and a false link."""
    reference_fixture(tmp_path, "markdown-relative-link")
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    guard = "    code_ends = markdown_code_span_ends(text)\n"
    assert source.count(guard) == 1
    for body, expected, mutated in (
        ("[a `]`](docs/azure-devops-support.md)", 1, 0),
        ("`[a](docs/azure-devops-support.md)`", 0, 1),
    ):
        write(
            tmp_path,
            "AGENTS.md",
            "Agents MUST validate.\nAgents MUST preserve authority.\n" + body,
        )
        path.write_text(source, encoding="utf-8", newline="\n")
        normal = run(tmp_path)
        assert normal.returncode == expected, normal.stdout + normal.stderr
        mutated_source = (
            source.replace(guard, "    code_ends = {}\n")
            if mutation == "remove-code-context"
            else restore_raw_link_candidate_scanner(source)
        )
        path.write_text(mutated_source, encoding="utf-8", newline="\n")
        mutant = run(tmp_path)
        assert mutant.returncode == mutated, mutant.stdout + mutant.stderr
        assert ("Stale protected-guide references" in mutant.stdout) is (mutated == 1)


def test_code_span_near_limit_runs_remain_bounded() -> None:
    """Indexed closers bound unmatched runs and skip large code literals once."""
    program = "\n".join(
        [
            "import sys",
            f"sys.path.insert(0, {str(ROOT / '.github/scripts')!r})",
            "import instruction_contract_core as c",
            "cases=[(' '.join('`'*n for n in range(1,1400)),0),",
            "       ('`'+('[x](hidden.md) '*65000)+'` [y](live.md)',1),",
            "       (('[a `]`](live.md) '*55000),55000),",
            "       ('[a '+('`'*500000)+'] '+('`'*500000)+'](live.md)',1)]",
            "for text, count in cases:",
            "    assert len(text.encode()) <= c.MAXIMUM_INPUT_BYTES",
            "    targets=c.markdown_link_targets_from_text(text)",
            "    assert len(targets) == count, (len(targets),count)",
            "    assert sum(len(target) for _,target in targets) <= len(text)",
        ]
    )
    result = subprocess.run(
        [sys.executable, "-c", program], capture_output=True, text=True, check=False, timeout=15
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize(("backslashes", "is_link"), BANG_PARITY_CASES)
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
@pytest.mark.parametrize("title", ["", '\n "Azure guide"'])
def test_escaped_bang_parity_preserves_exact_target_and_line(
    backslashes: int, is_link: bool, ending: str, title: str
) -> None:
    """Fixed CommonMark cases distinguish escaped bangs from actual image markers."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    body = "\\" * backslashes + f"![Guide](docs/azure-devops-support.md{title})"
    text = ("Heading\n\n" + body).replace("\n", ending)
    expected = ((3, "docs/azure-devops-support.md"),) if is_link else ()
    assert core.markdown_link_targets_from_text(text) == expected


@pytest.mark.parametrize(
    ("body", "expected"),
    [
        (r"!\[Guide](target.md)", ()),
        (r"\!\[Guide](target.md)", ()),
        ("! [Guide](target.md)", ((1, "target.md"),)),
        (r"![Image](image.md) \![Guide](target.md)", ((1, "target.md"),)),
        (r"\![Guide](target.md) ![Image](image.md)", ((1, "target.md"),)),
        (r"!\![Guide](target.md)", ((1, "target.md"),)),
        ("```markdown\n\\![Guide](target.md)\n```", ()),
    ],
)
def test_escaped_bang_boundaries_keep_independent_expected_targets(
    body: str, expected: tuple[tuple[int, str], ...]
) -> None:
    """Escapes do not activate brackets, examples, or adjacent image destinations."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(body) == expected


@pytest.mark.parametrize(("backslashes", "is_link"), BANG_PARITY_CASES)
def test_deployed_escaped_bang_obligation_uses_native_parity(
    tmp_path: Path, backslashes: int, is_link: bool
) -> None:
    """Only rendered links fail after deployment without sync-support files."""
    reference_fixture(tmp_path, "markdown-relative-link")
    append_reference(tmp_path, "\\" * backslashes + "![Guide](docs/azure-devops-support.md)")
    assert not (tmp_path / ".template-sync").exists()
    result = run(tmp_path)
    assert result.returncode == (1 if is_link else 0), result.stdout + result.stderr
    assert ("Stale protected-guide references" in result.stdout) is is_link


def test_deployed_escaped_bang_exact_exception_and_retained_target(tmp_path: Path) -> None:
    """A newly recognized link keeps module applicability and exact declarations."""
    test_deployed_multiline_reference_fails_without_sync_and_exact_exception_passes(
        tmp_path / "excepted", '\\![Guide](docs/azure-devops-support.md\n "Azure guide")'
    )
    retained = tmp_path / "retained"
    document, _ = reference_fixture(retained, "markdown-relative-link")
    document["modules"].append("azure-devops-platform")
    write(retained, ".github/instruction-profile.yml", yaml.safe_dump(document))
    append_reference(retained, r"\![Guide](docs/azure-devops-support.md)")
    result = run(retained)
    assert result.returncode == 0, result.stdout + result.stderr


def test_escaped_bang_native_mutants_have_independent_oracles(tmp_path: Path) -> None:
    """Old skipping, no image exclusion, and one-character escapes each fail controls."""
    reference_fixture(tmp_path, "markdown-relative-link")
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    cases = (
        (
            "if markdown_link_opening_is_image(text, opening):",
            'if opening > 0 and text[opening - 1] == "!":',
            r"\![Guide](docs/azure-devops-support.md)",
            1,
            0,
        ),
        (
            "if markdown_link_opening_is_image(text, opening):",
            "if False:",
            "![Image](docs/azure-devops-support.md)",
            0,
            1,
        ),
        (
            "return backslashes % 2 == 0",
            "return backslashes == 0",
            r"\\![Image](docs/azure-devops-support.md)",
            0,
            1,
        ),
    )
    for guard, replacement, body, expected_exit, mutant_exit in cases:
        assert source.count(guard) == 1
        write(
            tmp_path, "AGENTS.md", "Agents MUST validate.\nAgents MUST preserve authority.\n" + body
        )
        path.write_text(source, encoding="utf-8", newline="\n")
        normal = run(tmp_path)
        assert normal.returncode == expected_exit, normal.stdout + normal.stderr
        path.write_text(source.replace(guard, replacement), encoding="utf-8", newline="\n")
        mutant = run(tmp_path)
        assert mutant.returncode == mutant_exit, mutant.stdout + mutant.stderr
        assert ("Stale protected-guide references" in mutant.stdout) is (mutant_exit == 1)


def test_escaped_bang_near_limit_runs_remain_bounded() -> None:
    """Long and repeated parity runs finish without rescanning a growing prefix."""
    program = "\n".join(
        [
            "import sys",
            f"sys.path.insert(0, {str(ROOT / '.github/scripts')!r})",
            "import instruction_contract_core as c",
            "cases=[('\\\\'*1000001+'![x](a)',1),('\\\\'*1000000+'![x](a)',0),",
            "       (('\\\\'*3+'![x](a) ')*80000,80000),",
            "       (('\\\\'*2+'![x](a) ')*80000,0)]",
            "for text, count in cases:",
            "    assert len(text.encode()) <= c.MAXIMUM_INPUT_BYTES",
            "    targets = c.markdown_link_targets_from_text(text)",
            "    assert len(targets) == count",
            "    assert sum(len(target) for _, target in targets) <= len(text)",
        ]
    )
    result = subprocess.run(
        [sys.executable, "-c", program], capture_output=True, text=True, check=False, timeout=15
    )
    assert result.returncode == 0, result.stdout + result.stderr


DESTINATION_DECODING_CASES = (
    (r"docs/a\.md", "docs/a.md"),
    (r"docs\/a\(b\).md", "docs/a(b).md"),
    (r"docs/a\\.md", r"docs/a\.md"),
    (r"docs/a\z.md", r"docs/a\z.md"),
    (r"\ ", r"\ "),
    ("docs/a&period;md", "docs/a.md"),
    ("docs/a&#46;md", "docs/a.md"),
    ("docs/a&#x2E;md", "docs/a.md"),
    ("docs/a&#X2e;md", "docs/a.md"),
    ("docs/a&NotEqualTilde;md", "docs/a\u2242\u0338md"),
    (r"docs/a\&period;md", "docs/a&period;md"),
    ("docs/a&bsol;.md", r"docs/a\.md"),
    ("docs/a&#92;.md", r"docs/a\.md"),
    ("docs/a&amp;period;md", "docs/a&period;md"),
    ("docs/a%2Emd", "docs/a%2Emd"),
    ("docs/a%5C.md", "docs/a%5C.md"),
    ("docs/a%26period;md", "docs/a%26period;md"),
    ("&period", "&period"),
    ("&notAnEntity;", "&notAnEntity;"),
    ("&#00000046;", "&#00000046;"),
    ("&#x000002e;", "&#x000002e;"),
    ("&#46", "&#46"),
    ("&#x;", "&#x;"),
    ("&#0;", "\ufffd"),
    ("&#xD800;", "\ufffd"),
    ("&#x110000;", "\ufffd"),
    ("&#128;", "\u20ac"),
    ("&#11;", "\x0b"),
    ("&#xFFFF;", "\uffff"),
)


@pytest.mark.parametrize(("target", "expected"), DESTINATION_DECODING_CASES)
def test_destination_decoding_has_fixed_single_pass_oracles(target: str, expected: str) -> None:
    """Strict references and literal negatives follow fixed CommonMark expectations."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.decode_markdown_destination(target) == expected


@pytest.mark.parametrize("punctuation", "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~")
def test_destination_decoding_accepts_each_ascii_punctuation(punctuation: str) -> None:
    """Every punctuation escape loses exactly its leading backslash."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.decode_markdown_destination("\\" + punctuation) == punctuation


@pytest.mark.parametrize(
    ("target", "expected"),
    [
        (r"docs/a\.md?view=1#part", "docs/a.md"),
        ("docs/a&period;md&#63;view=1&#35;part", "docs/a.md"),
        ("docs/a%2Emd", "docs/a.md"),
        ("docs/a%252Emd", "docs/a%2Emd"),
        ("docs/a%23part.md", "docs/a#part.md"),
        ("docs/a%5C.md", r"docs/a\.md"),
        (r"docs/a\&period;md", "docs/a&period;md"),
        ("https&colon;//example.invalid/docs/a.md", None),
        (r"https\://example.invalid/docs/a.md", None),
        ("&sol;&sol;example.invalid/docs/a.md", None),
        ("&#35;part", None),
        ("&period;&period;/outside.md", None),
        ("%2e%2e/outside.md", None),
        ("&sol;absolute.md", None),
    ],
)
def test_destination_resolution_preserves_uri_boundaries(target: str, expected: str | None) -> None:
    """Markdown decoding precedes URI parsing and once-only path percent decoding."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.resolve_relative_markdown_target("AGENTS.md", target) == expected


ENCODED_AZURE_DESTINATIONS = (
    r"docs/azure-devops-support\.md",
    "docs/azure-devops-support&period;md",
    "docs/azure-devops-support&#46;md",
    "docs/azure-devops-support&#x2e;md",
)


@pytest.mark.parametrize("target", ENCODED_AZURE_DESTINATIONS)
@pytest.mark.parametrize("form", ["inline", "angle", "multiline", "definition"])
def test_deployed_encoded_destination_fails_with_original_identity(
    tmp_path: Path, target: str, form: str
) -> None:
    """Equivalent live links fail natively while extraction retains source spelling."""
    reference_fixture(tmp_path, "markdown-relative-link")
    clean = run(tmp_path)
    assert clean.returncode == 0, clean.stdout + clean.stderr
    if form == "definition":
        body = f"[Guide]: {target}\n\n[Guide]"
    elif form == "angle":
        body = f"[Guide](<{target}>)"
    elif form == "multiline":
        body = f'[Guide](\n {target}\n "Azure guide")'
    else:
        body = f"[Guide]({target})"
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text("Heading\n\n" + body) == ((3, target),)
    # Keep definition forms outside the helper's introductory paragraph.
    append_reference(tmp_path, "\n\n" + body)
    result = run(tmp_path)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Stale protected-guide references" in result.stdout
    assert target in result.stdout
    assert not (tmp_path / ".template-sync").exists()
    assert not (tmp_path / "pyproject.toml").exists()


@pytest.mark.parametrize(
    "target",
    [
        r"docs/azure-devops-support.\md",
        r"docs/azure-devops-support\\.md",
        r"docs/azure-devops-support\&period;md",
        "docs/azure-devops-support&bsol;.md",
        "docs/azure-devops-support&amp;period;md",
        "docs/azure-devops-support%5C.md",
        "docs/azure-devops-support%26period;md",
        "docs/azure-devops-support&periodmd",
    ],
)
def test_deployed_destination_literal_negatives_remain_distinct(
    tmp_path: Path, target: str
) -> None:
    """A second decode or a non-punctuation escape must not invent an excluded path."""
    reference_fixture(tmp_path, "markdown-relative-link")
    append_reference(tmp_path, f"[Guide]({target})")
    result = run(tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("control", ["fence", "retained-target", "excluded-source"])
def test_encoded_destination_preserves_module_and_fence_scope(tmp_path: Path, control: str) -> None:
    """Decoding changes destination comparison, not obligation applicability."""
    document, _ = reference_fixture(tmp_path, "markdown-relative-link")
    body = r"[Guide](docs/azure-devops-support\.md)"
    if control == "fence":
        body = "\n```markdown\n" + body + "\n```\n"
    append_reference(tmp_path, body)
    if control == "retained-target":
        document["modules"].append("azure-devops-platform")
        write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
    elif control == "excluded-source":
        path = tmp_path / ".github/instruction-contracts.yml"
        catalog = yaml.safe_load(path.read_text(encoding="utf-8"))
        catalog["instruction_contracts"][0]["requires_modules"] = ["agent-codex"]
        write(tmp_path, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
        (tmp_path / "AGENTS.md").unlink()
    result = run(tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("target", ENCODED_AZURE_DESTINATIONS[:2])
def test_encoded_destination_exception_keeps_raw_identity_and_exact_content(
    tmp_path: Path, target: str
) -> None:
    """Semantic path equivalence does not widen an exact source-target exception."""
    document, _ = reference_fixture(tmp_path, "markdown-relative-link")
    append_reference(tmp_path, f"[Guide]({target})")
    require_reference_failure(tmp_path)
    prefix = "reference:azure-reference:markdown-relative-link:"
    declaration = {
        "path": "AGENTS.md",
        "anchor": prefix + "docs/azure-devops-support.md",
        "content_sha256": hashlib.sha256((tmp_path / "AGENTS.md").read_bytes()).hexdigest(),
        "reason": "Exact destination fixture",
        "authorization_basis": "Explicit fixture authorization",
    }
    document["exceptions"] = [declaration]
    write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
    decoded_identity = run(tmp_path)
    assert decoded_identity.returncode == 1, decoded_identity.stdout + decoded_identity.stderr
    assert "Exception does not match" in decoded_identity.stderr
    declaration["anchor"] = prefix + target
    write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
    accepted = run(tmp_path)
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    assert prefix + target in accepted.stdout
    append_reference(tmp_path, "changed content")
    changed = run(tmp_path)
    assert changed.returncode == 1, changed.stdout + changed.stderr
    assert "Exception does not match" in changed.stderr


@pytest.mark.parametrize("target", ENCODED_AZURE_DESTINATIONS[:2])
def test_destination_decoder_removal_is_caught_by_native_oracle(
    tmp_path: Path, target: str
) -> None:
    """Removing one comparison guard restores the independently observed false pass."""
    reference_fixture(tmp_path, "markdown-relative-link")
    append_reference(tmp_path, f"[Guide]({target})")
    require_reference_failure(tmp_path)
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    guard = "    target = decode_markdown_destination(target)\n"
    assert source.count(guard) == 1
    path.write_text(source.replace(guard, ""), encoding="utf-8", newline="\n")
    mutant = run(tmp_path)
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    with pytest.raises(AssertionError):
        require_reference_failure(tmp_path)


@pytest.mark.parametrize("target", ENCODED_AZURE_DESTINATIONS[:2])
@pytest.mark.parametrize("label", ["Guide", "a `]`", "a\n `[` "])
def test_encoded_destination_marker_and_reporter_share_raw_findings(
    tmp_path: Path, target: str, label: str
) -> None:
    """Marker validation and cleanup reporting compare decoded paths consistently."""
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    import report_excluded_module_references as reporter
    import validate_downstream_adoption as adoption
    from template_sync_materialization_helpers import ManifestMapping

    mappings = (
        ManifestMapping("AGENTS.md", frozenset({"agent-instructions"}), frozenset()),
        ManifestMapping(
            "docs/azure-devops-support.md", frozenset({"azure-devops-platform"}), frozenset()
        ),
    )
    write(tmp_path, "AGENTS.md", f"Heading\n\n[{label}]({target})\n")
    failures = adoption.validate_retained_markdown_links(
        tmp_path, ("AGENTS.md",), mappings, {"agent-instructions"}, ()
    )
    assert len(failures) == 1
    assert (failures[0].line_number, failures[0].target, failures[0].target_path) == (
        3,
        target,
        "docs/azure-devops-support.md",
    )
    state = reporter.ReportState(
        source="explicit fixture",
        manifest_modules=frozenset({"agent-instructions", "azure-devops-platform"}),
        mappings=mappings,
        included_modules=frozenset({"agent-instructions"}),
        excluded_modules=frozenset({"azure-devops-platform"}),
        local_overrides=(),
        deferred_candidates=(),
        protected_decisions=(),
        present_files=frozenset({"AGENTS.md"}),
        safe_files=("AGENTS.md",),
    )
    findings = reporter.reference_link_findings(tmp_path, state)
    assert len(findings) == 1
    assert (findings[0].rule_id, findings[0].category, findings[0].line_number) == (
        "markdown-link.excluded-target",
        "protected_file_authorization_needed",
        3,
    )
    assert target in findings[0].detail
    upstream = "https://github.com/franklesniak/copilot-repo-template/blob/HEAD/" + target
    assert reporter.resolve_upstream_blob_target(upstream) == "docs/azure-devops-support.md"
    # contact_links are raw URL fields, so they must not decode Markdown syntax.
    assert reporter.resolve_github_blob_target(upstream) == target


@pytest.mark.parametrize("target", ENCODED_AZURE_DESTINATIONS[:2])
def test_encoded_reference_migration_preserves_raw_exception_and_noop(
    tmp_path: Path, target: str
) -> None:
    """Existing marker waivers translate without normalizing exception identities."""
    stage, previous = tmp_path / "stage", tmp_path / "previous"
    for root in (stage, previous):
        document, _ = reference_fixture(root, "markdown-relative-link")
        append_reference(root, f"[Guide]({target})")
    marker = {
        "template_sync": {
            "included_modules": document["modules"],
            "protected_guide_contract_waivers": [
                {
                    "path": "AGENTS.md",
                    "contract_key": "azure-reference",
                    "target_path": "docs/azure-devops-support.md",
                    "reason": "Exact encoded link",
                    "authorization_basis": "Fixture authorization",
                }
            ],
        }
    }
    migrated = run_migration_schema_control(stage, previous, marker_document=marker)
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    path = stage / ".github/instruction-profile.yml"
    before = path.read_bytes()
    generated = yaml.safe_load(before)
    assert generated["exceptions"][0]["anchor"] == (
        "reference:azure-reference:markdown-relative-link:" + target
    )
    accepted = run(stage)
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    write(previous, ".github/instruction-profile.yml", before.decode())
    repeated = run_migration_schema_control(stage, previous, marker_document=marker)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert path.read_bytes() == before


def test_destination_decoder_near_limit_input_remains_bounded() -> None:
    """Bounded entities and escapes finish without recursively expanding replacements."""
    program = "\n".join(
        [
            "import sys",
            f"sys.path.insert(0, {str(ROOT / '.github/scripts')!r})",
            "import instruction_contract_core as c",
            r"cases=[('a\\.'*300000,'a.'*300000),('&amp;period;'*80000,'&period;'*80000),",
            "       ('&notAnEntity;'*80000,'&notAnEntity;'*80000),",
            "       ('&#12345678;'*90000,'&#12345678;'*90000)]",
            "for text,expected in cases:",
            "    assert len(text.encode()) <= c.MAXIMUM_INPUT_BYTES",
            "    actual=c.decode_markdown_destination(text)",
            "    assert actual == expected",
            "    assert len(actual) <= len(text)",
        ]
    )
    result = subprocess.run(
        [sys.executable, "-c", program], capture_output=True, text=True, check=False, timeout=15
    )
    assert result.returncode == 0, result.stdout + result.stderr


WHITESPACE_DESTINATIONS = (
    ("docs/azure&Tab;-devops-support.md", "docs/azure\t-devops-support.md"),
    ("docs/azure&#10;-devops-support.md", "docs/azure\n-devops-support.md"),
    ("docs/azure&#13;-devops-support.md", "docs/azure\r-devops-support.md"),
    ("&Tab;docs/azure-devops-support.md", "\tdocs/azure-devops-support.md"),
    ("&#32;docs/azure-devops-support.md", " docs/azure-devops-support.md"),
    ("&#11;docs/azure-devops-support.md", "\x0bdocs/azure-devops-support.md"),
    ("&#13;docs/azure-devops-support.md", "\rdocs/azure-devops-support.md"),
    ("docs/azure&#127;-devops-support.md", "docs/azure\x7f-devops-support.md"),
    ("docs/azure%09-devops-support.md", "docs/azure\t-devops-support.md"),
)


@pytest.mark.parametrize(("target", "expected_path"), WHITESPACE_DESTINATIONS)
def test_destination_whitespace_survives_url_component_parsing(
    target: str, expected_path: str
) -> None:
    """URL parser cleanup must not silently change a Markdown destination's path."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    sys.path.insert(0, str(ROOT / ".template-sync/scripts"))
    import instruction_contract_core as core
    import report_excluded_module_references as reporter

    assert core.resolve_relative_markdown_target("AGENTS.md", target) == expected_path
    assert reporter.resolve_relative_markdown_target("AGENTS.md", target) == expected_path
    upstream = "https://github.com/franklesniak/copilot-repo-template/blob/HEAD/" + target
    assert reporter.resolve_upstream_blob_target(upstream) == expected_path


@pytest.mark.parametrize(("target", "expected_path"), WHITESPACE_DESTINATIONS)
def test_deployed_whitespace_destination_has_exact_target_identity(
    tmp_path: Path, target: str, expected_path: str
) -> None:
    """Distinct whitespace paths preserve matching and catalog-schema boundaries."""
    reference_fixture(tmp_path, "markdown-relative-link")
    append_reference(tmp_path, f"[Guide]({target})")
    distinct = run(tmp_path)
    assert distinct.returncode == 0, distinct.stdout + distinct.stderr
    path = tmp_path / ".github/instruction-contracts.yml"
    catalog = yaml.safe_load(path.read_text(encoding="utf-8"))
    catalog["protected_guide_reference_obligations"][0]["target_path"] = expected_path
    write(tmp_path, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
    if any(terminator in expected_path for terminator in "\n\r\u2028\u2029"):
        # The target_path schema rejects the ECMA-262 line terminators (LF, CR,
        # U+2028, U+2029) in both regex dialects before reference matching.
        invalid = run(tmp_path)
        assert invalid.returncode == 1, invalid.stdout + invalid.stderr
        assert "Schema validation failed" in invalid.stderr
    else:
        require_reference_failure(tmp_path)
    assert not (tmp_path / ".template-sync").exists()


@pytest.mark.parametrize(
    "target", ["docs/azure&Tab;-devops-support.md", "&#32;docs/azure-devops-support.md"]
)
def test_destination_whitespace_guard_removal_breaks_native_positive(
    tmp_path: Path, target: str
) -> None:
    """Removing percent protection makes a distinct path trigger the wrong obligation."""
    reference_fixture(tmp_path, "markdown-relative-link")
    append_reference(tmp_path, f"[Guide]({target})")
    correct = run(tmp_path)
    assert correct.returncode == 0, correct.stdout + correct.stderr
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    guard = "return urlsplit(protected_target)"
    assert source.count(guard) == 1
    path.write_text(
        source.replace(guard, "return urlsplit(target)"), encoding="utf-8", newline="\n"
    )
    mutant = run(tmp_path)
    assert mutant.returncode == 1, mutant.stdout + mutant.stderr
    assert "Stale protected-guide references" in mutant.stdout
    with pytest.raises(AssertionError):
        assert (
            mutant.returncode == 0
        ), "A distinct whitespace path must not trigger this obligation."


INDENTED_LINK_CASES = (
    ("    [x](target.md)\n", ()),
    ("\t[x](target.md)\n", ()),
    ("   [x](target.md)\n", ((1, "target.md"),)),
    ("Paragraph\n    [x](target.md)\n", ((2, "target.md"),)),
    ("Paragraph\n\n    [x](target.md)\n", ()),
    ("# Heading\n    [x](target.md)\n", ()),
    ("    [fake](ignored.md)\n[x](target.md)\n", ((2, "target.md"),)),
    ("- Item\n\n    [x](target.md)\n", ((3, "target.md"),)),
    ("- Item\n\n      [x](target.md)\n", ()),
    ("- Item\n      [x](target.md)\n", ((2, "target.md"),)),
    ("10. Item\n\n        [x](target.md)\n", ()),
    ("  10.  Item\n\n       [x](target.md)\n", ((3, "target.md"),)),
    ("  10.  Item\n\n           [x](target.md)\n", ()),
    ("-     [x](target.md)\n", ()),
    (">     [x](target.md)\n", ()),
    ("> Paragraph\n>     [x](target.md)\n", ((2, "target.md"),)),
    ("> - Item\n>\n>       [x](target.md)\n", ()),
    ("> - Item\n>\n>     [x](target.md)\n", ((3, "target.md"),)),
    ("- > Item\n  >\n  >     [x](target.md)\n", ()),
    ("- outer\n  - inner\n\n        [x](target.md)\n", ()),
    ("- outer\n  - inner\n\n      [x](target.md)\n", ((4, "target.md"),)),
    ("- Item\n\n\t[x](target.md)\n", ((3, "target.md"),)),
    ("- Item\n\n\t\t[x](target.md)\n", ()),
    ("- ```text\n  literal\n  ```\n\n    [x](target.md)\n", ((5, "target.md"),)),
    ("- ```text\n  literal\n  ```\n\n      [x](target.md)\n", ()),
    ("    - literal\n\n    [x](target.md)\n", ()),
    ("[x](\n\n    ignored.md\n)\n", ()),
    ("Paragraph\n    [x](<raw\ttarget.md>)\n", ((2, "raw\ttarget.md"),)),
    ("- - -\n    [x](target.md)\n", ()),
    ("- - Item\n\n        [x](target.md)\n", ()),
    ("Paragraph\n2. continues\n    [x](target.md)\n", ((3, "target.md"),)),
)


@pytest.mark.parametrize(("body", "expected"), INDENTED_LINK_CASES)
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
def test_indented_code_context_preserves_live_links_and_original_lines(
    body: str, expected: tuple[tuple[int, str], ...], ending: str
) -> None:
    """Fixed paired contexts distinguish literal indentation from live continuation."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(body.replace("\n", ending)) == expected


@pytest.mark.parametrize(("body", "expected"), INDENTED_LINK_CASES[:23])
def test_deployed_indented_code_has_contextual_native_result_without_sync(
    tmp_path: Path, body: str, expected: tuple[tuple[int, str], ...]
) -> None:
    """The installed checker accepts code and rejects live excluded references."""
    reference_fixture(tmp_path, "markdown-relative-link")
    write(
        tmp_path,
        "AGENTS.md",
        "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
        + body.replace("target.md", TOKENS["markdown-relative-link"]),
    )
    result = run(tmp_path)
    assert result.returncode == (1 if expected else 0), result.stdout + result.stderr
    if expected:
        assert "Stale protected-guide references" in result.stdout
    assert not (tmp_path / ".template-sync").exists()


def test_indented_code_context_removal_and_blanket_skip_mutants(tmp_path: Path) -> None:
    """Independent native expectations detect both false failures and false success."""
    reference_fixture(tmp_path, "markdown-relative-link")
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    start = source.index("def markdown_indented_code_lines(")
    stop = source.index("\ndef markdown_reference_definition_ends(", start)
    for body, expected_exit, replacement in (
        ("    [x](docs/azure-devops-support.md)\n", 0, "return set()"),
        (
            "Paragraph\n    [x](docs/azure-devops-support.md)\n",
            1,
            (
                "return {n for n, line in enumerate(markdown_lines(text), 1) "
                "if line.startswith(('    ', '\\t'))}"
            ),
        ),
    ):
        write(
            tmp_path,
            "AGENTS.md",
            "Agents MUST validate.\nAgents MUST preserve authority.\n\n" + body,
        )
        path.write_text(source, encoding="utf-8", newline="\n")
        positive = run(tmp_path)
        assert positive.returncode == expected_exit, positive.stdout + positive.stderr
        mutant = (
            source[:start]
            + "def markdown_indented_code_lines(text, definition_ends=None, *, include_indented=True, active_definitions=None, content_columns=None):\n    "
            + replacement
            + "\n\n"
            + source[stop:]
        )
        path.write_text(mutant, encoding="utf-8", newline="\n")
        negative = run(tmp_path)
        assert negative.returncode == 1 - expected_exit, negative.stdout + negative.stderr
        with pytest.raises(AssertionError):
            assert negative.returncode == expected_exit


def test_indented_container_scanning_is_bounded_and_rescan_mutant_times_out(
    tmp_path: Path,
) -> None:
    """Near-limit structural input has fixed output; repeated suffix work is caught."""
    reference_fixture(tmp_path, "markdown-relative-link")
    scripts = tmp_path / ".github/scripts"
    program = (
        "import sys; "
        f"sys.path.insert(0, {str(scripts)!r}); "
        "import instruction_contract_core as c; "
        "body='- '*450000+'[x](target.md)\\n'; "
        "assert c.markdown_link_targets_from_text(body)==((1,'target.md'),); "
        "body='- '*100000+'text\\n'+'\\n'*100000+'[x](target.md)\\n'; "
        "assert c.markdown_link_targets_from_text(body)==((100002,'target.md'),); "
        "body=' '*900000+'[x](ignored.md)\\n'; "
        "assert c.markdown_link_targets_from_text(body)==()"
    )
    positive = subprocess.run(
        [sys.executable, "-B", "-c", program],
        capture_output=True,
        text=True,
        check=False,
        timeout=15,
    )
    assert positive.returncode == 0, positive.stdout + positive.stderr
    path = scripts / "instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    guard = """        while not (
            thematic_start <= offset <= thematic_end
            and markdown_prefix_end(line, offset) - offset <= 3
        ):
"""
    assert source.count(guard) == 1
    old = '        while re.fullmatch(r" {0,3}(?:-[ ]*){3,}", line[offset:]) is None:\n'
    path.write_text(source.replace(guard, old), encoding="utf-8", newline="\n")
    with pytest.raises(subprocess.TimeoutExpired):
        subprocess.run(
            [sys.executable, "-B", "-c", program],
            capture_output=True,
            text=True,
            check=False,
            timeout=3,
        )


REFERENCE_CODE_BOUNDARIES = (
    ("[ref]: target.md\n    [x](ignored.md)\n", ((1, "target.md"),)),
    ("[ref]:\n target.md\n    [x](ignored.md)\n", ((1, "target.md"),)),
    ('[ref]: target.md\n "multiline\n title"\n    [x](ignored.md)\n', ((1, "target.md"),)),
    ("[ref]: target.md\n[ref]\n", ((1, "target.md"),)),
    ("> [ref]: target.md\n>     [x](ignored.md)\n", ((1, "target.md"),)),
    ("- [ref]: target.md\n      [x](ignored.md)\n", ((1, "target.md"),)),
    (
        "Paragraph\n[ref]: target.md\n    [x](live.md)\n",
        ((3, "live.md"),),
    ),
    ("[ref]: target.md trailing\n    [x](live.md)\n", ((2, "live.md"),)),
    ("[ref]: <unfinished\n    [x](live.md)\n", ((2, "live.md"),)),
    ("    [ref]: ignored.md\n    [x](ignored.md)\n", ()),
)


@pytest.mark.parametrize(("body", "expected"), REFERENCE_CODE_BOUNDARIES)
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
def test_complete_definitions_end_paragraphs_without_losing_target_inventory(
    body: str, expected: tuple[tuple[int, str], ...], ending: str
) -> None:
    """Complete definitions are nonparagraphs, but remain target-bearing inventory."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(body.replace("\n", ending)) == expected


def test_definition_boundary_removal_and_inventory_removal_have_native_oracles(
    tmp_path: Path,
) -> None:
    """Independent CLI checks preserve both code acceptance and live definitions."""
    reference_fixture(tmp_path, "markdown-relative-link")
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    anchor = "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
    write(
        tmp_path,
        "AGENTS.md",
        anchor + "[ref]: safe.md\n    [x](docs/azure-devops-support.md)\n",
    )
    positive = run(tmp_path)
    assert positive.returncode == 0, positive.stdout + positive.stderr
    guard = "definition_ends = markdown_reference_definition_ends(spans)"
    assert source.count(guard) == 1
    path.write_text(source.replace(guard, "definition_ends = {}"), encoding="utf-8", newline="\n")
    mutant = run(tmp_path)
    assert mutant.returncode == 1, mutant.stdout + mutant.stderr
    assert "Stale protected-guide references" in mutant.stdout
    path.write_text(source, encoding="utf-8", newline="\n")
    write(tmp_path, "AGENTS.md", anchor + "[ref]: docs/azure-devops-support.md\n[ref]\n")
    require_reference_failure(tmp_path)
    guard = "        if line_number in code_lines:\n"
    assert source.count(guard) == 1
    path.write_text(
        source.replace(guard, '        if line_number in code_lines or "]:" in line:\n'),
        encoding="utf-8",
        newline="\n",
    )
    missing = run(tmp_path)
    assert missing.returncode == 0, missing.stdout + missing.stderr
    with pytest.raises(AssertionError):
        assert missing.returncode == 1


HTML_LINK_CASES = (
    ("<!-- [x](target.md) -->\n", ()),
    ("<!--\n[x](ignored.md)\n-->\n[x](target.md)\n", ((4, "target.md"),)),
    ("<!--\n\n[x](ignored.md)\n", ()),
    ("<!--\n--> [x](ignored.md)\n[x](target.md)\n", ((3, "target.md"),)),
    ("<pre>\n[x](ignored.md)\n</pre>\n[x](target.md)\n", ((4, "target.md"),)),
    ("<SCRIPT>\n[x](ignored.md)\n</STYLE>\n[x](target.md)\n", ((4, "target.md"),)),
    ("<style>\n[x](ignored.md)\n</style>\n", ()),
    ("<textarea>\n[x](ignored.md)\n</textarea>\n", ()),
    ("<?process\n[x](ignored.md)\n?>\n[x](target.md)\n", ((4, "target.md"),)),
    ("<!decl\n[x](ignored.md)\n>\n[x](target.md)\n", ((4, "target.md"),)),
    ("<![CDATA[\n[x](ignored.md)\n]]>\n[x](target.md)\n", ((4, "target.md"),)),
    ("<div>\n[x](ignored.md)\n</div>\n[x](ignored.md)\n\n[x](target.md)\n", ((6, "target.md"),)),
    ("<custom key='value'>\n[x](ignored.md)\n\n[x](target.md)\n", ((4, "target.md"),)),
    ("Paragraph\n<custom>\n[x](target.md)\n", ((3, "target.md"),)),
    ("Paragraph\n<div>\n[x](ignored.md)\n", ()),
    ("<span>[x](target.md)</span>\n", ((1, "target.md"),)),
    ("Text <!-- [x](ignored.md) --> [x](target.md)\n", ((1, "target.md"),)),
    ('Text <span title="[x](ignored.md)"> [x](target.md)</span>\n', ((1, "target.md"),)),
    ("Text <? [x](ignored.md) ?> [x](target.md)\n", ((1, "target.md"),)),
    ("Text <![CDATA[[x](ignored.md)]]> [x](target.md)\n", ((1, "target.md"),)),
    ("Text <!decl [x](ignored.md)> [x](target.md)\n", ((1, "target.md"),)),
    ("Text <!--> [x](target.md)\n", ((1, "target.md"),)),
    ("Text <!---> [x](target.md)\n", ((1, "target.md"),)),
    ("[A <!-- ] --> B](target.md)\n", ((1, "target.md"),)),
    ('[A <span title="]"> B](target.md)\n', ((1, "target.md"),)),
    ('[A <span title="`"> B](target.md) `unclosed\n', ((1, "target.md"),)),
    ("Text <!-- unclosed [x](target.md)\n", ((1, "target.md"),)),
    ('Text <span title="unclosed [x](target.md)\n', ((1, "target.md"),)),
    (r"\<!-- [x](target.md) -->" + "\n", ((1, "target.md"),)),
    ("[x](raw<!--literal-->target.md)\n", ((1, "raw<!--literal-->target.md"),)),
    ('[x](target.md "<!-- literal -->")\n', ((1, "target.md"),)),
    ("`<!-- [x](ignored.md) -->` [x](target.md)\n", ((1, "target.md"),)),
    ("<!--\n```\n-->\n[x](target.md)\n", ((4, "target.md"),)),
    ("<pre>\n```\n</pre>\n[x](target.md)\n", ((4, "target.md"),)),
    ("```\n<!--\n```\n[x](target.md)\n", ((4, "target.md"),)),
    ("    <!--\n[x](target.md)\n", ((2, "target.md"),)),
    ("<!--\n[ref]: ignored.md\n-->\n[ref]: target.md\n", ((4, "target.md"),)),
    ("> <!--\n> [x](ignored.md)\n[x](target.md)\n", ((3, "target.md"),)),
    ("- <!--\n  [x](ignored.md)\n[x](target.md)\n", ((3, "target.md"),)),
    ("- > <pre>\n  > [x](ignored.md)\n  > </pre>\n\n[x](target.md)\n", ((5, "target.md"),)),
    ("   <!-- [x](ignored.md) -->\n", ()),
    ("Paragraph\n    <!-- [x](ignored.md) --> [x](target.md)\n", ((2, "target.md"),)),
)


@pytest.mark.parametrize(("body", "expected"), HTML_LINK_CASES)
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
def test_raw_html_preserves_live_links_components_and_original_lines(
    body: str, expected: tuple[tuple[int, str], ...], ending: str
) -> None:
    """Independent fixed targets distinguish opaque tokens from live inline text."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(body.replace("\n", ending)) == expected


@pytest.mark.parametrize(
    ("body", "expected_exit"),
    [
        ("<!-- [x](target.md) -->\n", 0),
        ("<pre>\n[x](target.md)\n</pre>\n", 0),
        ("<div>\n[x](target.md)\n\n", 0),
        ("Text <!-- [x](target.md) -->\n", 0),
        ('Text <span title="[x](target.md)"> ordinary text</span>\n', 0),
        ("<span>[x](target.md)</span>\n", 1),
        ("[A <!-- ] --> B](target.md)\n", 1),
        ('[A <span title="]"> B](target.md)\n', 1),
        ("<!--\n```\n-->\n[x](target.md)\n", 1),
    ],
)
def test_deployed_html_context_has_fixed_native_result_without_sync(
    tmp_path: Path, body: str, expected_exit: int
) -> None:
    """The installed checker accepts literals and rejects independently live links."""
    reference_fixture(tmp_path, "markdown-relative-link")
    write(
        tmp_path,
        "AGENTS.md",
        "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
        + body.replace("target.md", TOKENS["markdown-relative-link"]),
    )
    result = run(tmp_path)
    assert result.returncode == expected_exit, result.stdout + result.stderr
    if expected_exit:
        assert "Stale protected-guide references" in result.stdout
    assert not (tmp_path / ".template-sync").exists()


def test_html_block_and_inline_guard_removal_have_native_oracles(tmp_path: Path) -> None:
    """Removing either literal boundary violates a fixed deployed success oracle."""
    reference_fixture(tmp_path, "markdown-relative-link")
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    start = source.index("def markdown_link_html_block_end(")
    stop = source.index("\ndef markdown_thematic_suffix(", start)
    block_mutant = (
        source[:start]
        + "def markdown_link_html_block_end(content, *, paragraph_active):\n    return None\n\n"
        + source[stop:]
    )
    inline_guard = '        if character == "<" and index in html_ends:\n'
    assert source.count(inline_guard) == 1
    for body, mutant in (
        ("<div>\n[x](docs/azure-devops-support.md)\n", block_mutant),
        (
            "Text <!-- [x](docs/azure-devops-support.md) -->\n",
            source.replace(inline_guard, "        if False:\n"),
        ),
    ):
        write(
            tmp_path,
            "AGENTS.md",
            "Agents MUST validate.\nAgents MUST preserve authority.\n\n" + body,
        )
        path.write_text(source, encoding="utf-8", newline="\n")
        correct = run(tmp_path)
        assert correct.returncode == 0, correct.stdout + correct.stderr
        path.write_text(mutant, encoding="utf-8", newline="\n")
        incorrect = run(tmp_path)
        assert incorrect.returncode == 1, incorrect.stdout + incorrect.stderr
        assert "Stale protected-guide references" in incorrect.stdout
        with pytest.raises(AssertionError):
            assert incorrect.returncode == 0


def test_html_token_and_block_work_remains_bounded_near_input_limit() -> None:
    """Malformed prefixes and shared tag suffixes have fixed bounded output."""
    program = "\n".join(
        [
            "import sys",
            f"sys.path.insert(0, {str(ROOT / '.github/scripts')!r})",
            "import instruction_contract_core as c",
            "cases = [",
            "('Text '+'<!-- '*160000, ()),",
            "('<!--'+'a'*1000000+'\\n', ()),",
            "('Text '+'<a q=\\\"'*110000, ()),",
            "('Text '+'<a '*100000+'q '*300000, ()),",
            "('Text <!--'+'a'*900000+'--> [x](target.md)', ((1,'target.md'),)),",
            "('<pre>\\n'+'```\\n'*200000+'</pre>\\n[x](target.md)', ((200003,'target.md'),)),",
            "]",
            "for text, expected in cases:",
            "    assert len(text.encode('utf-8')) < c.MAXIMUM_INPUT_BYTES",
            "    assert c.markdown_link_targets_from_text(text) == expected",
        ]
    )
    result = subprocess.run(
        [sys.executable, "-B", "-c", program],
        capture_output=True,
        text=True,
        check=False,
        timeout=20,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_html_opacity_and_overbroad_element_mutants_keep_live_native_oracles(
    tmp_path: Path,
) -> None:
    """Fixed failures catch leaked fence state and accidental inline-content hiding."""
    reference_fixture(tmp_path, "markdown-relative-link")
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    guard = "        tuple(enumerate(markdown_lines(text), 1))\n"
    assert source.count(guard) == 1
    leaked_fence = source.replace(
        guard,
        "        lines_outside_markdown_fences(text, fence_context=fence_context)\n",
    )
    inline_guard = "            index = html_ends[index]\n"
    assert source.count(inline_guard) == 1
    hidden_text = source.replace(inline_guard, "            index = len(text)\n")
    for body, mutant in (
        ("<!--\n```\n-->\n[x](docs/azure-devops-support.md)\n", leaked_fence),
        ("<span>[x](docs/azure-devops-support.md)</span>\n", hidden_text),
    ):
        write(
            tmp_path,
            "AGENTS.md",
            "Agents MUST validate.\nAgents MUST preserve authority.\n\n" + body,
        )
        path.write_text(source, encoding="utf-8", newline="\n")
        correct = run(tmp_path)
        assert correct.returncode == 1, correct.stdout + correct.stderr
        assert "Stale protected-guide references" in correct.stdout
        path.write_text(mutant, encoding="utf-8", newline="\n")
        incorrect = run(tmp_path)
        assert incorrect.returncode == 0, incorrect.stdout + incorrect.stderr
        with pytest.raises(AssertionError):
            assert incorrect.returncode == 1


NESTED_LINK_CASES = (
    ("[outer [inner](target.md)](literal.md)", ((1, "target.md"),)),
    ("[a [b [inner](target.md)](middle.md)](literal.md)", ((1, "target.md"),)),
    (
        "[outer [a](target.md) [b](second.md)](literal.md) [c](after.md)",
        ((1, "target.md"), (1, "second.md"), (1, "after.md")),
    ),
    ("[outer [ordinary text]](target.md)", ((1, "target.md"),)),
    ("[outer [bad](not valid)](target.md)", ((1, "target.md"),)),
    ("[outer ![image](image.png)](target.md)", ((1, "target.md"),)),
    ("![outer [inner](target.md)](image.png)", ()),
    ("![outer ![inner](target.md)](image.png)", ()),
    ("[outer ![image [inner](literal.md)](image.png)](target.md)", ((1, "target.md"),)),
    ("![outer [inner](target.md)](not valid)", ((1, "target.md"),)),
    (r"\![outer [inner](target.md)](literal.md)", ((1, "target.md"),)),
    (r"\\![outer [inner](literal.md)](image.png)", ()),
    (r"[outer \[inner](target.md)](literal.md)", ((1, "target.md"),)),
    ("[outer `[fake](literal.md)`](target.md)", ((1, "target.md"),)),
    ("[outer [a `]`](target.md)](literal.md)", ((1, "target.md"),)),
    ("[outer <!-- [fake](literal.md) -->](target.md)", ((1, "target.md"),)),
    ("[outer [a <!-- ] --> b](target.md)](literal.md)", ((1, "target.md"),)),
    ("[outer <span>[inner](target.md)</span>](literal.md)", ((1, "target.md"),)),
    ("[outer](folder/[inner](literal.md))", ((1, "folder/[inner](literal.md)"),)),
    ('[outer](target.md "[inner](literal.md)")', ((1, "target.md"),)),
    ('[outer\n [inner](target.md\n "title")](literal.md)', ((2, "target.md"),)),
    ("> [outer [inner](target.md)](literal.md)", ((1, "target.md"),)),
    ("- [outer [inner](target.md)](literal.md)", ((1, "target.md"),)),
    ("[outer <https://example.invalid/>](target.md)", ((1, "target.md"),)),
)


@pytest.mark.parametrize(("body", "expected"), NESTED_LINK_CASES)
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
def test_nested_link_activity_preserves_images_components_and_physical_lines(
    body: str, expected: tuple[tuple[int, str], ...], ending: str
) -> None:
    """Fixed targets distinguish live inner labels from raw component nesting."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(body.replace("\n", ending)) == expected


REFERENCE_ACTIVITY_CASES = (
    ("[outer [inner][ref]](literal.md)\n\n[ref]: target.md", ((3, "target.md"),)),
    ("[outer [ref][]](literal.md)\n\n[ref]: target.md", ((3, "target.md"),)),
    ("[outer [ref]](literal.md)\n\n[ref]: target.md", ((3, "target.md"),)),
    ("[outer [missing]](target.md)", ((1, "target.md"),)),
    ("[outer [inner][missing]](target.md)", ((1, "target.md"),)),
    (
        "[outer [inner][missing]](target.md)\n\n[inner]: unused.md",
        ((1, "target.md"), (3, "unused.md")),
    ),
    ("[outer ![ref]](target.md)\n\n[ref]: image.png", ((1, "target.md"), (3, "image.png"))),
    ("![outer [inner](literal.md)][ref]\n\n[ref]: image.png", ((3, "image.png"),)),
    (
        "[outer ![image [inner](literal.md)][ref]](target.md)\n\n[ref]: image.png",
        ((1, "target.md"), (3, "image.png")),
    ),
    ("[outer [ẞ]](literal.md)\n\n[SS]: target.md", ((3, "target.md"),)),
    ("[outer [Foo\t bar]](literal.md)\n\n[foo bar]: target.md", ((3, "target.md"),)),
    ("[outer [Foo\n bar]](literal.md)\n\n[foo bar]: target.md", ((4, "target.md"),)),
    (r"[outer [ref\[]](literal.md)" + "\n\n" + r"[ref\[]: target.md", ((3, "target.md"),)),
    ("[outer [ref]](target.md)\n\n```\n[ref]: ignored.md\n```", ((1, "target.md"),)),
    ("[outer [ref]](target.md)\n\n<!--\n[ref]: ignored.md\n-->", ((1, "target.md"),)),
    ("[outer [ref]](target.md)\n\n    [ref]: ignored.md", ((1, "target.md"),)),
    (
        "[outer [ref]](target.md)\n\nParagraph\n[ref]: inventory.md",
        ((1, "target.md"),),
    ),
    (
        "[outer [ref]](literal.md)\n\n[ref]: first.md\n[REF]: second.md",
        ((3, "first.md"), (4, "second.md")),
    ),
    ("[unused]: target.md", ((1, "target.md"),)),
)


@pytest.mark.parametrize(("body", "expected"), REFERENCE_ACTIVITY_CASES)
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
def test_reference_activity_preserves_genuine_definition_inventory(
    body: str, expected: tuple[tuple[int, str], ...], ending: str
) -> None:
    """Only block definitions supply lookup labels and inventoried destinations."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(body.replace("\n", ending)) == expected


@pytest.mark.parametrize("length", [999, 1000])
def test_reference_activity_bounds_label_length_before_normalization(length: int) -> None:
    """Oversized labels do not deactivate outer links or lose raw inventory."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    label = "a" * length
    body = f"[outer [{label}]](outer.md)\n\n[{label}]: inventory.md"
    expected = ((3, "inventory.md"),) if length == 999 else ((1, "outer.md"), (3, "inventory.md"))
    assert core.markdown_link_targets_from_text(body) == expected


@pytest.mark.parametrize(
    ("body", "expected_exit"),
    [
        ("[outer [Guide](target.md)](kept.md)", 1),
        ("[outer [Guide](kept.md)](target.md)", 0),
        ("![outer [Guide](target.md)](image.png)", 0),
        ("[outer ![image [Guide](target.md)](image.png)](kept.md)", 0),
        ("[outer ![image](image.png)](target.md)", 1),
        ("[outer [Guide][ref]](target.md)\n\n[ref]: kept.md", 0),
        ("[outer [Guide][missing]](target.md)", 1),
    ],
)
def test_deployed_nested_link_outcomes_without_sync(
    tmp_path: Path, body: str, expected_exit: int
) -> None:
    """The native standalone route distinguishes nested labels and image alt text."""
    reference_fixture(tmp_path, "markdown-relative-link")
    write(
        tmp_path,
        "AGENTS.md",
        "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
        + body.replace("target.md", TOKENS["markdown-relative-link"]),
    )
    result = run(tmp_path)
    assert result.returncode == expected_exit, result.stdout + result.stderr
    if expected_exit:
        assert "Stale protected-guide references" in result.stdout
    assert not (tmp_path / ".template-sync").exists()


def test_nested_link_activity_image_and_reference_mutants_have_native_oracles(
    tmp_path: Path,
) -> None:
    """Three independent semantic guards are necessary for the fixed native cases."""
    reference_fixture(tmp_path, "markdown-relative-link")
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    activity = "            if not is_image and visible_links != link_checkpoint:\n"
    image = "                    del candidates[candidate_checkpoint:]\n"
    reference = "    reference_labels = reference_labels or set()\n"
    for guard in (activity, image, reference):
        assert source.count(guard) == 1
    cases = (
        (
            "[outer [Guide](target.md)](kept.md)",
            1,
            source.replace(activity, "            if False:\n"),
        ),
        (
            "[outer [Guide](kept.md)](target.md)",
            0,
            source.replace(activity, "            if False:\n"),
        ),
        ("![outer [Guide](target.md)](image.png)", 0, source.replace(image, "")),
        (
            "[outer [Guide][ref]](target.md)\n\n[ref]: kept.md",
            0,
            source.replace(reference, "    reference_labels = set()\n"),
        ),
    )
    for body, expected_exit, mutant in cases:
        write(
            tmp_path,
            "AGENTS.md",
            "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
            + body.replace("target.md", TOKENS["markdown-relative-link"]),
        )
        path.write_text(source, encoding="utf-8", newline="\n")
        correct = run(tmp_path)
        assert correct.returncode == expected_exit, correct.stdout + correct.stderr
        path.write_text(mutant, encoding="utf-8", newline="\n")
        incorrect = run(tmp_path)
        assert incorrect.returncode == 1 - expected_exit, incorrect.stdout + incorrect.stderr
        assert "Traceback" not in incorrect.stderr
        if incorrect.returncode:
            assert "Stale protected-guide references" in incorrect.stdout
        with pytest.raises(AssertionError):
            assert incorrect.returncode == expected_exit


def test_nested_link_and_image_frames_remain_bounded_near_input_limit() -> None:
    """Deep labels and images avoid ancestor rescans and preserve output bounds."""
    program = "\n".join(
        [
            "import sys",
            f"sys.path.insert(0, {str(ROOT / '.github/scripts')!r})",
            "import instruction_contract_core as c",
            "cases = [",
            "('['*70000+'[x](target.md)'+'](literal.md)'*70000, ((1,'target.md'),)),",
            "('!['*100000+'[x](ignored.md)'+'](i)'*100000+' [x](target.md)', ((1,'target.md'),)),",
            "('![image '+'[x](ignored.md) '*60000+'](i) [x](target.md)', ((1,'target.md'),)),",
            "('[outer '+'[ref] '*150000+'](literal.md)\\n\\n[ref]: target.md', ((3,'target.md'),)),",
            "]",
            "for text, expected in cases:",
            "    assert len(text.encode('utf-8')) < c.MAXIMUM_INPUT_BYTES",
            "    result=c.markdown_link_targets_from_text(text)",
            "    assert result == expected, (len(result), result[:2])",
            "    assert sum(len(target) for _,target in result) <= len(text)",
        ]
    )
    result = subprocess.run(
        [sys.executable, "-B", "-c", program],
        capture_output=True,
        text=True,
        check=False,
        timeout=20,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_nested_live_reference_preserves_exact_exception_digest(tmp_path: Path) -> None:
    """Nested syntax preserves the existing exact target and content-bound authority."""
    test_deployed_multiline_reference_fails_without_sync_and_exact_exception_passes(
        tmp_path, "[outer [Guide](docs/azure-devops-support.md)](kept.md)"
    )


PARAGRAPH_DEFINITION_CASES = (
    ("[Guide]\n[Guide]: target.md", ()),
    ("[Guide]\n\n[Guide]: target.md", ((3, "target.md"),)),
    ("> [Guide]\n> [Guide]: target.md", ()),
    ("- [Guide]\n  [Guide]: target.md", ()),
    ("> [Guide]\n[Guide]: target.md", ()),
    ("- [Guide]\n[Guide]: target.md", ()),
    ("Paragraph\n[Guide]: literal [live](target.md)", ((2, "target.md"),)),
    ('Paragraph\n[Guide]: kept.md "[live](target.md)"', ((2, "target.md"),)),
    ("[kept]: kept.md\n[Guide]: target.md\n[Guide]", ((1, "kept.md"), (2, "target.md"))),
    ("# Header\n[Guide]: target.md\n[Guide]", ((2, "target.md"),)),
    ("[kept]: kept.md\nParagraph\n[Guide]: target.md\n[Guide]", ((1, "kept.md"),)),
    ("Paragraph\n[Guide\n label]: target.md", ()),
    ("Paragraph\n---\n[Guide]: target.md", ((3, "target.md"),)),
    ("Paragraph\n***\n[Guide]: target.md", ((3, "target.md"),)),
    ("```text\nliteral\n```\n[Guide]: target.md", ((4, "target.md"),)),
    ("> Paragraph\n>\n> [Guide]: target.md", ((3, "target.md"),)),
    ("- Paragraph\n\n  [Guide]: target.md", ((3, "target.md"),)),
    ("Paragraph\n> [Guide]: target.md", ((2, "target.md"),)),
    ("Paragraph\n- [Guide]: target.md", ((2, "target.md"),)),
    ('[Guide\n label]:\n target.md\n "title"', ((1, "target.md"),)),
    ("[unused]: target.md\n[UNUSED]: other.md", ((1, "target.md"), (2, "other.md"))),
    ('Paragraph\n[Guide]: kept.md "[outer [live](target.md)](literal.md)"', ((2, "target.md"),)),
)


@pytest.mark.parametrize(("body", "expected"), PARAGRAPH_DEFINITION_CASES)
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
def test_definition_candidates_follow_block_context_and_keep_inline_links(
    body: str, expected: tuple[tuple[int, str], ...], ending: str
) -> None:
    """Definition-shaped paragraph text remains inline without creating false targets."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(body.replace("\n", ending)) == expected


@pytest.mark.parametrize(("body", "expected"), PARAGRAPH_DEFINITION_CASES)
def test_deployed_definition_context_has_fixed_native_results_without_sync(
    tmp_path: Path, body: str, expected: tuple[tuple[int, str], ...]
) -> None:
    """The copied standalone runtime distinguishes real definitions from continuations."""
    reference_fixture(tmp_path, "markdown-relative-link")
    write(
        tmp_path,
        "AGENTS.md",
        "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
        + body.replace("target.md", TOKENS["markdown-relative-link"]),
    )
    result = run(tmp_path)
    expected_exit = int(any(target == "target.md" for _, target in expected))
    assert result.returncode == expected_exit, result.stdout + result.stderr
    if expected_exit:
        assert "Stale protected-guide references" in result.stdout
    assert not (tmp_path / ".template-sync").exists()


def test_definition_context_guard_removal_has_both_native_failure_oracles(tmp_path: Path) -> None:
    """Removing block context restores false failures and hides genuine inline links."""
    reference_fixture(tmp_path, "markdown-relative-link")
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    guard = "            if definition_starts is None or line_start in definition_starts:\n"
    assert source.count(guard) == 1
    mutant = source.replace(guard, "            if True:\n")
    for body, expected_exit in (
        ("[Guide]\n[Guide]: target.md", 0),
        ('Paragraph\n[Guide]: kept.md "[live](target.md)"', 1),
    ):
        write(
            tmp_path,
            "AGENTS.md",
            "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
            + body.replace("target.md", TOKENS["markdown-relative-link"]),
        )
        path.write_text(source, encoding="utf-8", newline="\n")
        correct = run(tmp_path)
        assert correct.returncode == expected_exit, correct.stdout + correct.stderr
        path.write_text(mutant, encoding="utf-8", newline="\n")
        incorrect = run(tmp_path)
        assert incorrect.returncode == 1 - expected_exit, incorrect.stdout + incorrect.stderr
        assert "Traceback" not in incorrect.stderr
        if incorrect.returncode:
            assert "Stale protected-guide references" in incorrect.stdout
        with pytest.raises(AssertionError):
            assert incorrect.returncode == expected_exit


def test_repeated_paragraph_definition_lookalikes_remain_bounded() -> None:
    """Near-limit definition lookalikes neither leak targets nor hide a final live link."""
    program = "\n".join(
        [
            "import sys",
            f"sys.path.insert(0, {str(ROOT / '.github/scripts')!r})",
            "import instruction_contract_core as c",
            "text='Paragraph\\n'+'[Guide]: ignored.md\\n'*50000+'[live](target.md)'",
            "assert len(text.encode('utf-8')) < c.MAXIMUM_INPUT_BYTES",
            "assert c.markdown_link_targets_from_text(text) == ((50002, 'target.md'),)",
        ]
    )
    result = subprocess.run(
        [sys.executable, "-B", "-c", program],
        capture_output=True,
        text=True,
        check=False,
        timeout=20,
    )
    assert result.returncode == 0, result.stdout + result.stderr


EMPTY_LIST_MARKERS = (
    "-",
    "+",
    "*",
    "1.",
    "0)",
    "123456789.",
    "> -",
    "- -",
    "1. -",
    "> 1. -",
    "  -",
    "- ",
    "-\t",
)

EMPTY_LIST_LINK_CASES = (
    ("-\n  [x](target.md)", ((2, "target.md"),)),
    ("-\n      [x](target.md)", ()),
    ("    -\n    [x](target.md)", ()),
    ("Paragraph\n+\n    [x](target.md)", ((3, "target.md"),)),
    ("Paragraph\n1.\n    [x](target.md)", ((3, "target.md"),)),
    ("```text\n-\n```\n[x](target.md)", ((4, "target.md"),)),
    ("<!--\n-\n-->\n[x](target.md)", ((4, "target.md"),)),
    (">", ()),
    ("Paragraph\n- \n    [x](target.md)", ()),
    ("Paragraph\n-\n    [x](target.md)", ()),
    ("> Paragraph\n> - \n>     [x](target.md)", ()),
    ("- Paragraph\n  - \n      [x](target.md)", ()),
    ("> - Paragraph\n>   - \n>       [x](target.md)", ()),
    ("Paragraph\n   - \n    [x](target.md)", ()),
    ("[x](target.md)\n- ", ((1, "target.md"),)),
    ("Paragraph\n- \n[x](target.md)", ((3, "target.md"),)),
    ("> Paragraph\n- \n    [x](target.md)", ((3, "target.md"),)),
    ("- Paragraph\n- \n    [x](target.md)", ((3, "target.md"),)),
    ("Paragraph\n+ \n    [x](target.md)", ((3, "target.md"),)),
    ("Paragraph\n1. \n    [x](target.md)", ((3, "target.md"),)),
    ("Paragraph\n--\n    [x](target.md)", ()),
    ("Paragraph\n=\n    [x](target.md)", ()),
    ("- \n\n    [x](target.md)", ()),
    ("-\n\n    [x](target.md)", ()),
    ("+ \n\n    [x](target.md)", ()),
    ("* \n\n    [x](target.md)", ()),
    ("1. \n\n    [x](target.md)", ()),
    ("> - \n>\n>     [x](target.md)", ()),
    ("- - \n\n      [x](target.md)", ()),
    ("- - \n\n    [x](target.md)", ((3, "target.md"),)),
    ("- > - \n  >\n  >   [x](target.md)", ((3, "target.md"),)),
    ("- \n    [x](target.md)", ((2, "target.md"),)),
    ("-\n    [x](target.md)", ((2, "target.md"),)),
    ("- \n      [x](target.md)", ()),
    ("- \n\n  [x](target.md)", ((3, "target.md"),)),
    ("- Text\n\n    [x](target.md)", ((3, "target.md"),)),
    ("- Text\n\n      [x](target.md)", ()),
    ("- \n\n- [x](target.md)", ((3, "target.md"),)),
    ("> - \n\n    [x](target.md)", ()),
    ("- \n\n\n    [x](target.md)", ()),
    ("- \n  \n    [x](target.md)", ()),
    ("-\t\n\t\n\t[x](target.md)", ()),
)


@pytest.mark.parametrize("marker", EMPTY_LIST_MARKERS)
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
def test_empty_list_markers_preserve_virtual_margins_without_crashing(
    marker: str, ending: str
) -> None:
    """Valid bare markers have no destinations under every physical line ending."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(marker + ending) == ()


@pytest.mark.parametrize(
    ("line", "offset", "expected"),
    [
        ("", 0, None),
        ("-", 1, None),
        ("-", 2, None),
        ("-", 0, ("list", 2)),
        ("-    ", 0, ("list", 2)),
        ("123456789.", 0, ("list", 11)),
    ],
)
def test_empty_list_container_offsets_are_exhausted_or_preserve_required_margin(
    line: str, offset: int, expected: tuple[str, int] | None
) -> None:
    """An exhausted source offset and a virtual continuation margin are distinct."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_container(line, offset) == expected


@pytest.mark.parametrize(("body", "expected"), EMPTY_LIST_LINK_CASES)
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
def test_empty_list_boundaries_preserve_code_live_links_and_physical_lines(
    body: str, expected: tuple[tuple[int, str], ...], ending: str
) -> None:
    """Fixed rendered cases distinguish setext headings and empty item lifetimes."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(body.replace("\n", ending)) == expected


@pytest.mark.parametrize(
    ("body", "expected"),
    tuple((marker, ()) for marker in EMPTY_LIST_MARKERS) + EMPTY_LIST_LINK_CASES,
)
def test_deployed_empty_list_boundaries_have_fixed_native_results_without_sync(
    tmp_path: Path, body: str, expected: tuple[tuple[int, str], ...]
) -> None:
    """The copied CLI accepts inert blocks and rejects genuinely live excluded links."""
    reference_fixture(tmp_path, "markdown-relative-link")
    write(
        tmp_path,
        "AGENTS.md",
        "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
        + body.replace("target.md", TOKENS["markdown-relative-link"]),
    )
    result = run(tmp_path)
    assert result.returncode == int(bool(expected)), result.stdout + result.stderr
    assert "Traceback" not in result.stderr
    if expected:
        assert "Stale protected-guide references" in result.stdout
        assert "azure-reference" in result.stdout
    assert not (tmp_path / ".template-sync").exists()


@pytest.mark.parametrize("ending", ["\r\n", "\r"])
@pytest.mark.parametrize(
    ("body", "expected_exit"),
    [
        ("-\n", 0),
        ("Paragraph\n-\n    [x](target.md)", 0),
        ("-\n\n    [x](target.md)", 0),
        ("-\n  [x](target.md)", 1),
    ],
)
def test_deployed_empty_lists_keep_native_meaning_with_alternate_line_endings(
    tmp_path: Path, ending: str, body: str, expected_exit: int
) -> None:
    """Write physical CRLF/CR bytes so native tests do not normalize their inputs."""
    reference_fixture(tmp_path, "markdown-relative-link")
    text = "Agents MUST validate.\nAgents MUST preserve authority.\n\n" + body.replace(
        "target.md", TOKENS["markdown-relative-link"]
    )
    (tmp_path / "AGENTS.md").write_bytes(text.replace("\n", ending).encode("utf-8"))
    result = run(tmp_path)
    assert result.returncode == expected_exit, result.stdout + result.stderr
    assert "Traceback" not in result.stderr
    if expected_exit:
        assert "Stale protected-guide references" in result.stdout


@pytest.mark.parametrize(
    "body",
    ["-\n", "Paragraph\n-\n    [x](target.md)", "-\n\n    [x](target.md)"],
)
def test_materialization_accepts_empty_list_and_heading_code_controls(
    tmp_path: Path, body: str
) -> None:
    """The actual migration materializer accepts the same three inert native cases."""
    stage = tmp_path / "stage"
    target = tmp_path / "target"
    reference_fixture(stage, "markdown-relative-link")
    reference_fixture(target, "markdown-relative-link")
    write(
        stage,
        "AGENTS.md",
        "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
        + body.replace("target.md", TOKENS["markdown-relative-link"]),
    )
    result = run_migration_schema_control(stage, target)
    assert result.returncode == 0, result.stdout + result.stderr
    assert "Traceback" not in result.stderr


def test_empty_list_guard_removals_have_independent_native_failure_oracles(
    tmp_path: Path, monkeypatch: Any
) -> None:
    """Three independent positive oracles expose the original crash and false failures."""
    # Equal-size edits within one timestamp interval must execute fresh source.
    monkeypatch.setenv("PYTHONDONTWRITEBYTECODE", "1")
    reference_fixture(tmp_path, "markdown-relative-link")
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    cases = (
        (
            "    if start - offset > 3 or start >= len(line):\n",
            "    if start - offset > 3 or start == len(line):\n",
            "-\n",
            "IndexError: string index out of range",
        ),
        (
            "        if blank or setext:\n",
            "        if blank:\n",
            "Paragraph\n- \n    [x](target.md)",
            "Stale protected-guide references",
        ),
        (
            (
                "        if blank and empty_list is not None:\n"
                "            # An item starting empty cannot consume another initial blank line.\n"
                "            matched = min(matched, empty_list)\n"
            ),
            "",
            "- \n\n    [x](target.md)",
            "Stale protected-guide references",
        ),
    )
    for guard, replacement, body, diagnostic in cases:
        assert source.count(guard) == 1
        write(
            tmp_path,
            "AGENTS.md",
            "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
            + body.replace("target.md", TOKENS["markdown-relative-link"]),
        )
        path.write_text(source, encoding="utf-8", newline="\n")
        positive = run(tmp_path)
        assert positive.returncode == 0, positive.stdout + positive.stderr
        path.write_text(source.replace(guard, replacement), encoding="utf-8", newline="\n")
        mutant = run(tmp_path)
        assert mutant.returncode == 1, mutant.stdout + mutant.stderr
        assert diagnostic in mutant.stdout + mutant.stderr
        assert not list((path.parent / "__pycache__").glob("instruction_contract_core.*.pyc"))
        if diagnostic.startswith("Stale"):
            assert "Traceback" not in mutant.stderr
        with pytest.raises(AssertionError):
            assert mutant.returncode == 0, "This fixed inert case must pass native validation."


def test_repeated_and_nested_empty_list_boundaries_remain_bounded() -> None:
    """Near-limit container input preserves exact output without repeated suffix scans."""
    program = "\n".join(
        [
            "import sys",
            f"sys.path.insert(0, {str(ROOT / '.github/scripts')!r})",
            "import instruction_contract_core as c",
            "cases = (",
            "    ('> - '*150000+'-\\n', ()),",
            (
                "    ('-\\n\\n    [x](ignored.md)\\n\\n'*40000+'[live](target.md)',"
                " ((160001, 'target.md'),)),"
            ),
            (
                "    ('Text\\n-\\n    [x](ignored.md)\\n\\n'*30000+'[live](target.md)',"
                " ((120001, 'target.md'),)),"
            ),
            ")",
            "for text, expected in cases:",
            "    assert len(text.encode('utf-8')) < c.MAXIMUM_INPUT_BYTES",
            "    assert c.markdown_link_targets_from_text(text) == expected",
        ]
    )
    result = subprocess.run(
        [sys.executable, "-B", "-c", program],
        capture_output=True,
        text=True,
        check=False,
        timeout=20,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def obligation_identity_catalog(
    *,
    reference_key: str = "azure-obligation",
    reference_path: str = "AGENTS.md",
    reference_kind: str = "markdown-relative-link",
    reference_module: str = "azure-devops-platform",
) -> dict[str, Any]:
    """Build distinct obligation families whose declared identities may collide."""
    reference: dict[str, Any] = {
        "key": reference_key,
        "path": reference_path,
        "reference_kind": reference_kind,
        "target_modules": [reference_module],
        "target_path": "docs/azure-devops-support.md",
    }
    if reference_kind != "markdown-relative-link":
        reference["tokens"] = [TOKENS[reference_kind]]
    return {
        "instruction_contracts": [
            {
                "path": "AGENTS.md",
                "requires_modules": ["agent-instructions"],
                "required_phrases": ["Agents MUST validate."],
            }
        ],
        "protected_guide_section_obligations": [
            {
                "key": "azure-obligation",
                "path": "AGENTS.md",
                "target_modules": ["azure-devops-platform"],
                "stale_phrases": ["Legacy Azure section."],
            }
        ],
        "protected_guide_reference_obligations": [reference],
    }


@pytest.mark.parametrize("family", ["section", "reference"])
@pytest.mark.parametrize("kind", KINDS)
@pytest.mark.parametrize("reference_module", ["azure-devops-platform", "terraform"])
def test_cross_family_obligation_identity_is_rejected_by_either_parser(
    family: str, kind: str, reference_module: str
) -> None:
    """Module sets and reference kind cannot distinguish a type-less waiver key."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    catalog = obligation_identity_catalog(reference_kind=kind, reference_module=reference_module)
    parser = getattr(core, f"parse_protected_guide_{family}_obligations")
    with pytest.raises(
        core.InstructionContractValidationError,
        match=r"across section and reference types: \(AGENTS.md, azure-obligation\)",
    ):
        parser(catalog, {"agent-instructions", "azure-devops-platform", "terraform"})


@pytest.mark.parametrize("family", ["section", "reference"])
@pytest.mark.parametrize("distinct", ["key", "path", "empty-section", "empty-reference"])
def test_cross_family_obligation_identity_preserves_distinct_and_single_families(
    family: str, distinct: str
) -> None:
    """Valid per-path identities do not acquire a global key or path-only restriction."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    catalog = obligation_identity_catalog()
    if distinct == "key":
        catalog["protected_guide_reference_obligations"][0]["key"] = "separate-reference"
    elif distinct == "path":
        catalog["protected_guide_reference_obligations"][0]["path"] = "CLAUDE.md"
    else:
        catalog[f"protected_guide_{distinct.removeprefix('empty-')}_obligations"] = []
    parser = getattr(core, f"parse_protected_guide_{family}_obligations")
    parsed = parser(catalog, {"agent-instructions", "azure-devops-platform"})
    assert len(parsed) == len(catalog[f"protected_guide_{family}_obligations"])


@pytest.mark.parametrize("family", ["section", "reference"])
@pytest.mark.parametrize("invalid", ["duplicate", "mapping", "key", "path", "module"])
def test_cross_family_identity_preserves_existing_invalid_catalog_failures(
    family: str, invalid: str
) -> None:
    """The new shared identity pass does not weaken either family's existing guards."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    catalog = obligation_identity_catalog(reference_key="separate-reference")
    records = catalog[f"protected_guide_{family}_obligations"]
    if invalid == "duplicate":
        records.append(dict(records[0]))
    elif invalid == "mapping":
        records[0] = "not a mapping"
    elif invalid == "key":
        records[0]["key"] = 1
    elif invalid == "path":
        records[0]["path"] = "../AGENTS.md"
    else:
        records[0]["target_modules"] = ["unknown-module"]
    parser = getattr(core, f"parse_protected_guide_{family}_obligations")
    with pytest.raises(
        (core.InstructionContractValidationError, core.support.TemplateSyncMaterializationError)
    ):
        parser(catalog, {"agent-instructions", "azure-devops-platform"})


@pytest.mark.parametrize("kind", KINDS)
@pytest.mark.parametrize("retained", [False, True])
def test_deployed_cross_family_identity_rejects_inert_catalog_without_sync(
    tmp_path: Path, kind: str, retained: bool
) -> None:
    """An inert ambiguous catalog is invalid independently of current applicability."""
    document, _ = reference_fixture(tmp_path, kind)
    if retained:
        document["modules"].append("azure-devops-platform")
        write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
    write(
        tmp_path,
        ".github/instruction-contracts.yml",
        yaml.safe_dump(obligation_identity_catalog(reference_kind=kind)),
    )
    result = run(tmp_path)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Duplicate protected-guide obligation" in result.stdout + result.stderr
    assert "(AGENTS.md, azure-obligation)" in result.stdout + result.stderr
    assert "Traceback" not in result.stderr
    assert not (tmp_path / ".template-sync").exists()


def test_marker_section_only_parser_rejects_cross_family_identity(tmp_path: Path) -> None:
    """The actual marker CLI must inspect identities even when it checks only sections."""
    from tests.test_validate_instruction_contracts import (
        _marker,
        _run_validator,
        _write_common_contract_repo,
        _write_text,
        _write_yaml,
    )

    catalog = obligation_identity_catalog()
    for family in ("section", "reference"):
        catalog[f"protected_guide_{family}_obligations"][0]["target_modules"] = [
            "azure-devops-collaboration"
        ]
    _write_common_contract_repo(tmp_path, catalog)
    _write_text(tmp_path, "AGENTS.md", "Agents MUST validate.\n")
    _write_yaml(
        tmp_path,
        ".template-sync/marker.yml",
        _marker(["agent-instructions", "template-sync-support"]),
    )
    result = _run_validator(tmp_path, "--mode", "downstream", "--require-marker")
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Duplicate protected-guide obligation" in result.stdout + result.stderr
    assert "Traceback" not in result.stderr


def test_reference_only_downstream_report_rejects_cross_family_identity(tmp_path: Path) -> None:
    """Reference-only consumers cannot bypass the combined catalog invariant."""
    reference_fixture(tmp_path, "markdown-relative-link")
    write(
        tmp_path,
        ".github/instruction-contracts.yml",
        yaml.safe_dump(obligation_identity_catalog()),
    )
    program = "\n".join(
        [
            "import sys",
            "from pathlib import Path",
            f"sys.path.insert(0, {str(ROOT / '.template-sync/scripts')!r})",
            "import validate_downstream_adoption as v",
            "import instruction_contract_core as c",
            "root=Path(sys.argv[1])",
            "try:",
            "    v.protected_guide_reference_report_items(",
            "        repo_root=root, contracts_path=root/'.github/instruction-contracts.yml',",
            "        contracts_schema_path=root/'schemas/instruction-contracts.schema.json',",
            "        manifest_modules={'agent-instructions','azure-devops-platform'},",
            "        mappings=(), included_modules={'agent-instructions'},",
            "        protected_guide_waivers=())",
            "except c.InstructionContractValidationError as error:",
            "    print(error)",
            "    raise SystemExit(1)",
        ]
    )
    result = subprocess.run(
        [sys.executable, "-B", "-c", program, str(tmp_path)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Duplicate protected-guide obligation" in result.stdout
    assert "Traceback" not in result.stderr


def test_cross_family_identity_rejects_waiver_migration_before_profile_write(
    tmp_path: Path,
) -> None:
    """One marker record cannot be expanded into both stale and reference exceptions."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    for root in (stage, target):
        document, _ = reference_fixture(root, "markdown-relative-link")
        write(
            root, ".github/instruction-contracts.yml", yaml.safe_dump(obligation_identity_catalog())
        )
        write(
            root,
            "AGENTS.md",
            "Agents MUST validate.\n\nLegacy Azure section.\n"
            "[Guide](docs/azure-devops-support.md)\n",
        )
    before = (stage / ".github/instruction-profile.yml").read_bytes()
    marker = {
        "template_sync": {
            "included_modules": document["modules"],
            "protected_guide_contract_waivers": [
                {
                    "path": "AGENTS.md",
                    "contract_key": "azure-obligation",
                    "target_module": "azure-devops-platform",
                    "reason": "One selected obligation",
                    "authorization_basis": "Fixture declaration",
                }
            ],
        }
    }
    result = run_migration_schema_control(stage, target, marker_document=marker)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Duplicate protected-guide obligation" in result.stdout + result.stderr
    assert (stage / ".github/instruction-profile.yml").read_bytes() == before


@pytest.mark.parametrize("distinct", ["key", "path"])
def test_distinct_obligation_identities_keep_exact_waiver_migration_and_noop(
    tmp_path: Path, distinct: str
) -> None:
    """Two separately identified waiver records retain their exact typed exceptions."""
    reference_key = "separate-reference" if distinct == "key" else "azure-obligation"
    reference_path = "CLAUDE.md" if distinct == "path" else "AGENTS.md"
    catalog = obligation_identity_catalog(
        reference_key=reference_key, reference_path=reference_path
    )
    stage, target = tmp_path / "stage", tmp_path / "target"
    for root in (stage, target):
        document, _ = reference_fixture(root, "markdown-relative-link")
        write(root, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
        write(root, "AGENTS.md", "Agents MUST validate.\n\nLegacy Azure section.\n")
        path = root / reference_path
        before_reference = path.read_text(encoding="utf-8") if path.exists() else ""
        write(root, reference_path, before_reference + "[Guide](docs/azure-devops-support.md)\n")
    waivers = [
        {
            "path": "AGENTS.md",
            "contract_key": "azure-obligation",
            "target_module": "azure-devops-platform",
            "reason": "Selected section",
            "authorization_basis": "Fixture declaration",
        },
        {
            "path": reference_path,
            "contract_key": reference_key,
            "target_path": "docs/azure-devops-support.md",
            "reason": "Selected reference",
            "authorization_basis": "Separate fixture declaration",
        },
    ]
    marker = {
        "template_sync": {
            "included_modules": document["modules"],
            "protected_guide_contract_waivers": waivers,
        }
    }
    result = run_migration_schema_control(stage, target, marker_document=marker)
    assert result.returncode == 0, result.stdout + result.stderr
    profile_path = stage / ".github/instruction-profile.yml"
    before = profile_path.read_bytes()
    generated = yaml.safe_load(before)
    assert {(x["path"], x["anchor"]) for x in generated["exceptions"]} == {
        ("AGENTS.md", "stale:azure-obligation:phrase:Legacy Azure section."),
        (
            reference_path,
            f"reference:{reference_key}:markdown-relative-link:docs/azure-devops-support.md",
        ),
    }
    accepted = run(stage)
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    write(target, ".github/instruction-profile.yml", before.decode("utf-8"))
    repeated = run_migration_schema_control(stage, target, marker_document=marker)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert profile_path.read_bytes() == before


def test_cross_family_identity_mutants_have_independent_native_oracles(
    tmp_path: Path, monkeypatch: Any
) -> None:
    """Fixed invalid and valid catalogs detect missing and overbroad identity guards."""
    monkeypatch.setenv("PYTHONDONTWRITEBYTECODE", "1")
    reference_fixture(tmp_path, "markdown-relative-link")
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    for catalog, guard, replacement, expected_exit in (
        (
            obligation_identity_catalog(),
            "    duplicate_keys = family_keys[0] & family_keys[1]\n",
            "    duplicate_keys = set()\n",
            1,
        ),
        (
            obligation_identity_catalog(reference_path="CLAUDE.md"),
            "            keys.add((path, key))\n",
            '            keys.add(("", key))\n',
            0,
        ),
        (
            obligation_identity_catalog(reference_key="separate-reference"),
            "            keys.add((path, key))\n",
            '            keys.add((path, ""))\n',
            0,
        ),
    ):
        assert source.count(guard) == 1
        write(tmp_path, ".github/instruction-contracts.yml", yaml.safe_dump(catalog))
        write(tmp_path, "CLAUDE.md", "No stale reference.\n")
        path.write_text(source, encoding="utf-8", newline="\n")
        correct = run(tmp_path)
        assert correct.returncode == expected_exit, correct.stdout + correct.stderr
        path.write_text(source.replace(guard, replacement), encoding="utf-8", newline="\n")
        mutant = run(tmp_path)
        assert mutant.returncode == 1 - expected_exit, mutant.stdout + mutant.stderr
        assert "Traceback" not in mutant.stderr
        if mutant.returncode:
            assert "Duplicate protected-guide obligation" in mutant.stdout + mutant.stderr
        with pytest.raises(AssertionError):
            assert mutant.returncode == expected_exit


def test_combined_obligation_identity_scan_is_bounded_and_deterministic() -> None:
    """A near-limit inventory keeps valid keys and reports sorted cross-family pairs."""
    program = "\n".join(
        [
            "import json, sys",
            f"sys.path.insert(0, {str(ROOT / '.github/scripts')!r})",
            "import instruction_contract_core as c",
            "section={'path':'AGENTS.md','target_modules':['terraform'],'stale_phrases':['Old']}",
            (
                "reference={'path':'AGENTS.md','target_modules':['terraform'],"
                "'reference_kind':'prose-reference','tokens':['Old']}"
            ),
            "catalog={'protected_guide_section_obligations':",
            "    [dict(section,key=f'section-{i}') for i in range(3000)],",
            "    'protected_guide_reference_obligations':",
            "    [dict(reference,key=f'reference-{i}') for i in range(3000)]}",
            "assert len(json.dumps(catalog).encode()) < c.MAXIMUM_INPUT_BYTES",
            "assert len(c.parse_protected_guide_section_obligations(catalog,{'terraform'}))==3000",
            "assert len(c.parse_protected_guide_reference_obligations(catalog,{'terraform'}))==3000",
            "catalog['protected_guide_reference_obligations'][-2]['key']='section-9'",
            "catalog['protected_guide_reference_obligations'][-1]['key']='section-1'",
            "for parser in (c.parse_protected_guide_section_obligations,",
            "               c.parse_protected_guide_reference_obligations):",
            "    try: parser(catalog,{'terraform'})",
            "    except c.InstructionContractValidationError as error:",
            "        assert str(error).endswith('(AGENTS.md, section-1), (AGENTS.md, section-9)')",
            "    else: raise AssertionError('An ambiguous combined inventory must fail.')",
        ]
    )
    result = subprocess.run(
        [sys.executable, "-B", "-c", program],
        capture_output=True,
        text=True,
        check=False,
        timeout=15,
    )
    assert result.returncode == 0, result.stdout + result.stderr


NESTED_DEFINITION_CASES = (
    ("[Guide]: target.md\n\n[Guide]\n", ((1, "target.md"),)),
    ("- [Guide]: target.md\n\n[Guide]\n", ((1, "target.md"),)),
    ("- - [Guide]: target.md\n\n[Guide]\n", ((1, "target.md"),)),
    ("- + * [Guide]: target.md\n\n[Guide]\n", ((1, "target.md"),)),
    ("1. 2) [Guide]: target.md\n\n[Guide]\n", ((1, "target.md"),)),
    ("> - - [Guide]: target.md\n\n[Guide]\n", ((1, "target.md"),)),
    ("- > [Guide]: target.md\n\n[Guide]\n", ((1, "target.md"),)),
    ("> - > 1. [Guide]: target.md\n\n[Guide]\n", ((1, "target.md"),)),
    ("-\t[Guide]: target.md\n\n[Guide]\n", ((1, "target.md"),)),
    ("-\t-\t[Guide]: target.md\n\n[Guide]\n", ((1, "target.md"),)),
    ("- outer\n\n  - [Guide]: target.md\n\n[Guide]\n", ((3, "target.md"),)),
    ("- - outer\n\n    [Guide]: target.md\n\n[Guide]\n", ((3, "target.md"),)),
    ("12. outer\n\n    [Guide]: target.md\n\n[Guide]\n", ((3, "target.md"),)),
    ("- - [Guide]:\n    target.md\n\n[Guide]\n", ((1, "target.md"),)),
    ("- - [Gui\n    de]: target.md\n\n[Gui de]\n", ((1, "target.md"),)),
    ("- - paragraph\n    [Guide]: target.md\n\n[Guide]\n", ()),
    ("- - paragraph\n[Guide]: target.md\n\n[Guide]\n", ()),
    ("- -     [Guide]: target.md\n\n[Guide]\n", ()),
    ("    - - [Guide]: target.md\n\n[Guide]\n", ()),
    ("- - ```md\n    [Guide]: target.md\n    ```\n\n[Guide]\n", ()),
    ("- - <div>\n    [Guide]: target.md\n    </div>\n\n[Guide]\n", ()),
    ("`- - [Guide]: target.md`\n\n[Guide]\n", ()),
    ("- -[Guide]: target.md\n\n[Guide]\n", ()),
    ("\\- - [Guide]: target.md\n\n[Guide]\n", ()),
    ("- - -\n[Guide]: target.md\n\n[Guide]\n", ((2, "target.md"),)),
    ("- - [Guide]: target.md\n", ((1, "target.md"),)),
    ("- - [Guide](target.md)\n", ((1, "target.md"),)),
    ("> \t[Guide]: target.md\n\n[Guide]\n", ((1, "target.md"),)),
    (">\t[Guide]: target.md\n\n[Guide]\n", ((1, "target.md"),)),
    ("- - item\n\n\t[Guide]: target.md\n\n[Guide]\n", ((3, "target.md"),)),
    ("- - [Guide]: <folder\tname>\n\n[Guide]\n", ((1, "folder\tname"),)),
    ("- - [Guide]: target.md&#35;part\n\n[Guide]\n", ((1, "target.md&#35;part"),)),
    (
        "- - [Guide]: first.md\n    [GUIDE]: second.md\n\n[Guide]\n",
        ((1, "first.md"), (2, "second.md")),
    ),
)


@pytest.mark.parametrize(("body", "expected"), NESTED_DEFINITION_CASES)
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
def test_nested_definition_containers_keep_raw_targets_and_physical_lines(
    body: str, expected: tuple[tuple[int, str], ...], ending: str
) -> None:
    """Fixed CommonMark bodies distinguish real definitions from literal lookalikes."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_link_targets_from_text(body.replace("\n", ending)) == expected


@pytest.mark.parametrize(("body", "expected"), NESTED_DEFINITION_CASES[:30])
@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
def test_deployed_nested_definition_containers_have_fixed_native_outcomes(
    tmp_path: Path, body: str, expected: tuple[tuple[int, str], ...], ending: str
) -> None:
    """The no-sync CLI must reject real definitions and retain physical line reports."""
    reference_fixture(tmp_path, "markdown-relative-link")
    target = "docs/azure-devops-support.md"
    text = "Agents MUST validate.\nAgents MUST preserve authority.\n\n" + body.replace(
        "target.md", target
    )
    (tmp_path / "AGENTS.md").write_bytes(text.replace("\n", ending).encode("utf-8"))
    result = run(tmp_path)
    assert result.returncode == int(bool(expected)), result.stdout + result.stderr
    assert "Traceback" not in result.stderr
    if expected:
        assert "Stale protected-guide references" in result.stdout
        assert f"AGENTS.md:{expected[0][0] + 3}: azure-reference:" in result.stdout
    assert not (tmp_path / ".template-sync").exists()


@pytest.mark.parametrize(
    ("line", "column", "expected"),
    [
        ("-\t-\t[Guide]: <folder\tname>", 8, "[Guide]: <folder\tname>"),
        (">\t[Guide]: target.md", 2, "  [Guide]: target.md"),
        ("> \t[Guide]: target.md", 2, "  [Guide]: target.md"),
        ("    [Guide]: target.md", 4, "[Guide]: target.md"),
        ("-", 2, ""),
        ("- -", 4, ""),
        ("[Guide]: <folder\tname>", 0, "[Guide]: <folder\tname>"),
    ],
)
def test_nested_definition_source_columns_preserve_raw_content(
    line: str, column: int, expected: str
) -> None:
    """Only structural prefix columns are consumed or expanded; raw targets stay exact."""
    sys.path.insert(0, str(ROOT / ".github/scripts"))
    import instruction_contract_core as core

    assert core.markdown_source_content(line, column) == expected


@pytest.mark.parametrize("mutation", ["handoff", "raw-offset"])
def test_nested_definition_container_mutants_have_independent_native_oracles(
    tmp_path: Path, monkeypatch: Any, mutation: str
) -> None:
    """Removing context or confusing structural columns with indices recreates false passes."""
    monkeypatch.setenv("PYTHONDONTWRITEBYTECODE", "1")
    reference_fixture(tmp_path, "markdown-relative-link")
    prefix = "- - " if mutation == "handoff" else "-\t-\t"
    write(
        tmp_path,
        "AGENTS.md",
        "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
        + prefix
        + "[Guide]: docs/azure-devops-support.md\n\n[Guide]\n",
    )
    require_reference_failure(tmp_path)
    path = tmp_path / ".github/scripts/instruction_contract_core.py"
    source = path.read_text(encoding="utf-8")
    if mutation == "handoff":
        guard = (
            "            if content_columns is not None:\n"
            "                content = markdown_source_content(line, content_columns.get(line_number, 0))\n"
        )
        replacement = guard.replace("if content_columns is not None:", "if False:")
    else:
        guard = '    return " " * max(0, consumed - column) + line[offset:]\n'
        replacement = "    return line[column:]\n"
    assert source.count(guard) == 1
    path.write_text(source.replace(guard, replacement), encoding="utf-8", newline="\n")
    mutant = run(tmp_path)
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    assert "Traceback" not in mutant.stderr
    with pytest.raises(AssertionError):
        require_reference_failure(tmp_path)
    assert not list((path.parent / "__pycache__").glob("instruction_contract_core.*.pyc"))


@pytest.mark.parametrize("change", ["none", "content", "path", "key", "kind", "target", "retained"])
def test_nested_definition_exception_and_applicability_remain_exact(
    tmp_path: Path, change: str
) -> None:
    """Nested definitions use the same narrow raw identity and content hash as flat links."""
    document, _ = reference_fixture(tmp_path, "markdown-relative-link")
    target = "docs/azure-devops-support.md&#35;setup"
    text = (
        "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
        f"-\t-\t[Guide]: {target}\n\n[Guide]\n"
    )
    write(tmp_path, "AGENTS.md", text)
    require_reference_failure(tmp_path)
    declaration = {
        "path": "AGENTS.md",
        "anchor": f"reference:azure-reference:markdown-relative-link:{target}",
        "content_sha256": hashlib.sha256((tmp_path / "AGENTS.md").read_bytes()).hexdigest(),
        "reason": "Exact nested definition fixture",
        "authorization_basis": "Fixture declaration only",
    }
    if change == "content":
        write(tmp_path, "AGENTS.md", text + "Changed local content.\n")
    elif change == "path":
        declaration["path"] = "CLAUDE.md"
    elif change == "key":
        declaration["anchor"] = declaration["anchor"].replace("azure-reference", "other-key")
    elif change == "kind":
        declaration["anchor"] = declaration["anchor"].replace(
            "markdown-relative-link", "absolute-url"
        )
    elif change == "target":
        declaration["anchor"] = declaration["anchor"].replace("&#35;", "#")
    elif change == "retained":
        document["modules"].append("azure-devops-platform")
    document["exceptions"] = [] if change == "retained" else [declaration]
    write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
    result = run(tmp_path)
    assert result.returncode == (0 if change in {"none", "retained"} else 1), (
        result.stdout + result.stderr
    )
    assert "Traceback" not in result.stderr
    if change == "none":
        write(tmp_path, "AGENTS.md", text.replace("Agents MUST validate.\n", ""))
        declaration["content_sha256"] = hashlib.sha256(
            (tmp_path / "AGENTS.md").read_bytes()
        ).hexdigest()
        write(tmp_path, ".github/instruction-profile.yml", yaml.safe_dump(document))
        unrelated = run(tmp_path)
        assert unrelated.returncode == 1, unrelated.stdout + unrelated.stderr
        assert "missing required phrase" in unrelated.stdout


def test_nested_definition_migration_keeps_exact_exception_and_repeated_noop(
    tmp_path: Path,
) -> None:
    """A precise existing marker waiver remains a precise standalone declaration."""
    stage, target = tmp_path / "stage", tmp_path / "target"
    for root in (stage, target):
        document, _ = reference_fixture(root, "markdown-relative-link")
        write(
            root,
            "AGENTS.md",
            "Agents MUST validate.\nAgents MUST preserve authority.\n\n"
            "- - [Guide]: docs/azure-devops-support.md\n\n[Guide]\n",
        )
    marker = {
        "template_sync": {
            "included_modules": document["modules"],
            "protected_guide_contract_waivers": [
                {
                    "path": "AGENTS.md",
                    "contract_key": "azure-reference",
                    "target_path": "docs/azure-devops-support.md",
                    "reason": "Exact nested definition fixture",
                    "authorization_basis": "Fixture declaration only",
                }
            ],
        }
    }
    migrated = run_migration_schema_control(stage, target, marker_document=marker)
    assert migrated.returncode == 0, migrated.stdout + migrated.stderr
    path = stage / ".github/instruction-profile.yml"
    before = path.read_bytes()
    declarations = yaml.safe_load(before)["exceptions"]
    assert len(declarations) == 1
    assert declarations[0]["anchor"] == (
        "reference:azure-reference:markdown-relative-link:docs/azure-devops-support.md"
    )
    assert (
        declarations[0]["content_sha256"]
        == hashlib.sha256((stage / "AGENTS.md").read_bytes()).hexdigest()
    )
    accepted = run(stage)
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    write(target, ".github/instruction-profile.yml", before.decode("utf-8"))
    repeated = run_migration_schema_control(stage, target, marker_document=marker)
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert path.read_bytes() == before


def test_nested_definition_container_scanning_remains_bounded() -> None:
    """Deep and repeated container definitions preserve fixed inventories near the byte limit."""
    program = (
        "import sys; "
        f"sys.path.insert(0, {str(ROOT / '.github/scripts')!r}); "
        "import instruction_contract_core as c; "
        "body='- '*450000+'[Guide]: target.md\\n\\n[Guide]\\n'; "
        "assert len(body.encode()) < c.MAXIMUM_INPUT_BYTES; "
        "assert c.markdown_link_targets_from_text(body)==((1,'target.md'),); "
        "body='- - item\\n\\n    [Guide]: target.md\\n\\n'*20000; "
        "assert len(body.encode()) < c.MAXIMUM_INPUT_BYTES; "
        "targets=c.markdown_link_targets_from_text(body); "
        "assert len(targets)==20000 and targets[0]==(3,'target.md') "
        "and targets[-1]==(79999,'target.md')"
    )
    result = subprocess.run(
        [sys.executable, "-B", "-c", program],
        capture_output=True,
        text=True,
        check=False,
        timeout=20,
    )
    assert result.returncode == 0, result.stdout + result.stderr
