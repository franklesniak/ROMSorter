"""Validate bounded, explicitly owned GitHub workflow security contracts."""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import sys
import urllib.request
from collections.abc import Callable
from pathlib import Path, PurePosixPath
from typing import Any

import jsonschema
import yaml  # type: ignore[import-untyped]

_SHARED_SCRIPTS = str(Path(__file__).resolve().parent)
if _SHARED_SCRIPTS not in sys.path:
    sys.path.insert(0, _SHARED_SCRIPTS)

from instruction_contract_support import (  # noqa: E402
    LIST_MARKER_RE,
    MARKDOWN_FENCE_CONTEXT,
    MarkdownFence,
    active_fence_content,
    consume_blockquote_prefix,
    line_is_outside_list_item,
    parse_fence_close_from_content,
    parse_markdown_fence_open,
)

CONTRACT = ".github/workflow-security-contract.yml"
SCHEMA = "schemas/workflow-security-contract.schema.json"
LIMIT = 1024 * 1024
WORKFLOW = re.compile(r"\.github/workflows/[A-Za-z0-9_.-]+\.ya?ml\Z")
REFERENCE = re.compile(r"([A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+)(?:/[A-Za-z0-9_./-]+)?@([0-9a-f]{40})\Z")
ANNOTATION = re.compile(r"# (v[0-9]+\.[0-9]+\.[0-9]+)\s*\Z")
USES_LINE = re.compile(r"^\s*(?:#\s*)?(?:-\s*)?uses:\s*(\S+)")
MARKDOWN_QUOTE_PREFIX = re.compile(r"[ \t]*(?:(?:[-+*]|[0-9]{1,9}[.)])[ \t]+)*>[ \t]?")
EXAMPLE_FENCE = re.compile(r"^[ \t]*(?:(?:[-+*]|[0-9]{1,9}[.)])[ \t]+)?(`{3,}|~{3,})(.*)$")
EXAMPLE_ACTION = re.compile(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_./-]+@")
CONTROLS = (
    "name",
    "id",
    "if",
    "shell",
    "continue-on-error",
    "working-directory",
    "env",
    "with",
    "timeout-minutes",
    "background",
    "wait",
    "wait-all",
    "cancel",
)


class PolicyError(ValueError):
    """A policy input is unsafe, invalid, or inconsistent with the contract."""


class WorkflowLoader(yaml.SafeLoader):
    """Use YAML 1.2 booleans, unique keys, and bounded non-aliased trees."""

    yaml_implicit_resolvers = copy.deepcopy(yaml.SafeLoader.yaml_implicit_resolvers)
    for _key, _resolvers in yaml_implicit_resolvers.items():
        yaml_implicit_resolvers[_key] = [
            pair for pair in _resolvers if pair[0] != "tag:yaml.org,2002:bool"
        ]

    def compose_node(self, parent: Any, index: Any) -> Any:
        """Reject aliases, anchors, and excessive nesting before constructing values."""
        event = self.peek_event()
        if isinstance(event, yaml.AliasEvent) or getattr(event, "anchor", None):
            raise PolicyError("YAML anchors and aliases are not supported")
        depth = getattr(self, "policy_depth", 0)
        if depth >= 64:
            raise PolicyError("YAML nesting exceeds 64 levels")
        self.policy_depth = depth + 1
        try:
            return super().compose_node(parent, index)
        finally:
            self.policy_depth = depth

    def construct_mapping(self, node: Any, deep: bool = False) -> dict[Any, Any]:
        """Reject duplicate and non-string mapping keys, including merge keys."""
        result: dict[Any, Any] = {}
        for key_node, value_node in node.value:
            key = self.construct_object(key_node, deep=deep)
            if not isinstance(key, str) or key in result:
                raise PolicyError(f"Duplicate or non-string YAML key: {key!r}")
            result[key] = self.construct_object(value_node, deep=deep)
        return result


WorkflowLoader.add_implicit_resolver(
    "tag:yaml.org,2002:bool", re.compile(r"^(?:true|false|True|False|TRUE|FALSE)$"), list("tTfF")
)


def parse_yaml(text: str) -> dict[str, Any]:
    """Parse one bounded mapping without executing custom YAML tags."""
    if len(text.encode("utf-8")) > LIMIT:
        raise PolicyError("Input exceeds 1 MiB")
    try:
        document = yaml.load(text, Loader=WorkflowLoader)
    except yaml.YAMLError as error:
        raise PolicyError(f"Invalid YAML: {error}") from error
    if not isinstance(document, dict):
        raise PolicyError("Expected a YAML mapping")
    return document


def read_text(root: Path, relative: str) -> str:
    """Read a bounded regular UTF-8 file without following repository symlinks."""
    path = PurePosixPath(relative)
    if (
        not relative
        or "\\" in relative
        or ":" in relative
        or path.is_absolute()
        or any(part in {"", ".", ".."} for part in relative.split("/"))
    ):
        raise PolicyError(f"Unsafe repository path: {relative}")
    if root.is_symlink():
        raise PolicyError("Symlink repository root")
    candidate = root
    for part in path.parts:
        candidate = candidate / part
        if candidate.is_symlink():
            raise PolicyError(f"Symlink input: {relative}")
    if not candidate.is_file():
        raise PolicyError(f"Missing regular input: {relative}")
    with candidate.open("rb") as stream:
        data = stream.read(LIMIT + 1)
    if len(data) > LIMIT:
        raise PolicyError(f"Input exceeds 1 MiB: {relative}")
    return data.decode("utf-8").replace("\r\n", "\n")


def release_commit(repository: str, release: str) -> str:
    """Resolve a public upstream release with bounded, credential-free HTTPS reads."""
    url = f"https://api.github.com/repos/{repository}/git/ref/tags/{release}"
    for _attempt in range(5):
        request = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json"})
        with urllib.request.urlopen(request, timeout=20) as response:
            data = response.read(LIMIT + 1)
        if len(data) > LIMIT:
            raise PolicyError("Upstream release response exceeds 1 MiB")
        obj = json.loads(data)["object"]
        if obj["type"] == "commit" and re.fullmatch(r"[0-9a-f]{40}", obj["sha"]):
            return str(obj["sha"])
        if obj["type"] != "tag" or not re.fullmatch(r"[0-9a-f]{40}", obj["sha"]):
            break
        url = f"https://api.github.com/repos/{repository}/git/tags/{obj['sha']}"
    raise PolicyError(f"Cannot resolve upstream release: {repository} {release}")


def check_reference(
    reference: str, line: str, resolver: Callable[[str, str], str] | None = None
) -> None:
    """Require an immutable external reference and an exact release annotation."""
    match = REFERENCE.fullmatch(reference)
    annotation = ANNOTATION.search(line)
    if not match or not annotation:
        raise PolicyError(
            f"Expected external full SHA and same-line release annotation # vMAJOR.MINOR.PATCH: {reference}"
        )
    if any(part in {".", ".."} for part in reference.split("@", 1)[0].split("/")):
        raise PolicyError("Unsafe action subpath")
    if resolver and resolver(match[1], annotation[1]) != match[2]:
        raise PolicyError(f"Misleading release annotation: {reference} {annotation[1]}")


def describe_workflow(document: dict[str, Any]) -> dict[str, Any]:
    """Summarize reviewed execution controls without mirroring action identities."""
    jobs: dict[str, Any] = {}
    for key, job in document["jobs"].items():
        steps = []
        for step in job.get("steps", []):
            record = {field: step[field] for field in CONTROLS if field in step}
            if "uses" in step:
                record["action"] = step["uses"].split("@", 1)[0]
            if "run" in step:
                record["run_sha256"] = hashlib.sha256(
                    step["run"].replace("\r\n", "\n").encode("utf-8")
                ).hexdigest()
            steps.append(record)
        jobs[key] = {
            "controls": {
                field: job[field]
                for field in (
                    "name",
                    "concurrency",
                    "cache-mode",
                    "outputs",
                    "snapshot",
                    "permissions",
                    "if",
                    "continue-on-error",
                    "needs",
                    "defaults",
                    "env",
                    "runs-on",
                    "strategy",
                    "container",
                    "services",
                    "environment",
                    "timeout-minutes",
                    "with",
                    "secrets",
                )
                if field in job
            },
            "steps": steps,
        }
        if "uses" in job:
            jobs[key]["controls"]["action"] = job["uses"].split("@", 1)[0]
    return {
        "events": document["on"],
        "permissions": document["permissions"],
        "defaults": document.get("defaults", {}),
        "env": document.get("env", {}),
        **{field: document[field] for field in ("concurrency",) if field in document},
        **{field: document[field] for field in ("cache-mode",) if field in document},
        "jobs": jobs,
    }


def check_executable_annotations(text: str, resolver: Callable[[str, str], str] | None) -> Any:
    """Bind executable references to source lines and return the checked composition."""
    tree = yaml.compose(text, Loader=WorkflowLoader)
    lines = text.splitlines()

    def field(node: Any, name: str) -> Any:
        """Find a field in a composed mapping without constructing a second authority."""
        if not isinstance(node, yaml.MappingNode):
            return None
        return next((value for key, value in node.value if key.value == name), None)

    def check(node: Any) -> None:
        """Reject folded, flow-style, or quoted references lacking the selected literal form."""
        value = field(node, "uses")
        if value is None:
            return
        if not isinstance(value, yaml.ScalarNode):
            raise PolicyError("Action reference must be a scalar")
        line = lines[value.start_mark.line]
        match = USES_LINE.match(line)
        if line.lstrip().startswith("#") or not match or match[1] != value.value:
            raise PolicyError("Executable uses must have its own literal annotated uses line")
        check_reference(value.value, line, resolver)

    jobs = field(tree, "jobs")
    if not isinstance(jobs, yaml.MappingNode):
        raise PolicyError("Workflow must declare jobs")
    for _key, job in jobs.value:
        check(job)
        steps = field(job, "steps")
        if steps is not None:
            if not isinstance(steps, yaml.SequenceNode):
                raise PolicyError("Steps must be a sequence")
            for step in steps.value:
                check(step)
    return tree


def workflow_comment_lines(tree: Any, lines: list[str]) -> list[str]:
    """Separate real YAML comments from scalar content using the checked source marks."""
    pending = [tree]
    seen: set[int] = set()
    scalar_lines: set[int] = set()
    while pending:
        node = pending.pop()
        if id(node) in seen:
            continue
        seen.add(id(node))
        if isinstance(node, yaml.ScalarNode):
            # Column-zero end marks precede the next line; quoted closing lines
            # with a nonzero end column still belong to the scalar.
            end = node.end_mark.line + bool(node.end_mark.column)
            scalar_lines.update(range(node.start_mark.line, end))
        elif isinstance(node, yaml.MappingNode):
            for key, value in node.value:
                pending.extend((key, value))
        elif isinstance(node, yaml.SequenceNode):
            pending.extend(node.value)
    return ["" if number in scalar_lines else line for number, line in enumerate(lines)]


def validate_workflow(
    text: str, resolver: Callable[[str, str], str] | None = None
) -> dict[str, Any]:
    """Enforce universal rules before comparing reviewed required execution controls."""
    document = parse_yaml(text)
    tree = check_executable_annotations(text, resolver)
    events = document.get("on")
    names = {events} if isinstance(events, str) else set(events or [])
    if not names or names & {"pull_request_target", "workflow_run"}:
        raise PolicyError("Missing triggers or unsupported privileged workflow event")
    if document.get("permissions") not in ({}, {"contents": "read"}):
        raise PolicyError("Workflow permissions must be explicit and read-only")
    jobs = document.get("jobs")
    if not isinstance(jobs, dict) or not jobs:
        raise PolicyError("Workflow must declare jobs")
    for job in jobs.values():
        if not isinstance(job, dict) or job.get("permissions") not in ({}, {"contents": "read"}):
            raise PolicyError("Job permissions must be explicit and read-only")
        for step in job.get("steps", []):
            if not isinstance(step, dict):
                raise PolicyError("Workflow step must be a mapping")
            if "parallel" in step:
                raise PolicyError("Parallel step groups require recursive workflow policy support")
            reference = step.get("uses")
            if (
                reference
                and reference.split("@", 1)[0].casefold() == "actions/checkout"
                and step.get("with", {}).get("persist-credentials") is not False
            ):
                raise PolicyError("Checkout must set persist-credentials: false")
    lines = workflow_comment_lines(tree, text.splitlines())
    check_commented_examples(lines, resolver)
    return describe_workflow(document)


def load_contract(root: Path, *, schema_root: Path | None = None) -> dict[str, Any]:
    """Validate source contract data with an explicitly selected schema authority."""
    contract = parse_yaml(read_text(root, CONTRACT))
    schema = parse_yaml(read_text(schema_root if schema_root is not None else root, SCHEMA))
    jsonschema.Draft202012Validator.check_schema(schema)
    pending: list[Any] = [schema]
    while pending:
        node = pending.pop()
        if isinstance(node, dict):
            for key, value in node.items():
                if key in {"$ref", "$dynamicRef"} and not value.startswith("#"):
                    raise PolicyError("Only local schema references are supported")
                pending.append(value)
        elif isinstance(node, list):
            pending.extend(node)
    jsonschema.Draft202012Validator(schema).validate(contract)
    return contract


def markdown_example_content(line: str) -> str:
    """Peel quote containers without repeatedly copying the remaining input."""
    offset = 0
    while (match := MARKDOWN_QUOTE_PREFIX.match(line, offset)) is not None:
        offset = match.end()
    return line[offset:]


def example_fence_open(line: str) -> tuple[MarkdownFence, list[tuple[str, int]], int] | None:
    """Recognize a fence while retaining ordered quote/list continuation requirements."""
    containers: list[tuple[str, int]] = []
    last_quote = -1
    offset = 0
    while offset < len(line):
        depth, consumed = consume_blockquote_prefix(line[offset : offset + 5], max_depth=1)
        if depth:
            containers.append(("quote", 1))
            last_quote = len(containers) - 1
            offset += consumed
            continue
        # The shared grammar's longest prefix is 3 spaces + 9 digits + delimiter
        # + 4 spaces. One more character suffices without copying a long suffix.
        match = LIST_MARKER_RE.match(line[offset : offset + 18])
        if match is None:
            opened = parse_markdown_fence_open(line[offset:], MARKDOWN_FENCE_CONTEXT)
            return (opened, containers, last_quote) if opened is not None else None
        indent = (
            len(match.group("indent")) + len(match.group("marker")) + len(match.group("spaces"))
        )
        containers.append(("list", indent))
        offset += indent
    return None


def example_fence_content(
    line: str, fence: MarkdownFence, containers: list[tuple[str, int]], last_quote: int
) -> str | None:
    """Stop a literal range before a line leaves any ordered containing block."""
    if not line.strip(" \t"):
        return None if last_quote >= 0 else ""
    offset = 0
    for index, (kind, width) in enumerate(containers):
        if offset >= len(line):
            return None if last_quote >= index else ""
        if kind == "quote":
            depth, consumed = consume_blockquote_prefix(line[offset : offset + 5], max_depth=width)
            if depth < width:
                return None
            offset += consumed
        else:
            prefix = line[offset : offset + width]
            if prefix != " " * width:
                if line_is_outside_list_item(line[offset:], width):
                    return None
                return None if last_quote > index else ""
            offset += width
    return active_fence_content(line[offset:], fence)


def opaque_example_lines(lines: list[str]) -> set[int]:
    """Exclude non-YAML literals without granting invalid or ended fences opacity."""
    opaque: set[int] = set()
    active: tuple[MarkdownFence, list[tuple[str, int]], int] | None = None
    is_opaque = False
    for number, line in enumerate(lines):
        if active is not None:
            fence, containers, last_quote = active
            content = example_fence_content(line, fence, containers, last_quote)
            if content is not None:
                if is_opaque:
                    opaque.add(number)
                if parse_fence_close_from_content(
                    content,
                    fence_character=fence.character,
                    minimum_length=fence.length,
                    allow_arbitrary_indent=False,
                ):
                    active = None
                continue
            active = None
        active = example_fence_open(line)
        if active is not None:
            info = active[0].info.strip().lower().split(maxsplit=1)
            is_opaque = bool(info) and info[0] not in {"yaml", "yml"}
            if is_opaque:
                opaque.add(number)
    return opaque


class ExampleLoader(yaml.BaseLoader):
    """Compose inert documentation nodes with bounded depth and no object construction."""

    def compose_node(self, parent: Any, index: Any) -> Any:
        """Bound representation depth before composing potentially nested examples."""
        depth = getattr(self, "policy_depth", 0)
        if depth >= 64:
            raise PolicyError("Example YAML nesting exceeds 64 levels")
        self.policy_depth = depth + 1
        try:
            return super().compose_node(parent, index)
        finally:
            self.policy_depth = depth


def example_references(text: str) -> list[tuple[int, str]]:
    """Find decoded uses mappings without constructing tagged objects or following cycles."""
    tree = yaml.compose(text, Loader=ExampleLoader)
    pending = [tree]
    seen: set[int] = set()
    references: list[tuple[int, str]] = []
    used_lines: set[int] = set()
    while pending:
        node = pending.pop()
        if id(node) in seen:
            continue
        seen.add(id(node))
        if isinstance(node, yaml.MappingNode):
            for key, value in node.value:
                if isinstance(key, yaml.ScalarNode) and key.value == "uses":
                    if (
                        not isinstance(value, yaml.ScalarNode)
                        or not value.value
                        or key.start_mark.line != value.start_mark.line
                        or value.start_mark.line != value.end_mark.line
                        or key.start_mark.line != key.end_mark.line
                    ):
                        raise PolicyError(
                            "Example uses needs a scalar and annotation on one physical line"
                        )
                    line = value.start_mark.line
                    if line in used_lines:
                        raise PolicyError(
                            "Multiple example uses on one line have an ambiguous annotation"
                        )
                    used_lines.add(line)
                    references.append((line, value.value))
                pending.extend((key, value))
        elif isinstance(node, yaml.SequenceNode):
            pending.extend(node.value)
    return references


def check_examples(text: str, resolver: Callable[[str, str], str] | None = None) -> int:
    """Validate examples and return the number of governed action references."""
    raw_lines = text.splitlines()
    opaque_lines = opaque_example_lines(raw_lines)
    lines = [markdown_example_content(line) for line in raw_lines]
    references: set[tuple[int, str]] = set()
    fenced_lines: set[int] = set()
    fence: tuple[MarkdownFence, list[tuple[str, int]], int] | None = None
    block_start = 0
    indented_compatibility = False
    governed = False

    def collect_block(start: int, end: int, *, unclosed: bool = False) -> None:
        """Give complete YAML blocks authority over misleading isolated flow fragments."""
        block_lines = lines[start:end]
        if all(not line.strip() or line.lstrip().startswith("#") for line in block_lines):
            block_lines = [re.sub(r"^(\s*)#\s?", r"\1", line, count=1) for line in block_lines]
        block = "\n".join(block_lines)
        try:
            found = example_references(block)
        except yaml.YAMLError as error:
            if EXAMPLE_ACTION.search(block):
                raise PolicyError("Invalid YAML in an action-containing example fence") from error
            found = []
        if unclosed and (found or EXAMPLE_ACTION.search(block)):
            raise PolicyError("Unclosed action-containing example fence")
        references.update((start + offset, ref) for offset, ref in found)
        fenced_lines.update(range(start, end))

    for number, raw_line in enumerate(raw_lines):
        if number in opaque_lines:
            continue
        if fence is not None:
            opened, containers, last_quote = fence
            content = (
                lines[number]
                if indented_compatibility
                else example_fence_content(raw_line, opened, containers, last_quote)
            )
            if content is not None:
                if parse_fence_close_from_content(
                    content,
                    fence_character=opened.character,
                    minimum_length=opened.length,
                    allow_arbitrary_indent=indented_compatibility,
                ):
                    if governed:
                        collect_block(block_start, number)
                        fenced_lines.add(number)
                    fence = None
                elif governed:
                    # Removing containers must not change the physical line index.
                    lines[number] = markdown_example_content(content)
                continue
            if governed:
                collect_block(block_start, number, unclosed=True)
            fence = None
        opened_candidate = example_fence_open(raw_line)
        indented_compatibility = False
        if opened_candidate is None and EXAMPLE_FENCE.match(lines[number]):
            # Keep prior governed coverage for fences after indented list prose.
            # This compatibility branch does not grant non-YAML literal opacity.
            opened_candidate = example_fence_open(lines[number].lstrip(" \t"))
            indented_compatibility = opened_candidate is not None
        if opened_candidate is not None:
            info = opened_candidate[0].info.strip().lower().split(maxsplit=1)
            governed = not info or info[0] in {"yaml", "yml"}
            fence = opened_candidate
            block_start = number + 1
            if governed:
                fenced_lines.add(number)
    if fence is not None and governed:
        collect_block(block_start, len(lines), unclosed=True)
    block_reference_lines = {number for number, _reference in references}
    for number, candidate in enumerate(lines):
        if number in opaque_lines:
            continue
        if number in fenced_lines and (
            number in block_reference_lines or not candidate.lstrip().startswith("#")
        ):
            continue
        fragment = re.sub(r"^(\s*)#\s?", r"\1", candidate, count=1)
        try:
            references.update(
                (number + offset, ref) for offset, ref in example_references(fragment)
            )
        except yaml.YAMLError:
            pass  # Complete fences cover multiline YAML; surrounding prose is not YAML.
    semantic_lines = {number for number, _reference in references}
    for number, candidate in enumerate(lines):
        if (
            number not in opaque_lines
            and number not in semantic_lines
            and (match := USES_LINE.match(candidate))
        ):
            references.add((number, match[1]))
    for number, reference in sorted(references):
        check_reference(reference, lines[number], resolver)
    return len(references)


def check_commented_examples(
    lines: list[str], resolver: Callable[[str, str], str] | None = None
) -> None:
    """Validate contiguous commented YAML fragments without changing executable binding."""
    group: list[str] = []
    for line in [*lines, ""]:
        stripped = line.lstrip()
        if stripped.startswith("#") and stripped[1:].strip():
            group.append(stripped[1:])
        elif group:
            check_examples("```yaml\n" + "\n".join(group) + "\n```", resolver)
            group = []


def validate_repository(
    root: Path,
    *,
    strict: bool = False,
    verify_releases: bool = False,
    schema_root: Path | None = None,
) -> int:
    """Validate owned inputs, optionally including adopter workflows and upstream refs."""
    contract = load_contract(root, schema_root=schema_root)
    cache: dict[tuple[str, str], str] = {}

    def resolve(repository: str, release: str) -> str:
        """Avoid repeated public ref lookups in one verification run."""
        key = (repository, release)
        if key not in cache:
            cache[key] = release_commit(*key)
        return cache[key]

    resolver = resolve if verify_releases else None
    for path, expected in contract["workflows"].items():
        if not WORKFLOW.fullmatch(path):
            raise PolicyError(f"Out-of-scope workflow path: {path}")
        actual = validate_workflow(read_text(root, path), resolver)
        if actual != expected:
            raise PolicyError(
                f"Required workflow execution controls changed: {path}; review contract"
            )
    for path in contract["examples"]:
        text = read_text(root, path)
        check_examples(text, resolver)
    if strict:
        for path in (root / ".github/workflows").glob("*"):
            if path.suffix in {".yml", ".yaml"}:
                validate_workflow(read_text(root, path.relative_to(root).as_posix()), resolver)
    return len(contract["workflows"])


def main(argv: list[str] | None = None) -> int:
    """Return a nonzero native exit for invalid policy or unavailable verification."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo-root", type=Path, default=Path.cwd())
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--verify-releases", action="store_true")
    args = parser.parse_args(argv)
    try:
        count = validate_repository(
            args.repo_root.resolve(), strict=args.strict, verify_releases=args.verify_releases
        )
    except OSError as error:
        summary = f"{type(error).__name__}: {error.strerror or 'I/O error'}"
        print(f"Workflow security validation failed: {summary}", file=sys.stderr)
        return 1
    except (
        PolicyError,
        UnicodeError,
        jsonschema.ValidationError,
        jsonschema.SchemaError,
        ValueError,
        KeyError,
        TypeError,
    ) as error:
        print(f"Workflow security validation failed: {error}", file=sys.stderr)
        return 1
    print(f"Workflow security: {count} owned workflows passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
