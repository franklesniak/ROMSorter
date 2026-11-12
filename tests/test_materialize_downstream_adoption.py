"""Exercise deterministic first-adoption materialization."""

from __future__ import annotations

import contextlib
import fnmatch
import importlib
import importlib.util
import io
import json
import os
import re
import shutil
import subprocess
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

import yaml  # type: ignore[import-untyped]

from tests._pytest_compat import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_PATH = REPO_ROOT / ".template-sync" / "scripts" / "materialize_downstream_adoption.py"
REPORT_SCRIPT_PATH = (
    REPO_ROOT / ".template-sync" / "scripts" / "report_excluded_module_references.py"
)
DOWNSTREAM_ADOPTION_SCRIPT_PATH = (
    REPO_ROOT / ".template-sync" / "scripts" / "validate_downstream_adoption.py"
)
SCRIPT_DIR = SCRIPT_PATH.parent
NESTED_MARKDOWN_LINT_PATH = REPO_ROOT / ".github" / "scripts" / "lint-nested-markdown.js"
SOURCE_REPO = "https://github.com/franklesniak/copilot-repo-template.git"
FULL_SHA = "0123456789abcdef0123456789abcdef01234567"
INSTRUCTION_CONTRACTS_PATH = ".template-sync/instruction-contracts.yml"
UPSTREAM_STYLE_GUIDES_PATH = "docs/upstream-style-guides.md"
POWERSHELL_GUIDE_PATH = ".github/instructions/powershell.instructions.md"
TERRAFORM_GUIDE_PATH = ".github/instructions/terraform.instructions.md"
ISSUE_692_NO_PYTHON_MODULES = (
    "baseline",
    "agent-instructions",
    "github-platform",
    "github-actions",
    "github-templates",
    "template-sync-support",
    "markdown",
    "powershell",
)
AGENT_PLATFORM_MODULES = (
    "agent-copilot",
    "agent-codex",
    "agent-claude",
    "agent-cursor",
    "agent-gemini",
    "agent-hermes",
)
GITHUB_POWERSHELL_PROFILE_MODULES = (
    *ISSUE_692_NO_PYTHON_MODULES,
    *AGENT_PLATFORM_MODULES,
)
ISSUE_693_PARTIAL_DOC_MODULES = ISSUE_692_NO_PYTHON_MODULES
FULL_TEMPLATE_MODULES = (
    "baseline",
    "git-lfs",
    "agent-instructions",
    "instruction-enforcement",
    *AGENT_PLATFORM_MODULES,
    "github-platform",
    "azure-devops-platform",
    "github-actions",
    "azure-pipelines",
    "github-templates",
    "azure-devops-collaboration",
    "template-onboarding",
    "template-sync-support",
    "markdown",
    "powershell",
    "json",
    "yaml",
    "schema",
    "python",
    "terraform",
)
StructuredObject = dict[str, Any]
AZURE_TEMPLATE_MODULES = frozenset(
    ("azure-devops-platform", "azure-pipelines", "azure-devops-collaboration")
)
GITHUB_HOST_TEMPLATE_MODULES = frozenset(("github-platform", "github-actions", "github-templates"))
DOWNSTREAM_PYTEST_MODULES = tuple(
    module_name
    for module_name in FULL_TEMPLATE_MODULES
    if module_name
    not in {
        "agent-instructions",
        "instruction-enforcement",
        *AGENT_PLATFORM_MODULES,
        "git-lfs",
        "powershell",
        "terraform",
    }
)
OPTIONAL_PRUNING_FIXTURES: tuple[Any, ...] = (
    pytest.param(
        ("baseline", "python", "schema", "template-sync-support"),
        True,
        id="python-schema-template-sync-without-terraform",
    ),
    pytest.param(
        ("baseline", "template-sync-support"),
        True,
        id="template-sync-support-without-powershell",
    ),
    pytest.param(
        FULL_TEMPLATE_MODULES,
        True,
        id="full-upstream-module-set",
        marks=pytest.mark.upstream_template_only,
    ),
    pytest.param(
        DOWNSTREAM_PYTEST_MODULES,
        True,
        id="complete-downstream-pytest-module-set",
    ),
)
AZURE_PROVIDER_BASE_FIELDS: dict[str, str] = {
    "azure_devops_organization": "contoso",
    "azure_devops_project": "Template Adoption",
    "azure_devops_repository": "downstream-template",
    "azure_boards_policy": "manual-follow-up",
    "azure_repos_pr_template_policy": "materialize",
    "azure_branch_policy_reviewer_guidance": "manual-follow-up",
    "azure_security_intake_policy": "manual-follow-up",
    "azure_security_product_enablement": "none",
    "azure_dependency_update_policy": "manual-follow-up",
}
BASELINE_GITATTRIBUTES_LF_PIN_PATHS = (
    "example.ps1",
    "settings.psd1",
    "module.psm1",
    "main.tf",
    "terraform.tfvars",
    "example.tftpl",
    "backend.tfbackend",
)
GIT_ATTRIBUTES_TO_CHECK = ("text", "eol", "filter", "diff", "merge")
GIT_LFS_ATTRIBUTES_TO_CHECK = ("filter", "diff", "merge", "text")
GIT_LFS_MANAGED_ATTRIBUTE_PATHS = (
    "assets/source/poster.psd",
    "assets/source/poster.psb",
    "assets/source/identity.ai",
    "assets/source/layout.indd",
    "assets/source/wireframe.sketch",
    "assets/source/mockup.fig",
    "assets/source/scene.blend",
    "assets/source/model.fbx",
    "assets/source/plan.dwg",
)
OPTIONAL_STACK_OWNED_PATHS: dict[str, tuple[str, ...]] = {
    "powershell": (
        ".github/instructions/powershell.instructions.md",
        ".github/linting/PSScriptAnalyzerSettings.psd1",
        ".github/workflows/powershell-ci.yml",
        "src/tools/Resolve-PSScriptAnalyzerGate.ps1",
        "templates/powershell/Example.Tests.ps1",
    ),
    "terraform": (
        ".github/instructions/terraform.instructions.md",
        ".github/workflows/terraform-ci.yml",
        ".tflint.hcl",
        "docs/terraform/TERRAFORM_LINTING_GUIDE.md",
        "docs/terraform/TERRAFORM_TESTING_GUIDE.md",
        "templates/terraform/Example.tftest.hcl",
    ),
}
OPTIONAL_STACK_INLINE_MARKERS: dict[str, tuple[str, ...]] = {
    "powershell": ("powershell-reference-only",),
    "terraform": ("terraform-only", "terraform-reference-only"),
}
DEPENDABOT_BASELINE_NO_PYTHON_ECOSYSTEMS = {
    "github-actions",
    "npm",
    "pip",
    "pre-commit",
}
DEPENDABOT_FULL_ECOSYSTEMS = DEPENDABOT_BASELINE_NO_PYTHON_ECOSYSTEMS
ISSUE_693_BASELINE_DOCS = ("README.md", "CONTRIBUTING.md")
# Base module set with no data-file modules (json, yaml, schema) and no
# template-sync-support, used to exercise the baseline+GitHub data-CI reference
# relation and the single-module template-sync-support-reference-only blocks.
NO_DATA_NO_TEMPLATE_SYNC_MODULES = (
    "baseline",
    "agent-instructions",
    "github-platform",
    "github-actions",
    "github-templates",
    "markdown",
    "powershell",
)
# References that exist only inside the template-sync-support-reference-only block
# in README.md.
TEMPLATE_SYNC_SUPPORT_README_REFERENCES = (
    "`.template-sync/`",
    "schemas/template-sync-",
    "validate_downstream_adoption.py",
)
ISSUE_693_EXCLUDED_DOC_REFERENCES: dict[str, tuple[str, ...]] = {
    "README.md": (
        "pyproject.toml",
        ".github/workflows/python-ci.yml",
        ".github/workflows/terraform-ci.yml",
        ".github/instructions/python.instructions.md",
        ".github/instructions/terraform.instructions.md",
        ".github/instructions/json.instructions.md",
        ".github/instructions/yaml.instructions.md",
        ".tflint.hcl",
        ".yamllint.yml",
        "templates/python/",
        "templates/terraform/",
        "templates/json/",
        "templates/yaml/",
        "schemas/README.md",
        "schemas/example-config",
        "docs/azure-devops-support.md",
        "pre-commit run yamllint --all-files",
        "terraform test -verbose",
        "TFLint",
    ),
    "CONTRIBUTING.md": (
        "Python Version Requirements",
        "pyproject.toml",
        ".github/workflows/python-ci.yml",
        ".github/workflows/terraform-ci.yml",
        ".github/instructions/python.instructions.md",
        ".github/instructions/terraform.instructions.md",
        ".github/instructions/json.instructions.md",
        ".github/instructions/yaml.instructions.md",
        ".tflint.hcl",
        ".yamllint.yml",
        "schemas/README.md",
        "schemas/example-config",
        "docs/azure-devops-support.md",
        'pytest tests/ -m "not slow" -v --cov --cov-report=term-missing',
        "mypy src tests",
        "terraform-fmt",
        "terraform-validate",
        "terraform-tflint",
        "terraform_version",
        "tflint_version",
        "TFLint",
    ),
}
ISSUE_693_RETAINED_DOC_REFERENCES: dict[str, tuple[str, ...]] = {
    "README.md": (
        "npm run lint:md",
        "Invoke-Pester -Path tests/ -Output Detailed",
        "python .template-sync/scripts/validate_downstream_adoption.py --require-marker",
        "`check-json`",
        "`check-yaml`",
        "`actionlint`",
        "`check-jsonschema`",
        "`check-metaschema`",
    ),
    "CONTRIBUTING.md": (
        "pre-commit run --all-files",
        "a working Python 3 interpreter is required",
        "npm run lint:md",
        "Invoke-Pester -Path tests/ -Output Detailed",
        "python .template-sync/scripts/validate_downstream_adoption.py --require-marker",
        "`check-json`",
        "`check-yaml`",
        "`actionlint`",
        "`check-jsonschema`",
        "`check-metaschema`",
        "secrets",
    ),
}
NESTED_MARKDOWN_LINT_NODE_MODULES = (
    "glob",
    "jsonc-parser",
    "markdown-it",
    "markdownlint",
)

if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

materializer = cast(Any, importlib.import_module("materialize_downstream_adoption"))
excluded_reporter = cast(Any, importlib.import_module("report_excluded_module_references"))


def write_file(path: Path, content: str) -> None:
    """Write a UTF-8 fixture file with LF newlines on every platform.

    Newline translation is disabled so generated fixtures (notably the
    decisions.yml exercised by the yamllint assertions) keep LF endings on
    Windows and satisfy yamllint's unix new-lines rule, matching the
    repository's LF-normalized files.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")


class _IndentSequenceDumper(yaml.SafeDumper):
    """SafeDumper that indents block sequences under their mapping key.

    The repository ``.yamllint.yml`` enables ``indentation.indent-sequences``,
    so fixture YAML must nest sequence items one level deeper than the key that
    introduces them. The default PyYAML emitter writes indentless sequences,
    which would make generated fixtures such as ``decisions.yml`` violate the
    repository yamllint configuration.
    """

    def increase_indent(self, flow: bool = False, indentless: bool = False) -> None:
        """Force block sequences to use ordinary nested indentation."""
        super().increase_indent(flow=flow, indentless=False)

    def ignore_aliases(self, data: object) -> bool:
        """Keep generated fixture YAML free of anchors and aliases."""
        return True


def write_yaml(path: Path, data: dict[str, Any]) -> None:
    """Write a YAML fixture file using the repository's indent-sequence style."""
    write_file(
        path,
        yaml.dump(
            data,
            Dumper=_IndentSequenceDumper,
            sort_keys=False,
            default_flow_style=False,
            width=120,
        ),
    )


def write_json(path: Path, data: dict[str, Any]) -> None:
    """Write a strict JSON fixture file."""
    write_file(path, json.dumps(data, indent=2) + "\n")


def read_file(path: Path) -> str:
    """Read a UTF-8 fixture file."""
    return path.read_text(encoding="utf-8")


def snapshot_fixture_files(root: Path) -> dict[str, bytes]:
    """Capture stable fixture bytes while excluding Git and interpreter caches."""
    return {
        path.relative_to(root).as_posix(): path.read_bytes()
        for path in root.rglob("*")
        if path.is_file() and ".git" not in path.parts and "__pycache__" not in path.parts
    }


def run_in_process_cli(
    command: list[str],
    main: Callable[[list[str]], int],
) -> subprocess.CompletedProcess[str]:
    """Run an imported CLI main function and return a CompletedProcess-shaped result."""
    stdout = io.StringIO()
    stderr = io.StringIO()
    argv = command[2:]
    with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
        try:
            returncode = main(argv)
        except SystemExit as error:
            returncode = error.code if isinstance(error.code, int) else 1
    return subprocess.CompletedProcess(
        args=command,
        returncode=returncode,
        stdout=stdout.getvalue(),
        stderr=stderr.getvalue(),
    )


@pytest.mark.upstream_template_only
def test_full_template_fixture_matches_manifest_inventory() -> None:
    """Keep full-adoption coverage explicit and complete when modules are added."""
    _manifest, module_order, _mappings = materializer.load_validated_manifest_context(REPO_ROOT)

    assert len(FULL_TEMPLATE_MODULES) == len(set(FULL_TEMPLATE_MODULES))
    assert set(FULL_TEMPLATE_MODULES) == set(module_order)


def test_materializer_script_entrypoint_help_smoke() -> None:
    """The materializer remains executable through its script entry point."""
    result = subprocess.run(
        [sys.executable, str(SCRIPT_PATH), "--help"],
        check=False,
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    assert "usage: materialize_downstream_adoption.py" in result.stdout


def section_entries(output: str, heading: str) -> set[str]:
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


def copy_template_file(template_root: Path, relative_path: str) -> None:
    """Copy a repository template file into a fixture template root."""
    destination = template_root / relative_path
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(REPO_ROOT / relative_path, destination)


def nested_markdown_lint_prerequisites_available() -> bool:
    """Return whether the Node-based nested Markdown linter can run locally."""
    if shutil.which("node") is None:
        return False
    return all(
        (REPO_ROOT / "node_modules" / module_name).exists()
        for module_name in NESTED_MARKDOWN_LINT_NODE_MODULES
    )


def check_jsonschema_command() -> list[str] | None:
    """Resolve a ``check-jsonschema`` command for optional generated YAML validation."""
    executable = shutil.which("check-jsonschema")
    if executable is not None:
        return [executable]
    if importlib.util.find_spec("check_jsonschema") is not None:
        return [sys.executable, "-m", "check_jsonschema"]
    return None


def yamllint_command() -> list[str] | None:
    """Resolve a direct yamllint command for generated temporary files."""
    executable = shutil.which("yamllint")
    if executable is not None:
        return [executable]
    if importlib.util.find_spec("yamllint") is not None:
        return [sys.executable, "-m", "yamllint"]
    return None


def assert_yamllint_clean(*paths: Path) -> None:
    """Assert paths pass yamllint under the repository configuration."""
    command = yamllint_command()
    if command is None:
        pytest.skip("yamllint executable or importable module is required for this assertion")
    assert command is not None
    result = subprocess.run(
        [
            *command,
            "-c",
            str(REPO_ROOT / ".yamllint.yml"),
            *(str(path) for path in paths),
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def prepare_template(
    template_root: Path,
    path_mappings: list[dict[str, Any]],
    *,
    include_placeholder_helper: bool = False,
) -> None:
    """Create the minimal template-side contract files for a fixture run."""
    effective_path_mappings = [dict(mapping) for mapping in path_mappings]
    if include_placeholder_helper:
        placeholder_manifest = json.loads(
            (REPO_ROOT / ".github" / "template-placeholders.json").read_text(encoding="utf-8")
        )
        path_groups = placeholder_manifest["pathGroups"]
        placeholder_paths = {path for group_paths in path_groups.values() for path in group_paths}
        placeholder_paths.update(
            path
            for token_spec in placeholder_manifest.get("tokens", [])
            for path in token_spec.get("pathPatterns", [])
        )
        for relative_path in sorted(placeholder_paths):
            if any(
                fnmatch.fnmatchcase(relative_path, mapping["pattern"])
                for mapping in effective_path_mappings
            ):
                continue
            effective_path_mappings.append({"pattern": relative_path, "requires_all": ["baseline"]})

    manifest = {
        "template_manifest": {
            "version": 2,
            "modules": [
                {"name": "baseline", "description": "Baseline files."},
                {
                    "name": "agent-instructions",
                    "description": "Agent instruction files.",
                },
                {
                    "name": "template-sync-support",
                    "description": "Template sync support files.",
                },
                {
                    "name": "azure-devops-platform",
                    "description": "Azure DevOps platform files.",
                },
                {
                    "name": "azure-pipelines",
                    "description": "Azure Pipelines files.",
                },
                {
                    "name": "azure-devops-collaboration",
                    "description": "Azure DevOps collaboration files.",
                },
                {"name": "python", "description": "Python files."},
                {"name": "markdown", "description": "Markdown files."},
            ],
            "path_mappings": effective_path_mappings,
            "filtering": {
                "default_semantics": "AND",
                "requires_any_semantics": "OR",
                "path_matching": "most_specific_match_wins",
                "same_specificity_action": "union_modules",
                "unmapped_action": "surface_for_owner",
            },
            "notes": {
                "downstream_retention": "Downstream repositories keep marker data.",
            },
        }
    }
    write_yaml(template_root / ".template-sync" / "manifest.yml", manifest)
    schema_root = template_root / "schemas"
    schema_root.mkdir(parents=True)
    shutil.copyfile(
        REPO_ROOT / "schemas" / "template-sync-manifest.schema.json",
        schema_root / "template-sync-manifest.schema.json",
    )
    shutil.copyfile(
        REPO_ROOT / "schemas" / "template-sync-marker.schema.json",
        schema_root / "template-sync-marker.schema.json",
    )
    if include_placeholder_helper:
        helper_path = template_root / ".github" / "scripts" / "replace-template-placeholders.py"
        helper_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(
            REPO_ROOT / ".github" / "scripts" / "replace-template-placeholders.py",
            helper_path,
        )
        shutil.copyfile(
            REPO_ROOT / ".github" / "template-placeholders.json",
            template_root / ".github" / "template-placeholders.json",
        )
        shutil.copyfile(
            REPO_ROOT / "schemas" / "template-placeholders.schema.json",
            schema_root / "template-placeholders.schema.json",
        )


def prepare_security_reporting_template(template_root: Path) -> None:
    """Create a fixture template with the security-reporting placeholder surfaces."""
    prepare_template(
        template_root,
        [
            {"pattern": "SECURITY.md", "requires_all": ["baseline"]},
            {"pattern": "CODE_OF_CONDUCT.md", "requires_all": ["baseline"]},
            {"pattern": ".github/ISSUE_TEMPLATE/config.yml", "requires_all": ["baseline"]},
            {"pattern": ".github/ISSUE_TEMPLATE/bug_report.yml", "requires_all": ["baseline"]},
        ],
        include_placeholder_helper=True,
    )
    for relative_path in (
        "SECURITY.md",
        "CODE_OF_CONDUCT.md",
        ".github/ISSUE_TEMPLATE/config.yml",
        ".github/ISSUE_TEMPLATE/bug_report.yml",
    ):
        copy_template_file(template_root, relative_path)


def load_yaml(path: Path) -> object:
    """Load generated YAML after asserting it parses."""
    return yaml.safe_load(read_file(path))


def as_mapping(value: object, message: str) -> StructuredObject:
    """Return ``value`` after asserting it is a structured mapping."""
    assert isinstance(value, dict), message
    return cast(StructuredObject, value)


def as_list(value: object, message: str) -> list[Any]:
    """Return ``value`` after asserting it is a list."""
    assert isinstance(value, list), message
    return cast(list[Any], value)


def as_string_list(value: object, message: str) -> list[str]:
    """Return ``value`` after asserting it is a list of strings."""
    values = as_list(value, message)
    assert all(isinstance(item, str) for item in values), message
    return cast(list[str], values)


def dependabot_update_ecosystems(path: Path) -> set[str]:
    """Return Dependabot package ecosystems from a generated config."""
    document = as_mapping(load_yaml(path), "Dependabot config must be a mapping")
    updates = as_list(document.get("updates"), "Dependabot updates must be a list")
    ecosystems: set[str] = set()
    for update in updates:
        update_mapping = as_mapping(update, "Dependabot update must be a mapping")
        ecosystem = update_mapping.get("package-ecosystem")
        assert isinstance(ecosystem, str)
        ecosystems.add(ecosystem)
    return ecosystems


def validate_dependabot_vendor_schema(path: Path) -> None:
    """Validate generated Dependabot YAML against the built-in schema when available."""
    validator_command = check_jsonschema_command()
    if validator_command is None:
        return

    result = subprocess.run(
        [
            *validator_command,
            "--builtin-schema",
            "vendor.dependabot",
            str(path),
        ],
        check=False,
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stdout + result.stderr


def security_contact_link(config: object) -> dict[str, object]:
    """Return the security contact link from parsed issue-template config."""
    return contact_link_by_name(config, "Security Vulnerabilities")


def contact_link_by_name(config: object, name: str) -> dict[str, object]:
    """Return a named contact link from parsed issue-template config."""
    config_mapping = as_mapping(config, "issue-template config must be a mapping")
    contact_links = as_list(config_mapping["contact_links"], "contact_links must be a list")
    for contact_link in contact_links:
        contact_link_mapping = as_mapping(contact_link, "contact link must be a mapping")
        if contact_link_mapping.get("name") == name:
            return contact_link_mapping
    raise AssertionError(f"{name} contact link not found")


def assert_issue_form_shape(bug_report: object) -> None:
    """Assert the generated bug-report issue form has the expected structure."""
    bug_report_mapping = as_mapping(bug_report, "bug-report form must be a mapping")
    body = as_list(bug_report_mapping["body"], "bug-report body must be a list")
    assert body
    for item in body:
        item_mapping = as_mapping(item, "bug-report body item must be a mapping")
        assert isinstance(item_mapping.get("type"), str)
        assert isinstance(item_mapping.get("attributes"), dict)


def run_materialize(
    template_root: Path,
    target_root: Path,
    *args: str,
) -> subprocess.CompletedProcess[str]:
    """Run the materialization helper against fixture repositories."""
    command = [
        sys.executable,
        str(SCRIPT_PATH),
        "--template-root",
        str(template_root),
        "--target-root",
        str(target_root),
        *args,
    ]
    return run_in_process_cli(command, materializer.main)


def retained_protected_paths_for_modules(included_modules: tuple[str, ...]) -> tuple[str, ...]:
    """Return protected governance files retained by a materialized module set."""
    _manifest, _module_order, mappings = materializer.load_validated_manifest_context(REPO_ROOT)
    template_paths, skipped_symlinks = materializer.iter_safe_repository_files(REPO_ROOT)
    assert not skipped_symlinks, f"fixture source has skipped symlink(s): {skipped_symlinks}"

    protected_paths: list[str] = []
    for relative_path in template_paths:
        relation = materializer.selected_relation_for_path(relative_path, mappings)
        if relation is None or not relation.is_retained_by(included_modules):
            continue
        if materializer.is_protected_instruction_path(relative_path):
            protected_paths.append(relative_path)
    return tuple(sorted(protected_paths))


def protected_take_decisions_for_modules(
    included_modules: tuple[str, ...],
) -> list[dict[str, str]]:
    """Return marker decisions that allow a full protected-governance fixture."""
    return [
        {
            "path": relative_path,
            "decision": "TAKE",
            "adoption_mode": "minimal-preservation",
            "authorization_basis": f"Regression fixture authorizes taking {relative_path}.",
            "authorized_scope": f"{relative_path} only.",
        }
        for relative_path in retained_protected_paths_for_modules(included_modules)
    ]


def materialize_module_fixture(
    tmp_path: Path,
    included_modules: tuple[str, ...],
    *,
    authorize_protected_files: bool = False,
    authorize_workflow_contract: bool = True,
) -> Path:
    """Materialize a real downstream fixture from the requested module set."""
    target_root = tmp_path / "downstream"
    target_root.mkdir()
    marker_fields: dict[str, Any] = {}
    marker_fields.update(azure_provider_fields_for_modules(included_modules))
    if authorize_protected_files:
        marker_fields["protected_file_decisions"] = protected_take_decisions_for_modules(
            included_modules
        )
    elif authorize_workflow_contract and "github-actions" in included_modules:
        marker_fields["protected_file_decisions"] = [
            decision
            for decision in protected_take_decisions_for_modules(included_modules)
            if decision["path"] == ".github/workflow-security-contract.yml"
        ]
    write_yaml(
        target_root / "decisions.yml",
        marker_document(list(included_modules), **marker_fields),
    )
    placeholder_args = azure_provider_cli_args_for_modules(included_modules)

    result = run_materialize(
        REPO_ROOT,
        target_root,
        "--decisions-file",
        "decisions.yml",
        *placeholder_args,
    )

    assert result.returncode == 0, result.stdout + result.stderr
    if "template-sync-support" in included_modules:
        assert (
            authorize_protected_files
        ), "support-retaining fixtures must explicitly authorize the protected catalog"
        assert (target_root / INSTRUCTION_CONTRACTS_PATH).read_bytes() == (
            REPO_ROOT / INSTRUCTION_CONTRACTS_PATH
        ).read_bytes()
    return target_root


def finish_review_profile_link_cleanup(target: Path, profile: str) -> None:
    """Model the adopter's authorized link cleanup after materialization.

    Materialization deliberately reports manual cleanup separately. This fixture
    owner authorizes removal of links to excluded instruction and template docs only;
    link labels and all unrelated links remain. No validator waiver is added.
    """
    if profile not in {"neither", "no-markdown", "no-yaml"}:
        return
    excluded = {".github/instructions/yaml.instructions.md", "templates/yaml/README.md"}
    if profile == "no-markdown":
        excluded = {".github/instructions/docs.instructions.md"}
    if profile == "neither":
        excluded = {
            ".github/copilot-instructions.md",
            "COPILOT_CHAT_PROMPTS.md",
            *(
                f".github/instructions/{name}.instructions.md"
                for name in (
                    "docs",
                    "gitattributes",
                    "json",
                    "powershell",
                    "python",
                    "terraform",
                    "yaml",
                )
            ),
        }
    paths = subprocess.run(
        ["git", "ls-files", "-z", "--", "*.md"],
        cwd=target,
        check=True,
        capture_output=True,
        text=True,
    ).stdout.split("\0")
    for relative_path in filter(None, paths):
        document = target / relative_path
        assert not document.is_symlink()
        assert document.resolve().is_relative_to(target.resolve())

        def remove_excluded_link(match: re.Match[str], document: Path = document) -> str:
            link = match.group(2).split("#", 1)[0]
            if not link or ":" in link or link.startswith("/"):
                return match.group(0)
            resolved = (document.parent / link).resolve()
            if not resolved.is_relative_to(target.resolve()):
                return match.group(0)
            if resolved.relative_to(target.resolve()).as_posix() in excluded:
                assert not resolved.exists()
                return match.group(1)
            return match.group(0)

        original = read_file(document)
        cleaned = re.sub(r"\[([^\]\n]+)\]\(([^()\s]+)\)", remove_excluded_link, original)
        if cleaned != original:
            write_file(document, cleaned)


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    "profile",
    ["both", "codex-only", "claude-only", "neither", "no-markdown", "no-yaml"],
)
def test_materialized_review_governance_profiles(tmp_path: Path, profile: str) -> None:
    """Actual retained policy passes direct/aggregate CLIs and rejects a weakened gate."""
    modules = tuple(
        module
        for module in FULL_TEMPLATE_MODULES
        if not (
            profile == "neither"
            and module in {"agent-instructions", "instruction-enforcement", *AGENT_PLATFORM_MODULES}
        )
        and not (
            profile in {"codex-only", "claude-only"}
            and module in AGENT_PLATFORM_MODULES
            and module != f"agent-{profile.removesuffix('-only')}"
        )
        and not (profile == "no-markdown" and module == "markdown")
        and not (profile == "no-yaml" and module == "yaml")
    )
    target = tmp_path / "profile"
    target.mkdir()
    decisions = protected_take_decisions_for_modules(modules)
    removed = {"codex-only": "CLAUDE.md", "claude-only": "AGENTS.md"}.get(profile)
    fields: dict[str, Any] = azure_provider_fields_for_modules(modules)
    fields["protected_file_decisions"] = decisions
    write_yaml(target / "decisions.yml", marker_document(list(modules), **fields))
    result = run_materialize(
        REPO_ROOT,
        target,
        "--decisions-file",
        "decisions.yml",
        *azure_provider_cli_args_for_modules(modules),
    )
    assert result.returncode == 0, result.stdout + result.stderr
    catalog_path = target / INSTRUCTION_CONTRACTS_PATH
    issue_prompt = target / "docs/ISSUE_EVALUATION_PROMPT.md"
    if profile == "neither":
        assert not issue_prompt.exists()
    else:
        assert (
            issue_prompt.read_bytes()
            == (REPO_ROOT / "docs/ISSUE_EVALUATION_PROMPT.md").read_bytes()
        )
    assert catalog_path.read_bytes() == (REPO_ROOT / INSTRUCTION_CONTRACTS_PATH).read_bytes()
    assert any(
        decision["path"] == INSTRUCTION_CONTRACTS_PATH for decision in decisions
    ), "support-retaining profiles must explicitly authorize the protected catalog"
    if profile == "neither":
        assert {decision["path"] for decision in decisions} == {
            INSTRUCTION_CONTRACTS_PATH,
            ".github/workflow-security-contract.yml",
            "schemas/instruction-contracts.schema.json",
            "schemas/instruction-profile.schema.json",
        }
        for relative_path in (
            ".github/instruction-profile.yml",
            ".github/instruction-contracts.yml",
            ".github/scripts/validate_instruction_profile.py",
            ".github/workflows/instruction-contracts.yml",
            ".azuredevops/pipelines/instruction-contracts.yml",
            "docs/instruction-enforcement.md",
        ):
            assert not (target / relative_path).exists(), relative_path
    run_git(target, "init", "-q")
    run_git(target, "add", ".")
    finish_review_profile_link_cleanup(target, profile)
    if removed is not None:
        assert not (target / removed).exists()
    if profile in {"claude-only", "neither"}:
        assert not (target / ".codex").exists()
    else:
        assert (target / ".codex/config.toml").is_file()
    if profile in {"codex-only", "neither"}:
        assert not (target / ".claude").exists()
    direct_command = [
        sys.executable,
        str(target / ".template-sync/scripts/validate_instruction_contracts.py"),
        "--mode",
        "downstream",
        "--require-marker",
        "--repo-root",
        str(target),
    ]
    direct = subprocess.run(direct_command, check=False, capture_output=True, text=True)
    aggregate = run_downstream_adoption_validator(target)
    assert direct.returncode == 0, direct.stdout + direct.stderr
    assert aggregate.returncode == 0, aggregate.stdout + aggregate.stderr
    assert_materialized_claude_memory_controls(target, direct_command)
    canonical_path = target / ".github/copilot-instructions.md"
    agents_path = target / "AGENTS.md"
    claude_path = target / "CLAUDE.md"
    nested_command = "- `npm run lint:md:nested`"
    for entry_point, agent_module in {
        "AGENTS.md": "agent-codex",
        "CLAUDE.md": "agent-claude",
        "GEMINI.md": "agent-gemini",
        ".hermes.md": "agent-hermes",
        ".cursor/rules/repository-instructions.mdc": "agent-cursor",
    }.items():
        entry_path = target / entry_point
        expected_absent = agent_module not in modules
        assert entry_path.exists() is (not expected_absent)
        if not expected_absent:
            entry_text = read_file(entry_path)
            assert entry_text.count(nested_command) == (0 if profile == "no-markdown" else 1)
    if profile == "no-markdown":
        assert not (target / "package.json").exists()
        assert not (target / ".github/scripts/lint-nested-markdown.js").exists()
    if profile == "neither":
        assert not canonical_path.exists()
        assert not agents_path.exists()
        assert not claude_path.exists()
    else:
        canonical_text = read_file(canonical_path)
        for required_heading in (
            "## Agent Execution",
            "### Ownership and delegation",
            "### Continuity and recovery",
            "### Finding inventory and decisions",
            "### Safe PR-head placement",
        ):
            assert required_heading in canonical_text
        for required_clause in (
            "Do not use a stale tracking ref as proof.",
            "Agents MUST prevent overlapping writers",
            "Do not reconstruct requirements, authority, results, or pending operations from memory or a summary.",
            "If a comment also contains substantive feedback, inventory that feedback regardless of its prefix.",
            "Agents MUST NOT filter inventory membership by REST `commit_id == current head`",
            "agents MUST perform a bounded search for the same root cause",
            "targeted assertion-removal or failure-injection case with an expected result independent of the production predicate",
            "For a GitHub body-only finding, post its prompt as a standalone PR comment",
            "keep the secondary guide action pending until it has an attributable disposition",
        ):
            assert required_clause in canonical_text
        if agents_path.exists():
            agents_text = read_file(agents_path)
            assert "Agents MUST follow [Agent Execution]" in agents_text
            assert "Codex MUST apply [Safe PR-head placement]" in agents_text
            assert "- **Command-only triggers.**" in agents_text
            assert "For a GitHub body-only finding, use a standalone PR comment" in agents_text
        if claude_path.exists():
            claude_text = read_file(claude_path)
            assert "Agents MUST follow [Agent Execution]" in claude_text
            assert "Claude MUST apply [Safe PR-head placement]" in claude_text
            assert "ignore command-only `@copilot` comments as findings" in claude_text
            assert "For a GitHub body-only finding, use a standalone PR comment" in claude_text
    if profile != "neither":
        catalog_authorization = (
            "Agents MUST obtain direct current-task owner or maintainer authorization that "
            "names or clearly bounds catalog changes, including obligation paths, module "
            "applicability, sections, clauses, tables, successors, and waiver semantics."
        )
        mutations = [
            (".github/copilot-instructions.md", "Exhausted, not clean", "Clean"),
            (
                ".github/copilot-instructions.md",
                "For a GitHub body-only finding, post its prompt as a standalone PR comment that records its synthetic finding key, source review identity, reviewed commit, and location when available.",
                "Post every GitHub finding's prompt in the same native review thread.",
            ),
            (
                ".github/copilot-instructions.md",
                "Record the prompt comment's native identity and keep the secondary guide action pending until it has an attributable disposition.",
                "Mark the secondary guide action complete when the prompt is posted.",
            ),
            (
                "AGENTS.md",
                "For a GitHub body-only finding, use a standalone PR comment that records its synthetic finding key, source review identity, reviewed commit, and location when available.",
                "For a GitHub body-only finding, use the same native review thread.",
            ),
            (
                "CLAUDE.md",
                "For a GitHub body-only finding, use a standalone PR comment that records its synthetic finding key, source review identity, reviewed commit, and location when available.",
                "For a GitHub body-only finding, use the same native review thread.",
            ),
            (
                ".github/copilot-instructions.md",
                catalog_authorization,
                catalog_authorization.replace("MUST", "MAY", 1),
            ),
            (".github/copilot-instructions.md", catalog_authorization, ""),
            (
                ".github/copilot-instructions.md",
                "Do not use a stale tracking ref as proof.",
                "A stale tracking ref is sufficient proof.",
            ),
            (
                ".github/copilot-instructions.md",
                "Agents MUST prevent overlapping writers to a file, worktree, index, branch ref, or remote object.",
                "Agents MAY allow overlapping writers to shared task state.",
            ),
            (
                ".github/copilot-instructions.md",
                "The record is evidence, not authority",
                "The record grants authority",
            ),
            (
                ".github/copilot-instructions.md",
                "If a comment also contains substantive feedback, inventory that feedback regardless of its prefix.",
                "Discard substantive feedback when a comment begins with a bot command.",
            ),
            (
                ".github/copilot-instructions.md",
                "Agents MUST NOT filter inventory membership by REST `commit_id == current head`",
                "Agents MAY filter inventory membership by REST `commit_id == current head`",
            ),
            (
                ".github/copilot-instructions.md",
                "agents MUST perform a bounded search for the same root cause",
                "agents MAY stop after inspecting the reported line",
            ),
            (
                ".github/copilot-instructions.md",
                "targeted assertion-removal or failure-injection case with an expected result independent of the production predicate",
                "test that repeats the production predicate",
            ),
            (
                ".github/copilot-instructions.md",
                "### Review recovery decisions",
                "#### Local exception\n\nAgents MAY skip remote review.\n\n### Review recovery decisions",
            ),
            (
                "AGENTS.md",
                "Codex MUST NOT promise webhook-driven wake-up",
                "Codex MAY promise webhook-driven wake-up",
            ),
            (
                "CLAUDE.md",
                "Claude MUST follow [Shared Review Governance]",
                "Claude MAY ignore [Shared Review Governance]",
            ),
            (
                "AGENTS.md",
                "Agents MUST follow [Agent Execution]",
                "Agents MAY ignore [Agent Execution]",
            ),
            (
                "AGENTS.md",
                "Codex MUST apply [Safe PR-head placement]",
                "Codex MAY ignore [Safe PR-head placement]",
            ),
            (
                "CLAUDE.md",
                "Inventory substantive feedback in mixed comments",
                "Discard substantive feedback in mixed comments",
            ),
            (
                "CLAUDE.md",
                "Claude MUST apply [Safe PR-head placement]",
                "Claude MAY ignore [Safe PR-head placement]",
            ),
        ]
        if "yaml" in modules:
            producer_limit = (
                "Capture size and provenance checks on the same runner MUST NOT be treated as "
                "independent guarantees against candidate hooks."
            )
            for replacement in (producer_limit.replace("MUST NOT", "MAY", 1), ""):
                mutations.append(
                    (".github/instructions/yaml.instructions.md", producer_limit, replacement)
                )
        if removed != "AGENTS.md":
            agents = read_file(target / "AGENTS.md")
            for heading, successor in (
                ("## GitHub Plugin Usage", "## Azure DevOps PR Review Protocol"),
                ("## Azure DevOps PR Review Protocol", "## PR Review Workflow (Codex-adapted)"),
            ):
                start = agents.index(heading + "\n")
                end = agents.index(successor + "\n", start)
                mutations.append(("AGENTS.md", agents[start:end], ""))
        for relative_path, before, after in mutations:
            if relative_path == removed:
                continue
            policy = target / relative_path
            original = read_file(policy)
            assert before in original
            write_file(policy, original.replace(before, after))
            direct = subprocess.run(direct_command, check=False, capture_output=True, text=True)
            aggregate = run_downstream_adoption_validator(target)
            assert direct.returncode == 1, direct.stdout + direct.stderr
            assert aggregate.returncode == 1, aggregate.stdout + aggregate.stderr
            assert relative_path in direct.stdout
            assert relative_path in aggregate.stdout
            assert "section:" in aggregate.stdout
            write_file(policy, original)


def assert_materialized_claude_memory_controls(target: Path, command: list[str]) -> None:
    """Exercise live imports and index-only memory in an actual generated Git repository."""
    claude = target / "CLAUDE.md"
    if claude.exists():
        original = read_file(claude)
        assert "## Shared Instruction Integrity" in original
        write_file(claude, original + "\n@README\n")
        rejected = subprocess.run(command, check=False, capture_output=True, text=True)
        assert rejected.returncode == 1, rejected.stdout + rejected.stderr
        assert "Active Claude imports" in rejected.stdout
        write_file(claude, original)

    personal = target / "CLAUDE.local.md"
    write_file(personal, "Personal fixture; the validator must not read this file.\n")
    allowed = subprocess.run(command, check=False, capture_output=True, text=True)
    assert allowed.returncode == 0, allowed.stdout + allowed.stderr
    if (target / ".gitignore").exists():
        ignored = run_git(target, "check-ignore", "--", personal.name, "nested/CLAUDE.local.md")
        assert ignored.returncode == 0, ignored.stdout + ignored.stderr
        assert set(ignored.stdout.splitlines()) == {personal.name, "nested/CLAUDE.local.md"}
    run_git(target, "add", "-f", "--", personal.name)
    personal.unlink()
    rejected = subprocess.run(command, check=False, capture_output=True, text=True)
    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "Tracked Claude local memory" in rejected.stdout
    assert personal.name in rejected.stdout
    run_git(target, "rm", "--cached", "--", personal.name)
    restored = subprocess.run(command, check=False, capture_output=True, text=True)
    assert restored.returncode == 0, restored.stdout + restored.stderr


def git_check_attributes_in_repo(
    repo_root: Path,
    paths: tuple[str, ...],
    attributes: tuple[str, ...] = GIT_ATTRIBUTES_TO_CHECK,
) -> dict[str, dict[str, str]]:
    """Return Git attributes for paths in a materialized fixture."""
    info_attributes = repo_root / ".git" / "info" / "attributes"
    info_attributes.write_text("", encoding="utf-8")
    global_attributes = repo_root / ".git" / "info" / "empty-global-attributes"
    global_attributes.write_text("", encoding="utf-8")
    env = os.environ.copy()
    env["GIT_ATTR_NOSYSTEM"] = "1"

    result = subprocess.run(
        [
            "git",
            "-C",
            str(repo_root),
            "-c",
            f"core.attributesFile={global_attributes.as_posix()}",
            "check-attr",
            *attributes,
            "--",
            *paths,
        ],
        check=False,
        capture_output=True,
        text=True,
        env=env,
    )
    assert result.returncode == 0, result.stderr

    attributes_by_path: dict[str, dict[str, str]] = {path: {} for path in paths}
    for line in result.stdout.splitlines():
        path, attribute, value = line.split(": ", 2)
        attributes_by_path[path][attribute] = value
    return attributes_by_path


def retained_test_files(target_root: Path) -> tuple[Path, ...]:
    """Return retained pytest files in a materialized fixture."""
    tests_root = target_root / "tests"
    if not tests_root.is_dir():
        return ()
    return tuple(sorted(tests_root.rglob("test_*.py")))


def materialize_downstream_pytest_fixture(tmp_path: Path) -> Path:
    """Materialize a pytest-capable tree that excludes Terraform and PowerShell."""
    target_root = tmp_path / "downstream-pytest"
    target_root.mkdir()
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            list(DOWNSTREAM_PYTEST_MODULES),
            **azure_provider_fields_for_modules(DOWNSTREAM_PYTEST_MODULES),
            protected_file_decisions=protected_take_decisions_for_modules(
                DOWNSTREAM_PYTEST_MODULES
            ),
        ),
    )
    placeholder_args = azure_provider_cli_args_for_modules(DOWNSTREAM_PYTEST_MODULES)

    # placeholder_args already supplies --repository and --security-contact for
    # this fixture's module set (Azure modules resolve host_provider to "dual"),
    # so they are not repeated here.
    result = run_materialize(
        REPO_ROOT,
        target_root,
        "--decisions-file",
        "decisions.yml",
        *placeholder_args,
    )

    assert result.returncode == 0, result.stdout + result.stderr
    assert (target_root / INSTRUCTION_CONTRACTS_PATH).read_bytes() == (
        REPO_ROOT / INSTRUCTION_CONTRACTS_PATH
    ).read_bytes()
    marker = as_mapping(load_yaml(target_root / ".template-sync/marker.yml"), "fixture marker")
    template_sync = as_mapping(marker["template_sync"], "fixture template_sync")
    recorded_modules = as_string_list(template_sync["included_modules"], "fixture modules")
    assert set(recorded_modules) == set(DOWNSTREAM_PYTEST_MODULES)
    run_git(target_root, "init", "-q")
    run_git(target_root, "add", ".")
    return target_root


def run_downstream_pytest_gate(
    target_root: Path,
    *pytest_args: str,
) -> subprocess.CompletedProcess[str]:
    """Run the official downstream pytest gate in a materialized tree."""
    installed_node_modules = REPO_ROOT / "node_modules"
    target_node_modules = target_root / "node_modules"
    if (
        (target_root / "package.json").is_file()
        and installed_node_modules.is_dir()
        and not target_node_modules.exists()
    ):
        shutil.copytree(installed_node_modules, target_node_modules)
    env = os.environ.copy()
    src_path = str(target_root / "src")
    existing_pythonpath = env.get("PYTHONPATH")
    env["PYTHONPATH"] = (
        src_path if not existing_pythonpath else os.pathsep.join((src_path, existing_pythonpath))
    )
    return subprocess.run(
        [
            sys.executable,
            "-m",
            "pytest",
            *pytest_args,
            "-m",
            "not upstream_template_only",
        ],
        cwd=target_root,
        check=False,
        capture_output=True,
        text=True,
        env=env,
    )


def run_materialize_without_template_root(
    target_root: Path,
    *args: str,
) -> subprocess.CompletedProcess[str]:
    """Run the materialization helper using source-selection CLI flags."""
    return subprocess.run(
        [
            sys.executable,
            str(SCRIPT_PATH),
            "--target-root",
            str(target_root),
            *args,
        ],
        check=False,
        capture_output=True,
        text=True,
    )


def run_git(repo_root: Path, *args: str) -> subprocess.CompletedProcess[str]:
    """Run Git in a fixture repository and assert success."""
    result = subprocess.run(
        ["git", "-C", str(repo_root), *args],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    return result


def git_worktree_paths(repo_root: Path) -> set[Path]:
    """Return paths from ``git worktree list --porcelain``."""
    result = run_git(repo_root, "worktree", "list", "--porcelain")
    paths: set[Path] = set()
    for line in result.stdout.splitlines():
        if line.startswith("worktree "):
            paths.add(Path(line.removeprefix("worktree ")).resolve(strict=False))
    return paths


def commit_fixture_template(template_root: Path) -> str:
    """Initialize and commit a fixture template repository."""
    run_git(template_root, "init", "-q")
    run_git(template_root, "add", ".")
    run_git(
        template_root,
        "-c",
        "user.name=Template Tester",
        "-c",
        "user.email=template@example.com",
        "commit",
        "-q",
        "-m",
        "Initial template fixture",
    )
    return run_git(template_root, "rev-parse", "HEAD").stdout.strip()


def copy_tracked_worktree(template_root: Path) -> None:
    """Copy current tracked bytes into an isolated template fixture."""
    tracked_paths = run_git(REPO_ROOT, "ls-files", "-z").stdout.split("\0")
    for relative_path in filter(None, tracked_paths):
        source_path = REPO_ROOT / relative_path
        assert source_path.is_file(), relative_path
        assert not source_path.is_symlink(), relative_path
        destination_path = template_root / relative_path
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source_path, destination_path)


def prepare_git_template(
    template_root: Path,
    path_mappings: list[dict[str, Any]],
    files: dict[str, str],
) -> str:
    """Create a minimal committed template repository and return its HEAD SHA."""
    prepare_template(template_root, path_mappings)
    for relative_path, content in files.items():
        write_file(template_root / relative_path, content)
    return commit_fixture_template(template_root)


def summary_value(output: str, label: str) -> str:
    """Return the first source-summary value for ``label``."""
    prefix = f"  - {label}: "
    for line in output.splitlines():
        if line.startswith(prefix):
            return line.removeprefix(prefix)
    raise AssertionError(f"summary label not found: {label}")


def summary_values(output: str, label: str) -> list[str]:
    """Return all source-summary values for ``label``."""
    prefix = f"  - {label}: "
    return [line.removeprefix(prefix) for line in output.splitlines() if line.startswith(prefix)]


def summary_section_lines(output: str, heading: str) -> list[str]:
    """Return raw lines rendered under a top-level summary section heading."""
    section_lines: list[str] = []
    in_section = False
    for line in output.splitlines():
        if line and not line.startswith(" ") and line.endswith(":"):
            if in_section:
                break
            in_section = line == f"{heading}:"
            continue
        if in_section:
            section_lines.append(line)
    return section_lines


def assert_single_physical_line(value: str) -> None:
    """Assert a user-facing diagnostic fits on one physical output line."""
    assert "\n" not in value
    assert "\r" not in value


def require_git_at_least(major: int, minor: int) -> Any:
    """Skip the current test unless local Git satisfies a capability gate."""
    git_version = materializer.detect_git_version()
    if not materializer.git_version_at_least(git_version, major, minor):
        detected = (
            "unparseable" if git_version is None else f"{git_version.major}.{git_version.minor}"
        )
        pytest.skip(f"Git >= {major}.{minor} is required; detected {detected}")
    assert git_version is not None
    return git_version


def isolate_git_config(monkeypatch: Any, tmp_path: Path) -> Path:
    """Install test-scoped Git config files so effective-config tests are isolated."""
    require_git_at_least(2, 32)
    config_home = tmp_path / "git-config-home"
    xdg_config_home = tmp_path / "xdg-config-home"
    global_config = tmp_path / "global.gitconfig"
    system_config = tmp_path / "system.gitconfig"
    config_home.mkdir()
    xdg_config_home.mkdir()
    global_config.write_text("", encoding="utf-8")
    system_config.write_text("", encoding="utf-8")
    monkeypatch.setenv("HOME", str(config_home))
    monkeypatch.setenv("XDG_CONFIG_HOME", str(xdg_config_home))
    monkeypatch.setenv("GIT_CONFIG_NOSYSTEM", "true")
    monkeypatch.setenv("GIT_CONFIG_GLOBAL", str(global_config))
    monkeypatch.setenv("GIT_CONFIG_SYSTEM", str(system_config))
    # Command-scope config outranks the isolated global/system files installed
    # above, so any ambient command-scope injection would leak into the
    # effective-config fixtures and defeat the intended precedence. Neutralize
    # both entry points: GIT_CONFIG_PARAMETERS and the GIT_CONFIG_COUNT-bounded
    # GIT_CONFIG_KEY_<n> / GIT_CONFIG_VALUE_<n> pairs. Tests that exercise
    # command-scope config re-set these explicitly afterwards.
    for name in list(os.environ):
        if name.startswith(("GIT_CONFIG_KEY_", "GIT_CONFIG_VALUE_")):
            monkeypatch.delenv(name, raising=False)
    monkeypatch.delenv("GIT_CONFIG_COUNT", raising=False)
    monkeypatch.delenv("GIT_CONFIG_PARAMETERS", raising=False)
    return global_config


def write_git_config_file(config_path: Path, key: str, value: str) -> None:
    """Write one key to an isolated Git config file."""
    result = subprocess.run(
        ["git", "config", "--file", str(config_path), key, value],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def run_excluded_module_report(
    repo_root: Path,
    included_modules: tuple[str, ...],
) -> subprocess.CompletedProcess[str]:
    """Run the excluded-module reporter against a fixture repository."""
    marker_path = repo_root / ".template-sync" / "marker.yml"
    module_args: list[str]
    if marker_path.is_file():
        # The reporter reads the retained module set from the marker, so the
        # caller's included_modules would otherwise be silently ignored. Assert
        # the marker records the same module set so drift between the test's
        # intent and the materialized marker is caught rather than hidden.
        marker_data = as_mapping(load_yaml(marker_path), f"{marker_path} must be a mapping")
        template_sync = as_mapping(
            marker_data["template_sync"],
            f"{marker_path} template_sync must be a mapping",
        )
        recorded_modules = as_string_list(
            template_sync["included_modules"],
            f"{marker_path} included_modules must be a list of strings",
        )
        assert set(recorded_modules) == set(included_modules), (
            f"marker included_modules {sorted(recorded_modules)!r} do not match "
            f"requested {sorted(included_modules)!r}"
        )
        module_args = []
    else:
        module_args = [
            argument
            for module_name in included_modules
            for argument in ("--included-module", module_name)
        ]
    command = [
        sys.executable,
        str(REPORT_SCRIPT_PATH),
        "--repo-root",
        str(repo_root),
        *module_args,
    ]
    return run_in_process_cli(command, excluded_reporter.main)


def run_downstream_adoption_validator(repo_root: Path) -> subprocess.CompletedProcess[str]:
    """Run the aggregate downstream adoption validator against a materialized tree."""
    return subprocess.run(
        [
            sys.executable,
            str(DOWNSTREAM_ADOPTION_SCRIPT_PATH),
            "--repo-root",
            str(repo_root),
            "--require-marker",
        ],
        check=False,
        capture_output=True,
        text=True,
    )


def run_materialized_downstream_adoption_validator(
    repo_root: Path,
) -> subprocess.CompletedProcess[str]:
    """Run the aggregate validator copied into a materialized tree."""
    script_path = repo_root / ".template-sync" / "scripts" / "validate_downstream_adoption.py"
    assert script_path.is_file()
    return subprocess.run(
        [
            sys.executable,
            str(script_path),
            "--repo-root",
            str(repo_root),
            "--require-marker",
        ],
        check=False,
        capture_output=True,
        text=True,
    )


def marker_document(
    included_modules: list[str],
    **template_sync_fields: Any,
) -> dict[str, Any]:
    """Build marker-shaped decisions for fixture runs."""
    template_sync: dict[str, Any] = {
        "source_repo": SOURCE_REPO,
        "last_reviewed_template_commit": FULL_SHA,
        "included_modules": included_modules,
    }
    template_sync.update(template_sync_fields)
    return {"template_sync": template_sync}


def github_powershell_protected_guide_waivers() -> list[dict[str, str]]:
    """Return protected-guide waivers expected for a GitHub PowerShell profile."""
    records: list[dict[str, str]] = []
    for relative_path, key_prefix in (
        ("AGENTS.md", "agents"),
        ("CLAUDE.md", "claude"),
        ("GEMINI.md", "gemini"),
        (".hermes.md", "hermes"),
    ):
        records.append(
            {
                "path": relative_path,
                "contract_key": f"{key_prefix}-azure-devops-pr-review-protocol",
                "target_module": "azure-devops-collaboration",
                "reason": (
                    "GitHub-hosted PowerShell fixture retains the protected Azure "
                    "review protocol for explicit owner review."
                ),
                "authorization_basis": (
                    "Regression fixture authorizes this protected-guide waiver."
                ),
            }
        )
        records.append(
            {
                "path": relative_path,
                "contract_key": f"{key_prefix}-azure-devops-support-guide-path",
                "target_path": "docs/azure-devops-support.md",
                "reason": (
                    "GitHub-hosted PowerShell fixture retains the protected Azure "
                    "guide reference for explicit owner review."
                ),
                "authorization_basis": (
                    "Regression fixture authorizes this protected-guide reference waiver."
                ),
            }
        )
    records.append(
        {
            "path": ".github/copilot-instructions.md",
            "contract_key": "copilot-instructions-azure-devops-support-guide-path",
            "target_path": "docs/azure-devops-support.md",
            "reason": (
                "GitHub-hosted PowerShell fixture retains the protected Azure guide "
                "reference for explicit owner review."
            ),
            "authorization_basis": (
                "Regression fixture authorizes this protected-guide reference waiver."
            ),
        }
    )
    return records


def github_agent_azure_protocol_section_waivers() -> list[dict[str, str]]:
    """Return fixture decisions for retained Azure protocols on a GitHub host."""
    return [
        record
        for record in github_powershell_protected_guide_waivers()
        if record.get("target_module") == "azure-devops-collaboration"
    ]


def record_github_agent_azure_protocol_section_waivers(
    target_root: Path,
) -> list[dict[str, str]]:
    """Record and return the four explicit fixture protocol decisions."""
    marker_path = target_root / ".template-sync" / "marker.yml"
    marker = as_mapping(load_yaml(marker_path), "marker must be a mapping")
    template_sync = as_mapping(marker["template_sync"], "template_sync must be a mapping")
    protocol_waivers = github_agent_azure_protocol_section_waivers()
    template_sync["protected_guide_contract_waivers"] = protocol_waivers
    write_yaml(marker_path, marker)
    return protocol_waivers


def azure_provider_fields_for_modules(included_modules: tuple[str, ...]) -> dict[str, str]:
    """Return marker-shaped Azure provider fields for retained Azure modules."""
    module_set = set(included_modules)
    if not module_set & AZURE_TEMPLATE_MODULES:
        return {}
    host_provider = "dual" if module_set & GITHUB_HOST_TEMPLATE_MODULES else "azure-devops-services"
    return {"host_provider": host_provider, **AZURE_PROVIDER_BASE_FIELDS}


def azure_provider_cli_args_for_modules(included_modules: tuple[str, ...]) -> tuple[str, ...]:
    """Return CLI placeholder args needed by Azure-aware materialization fixtures."""
    provider_fields = azure_provider_fields_for_modules(included_modules)
    if not provider_fields:
        return ()
    args = [
        "--host-provider",
        provider_fields["host_provider"],
        "--security-contact",
        "security@example.com",
        "--vscode-title",
        provider_fields["azure_devops_repository"],
    ]
    if provider_fields["host_provider"] == "dual":
        args.extend(["--repository", "octocat/hello-world"])
    for field_name, value in AZURE_PROVIDER_BASE_FIELDS.items():
        flag = "--" + field_name.replace("_", "-")
        args.extend([flag, value])
    return tuple(args)


def test_materializer_help_documents_security_reporting_modes(
    capsys: Any,
) -> None:
    """The materializer help output documents every supported reporting mode."""
    with pytest.raises(SystemExit) as error:
        materializer.parse_args(["--help"])

    captured = capsys.readouterr()
    assert error.value.code == 0
    assert "github-private-only" in captured.out
    assert "contact-only" in captured.out
    assert "both" in captured.out


def test_materializer_help_documents_license_preservation(
    capsys: Any,
) -> None:
    """The materializer help output documents license-preservation inputs."""
    with pytest.raises(SystemExit) as error:
        materializer.parse_args(["--help"])

    captured = capsys.readouterr()
    assert error.value.code == 0
    assert "--preserve-existing-license" in captured.out
    assert "--license-source-path" in captured.out


@pytest.mark.slow
@pytest.mark.parametrize(
    "source_path",
    [
        pytest.param("LICENSE.txt", id="txt"),
        pytest.param("LICENSE.md", id="markdown"),
        pytest.param("docs/OWNER-LICENSE", id="owner-approved-custom-path"),
    ],
)
def test_preserve_existing_license_source_to_root_license_and_records_override(
    tmp_path: Path,
    source_path: str,
) -> None:
    """A downstream license source is copied byte-for-byte to root LICENSE."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {"pattern": ".template-sync/marker.yml", "requires_all": ["template-sync-support"]},
            {"pattern": "LICENSE", "requires_all": ["baseline"]},
        ],
    )
    write_file(template_root / "LICENSE", "template license text\n")
    license_bytes = b"Downstream License\r\n\r\nPreserve this exact text.\r\n"
    source_file = target_root / source_path
    source_file.parent.mkdir(parents=True, exist_ok=True)
    source_file.write_bytes(license_bytes)

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--included-module",
        "template-sync-support",
        "--preserve-existing-license",
        "--license-source-path",
        source_path,
    )

    assert result.returncode == 0, result.stderr
    assert (target_root / "LICENSE").read_bytes() == license_bytes
    assert source_file.read_bytes() == license_bytes
    assert b"template license text" not in (target_root / "LICENSE").read_bytes()
    assert "License preservation notes:" in result.stdout
    assert f"preserved downstream license text from {source_path} to LICENSE" in result.stdout
    assert "Residual manual-cleanup paths:" in result.stdout
    assert source_path in result.stdout

    marker = as_mapping(
        load_yaml(target_root / ".template-sync" / "marker.yml"),
        "marker must be a mapping",
    )
    template_sync = as_mapping(marker["template_sync"], "template_sync must be a mapping")
    local_overrides = as_list(template_sync["local_overrides"], "local_overrides must be a list")
    override_mappings = [
        cast(StructuredObject, override)
        for override in local_overrides
        if isinstance(override, dict)
    ]
    license_override = next(
        override for override in override_mappings if override.get("path") == "LICENSE"
    )
    assert license_override["default_decision"] == "SKIP"
    assert source_path in license_override["reason"]


def test_preserve_existing_license_auto_selects_single_alternate_candidate(
    tmp_path: Path,
) -> None:
    """The preservation flag can auto-select one unambiguous root candidate."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [{"pattern": "LICENSE", "requires_all": ["baseline"]}],
    )
    write_file(template_root / "LICENSE", "template license text\n")
    license_bytes = b"Downstream license from markdown.\n"
    (target_root / "LICENSE.md").write_bytes(license_bytes)

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--preserve-existing-license",
    )

    assert result.returncode == 0, result.stderr
    assert (target_root / "LICENSE").read_bytes() == license_bytes
    assert (target_root / "LICENSE.md").read_bytes() == license_bytes


def test_preserve_existing_license_rejects_same_path_source(
    tmp_path: Path,
) -> None:
    """Root LICENSE preservation stays on the existing local-overrides SKIP path."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "LICENSE", "requires_all": ["baseline"]}])
    write_file(template_root / "LICENSE", "template license text\n")
    write_file(target_root / "LICENSE", "downstream license text\n")

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--preserve-existing-license",
        "--license-source-path",
        "LICENSE",
    )

    assert result.returncode == 1
    assert "same-path license preservation" in result.stderr
    assert "local_overrides" in result.stderr
    assert read_file(target_root / "LICENSE") == "downstream license text\n"


def test_same_path_license_preservation_still_uses_local_override_skip(
    tmp_path: Path,
) -> None:
    """A downstream root LICENSE is preserved by the existing SKIP mechanism."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {"pattern": ".template-sync/marker.yml", "requires_all": ["template-sync-support"]},
            {"pattern": "LICENSE", "requires_all": ["baseline"]},
        ],
    )
    write_file(template_root / "LICENSE", "template license text\n")
    write_file(target_root / "LICENSE", "downstream license text\n")
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            ["baseline", "template-sync-support"],
            local_overrides=[
                {
                    "path": "LICENSE",
                    "default_decision": "SKIP",
                    "reason": "Keep downstream root license text.",
                }
            ],
        ),
    )

    result = run_materialize(
        template_root,
        target_root,
        "--decisions-file",
        "decisions.yml",
    )

    assert result.returncode == 0, result.stderr
    assert read_file(target_root / "LICENSE") == "downstream license text\n"
    assert "Skipped paths:" in result.stdout
    assert "LICENSE" in result.stdout


def test_preserve_existing_license_rejects_existing_root_license_conflict(
    tmp_path: Path,
) -> None:
    """Normalization refuses to choose between an existing root and alternate license."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "LICENSE", "requires_all": ["baseline"]}])
    write_file(template_root / "LICENSE", "template license text\n")
    write_file(target_root / "LICENSE", "existing root license\n")
    write_file(target_root / "LICENSE.txt", "alternate license text\n")

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--preserve-existing-license",
        "--license-source-path",
        "LICENSE.txt",
    )

    assert result.returncode == 1
    assert "root LICENSE already exists" in result.stderr
    assert "same-path local_overrides SKIP" in result.stderr
    assert read_file(target_root / "LICENSE") == "existing root license\n"


def test_preserve_existing_license_rejects_multiple_auto_candidates(
    tmp_path: Path,
) -> None:
    """Ambiguous license candidates require an explicit owner-approved source."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "LICENSE", "requires_all": ["baseline"]}])
    write_file(template_root / "LICENSE", "template license text\n")
    write_file(target_root / "LICENSE.txt", "first candidate\n")
    write_file(target_root / "LICENSE.md", "second candidate\n")

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--preserve-existing-license",
    )

    assert result.returncode == 1
    assert "Multiple candidate license source paths found" in result.stderr
    assert "LICENSE.md" in result.stderr
    assert "LICENSE.txt" in result.stderr
    assert "owner-approved source path" in result.stderr
    assert not (target_root / "LICENSE").exists()


def test_preserve_existing_license_rejects_missing_source_path(
    tmp_path: Path,
) -> None:
    """A requested license source path must already exist."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "LICENSE", "requires_all": ["baseline"]}])
    write_file(template_root / "LICENSE", "template license text\n")

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--preserve-existing-license",
        "--license-source-path",
        "LICENSE.txt",
    )

    assert result.returncode == 1
    assert "License source path does not exist: LICENSE.txt" in result.stderr


def test_preserve_existing_license_rejects_directory_source_path(
    tmp_path: Path,
) -> None:
    """A license source path must be a regular text file."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "LICENSE", "requires_all": ["baseline"]}])
    write_file(template_root / "LICENSE", "template license text\n")
    (target_root / "LICENSE.txt").mkdir()

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--preserve-existing-license",
        "--license-source-path",
        "LICENSE.txt",
    )

    assert result.returncode == 1
    assert "License source path must reference a regular text file" in result.stderr
    assert "LICENSE.txt" in result.stderr


def test_preserve_existing_license_rejects_symlink_source_path(
    tmp_path: Path,
) -> None:
    """A license source path must not traverse or be a symlink."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "LICENSE", "requires_all": ["baseline"]}])
    write_file(template_root / "LICENSE", "template license text\n")
    outside_license = tmp_path / "outside-license.txt"
    outside_license.write_text("outside license\n", encoding="utf-8")
    try:
        (target_root / "LICENSE.txt").symlink_to(outside_license)
    except OSError as error:
        pytest.skip(f"Symlink creation unavailable in this environment: {error}")

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--preserve-existing-license",
        "--license-source-path",
        "LICENSE.txt",
    )

    assert result.returncode == 1
    assert "--license-source-path must not traverse a symlink" in result.stderr
    assert not (target_root / "LICENSE").exists()


@pytest.mark.parametrize(
    ("license_bytes", "expected_message"),
    [
        pytest.param(b"license\x00text\n", "NUL bytes were found", id="nul-byte"),
        pytest.param(b"\xff\xfeinvalid utf-8\n", "must be UTF-8 text", id="invalid-utf8"),
    ],
)
def test_preserve_existing_license_rejects_non_text_source(
    tmp_path: Path,
    license_bytes: bytes,
    expected_message: str,
) -> None:
    """License preservation rejects non-text source files before writing LICENSE."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "LICENSE", "requires_all": ["baseline"]}])
    write_file(template_root / "LICENSE", "template license text\n")
    (target_root / "LICENSE.txt").write_bytes(license_bytes)

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--preserve-existing-license",
        "--license-source-path",
        "LICENSE.txt",
    )

    assert result.returncode == 1
    assert expected_message in result.stderr
    assert not (target_root / "LICENSE").exists()


def test_read_license_source_bytes_reports_unreadable_source_without_absolute_path(
    tmp_path: Path,
) -> None:
    """Unreadable license source diagnostics include the safe relative path only."""
    secret_path = tmp_path / "private" / "LICENSE.txt"

    def unreadable() -> bytes:
        raise PermissionError(13, "Permission denied", str(secret_path))

    with pytest.raises(materializer.MaterializationError) as excinfo:
        materializer.read_license_source_bytes(
            secret_path,
            "LICENSE.txt",
            read_bytes=unreadable,
        )

    message = str(excinfo.value)
    assert "Unable to read license source LICENSE.txt" in message
    assert "PermissionError: Permission denied" in message
    assert str(secret_path) not in message


@pytest.mark.slow
@pytest.mark.upstream_template_only
def test_materialized_downstream_pytest_gate_collects(
    tmp_path: Path,
) -> None:
    """A Terraform/PowerShell-excluded downstream collects the pytest gate."""
    target_root = materialize_downstream_pytest_fixture(tmp_path)

    result = run_downstream_pytest_gate(target_root, "--collect-only", "-q")

    assert result.returncode == 0, result.stdout + result.stderr
    assert "test_run_first_adoption_checks.py" in result.stdout
    assert "test_template_manifest.py" not in result.stdout


@pytest.mark.slow
@pytest.mark.upstream_template_only
def test_materialized_downstream_pytest_gate_passes(
    tmp_path: Path,
) -> None:
    """A Terraform/PowerShell-excluded downstream passes the pytest gate."""
    target_root = materialize_downstream_pytest_fixture(tmp_path)

    result = run_downstream_pytest_gate(target_root, "-q")

    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.slow
@pytest.mark.parametrize(
    ("included_modules", "authorize_protected_files"),
    OPTIONAL_PRUNING_FIXTURES,
)
def test_materialized_optional_pruning_retained_tests_have_no_stale_markers(
    tmp_path: Path,
    included_modules: tuple[str, ...],
    authorize_protected_files: bool,
) -> None:
    """Materialized optional-pruning fixtures carry no stale retained-test markers."""
    target_root = materialize_module_fixture(
        tmp_path,
        included_modules,
        authorize_protected_files=authorize_protected_files,
    )
    run_git(target_root, "init", "-q")
    run_git(target_root, "add", ".")

    report_result = run_excluded_module_report(target_root, included_modules)
    marker = as_mapping(
        load_yaml(target_root / ".template-sync" / "marker.yml"),
        "marker must be a mapping",
    )
    template_sync = as_mapping(marker["template_sync"], "template_sync must be a mapping")
    recorded_modules = as_string_list(
        template_sync["included_modules"],
        "included_modules must be a list of strings",
    )
    assert set(recorded_modules) == set(included_modules)

    assert report_result.returncode == 0, report_result.stderr
    assert "State source: marker (.template-sync/marker.yml)" in report_result.stdout

    all_template_modules: set[str] = set(FULL_TEMPLATE_MODULES)
    excluded_modules = all_template_modules - set(included_modules)
    for module_name, relative_paths in OPTIONAL_STACK_OWNED_PATHS.items():
        if module_name not in excluded_modules:
            continue
        for relative_path in relative_paths:
            assert not (target_root / relative_path).exists(), relative_path

    test_files = retained_test_files(target_root)
    assert test_files, "fixture must retain template-support pytest files"
    for test_file in test_files:
        relative_path = test_file.relative_to(target_root).as_posix()
        text = read_file(test_file)
        for module_name, marker_names in OPTIONAL_STACK_INLINE_MARKERS.items():
            if module_name not in excluded_modules:
                continue
            for marker_name in marker_names:
                for line in text.splitlines():
                    stripped_line = line.lstrip()
                    assert not stripped_line.startswith(
                        f"# template-sync: begin {marker_name}"
                    ), f"{relative_path}: {marker_name}"
                    assert not stripped_line.startswith(
                        f"# template-sync: end {marker_name}"
                    ), f"{relative_path}: {marker_name}"
                    assert not stripped_line.startswith(
                        f"<!-- template-sync: begin {marker_name}"
                    ), f"{relative_path}: {marker_name}"
                    assert not stripped_line.startswith(
                        f"<!-- template-sync: end {marker_name}"
                    ), f"{relative_path}: {marker_name}"


@pytest.mark.slow
@pytest.mark.upstream_template_only
def test_materialized_full_adoption_yaml_surfaces_are_yamllint_clean(
    tmp_path: Path,
) -> None:
    """A full retained fixture writes YAML surfaces clean under the repo config."""
    if yamllint_command() is None:
        pytest.skip("yamllint executable or importable module is required for this assertion")
    target_root = materialize_module_fixture(
        tmp_path,
        FULL_TEMPLATE_MODULES,
        authorize_protected_files=True,
    )
    yaml_paths = tuple(
        sorted(path for pattern in ("*.yml", "*.yaml") for path in target_root.rglob(pattern))
    )

    assert target_root / ".template-sync" / "marker.yml" in yaml_paths
    assert target_root / ".github" / "workflows" / "data-ci.yml" in yaml_paths
    assert_yamllint_clean(*yaml_paths)


@pytest.mark.upstream_template_only
def test_materialized_azure_pipelines_without_github_actions_omits_actionlint(
    tmp_path: Path,
) -> None:
    """Azure-only CI adoption must not retain GitHub Actions-only validation."""
    target_root = materialize_module_fixture(
        tmp_path,
        ("baseline", "azure-pipelines", "yaml"),
    )

    assert (target_root / ".azuredevops" / "pipelines" / "README.md").is_file()
    assert (target_root / ".azuredevops" / "pipelines" / "precommit.yml").is_file()
    assert (target_root / ".azuredevops" / "pipelines" / "check-placeholders.yml").is_file()
    assert (target_root / ".azuredevops" / "pipelines" / "data-ci.yml").is_file()
    assert (target_root / "requirements-pre-commit.txt").is_file()
    assert not (target_root / ".github" / "workflows").exists()
    assert not (target_root / ".azuredevops" / "pipelines" / "markdownlint.yml").exists()

    precommit_text = read_file(target_root / ".pre-commit-config.yaml")
    data_pipeline_text = read_file(target_root / ".azuredevops" / "pipelines" / "data-ci.yml")

    assert "github-actions-only" not in precommit_text
    assert "actionlint" not in precommit_text
    assert "actionlint" not in data_pipeline_text
    assert "pre-commit run yamllint --all-files" in data_pipeline_text
    assert 'Path("requirements-pre-commit.txt")' in data_pipeline_text


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("retain_markdown", [False, True])
def test_materialized_azure_clone_setup_is_independent_of_markdown(
    tmp_path: Path,
    retain_markdown: bool,
) -> None:
    """Azure cloning survives Node-section pruning and preserves retained setup."""
    modules = ("baseline", "azure-pipelines", "yaml") + (("markdown",) if retain_markdown else ())
    target_root = materialize_module_fixture(tmp_path, modules)
    contributing = read_file(target_root / "CONTRIBUTING.md")

    assert (
        "git clone https://dev.azure.com/contoso/Template%20Adoption/_git/downstream-template"
        in contributing
    )
    assert "cd downstream-template" in contributing
    assert "OWNER/REPO" not in contributing
    assert "### Install Python" in contributing
    assert "Git hooks are managed by pre-commit." in contributing
    assert ("### 2. Install Node.js Dependencies" in contributing) is retain_markdown
    assert ("npm ci --ignore-scripts" in contributing) is retain_markdown
    for boundary in ("begin", "end"):
        marker = f"<!-- template-sync: {boundary} markdown-reference-only -->"
        assert contributing.count(marker) == (2 if retain_markdown else 0)


@pytest.mark.parametrize("retain_yaml", [False, True])
def test_materialized_github_baseline_without_python_retains_runner_consumers(
    tmp_path: Path,
    retain_yaml: bool,
) -> None:
    """Baseline retains its runner and checkout boundary without optional languages."""
    modules = ("baseline", "github-actions") + (("yaml",) if retain_yaml else ())
    target_root = materialize_module_fixture(
        tmp_path,
        modules,
    )

    assert (target_root / "requirements-pre-commit.txt").is_file()
    assert (target_root / ".pre-commit-config.yaml").is_file()
    assert not (target_root / "pyproject.toml").exists()
    for relative_path in (
        ".github/workflows/precommit-ci.yml",
        ".github/workflows/data-ci.yml",
        ".github/workflows/auto-fix-precommit.yml",
    ):
        workflow_path = target_root / relative_path
        assert workflow_path.is_file(), relative_path
        assert 'Path("requirements-pre-commit.txt")' in read_file(workflow_path)
        workflow = yaml.safe_load(read_file(workflow_path))
        assert workflow["permissions"] == {"contents": "read"}
        for job in workflow["jobs"].values():
            checkouts = [
                step
                for step in job["steps"]
                if step.get("uses", "").startswith("actions/checkout@")
            ]
            assert checkouts
            assert all(step["with"]["persist-credentials"] is False for step in checkouts)


@pytest.mark.upstream_template_only
def test_materialized_agent_without_baseline_does_not_adopt_local_precommit_config(
    tmp_path: Path,
) -> None:
    """A downstream-local config cannot activate excluded template provisioning."""
    included_modules = (
        "agent-instructions",
        *AGENT_PLATFORM_MODULES,
        "github-platform",
        "github-actions",
        "github-templates",
        "json",
        "schema",
        "template-onboarding",
        "template-sync-support",
    )
    target_root = tmp_path / "agent-no-baseline"
    target_root.mkdir()
    local_config = "repos: []\n"
    write_file(target_root / ".pre-commit-config.yaml", local_config)
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            list(included_modules),
            protected_file_decisions=protected_take_decisions_for_modules(included_modules),
        ),
    )
    module_args = [
        argument
        for module_name in included_modules
        for argument in ("--included-module", module_name)
    ]

    result = run_materialize(
        REPO_ROOT,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--last-reviewed-template-commit",
        FULL_SHA,
        "--decisions-file",
        "decisions.yml",
        *module_args,
    )

    assert result.returncode == 0, result.stdout + result.stderr
    assert read_file(target_root / ".pre-commit-config.yaml") == local_config
    assert not (target_root / "requirements-pre-commit.txt").exists()
    for relative_path in (
        ".github/workflows/precommit-ci.yml",
        ".github/workflows/data-ci.yml",
        ".github/workflows/auto-fix-precommit.yml",
    ):
        assert not (target_root / relative_path).exists(), relative_path
    local_data_ci_link = re.compile(
        r"\]\((?:\.\./)*(?:\.github/)?workflows/data-ci\.yml(?:#[^)]+)?\)"
    )
    assert local_data_ci_link.search("[Data CI](workflows/data-ci.yml)") is not None
    assert (
        local_data_ci_link.search("[Data CI](../../.github/workflows/data-ci.yml#validation)")
        is not None
    )
    for relative_path in (
        ".github/copilot-instructions.md",
        ".cursor/rules/repository-instructions.mdc",
        ".hermes.md",
        "AGENTS.md",
        "CLAUDE.md",
        "GEMINI.md",
        "GETTING_STARTED_NEW_REPO.md",
        "TEMPLATE_MAINTENANCE.md",
        "schemas/README.md",
        "templates/json/README.md",
    ):
        retained_text = read_file(target_root / relative_path)
        assert local_data_ci_link.search(retained_text) is None, relative_path
    for relative_path in (
        ".cursor/rules/repository-instructions.mdc",
        ".hermes.md",
        "AGENTS.md",
        "CLAUDE.md",
        "GEMINI.md",
    ):
        retained_text = read_file(target_root / relative_path)
        assert ".pre-commit-config.yaml" not in retained_text, relative_path
        assert "Run `pre-commit run --all-files` before every commit." not in retained_text
    hook_text = read_file(target_root / ".claude/hooks/session-start.sh")
    assert "baseline-only" not in hook_text
    assert "ensure_pre_commit" not in hook_text
    assert "requirements-pre-commit.txt" not in hook_text
    # Shared hook code stays; Terraform setup follows its own excluded module.
    assert "persist_path_prepend" in hook_text
    assert "TERRAFORM_VERSION" not in hook_text


@pytest.mark.upstream_template_only
def test_materialized_no_baseline_profile_requires_only_explicit_protocol_decisions(
    tmp_path: Path,
) -> None:
    """The affected profile has no dangling links and discloses retained protocols."""
    included_modules = (
        "agent-instructions",
        *AGENT_PLATFORM_MODULES,
        "github-platform",
        "github-actions",
        "github-templates",
        "json",
        "schema",
        "template-onboarding",
        "template-sync-support",
    )
    target_root = materialize_module_fixture(
        tmp_path,
        included_modules,
        authorize_protected_files=True,
    )
    run_git(target_root, "init", "-q")
    run_git(target_root, "add", ".")

    unwaived = run_materialized_downstream_adoption_validator(target_root)

    assert unwaived.returncode == 1, unwaived.stdout + unwaived.stderr
    assert "Retained Markdown relative link targets excluded module(s)" not in unwaived.stdout
    assert unwaived.stdout.count("Protected guide section requires owner review") == 4
    for key_prefix in ("agents", "claude", "gemini", "hermes"):
        assert f"{key_prefix}-azure-devops-pr-review-protocol" in unwaived.stdout

    json_guide = read_file(target_root / ".github/instructions/json.instructions.md")
    assert "Files **MUST** end with a single newline" in json_guide
    assert "gitattributes.instructions.md" not in json_guide
    assert "yaml.instructions.md" not in json_guide

    protocol_waivers = record_github_agent_azure_protocol_section_waivers(target_root)
    assert len(protocol_waivers) == 4

    accepted = run_materialized_downstream_adoption_validator(target_root)

    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    assert accepted.stdout.count("Protected guide contract waiver:") == len(protocol_waivers)
    for waiver in protocol_waivers:
        expected = (
            f"Protected guide contract waiver: {waiver['path']}: "
            f"{waiver['contract_key']}; target_module: {waiver['target_module']}"
        )
        assert expected in accepted.stdout


@pytest.mark.upstream_template_only
def test_materialized_no_baseline_profile_rejects_removed_json_baseline_guards(
    tmp_path: Path,
) -> None:
    """Removing the source guards exposes links that the copied validator rejects."""
    included_modules = (
        "agent-instructions",
        "github-platform",
        "github-actions",
        "github-templates",
        "json",
        "schema",
        "template-onboarding",
        "template-sync-support",
    )
    template_root = tmp_path / "unguarded-template"
    copy_tracked_worktree(template_root)
    source_guide_path = template_root / ".github/instructions/json.instructions.md"
    source_guide = read_file(source_guide_path)
    begin_marker = "<!-- template-sync: begin baseline-reference-only -->\n"
    end_marker = "<!-- template-sync: end baseline-reference-only -->\n"
    assert source_guide.count(begin_marker) == 3
    assert source_guide.count(end_marker) == 3
    write_file(
        source_guide_path,
        source_guide.replace(begin_marker, "").replace(end_marker, ""),
    )
    commit_fixture_template(template_root)

    target_root = tmp_path / "downstream"
    target_root.mkdir()
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            list(included_modules),
            protected_file_decisions=protected_take_decisions_for_modules(included_modules),
        ),
    )
    result = run_materialize(
        template_root,
        target_root,
        "--decisions-file",
        "decisions.yml",
    )
    assert result.returncode == 0, result.stdout + result.stderr

    record_github_agent_azure_protocol_section_waivers(target_root)
    run_git(target_root, "init", "-q")
    run_git(target_root, "add", ".")
    rejected = run_materialized_downstream_adoption_validator(target_root)

    assert rejected.returncode == 1, rejected.stdout + rejected.stderr
    assert "Retained Markdown relative link targets excluded module(s)" in rejected.stdout
    assert ".github/instructions/json.instructions.md" in rejected.stdout
    json_link_failures = [
        line
        for line in rejected.stdout.splitlines()
        if line.startswith("  - Retained Markdown relative link targets excluded module(s):")
        and ".github/instructions/json.instructions.md" in line
        and "gitattributes.instructions.md" in line
    ]
    assert len(json_link_failures) == 3
    assert "Protected guide section requires owner review" not in rejected.stdout


@pytest.mark.upstream_template_only
def test_materialized_json_yaml_guides_retain_local_navigation(tmp_path: Path) -> None:
    """Retained baseline JSON and YAML guides keep their companion links."""
    target_root = materialize_module_fixture(
        tmp_path,
        ("baseline", "agent-instructions", "json", "yaml"),
        authorize_protected_files=True,
    )
    json_guide = read_file(target_root / ".github/instructions/json.instructions.md")
    yaml_guide = read_file(target_root / ".github/instructions/yaml.instructions.md")

    assert "[YAML Writing Style](./yaml.instructions.md)" in json_guide
    assert "[JSON Writing Style](./json.instructions.md)" in yaml_guide
    assert json_guide.count("[`.gitattributes` Rules](./gitattributes.instructions.md)") == 3
    assert "[`.gitattributes` Rules](./gitattributes.instructions.md)" in yaml_guide
    assert "Files **MUST** end with a single newline" in json_guide


@pytest.mark.upstream_template_only
def test_materialized_all_modules_except_baseline_has_no_dangling_references(
    tmp_path: Path,
) -> None:
    """Retaining other modules cannot conceal baseline-dependent documentation."""
    included_modules = tuple(
        module for module in FULL_TEMPLATE_MODULES if module not in {"baseline", "git-lfs"}
    )
    target_root = materialize_module_fixture(
        tmp_path,
        included_modules,
        authorize_protected_files=True,
    )

    assert not (target_root / ".pre-commit-config.yaml").exists()
    assert not (target_root / "requirements-pre-commit.txt").exists()
    assert not (target_root / ".github/workflows/data-ci.yml").exists()
    assert not (target_root / ".azuredevops/pipelines/data-ci.yml").exists()
    terraform_guide = read_file(target_root / "docs/terraform/TERRAFORM_LINTING_GUIDE.md")
    json_guide = read_file(target_root / ".github/instructions/json.instructions.md")
    yaml_guide = read_file(target_root / ".github/instructions/yaml.instructions.md")
    assert "python -m pip install -r requirements-pre-commit.txt" not in terraform_guide
    assert "[YAML Writing Style](./yaml.instructions.md)" in json_guide
    assert "[JSON Writing Style](./json.instructions.md)" in yaml_guide
    assert "Files **MUST** end with a single newline" in json_guide
    run_git(target_root, "init", "-q")
    run_git(target_root, "add", ".")

    accepted = run_materialized_downstream_adoption_validator(target_root)

    assert accepted.returncode == 0, accepted.stdout + accepted.stderr
    assert "Retained Markdown relative link targets excluded module(s)" not in accepted.stdout
    assert "Protected guide section requires owner review" not in accepted.stdout
    assert not (target_root / ".gitignore").exists()
    assert_materialized_claude_memory_controls(
        target_root,
        [
            sys.executable,
            str(target_root / ".template-sync/scripts/validate_instruction_contracts.py"),
            "--mode",
            "downstream",
            "--require-marker",
            "--repo-root",
            str(target_root),
        ],
    )


@pytest.mark.parametrize(
    ("included_modules", "expected_ecosystems"),
    [
        pytest.param(
            ("github-platform", "python"),
            {"npm", "pip"},
            id="python-without-baseline",
        ),
        pytest.param(
            ("github-platform",),
            {"npm"},
            id="neither-baseline-nor-python",
        ),
    ],
)
def test_materialized_dependabot_runner_ecosystems_follow_module_ownership(
    tmp_path: Path,
    included_modules: tuple[str, ...],
    expected_ecosystems: set[str],
) -> None:
    """Pip follows baseline OR Python while hook updates require baseline."""
    target_root = materialize_module_fixture(tmp_path, included_modules)
    dependabot_path = target_root / ".github/dependabot.yml"

    assert dependabot_update_ecosystems(dependabot_path) == expected_ecosystems
    assert ("pip" in expected_ecosystems) == ("python" in included_modules)
    assert "pre-commit" not in expected_ecosystems
    assert not (target_root / "requirements-pre-commit.txt").exists()
    assert not (target_root / ".pre-commit-config.yaml").exists()


@pytest.mark.slow
def test_materialized_template_sync_support_only_first_adoption_plan_omits_powershell(
    tmp_path: Path,
) -> None:
    """A no-PowerShell support fixture does not plan stale PowerShell checks."""
    target_root = materialize_module_fixture(
        tmp_path,
        ("baseline", "template-sync-support"),
        authorize_protected_files=True,
    )
    run_git(target_root, "init", "-q")
    run_git(target_root, "add", ".")

    result = subprocess.run(
        [
            sys.executable,
            ".template-sync/scripts/run_first_adoption_checks.py",
            "--repo-root",
            str(target_root),
            "--plan-only",
        ],
        cwd=target_root,
        check=False,
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stdout + result.stderr
    assert "first_adoption_quality_reports.py line-endings" in result.stdout
    assert "first_adoption_quality_reports.py path-references" in result.stdout
    assert "first_adoption_quality_reports.py powershell" not in result.stdout


@pytest.mark.slow
def test_materialized_gitattributes_baseline_lf_pins_survive_optional_stack_exclusion(
    tmp_path: Path,
) -> None:
    """Baseline-owned LF pins remain after Terraform and PowerShell are excluded."""
    target_root = materialize_module_fixture(
        tmp_path,
        DOWNSTREAM_PYTEST_MODULES,
        authorize_protected_files=True,
    )
    run_git(target_root, "init", "-q")

    attributes_by_path = git_check_attributes_in_repo(
        target_root,
        BASELINE_GITATTRIBUTES_LF_PIN_PATHS,
    )

    for path, attributes in attributes_by_path.items():
        assert attributes["text"] == "set", path
        assert attributes["eol"] == "lf", path


def test_materialized_gitattributes_strips_lfs_when_git_lfs_module_excluded(
    tmp_path: Path,
) -> None:
    """Repositories that do not adopt git-lfs receive no LFS attributes."""
    target_root = materialize_module_fixture(tmp_path, ("baseline",))
    run_git(target_root, "init", "-q")

    gitattributes_text = read_file(target_root / ".gitattributes")
    assert "git-lfs-only" not in gitattributes_text
    assert "filter=lfs" not in gitattributes_text

    attributes_by_path = git_check_attributes_in_repo(
        target_root,
        GIT_LFS_MANAGED_ATTRIBUTE_PATHS,
        GIT_LFS_ATTRIBUTES_TO_CHECK,
    )

    for path, attributes in attributes_by_path.items():
        assert attributes["filter"] == "unspecified", path
        assert attributes["diff"] == "unspecified", path
        assert attributes["merge"] == "unspecified", path
        assert attributes["text"] == "unspecified", path


@pytest.mark.upstream_template_only
def test_materialized_gitattributes_retains_lfs_when_git_lfs_module_included(
    tmp_path: Path,
) -> None:
    """Repositories that adopt git-lfs receive the opt-in LFS attributes."""
    target_root = materialize_module_fixture(tmp_path, ("baseline", "git-lfs"))
    run_git(target_root, "init", "-q")

    gitattributes_text = read_file(target_root / ".gitattributes")
    assert "git-lfs-only" in gitattributes_text
    assert "filter=lfs diff=lfs merge=lfs -text" in gitattributes_text

    attributes_by_path = git_check_attributes_in_repo(
        target_root,
        GIT_LFS_MANAGED_ATTRIBUTE_PATHS,
        GIT_LFS_ATTRIBUTES_TO_CHECK,
    )

    for path, attributes in attributes_by_path.items():
        assert attributes["filter"] == "lfs", path
        assert attributes["diff"] == "lfs", path
        assert attributes["merge"] == "lfs", path
        assert attributes["text"] == "unset", path


def test_materializer_materializes_from_local_template_ref_and_cleans_worktree(
    tmp_path: Path,
) -> None:
    """A locally available template ref is checked out privately and removed."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_repo,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    run_git(template_repo, "branch", "template-main", resolved_sha)

    result = run_materialize_without_template_root(
        target_root,
        "--template-ref",
        "template-main",
        "--template-repo",
        str(template_repo),
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 0, result.stderr
    assert read_file(target_root / "README.md") == "template readme\n"
    assert summary_value(result.stdout, "source ref") == "template-main"
    assert summary_value(result.stdout, "resolved source SHA") == resolved_sha
    assert summary_value(result.stdout, "source repository") == str(template_repo.resolve())
    temporary_checkout_path = Path(summary_value(result.stdout, "temporary checkout path"))
    assert not temporary_checkout_path.exists()
    assert not materializer.is_same_or_descendant(temporary_checkout_path, target_root)
    assert temporary_checkout_path not in git_worktree_paths(template_repo)
    assert "cleanup status: removed" in result.stdout
    assert "\n  - .git\n" not in result.stdout


def test_materializer_materializes_from_full_template_revision(tmp_path: Path) -> None:
    """A full SHA input is resolved, checked out, and reported as a revision."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_repo,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )

    result = run_materialize_without_template_root(
        target_root,
        "--template-revision",
        resolved_sha,
        "--template-repo",
        str(template_repo),
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 0, result.stderr
    assert read_file(target_root / "README.md") == "template readme\n"
    assert summary_value(result.stdout, "source revision") == resolved_sha
    assert summary_value(result.stdout, "resolved source SHA") == resolved_sha
    assert "cleanup status: removed" in result.stdout


def test_git_version_parser_handles_supported_git_version_shapes() -> None:
    """Git capability parsing accepts common platform-specific version strings."""
    assert materializer.parse_git_version("git version 2.43.0") == materializer.GitVersion(
        2,
        43,
    )
    assert materializer.parse_git_version(
        "git version 2.55.0.windows.1"
    ) == materializer.GitVersion(2, 55)
    assert materializer.parse_git_version(
        "git version 2.39.5 (Apple Git-154)"
    ) == materializer.GitVersion(2, 39)
    assert materializer.parse_git_version("git something unexpected") is None


def test_sha256_template_root_is_reported_not_stampable_when_git_supports_it(
    tmp_path: Path,
) -> None:
    """Real SHA-256 repositories are rejected by the current SHA-1 marker contract."""
    template_root = tmp_path / "template"
    template_root.mkdir()
    init_result = subprocess.run(
        [
            "git",
            "-C",
            str(template_root),
            "init",
            "--object-format=sha256",
            "-q",
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    if init_result.returncode != 0:
        pytest.skip(f"Git SHA-256 repositories are unavailable: {init_result.stderr}")
    prepare_template(template_root, [{"pattern": "README.md", "requires_all": ["baseline"]}])
    write_file(template_root / "README.md", "template readme\n")
    run_git(template_root, "add", ".")
    run_git(
        template_root,
        "-c",
        "user.name=Template Tester",
        "-c",
        "user.email=template@example.com",
        "commit",
        "-q",
        "-m",
        "Initial SHA-256 fixture",
    )
    resolved_sha = run_git(template_root, "rev-parse", "HEAD").stdout.strip()
    assert len(resolved_sha) == 64

    detection = materializer.detect_local_template_source(
        template_root,
        materializer.detect_git_version(),
    )

    assert detection.observed_source_sha == resolved_sha
    assert detection.stampable_sha is None
    assert "SHA-256 repositories are not stampable" in cast(
        str,
        detection.not_stampable_reason,
    )


def test_source_git_commands_receive_no_fetch_no_optional_locks_overlay(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Source Git helper calls inherit the environment plus source-safety overrides."""
    captured: dict[str, Any] = {}

    def fake_run(
        command: list[str],
        *,
        check: bool,
        capture_output: bool,
        text: bool,
        env: dict[str, str] | None = None,
    ) -> subprocess.CompletedProcess[str]:
        captured["command"] = command
        captured["env"] = env
        assert check is False
        assert capture_output is True
        assert text is True
        return subprocess.CompletedProcess(command, 0, "", "")

    monkeypatch.setattr(materializer.subprocess, "run", fake_run)

    materializer.run_source_git(tmp_path, ["status"])

    command = cast(list[str], captured["command"])
    env = cast(dict[str, str], captured["env"])
    assert env["GIT_NO_LAZY_FETCH"] == "1"
    assert env["GIT_OPTIONAL_LOCKS"] == "0"
    assert "PATH" in env
    assert "--no-lazy-fetch" not in command
    assert "--no-optional-locks" not in command


def test_template_root_detection_displays_native_windows_worktree_path(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Windows path comparison may casefold, but display keeps the native path."""
    if os.name != "nt":
        pytest.skip("Windows path-display normalization is platform-specific")

    template_root = tmp_path / "MiXeDCaseTemplate"
    template_root.mkdir()
    native_root = str(template_root.resolve(strict=False))
    posix_root = template_root.resolve(strict=False).as_posix()

    def fake_run_source_git(
        _repo_root: Path,
        args: list[str],
        *,
        text: bool = True,
    ) -> subprocess.CompletedProcess[Any]:
        if args == ["rev-parse", "--is-inside-work-tree"]:
            return subprocess.CompletedProcess(args, 0, "true\n", "")
        if args == ["rev-parse", "--show-toplevel"]:
            return subprocess.CompletedProcess(args, 0, f"{posix_root}\n", "")
        if args == ["config", "--local", "--get", "extensions.partialClone"]:
            return subprocess.CompletedProcess(args, 1, "", "")
        if args[:3] == ["config", "--name-only", "--get-regexp"]:
            return subprocess.CompletedProcess(args, 1, "", "")
        if args == ["rev-parse", "--verify", "--end-of-options", "HEAD^{commit}"]:
            return subprocess.CompletedProcess(args, 0, f"{FULL_SHA}\n", "")
        if args == ["-c", "core.fsmonitor=false", "ls-files", "-z", "-v", "--full-name"]:
            assert text is False
            return subprocess.CompletedProcess(args, 0, b"H README.md\0", b"")
        if args == ["-c", "core.fsmonitor=false", "ls-files", "-z", "--stage", "--full-name"]:
            assert text is False
            stdout = f"100644 {FULL_SHA} 0\tREADME.md".encode() + b"\0"
            return subprocess.CompletedProcess(args, 0, stdout, b"")
        if args[-4:] == [
            "status",
            "--porcelain=v1",
            "--untracked-files=all",
            "--ignore-submodules=none",
        ]:
            return subprocess.CompletedProcess(args, 0, "", "")
        raise AssertionError(f"unexpected Git args: {args!r}")

    monkeypatch.setattr(materializer, "run_source_git", fake_run_source_git)

    detection = materializer.detect_local_template_source(
        template_root,
        materializer.GitVersion(2, 45),
    )

    assert detection.stampable_sha == FULL_SHA
    assert detection.source_worktree_root == native_root
    assert detection.source_worktree_root != native_root.casefold()


def test_template_root_top_level_failure_omits_worktree_root_from_summary(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """When top-level discovery fails, the source-root summary line is omitted."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "README.md", "requires_all": ["baseline"]}])
    write_file(template_root / "README.md", "template readme\n")

    monkeypatch.setattr(
        materializer,
        "detect_local_template_source",
        lambda *_args: materializer.LocalSourceDetection(
            not_stampable_reason="unable to identify source worktree root: simulated failure"
        ),
    )

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 0, result.stderr
    assert summary_values(result.stdout, "source worktree root") == []
    source_commit = summary_value(result.stdout, "source commit")
    assert source_commit.startswith("not stampable (unable to identify source worktree root")
    assert_single_physical_line(source_commit)


def test_local_source_diagnostics_keep_multiline_git_stderr_single_line(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Different Git failures surface one physical diagnostic line."""
    template_root = tmp_path / "template"
    template_root.mkdir()

    def fake_top_level_failure(
        _repo_root: Path,
        args: list[str],
        *,
        text: bool = True,
    ) -> subprocess.CompletedProcess[str]:
        assert text is True
        if args == ["rev-parse", "--is-inside-work-tree"]:
            return subprocess.CompletedProcess(args, 0, "true\n", "")
        if args == ["rev-parse", "--show-toplevel"]:
            return subprocess.CompletedProcess(args, 128, "", "first line\nsecond line\n")
        raise AssertionError(f"unexpected Git args: {args!r}")

    monkeypatch.setattr(materializer, "run_source_git", fake_top_level_failure)
    detection = materializer.detect_local_template_source(
        template_root,
        materializer.GitVersion(2, 45),
    )

    top_level_reason = cast(str, detection.not_stampable_reason)
    assert "first line" in top_level_reason
    assert "second line" not in top_level_reason
    assert_single_physical_line(top_level_reason)

    def fake_completeness_failure(
        _repo_root: Path,
        args: list[str],
        *,
        text: bool = True,
    ) -> subprocess.CompletedProcess[bytes]:
        assert args == ["ls-files", "-z", "-v", "--full-name"]
        assert text is False
        return subprocess.CompletedProcess(args, 128, b"", b"alpha\nbeta\n")

    monkeypatch.setattr(materializer, "run_source_git", fake_completeness_failure)
    completeness_reason = cast(
        str, materializer.source_completeness_reason(template_root, git_args_prefix=())
    )

    assert "alpha" in completeness_reason
    assert "beta" not in completeness_reason
    assert_single_physical_line(completeness_reason)


def test_template_root_default_reports_observed_source_without_stamping(
    tmp_path: Path,
) -> None:
    """A clean local template root reports HEAD but does not infer reviewed state."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--included-module",
        "template-sync-support",
    )

    assert result.returncode == 0, result.stderr
    assert summary_values(result.stdout, "source commit") == [
        f"observed {resolved_sha}; not accepted as reviewed state"
    ]
    assert summary_value(result.stdout, "source commit") == (
        f"observed {resolved_sha}; not accepted as reviewed state"
    )
    source_root = summary_value(result.stdout, "source worktree root")
    assert source_root == str(template_root.resolve())
    assert_single_physical_line(source_root)
    assert "resolved source SHA: (not resolved)" not in result.stdout
    assert summary_value(
        result.stdout,
        "accepted reviewed commit for computed marker",
    ).startswith("(none; complete review")
    marker = as_mapping(
        load_yaml(target_root / ".template-sync" / "marker.yml"),
        "marker must be a mapping",
    )
    template_sync = as_mapping(marker["template_sync"], "template_sync must be a mapping")
    assert "last_reviewed_template_commit" not in template_sync


def test_stamp_resolved_template_root_source_persists_reviewed_commit(
    tmp_path: Path,
) -> None:
    """The explicit assertion stamps a clean local source into the written marker."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--included-module",
        "template-sync-support",
        "--stamp-resolved-source-as-reviewed",
    )

    assert result.returncode == 0, result.stderr
    assert summary_values(result.stdout, "source commit") == [
        f"observed {resolved_sha}; accepted for reviewed-state assertion"
    ]
    assert summary_value(result.stdout, "source commit") == (
        f"observed {resolved_sha}; accepted for reviewed-state assertion"
    )
    assert summary_value(result.stdout, "accepted reviewed commit for computed marker") == (
        f"{resolved_sha} (--stamp-resolved-source-as-reviewed)"
    )
    source_section = "\n".join(summary_section_lines(result.stdout, "Source"))
    assert "persist" not in source_section.lower()
    assert "written" not in source_section.lower()
    assert "Marker:" in result.stdout
    marker = as_mapping(
        load_yaml(target_root / ".template-sync" / "marker.yml"),
        "marker must be a mapping",
    )
    template_sync = as_mapping(marker["template_sync"], "template_sync must be a mapping")
    assert template_sync["last_reviewed_template_commit"] == resolved_sha


def test_stamp_resolved_template_root_source_appears_in_preview_without_support(
    tmp_path: Path,
) -> None:
    """Stamping still affects the computed-marker preview when support is excluded."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--stamp-resolved-source-as-reviewed",
    )

    assert result.returncode == 0, result.stderr
    assert not (target_root / ".template-sync" / "marker.yml").exists()
    assert "preview-only: template-sync-support is not included" in result.stdout
    assert f"last_reviewed_template_commit: {resolved_sha}" in result.stdout


def test_stamp_resolved_template_ref_source_persists_reviewed_commit(
    tmp_path: Path,
) -> None:
    """Ref materialization can opt in to stamping its resolved source SHA."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_repo,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    run_git(template_repo, "branch", "template-main", resolved_sha)

    result = run_materialize_without_template_root(
        target_root,
        "--template-ref",
        "template-main",
        "--template-repo",
        str(template_repo),
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--included-module",
        "template-sync-support",
        "--stamp-resolved-source-as-reviewed",
    )

    assert result.returncode == 0, result.stderr
    assert summary_value(result.stdout, "resolved source SHA") == resolved_sha
    assert summary_value(result.stdout, "accepted reviewed commit for computed marker") == (
        f"{resolved_sha} (--stamp-resolved-source-as-reviewed)"
    )
    marker = as_mapping(
        load_yaml(target_root / ".template-sync" / "marker.yml"),
        "marker must be a mapping",
    )
    template_sync = as_mapping(marker["template_sync"], "template_sync must be a mapping")
    assert template_sync["last_reviewed_template_commit"] == resolved_sha


def test_temporary_source_checkout_disables_sparse_checkout_before_worktree_add(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Temporary source checkout creation uses the required top-level -c shape."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    temp_root = tmp_path / "temp"
    template_repo.mkdir()
    target_root.mkdir()
    temp_root.mkdir()
    calls: list[list[str]] = []

    def fake_run_source_git(
        _repo_root: Path,
        args: list[str],
        *,
        text: bool = True,
    ) -> subprocess.CompletedProcess[str]:
        assert text is True
        calls.append(args)
        return subprocess.CompletedProcess(args, 0, "", "")

    monkeypatch.setattr(materializer, "run_source_git", fake_run_source_git)
    monkeypatch.setattr(materializer, "partial_promisor_guard_reason", lambda *_args: None)
    monkeypatch.setattr(
        materializer,
        "verify_source_worktree_stampable",
        lambda *_args, **_kwargs: True,
    )

    source_checkout = materializer.create_temporary_source_checkout(
        template_repo=template_repo,
        target_root=target_root,
        source_mode="template-ref",
        source_value="template-main",
        resolved_source_sha=FULL_SHA,
        temp_root=temp_root,
        git_version=materializer.GitVersion(2, 45),
    )

    worktree_add_args = next(args for args in calls if "worktree" in args)
    assert worktree_add_args[:5] == [
        "-c",
        "core.sparseCheckout=false",
        "worktree",
        "add",
        "--detach",
    ]
    assert worktree_add_args[5] == str(source_checkout.temporary_checkout_path)
    assert worktree_add_args[6] == FULL_SHA


def test_temporary_source_checkout_is_verified_after_worktree_add(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """The detached checkout is verified before the materializer trusts its files."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    temp_root = tmp_path / "temp"
    template_repo.mkdir()
    target_root.mkdir()
    temp_root.mkdir()
    calls: list[str] = []

    def fake_run_source_git(
        _repo_root: Path,
        args: list[str],
        *,
        text: bool = True,
    ) -> subprocess.CompletedProcess[str]:
        assert text is True
        if args[:5] == [
            "-c",
            "core.sparseCheckout=false",
            "worktree",
            "add",
            "--detach",
        ]:
            calls.append("worktree-add")
            return subprocess.CompletedProcess(args, 0, "", "")
        raise AssertionError(f"unexpected Git args: {args!r}")

    def fake_verify(_source_worktree: Path, _git_version: Any, *, fatal: bool) -> bool:
        assert fatal is True
        calls.append("verify")
        return True

    monkeypatch.setattr(materializer, "run_source_git", fake_run_source_git)
    monkeypatch.setattr(materializer, "partial_promisor_guard_reason", lambda *_args: None)
    monkeypatch.setattr(materializer, "verify_source_worktree_stampable", fake_verify)

    materializer.create_temporary_source_checkout(
        template_repo=template_repo,
        target_root=target_root,
        source_mode="template-ref",
        source_value="template-main",
        resolved_source_sha=FULL_SHA,
        temp_root=temp_root,
        git_version=materializer.GitVersion(2, 45),
    )

    assert calls == ["worktree-add", "verify"]


def test_template_ref_materializes_full_checkout_from_sparse_backing_worktree(
    tmp_path: Path,
) -> None:
    """Sparse backing worktrees still produce complete temporary source checkouts."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_repo,
        [
            {"pattern": "README.md", "requires_all": ["baseline"]},
            {"pattern": "docs/template.md", "requires_all": ["baseline"]},
        ],
        {
            "README.md": "template readme\n",
            "docs/template.md": "full checkout content\n",
        },
    )
    sparse_result = subprocess.run(
        [
            "git",
            "-C",
            str(template_repo),
            "sparse-checkout",
            "init",
            "--cone",
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    if sparse_result.returncode != 0:
        pytest.skip(f"Git sparse-checkout is unavailable: {sparse_result.stderr}")
    run_git(template_repo, "sparse-checkout", "set", ".template-sync", "schemas")
    assert not (template_repo / "docs" / "template.md").exists()
    run_git(template_repo, "branch", "template-main", resolved_sha)

    result = run_materialize_without_template_root(
        target_root,
        "--template-ref",
        "template-main",
        "--template-repo",
        str(template_repo),
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 0, result.stderr
    assert read_file(target_root / "docs" / "template.md") == "full checkout content\n"
    assert summary_value(result.stdout, "resolved source SHA") == resolved_sha


def test_stamp_reviewed_helper_aligns_top_level_and_marker_decisions() -> None:
    """Source stamping updates both frozen decision records."""
    decisions = materializer.Decisions(
        source_repo=SOURCE_REPO,
        last_reviewed_template_commit=None,
        included_modules=frozenset({"baseline"}),
        marker_data=materializer.MarkerDecisionData(
            last_reviewed_template_commit=None,
            included_modules=frozenset({"baseline"}),
            local_overrides=(),
            local_path_ownership=(),
            deferred_candidates=(),
            protected_decisions=(),
            protected_guide_contract_waivers=(),
        ),
        raw_marker_fields={},
    )

    stamped = materializer.decisions_with_reviewed_commit(decisions, FULL_SHA)

    assert stamped.last_reviewed_template_commit == FULL_SHA
    assert stamped.marker_data.last_reviewed_template_commit == FULL_SHA


def test_template_root_stamp_fails_when_source_is_not_stampable(
    tmp_path: Path,
) -> None:
    """The opt-in assertion is fatal when no explicit reviewed commit can stand."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    write_file(template_root / "UNTRACKED.txt", "not ignored\n")

    default_result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert default_result.returncode == 0, default_result.stderr
    assert summary_value(default_result.stdout, "source commit") == (
        f"observed {resolved_sha}; not stampable "
        "(the template source has dirty tracked files or untracked non-ignored files)"
    )
    assert summary_value(default_result.stdout, "source worktree root") == str(
        template_root.resolve()
    )

    stamped_result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--stamp-resolved-source-as-reviewed",
    )

    assert stamped_result.returncode == 1
    assert "--stamp-resolved-source-as-reviewed requires a trusted source SHA" in (
        stamped_result.stderr
    )
    assert "dirty tracked files or untracked non-ignored files" in stamped_result.stderr


def test_ignored_untracked_files_do_not_block_template_root_stampability(
    tmp_path: Path,
) -> None:
    """Ignored untracked files are outside the local source stampability gate."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {
            ".gitignore": "ignored-output/\n",
            "README.md": "template readme\n",
        },
    )
    write_file(template_root / "ignored-output" / "cache.tmp", "ignored\n")

    detection = materializer.detect_local_template_source(
        template_root,
        materializer.detect_git_version(),
    )

    assert detection.stampable_sha == resolved_sha

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--stamp-resolved-source-as-reviewed",
    )

    assert result.returncode == 0, result.stderr
    assert summary_value(result.stdout, "source commit") == (
        f"observed {resolved_sha}; accepted for reviewed-state assertion"
    )


def test_template_root_explicit_reviewed_commit_stands_when_source_not_stampable(
    tmp_path: Path,
) -> None:
    """An explicit reviewed commit can stand when a local root cannot be cross-checked."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    write_file(template_root / "UNTRACKED.txt", "not ignored\n")

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--last-reviewed-template-commit",
        FULL_SHA,
        "--included-module",
        "baseline",
        "--included-module",
        "template-sync-support",
        "--stamp-resolved-source-as-reviewed",
    )

    assert result.returncode == 0, result.stderr
    assert summary_value(result.stdout, "accepted reviewed commit for computed marker") == (
        f"{FULL_SHA} (explicit reviewed commit)"
    )
    assert "source HEAD could not be cross-checked" in result.stdout
    marker = as_mapping(
        load_yaml(target_root / ".template-sync" / "marker.yml"),
        "marker must be a mapping",
    )
    template_sync = as_mapping(marker["template_sync"], "template_sync must be a mapping")
    assert template_sync["last_reviewed_template_commit"] == FULL_SHA


def test_template_root_subdirectory_is_not_silently_promoted(tmp_path: Path) -> None:
    """A subdirectory inside a Git worktree is outside local source stampability."""
    template_root = tmp_path / "template"
    resolved_sha = prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n", "nested/README.md": "nested\n"},
    )
    detection = materializer.detect_local_template_source(
        template_root / "nested",
        materializer.detect_git_version(),
    )

    assert detection.observed_source_sha is None
    assert detection.stampable_sha is None
    assert detection.source_worktree_root == str(template_root.resolve())
    assert resolved_sha
    assert "not the Git worktree top level" in cast(str, detection.not_stampable_reason)
    assert_single_physical_line(cast(str, detection.not_stampable_reason))


def test_non_git_template_root_is_not_stampable(tmp_path: Path) -> None:
    """Non-Git roots are classified by Git exit code, not stdout text."""
    template_root = tmp_path / "template"
    template_root.mkdir()

    detection = materializer.detect_local_template_source(
        template_root,
        materializer.detect_git_version(),
    )

    assert detection.observed_source_sha is None
    assert detection.stampable_sha is None
    assert detection.not_stampable_reason == "the template root is not a Git worktree"


def test_unborn_template_root_head_has_generic_unresolvable_reason(tmp_path: Path) -> None:
    """An unborn HEAD is reported generically without matching Git stderr text."""
    template_root = tmp_path / "template"
    template_root.mkdir()
    run_git(template_root, "init", "-q")

    detection = materializer.detect_local_template_source(
        template_root,
        materializer.detect_git_version(),
    )

    assert detection.observed_source_sha is None
    assert detection.stampable_sha is None
    assert detection.not_stampable_reason == ("template source HEAD is not resolvable to a commit")


def test_sparse_checkout_hidden_tracked_files_are_not_stampable(tmp_path: Path) -> None:
    """Sparse checkout can hide tracked files from status but not from ls-files -v."""
    template_root = tmp_path / "template"
    prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {
            "README.md": "template readme\n",
            "docs/omitted.md": "omitted tracked content\n",
        },
    )
    sparse_result = subprocess.run(
        [
            "git",
            "-C",
            str(template_root),
            "sparse-checkout",
            "init",
            "--cone",
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    if sparse_result.returncode != 0:
        pytest.skip(f"Git sparse-checkout is unavailable: {sparse_result.stderr}")
    run_git(template_root, "sparse-checkout", "set", ".template-sync", "schemas")

    status_result = run_git(
        template_root,
        "status",
        "--porcelain=v1",
        "--untracked-files=all",
        "--ignore-submodules=none",
    )
    assert status_result.stdout == ""

    detection = materializer.detect_local_template_source(
        template_root,
        materializer.detect_git_version(),
    )

    assert detection.stampable_sha is None
    assert "skip-worktree" in cast(str, detection.not_stampable_reason)


@pytest.mark.parametrize(
    ("index_flag", "expected_reason"),
    [
        pytest.param("--skip-worktree", "skip-worktree", id="skip-worktree"),
        pytest.param("--assume-unchanged", "assume-unchanged", id="assume-unchanged"),
    ],
)
def test_source_completeness_rejects_hidden_tracked_file_states(
    tmp_path: Path,
    index_flag: str,
    expected_reason: str,
) -> None:
    """Skip-worktree and assume-unchanged entries are not stampable."""
    template_root = tmp_path / "template"
    prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    run_git(template_root, "update-index", index_flag, "README.md")
    if index_flag == "--skip-worktree":
        (template_root / "README.md").unlink()
    else:
        write_file(template_root / "README.md", "hidden local change\n")

    status_result = run_git(
        template_root,
        "status",
        "--porcelain=v1",
        "--untracked-files=all",
        "--ignore-submodules=none",
    )
    assert status_result.stdout == ""

    detection = materializer.detect_local_template_source(
        template_root,
        materializer.detect_git_version(),
    )

    assert detection.stampable_sha is None
    assert expected_reason in cast(str, detection.not_stampable_reason)


def test_source_completeness_probe_uses_binary_v_records_without_f(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Completeness parsing uses NUL-delimited binary -v output, never -f tags."""
    calls: list[tuple[list[str], bool]] = []
    non_ascii_path = "café file.txt"

    def fake_run_source_git(
        _repo_root: Path,
        args: list[str],
        *,
        text: bool = True,
    ) -> subprocess.CompletedProcess[bytes]:
        calls.append((args, text))
        if args == ["ls-files", "-z", "-v", "--full-name"]:
            assert text is False
            return subprocess.CompletedProcess(
                args,
                0,
                f"H {non_ascii_path}".encode() + b"\0",
                b"",
            )
        if args == ["ls-files", "-z", "--stage", "--full-name"]:
            assert text is False
            return subprocess.CompletedProcess(
                args,
                0,
                f"100644 {FULL_SHA} 0\t{non_ascii_path}".encode() + b"\0",
                b"",
            )
        raise AssertionError(f"unexpected Git args: {args!r}")

    monkeypatch.setattr(materializer, "run_source_git", fake_run_source_git)

    assert materializer.source_completeness_reason(tmp_path, git_args_prefix=()) is None
    assert (["ls-files", "-z", "-v", "--full-name"], False) in calls
    for args, _text in calls:
        assert "-f" not in args


def test_source_completeness_rejects_gitlinks(tmp_path: Path) -> None:
    """A source index that records submodule gitlinks is outside stampability."""
    template_root = tmp_path / "template"
    prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    run_git(
        template_root,
        "update-index",
        "--add",
        "--cacheinfo",
        f"160000,{FULL_SHA},sub modules/template lib",
    )
    run_git(
        template_root,
        "-c",
        "user.name=Template Tester",
        "-c",
        "user.email=template@example.com",
        "commit",
        "-q",
        "-m",
        "Add gitlink",
    )
    (template_root / "sub modules" / "template lib").mkdir(parents=True)
    status_result = run_git(
        template_root,
        "status",
        "--porcelain=v1",
        "--untracked-files=all",
        "--ignore-submodules=none",
    )
    assert status_result.stdout == ""

    detection = materializer.detect_local_template_source(
        template_root,
        materializer.detect_git_version(),
    )

    assert detection.stampable_sha is None
    assert "submodule gitlinks" in cast(str, detection.not_stampable_reason)


def test_template_ref_rejects_commit_that_records_gitlink(tmp_path: Path) -> None:
    """A selected ref whose tree records a gitlink is fatal for temporary checkouts."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_repo,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    run_git(
        template_repo,
        "update-index",
        "--add",
        "--cacheinfo",
        f"160000,{resolved_sha},sub modules/template lib",
    )
    run_git(
        template_repo,
        "-c",
        "user.name=Template Tester",
        "-c",
        "user.email=template@example.com",
        "commit",
        "-q",
        "-m",
        "Add gitlink",
    )
    gitlink_sha = run_git(template_repo, "rev-parse", "HEAD").stdout.strip()
    run_git(template_repo, "branch", "template-main", gitlink_sha)

    result = run_materialize_without_template_root(
        target_root,
        "--template-ref",
        "template-main",
        "--template-repo",
        str(template_repo),
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 1
    assert "submodule gitlinks" in result.stderr


@pytest.mark.parametrize(
    ("stage_stdout", "expected_reason"),
    [
        pytest.param(
            b"100644 abcdef 0 path-without-tab\0",
            "unable to parse source index metadata",
            id="missing-tab",
        ),
        pytest.param(
            b"100644\tREADME.md\0",
            "unable to parse source index metadata",
            id="short-metadata",
        ),
        pytest.param(
            b"not-a-mode abcdef 0\tREADME.md\0",
            "unable to parse source index mode",
            id="unreadable-mode",
        ),
    ],
)
def test_source_gitlink_stage_parser_treats_malformed_records_as_indeterminate(
    tmp_path: Path,
    monkeypatch: Any,
    stage_stdout: bytes,
    expected_reason: str,
) -> None:
    """Malformed NUL-delimited stage records are indeterminate, not silently clean."""

    def fake_run_source_git(
        _repo_root: Path,
        args: list[str],
        *,
        text: bool = True,
    ) -> subprocess.CompletedProcess[bytes]:
        if args == ["ls-files", "-z", "-v", "--full-name"]:
            assert text is False
            return subprocess.CompletedProcess(args, 0, b"H README.md\0", b"")
        if args == ["ls-files", "-z", "--stage", "--full-name"]:
            assert text is False
            return subprocess.CompletedProcess(args, 0, stage_stdout, b"")
        raise AssertionError(f"unexpected Git args: {args!r}")

    monkeypatch.setattr(materializer, "run_source_git", fake_run_source_git)

    reason = cast(str, materializer.source_completeness_reason(tmp_path, git_args_prefix=()))

    assert expected_reason in reason
    assert_single_physical_line(reason)


def test_stampability_backstop_runs_before_status_probe(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Completeness and gitlink backstops are evaluated before ordinary status."""
    calls: list[str] = []

    def fake_partial_guard(_repo_root: Path, _git_version: Any) -> None:
        calls.append("partial")

    def fake_safety(_repo_root: Path, _git_version: Any) -> tuple[tuple[str, ...], None]:
        calls.append("fsmonitor")
        return ("-c", "core.fsmonitor=false"), None

    def fake_completeness(_source_worktree: Path, *, git_args_prefix: Any) -> str:
        assert git_args_prefix == ("-c", "core.fsmonitor=false")
        calls.append("completeness")
        return "the template source records submodule gitlinks, which are not supported"

    def fake_status(_repo_root: Path, *, git_args_prefix: Any) -> str | None:
        calls.append("status")
        return "ordinary status should not decide this fixture"

    monkeypatch.setattr(materializer, "partial_promisor_guard_reason", fake_partial_guard)
    monkeypatch.setattr(materializer, "source_fsmonitor_safety", fake_safety)
    monkeypatch.setattr(materializer, "source_completeness_reason", fake_completeness)
    monkeypatch.setattr(materializer, "status_probe_not_stampable_reason", fake_status)

    assert (
        materializer.verify_source_worktree_stampable(
            tmp_path,
            materializer.GitVersion(2, 45),
            fatal=False,
        )
        is False
    )

    assert calls == ["partial", "fsmonitor", "completeness"]


def copy_materializer_runtime_for_mutation(root: Path) -> Path:
    """Copy the real split runtime into an isolated repository-shaped tool tree."""
    script_dir = root / ".template-sync/scripts"
    script_dir.mkdir(parents=True)
    for source in SCRIPT_DIR.glob("*.py"):
        shutil.copyfile(source, script_dir / source.name)
    shared_dir = root / ".github/scripts"
    shared_dir.mkdir(parents=True)
    for name in ("instruction_contract_core.py", "instruction_contract_support.py"):
        shutil.copyfile(REPO_ROOT / ".github/scripts" / name, shared_dir / name)
    return script_dir


@pytest.mark.parametrize("entry", ["detect", "verify", "completeness-mutant"])
def test_source_index_queries_disable_native_fsmonitor(
    tmp_path: Path, monkeypatch: Any, entry: str
) -> None:
    """Completeness and status cannot execute a source-local filesystem monitor."""
    template_root = tmp_path / "template"
    expected_sha = prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    version = materializer.detect_git_version()
    if not materializer.git_version_at_least(version, 2, 36):
        pytest.skip("The modern Git override requires Git >=2.36")
    hook = template_root / ".git" / "test-fsmonitor"
    hook.write_text(
        "#!/bin/sh\nprintf invoked > .git/fsmonitor-invoked\nexit 1\n", encoding="utf-8"
    )
    hook.chmod(0o755)
    run_git(template_root, "config", "core.fsmonitor", ".git/test-fsmonitor")
    sentinel = template_root / ".git" / "fsmonitor-invoked"
    run_git(template_root, "ls-files")
    assert sentinel.read_text(encoding="utf-8") == "invoked"
    sentinel.unlink()
    real_safety = materializer.source_fsmonitor_safety
    safety_calls: list[Path] = []

    def safety(root: Path, git_version: Any) -> Any:
        safety_calls.append(root)
        return real_safety(root, git_version)

    monkeypatch.setattr(materializer, "source_fsmonitor_safety", safety)
    if entry == "verify":
        assert materializer.verify_source_worktree_stampable(template_root, version, fatal=True)
    else:
        detection = materializer.detect_local_template_source(template_root, version)
        assert detection.stampable_sha == expected_sha, detection.not_stampable_reason
    assert safety_calls == [template_root]
    assert not sentinel.exists()
    if entry == "completeness-mutant":
        # Remove only completeness protection from an isolated copy; leave status intact.
        mutant_dir = copy_materializer_runtime_for_mutation(tmp_path / "mutant")
        mutant = mutant_dir / SCRIPT_PATH.name
        source = mutant.read_text(encoding="utf-8")
        guard = '[*git_args_prefix, "ls-files",'
        assert source.count(guard) == 2
        mutant.write_text(source.replace(guard, '["ls-files",'), encoding="utf-8")
        program = (
            "import sys; from pathlib import Path; sys.path.insert(0, sys.argv[1]); "
            "import materialize_downstream_adoption as m; "
            "result=m.detect_local_template_source(Path(sys.argv[2]), m.detect_git_version()); "
            "assert result.stampable_sha, result.not_stampable_reason"
        )
        result = subprocess.run(
            [sys.executable, "-c", program, str(mutant_dir), str(template_root)],
            check=False,
            capture_output=True,
            text=True,
        )
        assert result.returncode == 0, result.stdout + result.stderr
        assert sentinel.exists()


@pytest.mark.parametrize("version", [materializer.GitVersion(2, 35), None])
@pytest.mark.parametrize("setting", [None, "", "false", ".git/test-fsmonitor"])
@pytest.mark.parametrize("fatal", [False, True])
def test_source_legacy_fsmonitor_gate_precedes_index(
    tmp_path: Path, monkeypatch: Any, version: Any, setting: str | None, fatal: bool
) -> None:
    """All configured legacy values block before completeness; absence remains supported."""
    template_root = tmp_path / "template"
    prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    if setting is not None:
        run_git(template_root, "config", "core.fsmonitor", setting)
    real_runner = materializer.run_source_git
    index_calls: list[list[str]] = []

    def runner(root: Path, args: list[str], **kwargs: Any) -> Any:
        if "ls-files" in args or "status" in args:
            index_calls.append(args)
        return real_runner(root, args, **kwargs)

    monkeypatch.setattr(materializer, "run_source_git", runner)
    if setting is not None and fatal:
        with pytest.raises(materializer.MaterializationError, match="core.fsmonitor is configured"):
            materializer.verify_source_worktree_stampable(template_root, version, fatal=fatal)
    else:
        assert materializer.verify_source_worktree_stampable(
            template_root, version, fatal=fatal
        ) == (setting is None)
    assert bool(index_calls) == (setting is None)


def test_old_git_with_configured_fsmonitor_is_not_stampable(tmp_path: Path) -> None:
    """Configured fsmonitor is not trusted when Git cannot be parsed as >=2.36."""
    template_root = tmp_path / "template"
    prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    run_git(template_root, "config", "core.fsmonitor", "true")

    detection = materializer.detect_local_template_source(
        template_root,
        materializer.GitVersion(2, 35),
    )

    assert detection.stampable_sha is None
    reason = cast(str, detection.not_stampable_reason)
    assert "Git to be detected as >=2.36" in reason
    assert "--last-reviewed-template-commit FULL_SHA" in reason


@pytest.mark.parametrize(
    "git_version",
    [
        pytest.param(materializer.GitVersion(2, 35), id="old-git"),
        pytest.param(None, id="unparseable-git"),
    ],
)
def test_old_or_unparseable_git_reads_effective_global_fsmonitor(
    tmp_path: Path,
    monkeypatch: Any,
    git_version: Any,
) -> None:
    """The fsmonitor safety check reads effective config, including global config."""
    global_config = isolate_git_config(monkeypatch, tmp_path)
    write_git_config_file(global_config, "core.fsmonitor", "true")
    template_root = tmp_path / "template"
    prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )

    detection = materializer.detect_local_template_source(template_root, git_version)

    assert detection.stampable_sha is None
    reason = cast(str, detection.not_stampable_reason)
    assert "Git to be detected as >=2.36" in reason
    assert "--last-reviewed-template-commit FULL_SHA" in reason
    assert_single_physical_line(reason)


def test_partial_promisor_guard_uses_effective_remote_config(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """The partial/promisor guard follows final effective remote config values."""
    isolate_git_config(monkeypatch, tmp_path)
    template_root = tmp_path / "template"
    prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    old_git = materializer.GitVersion(2, 44)

    assert materializer.partial_promisor_guard_reason(template_root, old_git) is None

    run_git(template_root, "config", "--add", "remote.origin.promisor", "true")
    assert "remote.origin.promisor=true" in cast(
        str,
        materializer.partial_promisor_guard_reason(template_root, old_git),
    )
    run_git(template_root, "config", "--add", "remote.origin.promisor", "false")
    assert materializer.partial_promisor_guard_reason(template_root, old_git) is None

    run_git(template_root, "config", "--add", "remote.origin.partialCloneFilter", "blob:none")
    filter_reason = cast(
        str,
        materializer.partial_promisor_guard_reason(template_root, old_git),
    )
    assert "remote.origin.partialclonefilter" in filter_reason.lower()
    run_git(template_root, "config", "--add", "remote.origin.partialCloneFilter", "")
    assert materializer.partial_promisor_guard_reason(template_root, old_git) is None


def test_partial_promisor_guard_honors_isolated_config_precedence(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Higher-priority effective config overrides lower promisor/filter config."""
    global_config = isolate_git_config(monkeypatch, tmp_path)
    template_root = tmp_path / "template"
    prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    old_git = materializer.GitVersion(2, 44)

    write_git_config_file(global_config, "remote.origin.promisor", "true")
    assert "remote.origin.promisor=true" in cast(
        str,
        materializer.partial_promisor_guard_reason(template_root, old_git),
    )
    run_git(template_root, "config", "remote.origin.promisor", "false")
    assert materializer.partial_promisor_guard_reason(template_root, old_git) is None

    write_git_config_file(global_config, "remote.blobs.partialCloneFilter", "blob:none")
    assert (
        "remote.blobs.partialclonefilter"
        in cast(
            str,
            materializer.partial_promisor_guard_reason(template_root, old_git),
        ).lower()
    )
    run_git(template_root, "config", "remote.blobs.partialCloneFilter", "")
    assert materializer.partial_promisor_guard_reason(template_root, old_git) is None

    run_git(template_root, "config", "remote.bad.promisor", "not-a-bool")
    malformed_reason = cast(
        str,
        materializer.partial_promisor_guard_reason(template_root, old_git),
    )
    assert "unable to determine promisor remote state from remote.bad.promisor" in (
        malformed_reason
    )
    assert_single_physical_line(malformed_reason)


def test_partial_promisor_guard_reads_command_scope_config(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """GIT_CONFIG_COUNT command-scope values are visible and highest priority."""
    global_config = isolate_git_config(monkeypatch, tmp_path)
    template_root = tmp_path / "template"
    prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    old_git = materializer.GitVersion(2, 44)

    write_git_config_file(global_config, "remote.origin.promisor", "true")
    monkeypatch.setenv("GIT_CONFIG_COUNT", "1")
    monkeypatch.setenv("GIT_CONFIG_KEY_0", "remote.origin.promisor")
    monkeypatch.setenv("GIT_CONFIG_VALUE_0", "false")
    assert materializer.partial_promisor_guard_reason(template_root, old_git) is None

    run_git(template_root, "config", "remote.command.promisor", "false")
    monkeypatch.setenv("GIT_CONFIG_COUNT", "2")
    monkeypatch.setenv("GIT_CONFIG_KEY_1", "remote.command.promisor")
    monkeypatch.setenv("GIT_CONFIG_VALUE_1", "true")
    command_reason = cast(
        str,
        materializer.partial_promisor_guard_reason(template_root, old_git),
    )
    assert "remote.command.promisor=true" in command_reason

    monkeypatch.setenv("GIT_CONFIG_KEY_1", "remote.command.partialCloneFilter")
    monkeypatch.setenv("GIT_CONFIG_VALUE_1", "blob:none")
    filter_reason = cast(
        str,
        materializer.partial_promisor_guard_reason(template_root, old_git),
    )
    assert "remote.command.partialclonefilter" in filter_reason.lower()


def test_partial_promisor_guard_rejects_local_partial_clone_extension(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Repository-local extensions.partialClone independently triggers the guard."""
    isolate_git_config(monkeypatch, tmp_path)
    template_root = tmp_path / "template"
    prepare_git_template(
        template_root,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    run_git(template_root, "config", "extensions.partialClone", "origin")

    reason = cast(
        str,
        materializer.partial_promisor_guard_reason(
            template_root,
            materializer.GitVersion(2, 44),
        ),
    )

    assert "extensions.partialClone" in reason
    assert_single_physical_line(reason)


def test_backing_template_repository_partial_guard_runs_before_ref_resolution(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Backing repositories are guarded before selected-ref lookup or worktree add."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_git_template(
        template_repo,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    calls: list[str] = []

    def fake_partial_guard(_repo_root: Path, _git_version: Any) -> str:
        calls.append("partial")
        return "the template source has remote.origin.promisor=true"

    def fail_resolve(*_args: Any, **_kwargs: Any) -> str:
        calls.append("resolve")
        raise AssertionError("ref resolution must not run after partial guard failure")

    def fail_checkout(*_args: Any, **_kwargs: Any) -> Any:
        calls.append("checkout")
        raise AssertionError("worktree add must not run after partial guard failure")

    monkeypatch.setattr(materializer, "detect_git_version", lambda: materializer.GitVersion(2, 44))
    monkeypatch.setattr(materializer, "partial_promisor_guard_reason", fake_partial_guard)
    monkeypatch.setattr(materializer, "resolve_template_commit", fail_resolve)
    monkeypatch.setattr(materializer, "create_temporary_source_checkout", fail_checkout)

    args = materializer.parse_args(
        [
            "--target-root",
            str(target_root),
            "--template-ref",
            "template-main",
            "--template-repo",
            str(template_repo),
            "--source-repo",
            SOURCE_REPO,
            "--included-module",
            "baseline",
        ]
    )

    with pytest.raises(materializer.MaterializationError) as error:
        materializer.resolve_template_source(args, target_root)

    assert "remote.origin.promisor=true" in str(error.value)
    assert calls == ["partial"]


def test_args_file_stamp_true_can_be_overridden_by_cli_no(tmp_path: Path) -> None:
    """The BooleanOptionalAction false form wins over an args-file true value."""
    args_file = tmp_path / "materialize.args.json"
    write_json(
        args_file,
        {
            "target_root": str(tmp_path / "target"),
            "source_repo": SOURCE_REPO,
            "included_modules": ["baseline"],
            "stamp_resolved_source_as_reviewed": True,
        },
    )

    args = materializer.parse_args(
        [
            "--args-file",
            str(args_file),
            "--no-stamp-resolved-source-as-reviewed",
        ]
    )

    assert args.stamp_resolved_source_as_reviewed is False


def test_materializer_rejects_invalid_template_ref(tmp_path: Path) -> None:
    """The materializer resolves refs locally and fails without fetching."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_git_template(
        template_repo,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )

    result = run_materialize_without_template_root(
        target_root,
        "--template-ref",
        "missing-template-ref",
        "--template-repo",
        str(template_repo),
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 1
    assert "Unable to resolve --template-ref 'missing-template-ref'" in result.stderr


def test_materializer_rejects_template_root_with_ref(tmp_path: Path) -> None:
    """Explicit source root and source ref modes are mutually exclusive."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_repo, [{"pattern": "README.md", "requires_all": ["baseline"]}])

    result = run_materialize_without_template_root(
        target_root,
        "--template-root",
        str(template_repo),
        "--template-ref",
        "template-main",
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 2
    assert "not allowed with argument" in result.stderr


def test_materializer_rejects_template_temp_root_inside_target(tmp_path: Path) -> None:
    """Operator-supplied temporary checkout roots must stay outside the target."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_repo,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    run_git(template_repo, "branch", "template-main", resolved_sha)

    result = run_materialize_without_template_root(
        target_root,
        "--template-ref",
        "template-main",
        "--template-repo",
        str(template_repo),
        "--template-temp-root",
        str(target_root / ".tmp"),
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 1
    assert "--template-temp-root must not be inside --target-root" in result.stderr


def test_materializer_rejects_reviewed_sha_mismatch(tmp_path: Path) -> None:
    """Reviewed commits must match the resolved source SHA when both are supplied."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_repo,
        [{"pattern": "README.md", "requires_all": ["baseline"]}],
        {"README.md": "template readme\n"},
    )
    run_git(template_repo, "branch", "template-main", resolved_sha)
    mismatched_sha = "f" * 40
    assert mismatched_sha != resolved_sha

    result = run_materialize_without_template_root(
        target_root,
        "--template-ref",
        "template-main",
        "--template-repo",
        str(template_repo),
        "--source-repo",
        SOURCE_REPO,
        "--last-reviewed-template-commit",
        mismatched_sha,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 1
    assert "does not match the resolved source SHA" in result.stderr
    assert "Omit the reviewed value until review is complete" in result.stderr


def test_materializer_failure_cleans_temporary_worktree(tmp_path: Path) -> None:
    """A materialization failure still removes the tool-created source worktree."""
    template_repo = tmp_path / "template-repo"
    target_root = tmp_path / "target"
    target_root.mkdir()
    resolved_sha = prepare_git_template(
        template_repo,
        [{"pattern": "SECURITY.md", "requires_all": ["baseline"]}],
        {"SECURITY.md": "Email [security contact email]\n"},
    )
    run_git(template_repo, "branch", "template-main", resolved_sha)
    before_paths = git_worktree_paths(template_repo)

    result = run_materialize_without_template_root(
        target_root,
        "--template-ref",
        "template-main",
        "--template-repo",
        str(template_repo),
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--repository",
        "octo/widget",
        "--security-contact",
        "security@example.com",
    )

    assert result.returncode == 1
    assert "Placeholder helper failed" in result.stderr
    assert git_worktree_paths(template_repo) == before_paths


def test_cleanup_retries_once_and_verifies_worktree_absence(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Cleanup retries the remove command once before verifying worktree state."""
    template_repo = tmp_path / "template-repo"
    template_repo.mkdir()
    checkout_path = tmp_path / "source-checkout"
    source_checkout = materializer.SourceCheckout(
        template_root=checkout_path,
        template_repo=template_repo,
        temporary_parent=tmp_path,
        temporary_checkout_path=checkout_path,
        summary=materializer.SourceSummary(
            target_root=str(tmp_path / "target"),
            template_root=str(checkout_path),
            source_mode="template-ref",
            source_value="template-main",
            resolved_source_sha="a" * 40,
            source_repository=str(template_repo),
            temporary_checkout_path=str(checkout_path),
            cleanup_status="pending",
            manual_cleanup_command="git worktree remove --force source-checkout",
        ),
    )
    remove_calls = 0

    def fake_run_git(
        _repo_root: Path,
        args: list[str],
        **_kwargs: object,
    ) -> subprocess.CompletedProcess[str]:
        nonlocal remove_calls
        if args[:3] == ["worktree", "remove", "--force"]:
            remove_calls += 1
            if remove_calls == 1:
                return subprocess.CompletedProcess(args, 1, "", "locked")
            return subprocess.CompletedProcess(args, 0, "", "")
        if args == ["worktree", "list", "--porcelain"]:
            return subprocess.CompletedProcess(args, 0, "worktree /other/path\n", "")
        raise AssertionError(f"unexpected git args: {args}")

    monkeypatch.setattr(materializer, "run_git", fake_run_git)

    failure = materializer.cleanup_source_checkout(source_checkout)

    assert failure is None
    assert remove_calls == 2
    assert source_checkout.summary.cleanup_status == "removed after retry"


def test_successful_materialization_cleanup_failure_exits_dedicated_code(
    tmp_path: Path,
    monkeypatch: Any,
    capsys: Any,
) -> None:
    """Successful materialization plus cleanup failure has a dedicated exit code."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "README.md", "requires_all": ["baseline"]}])
    write_file(template_root / "README.md", "template readme\n")

    def fake_cleanup(source_checkout: Any) -> Any:
        source_checkout.summary.cleanup_status = "failed"
        source_checkout.summary.cleanup_failure = "simulated cleanup failure"
        source_checkout.summary.temporary_checkout_path = str(tmp_path / "stale-worktree")
        source_checkout.summary.source_repository = str(template_root)
        source_checkout.summary.manual_cleanup_command = "git worktree remove --force stale"
        return materializer.CleanupFailure(
            detail="simulated cleanup failure",
            residual_worktree_path=tmp_path / "stale-worktree",
            source_repository=template_root,
            manual_cleanup_command="git worktree remove --force stale",
        )

    monkeypatch.setattr(materializer, "cleanup_source_checkout", fake_cleanup)

    exit_code = materializer.main(
        [
            "--template-root",
            str(template_root),
            "--target-root",
            str(target_root),
            "--source-repo",
            SOURCE_REPO,
            "--included-module",
            "baseline",
        ]
    )
    captured = capsys.readouterr()

    assert exit_code == materializer.EXIT_CLEANUP_FAILURE
    assert read_file(target_root / "README.md") == "template readme\n"
    assert "cleanup status: failed" in captured.out
    assert "Target tree was materialized successfully" in captured.err
    assert "Manual cleanup command" in captured.err


def test_materialization_failure_plus_cleanup_failure_preserves_primary_error(
    tmp_path: Path,
    monkeypatch: Any,
    capsys: Any,
) -> None:
    """Cleanup failure diagnostics do not replace the materialization failure."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "SECURITY.md", "requires_all": ["baseline"]}])
    write_file(template_root / "SECURITY.md", "Email [security contact email]\n")

    def fake_cleanup(source_checkout: Any) -> Any:
        source_checkout.summary.cleanup_status = "failed"
        return materializer.CleanupFailure(
            detail="simulated cleanup failure",
            residual_worktree_path=tmp_path / "stale-worktree",
            source_repository=template_root,
            manual_cleanup_command="git worktree remove --force stale",
        )

    monkeypatch.setattr(materializer, "cleanup_source_checkout", fake_cleanup)

    exit_code = materializer.main(
        [
            "--template-root",
            str(template_root),
            "--target-root",
            str(target_root),
            "--source-repo",
            SOURCE_REPO,
            "--included-module",
            "baseline",
            "--repository",
            "octo/widget",
            "--security-contact",
            "security@example.com",
        ]
    )
    captured = capsys.readouterr()

    assert exit_code == materializer.EXIT_RUNTIME_FAILURE
    assert "Placeholder helper failed" in captured.err
    assert captured.err.index("Placeholder helper failed") < captured.err.index(
        "Temporary source checkout cleanup also failed"
    )


def test_retained_files_excluded_files_inline_blocks_and_unmapped_paths(
    tmp_path: Path,
) -> None:
    """Retained module files are copied, excluded modules are absent, and blocks prune."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {"pattern": "README.md", "requires_all": ["baseline"]},
            {"pattern": "CONTRIBUTING.md", "requires_all": ["baseline"]},
            {"pattern": "templates/python/**", "requires_all": ["python"]},
        ],
    )
    write_file(
        template_root / "README.md",
        "top\n# template-sync: begin python-only\npython\n# template-sync: end python-only\nbottom\n",
    )
    write_file(template_root / "CONTRIBUTING.md", "contributing\n")
    write_file(template_root / "templates" / "python" / "app.py", "print('hi')\n")
    write_file(template_root / "UNMAPPED.txt", "unmapped\n")

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 0, result.stderr
    assert read_file(target_root / "README.md") == "top\nbottom\n"
    assert read_file(target_root / "CONTRIBUTING.md") == "contributing\n"
    assert not (target_root / "templates" / "python" / "app.py").exists()
    assert "templates/python/app.py" in result.stdout
    assert "UNMAPPED.txt" in result.stdout


@pytest.mark.parametrize(
    ("included_modules", "case_name"),
    [
        pytest.param(
            (
                "baseline",
                "agent-instructions",
                "github-platform",
                "github-actions",
                "github-templates",
                "template-sync-support",
                "markdown",
                "powershell",
            ),
            "issue-690-powershell-retained",
            id="issue-690-powershell-retained",
        ),
        pytest.param(
            (
                "baseline",
                "agent-instructions",
                "github-platform",
                "github-actions",
                "github-templates",
                "template-sync-support",
                "markdown",
            ),
            "powershell-reference-stripped",
            id="powershell-reference-stripped",
        ),
    ],
)
@pytest.mark.slow
def test_materialized_template_update_procedure_passes_nested_markdown_lint(
    tmp_path: Path,
    included_modules: tuple[str, ...],
    case_name: str,
) -> None:
    """Partial materialization produces a nested-lint-clean sync procedure."""
    if not nested_markdown_lint_prerequisites_available():
        pytest.skip("Run npm ci --ignore-scripts before this generated-output lint test.")

    target_root = tmp_path / case_name
    target_root.mkdir()
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            list(included_modules),
            protected_file_decisions=protected_take_decisions_for_modules(included_modules),
        ),
    )
    module_args = [
        argument
        for module_name in included_modules
        for argument in ("--included-module", module_name)
    ]

    result = run_materialize(
        REPO_ROOT,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--last-reviewed-template-commit",
        FULL_SHA,
        "--repository",
        "octocat/hello-world",
        "--security-contact",
        "security@example.com",
        "--decisions-file",
        "decisions.yml",
        *module_args,
    )

    assert result.returncode == 0, result.stderr
    generated_procedure = target_root / "TEMPLATE_UPDATE_PROCEDURE.md"
    assert generated_procedure.is_file(), result.stdout

    lint_result = subprocess.run(
        [
            "node",
            str(target_root / ".github/scripts/lint-nested-markdown.js"),
            str(generated_procedure),
        ],
        cwd=target_root,
        env={**os.environ, "NODE_PATH": str(REPO_ROOT / "node_modules")},
        check=False,
        capture_output=True,
        text=True,
    )

    assert lint_result.returncode == 0, lint_result.stdout + lint_result.stderr


@pytest.mark.upstream_template_only
def test_materialized_github_powershell_profile_records_protected_guide_waivers(
    tmp_path: Path,
) -> None:
    """A GitHub-hosted PowerShell profile validates with explicit protected-guide waivers."""
    target_root = tmp_path / "github-powershell"
    target_root.mkdir()
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            list(GITHUB_POWERSHELL_PROFILE_MODULES),
            protected_file_decisions=protected_take_decisions_for_modules(
                GITHUB_POWERSHELL_PROFILE_MODULES
            ),
            protected_guide_contract_waivers=github_powershell_protected_guide_waivers(),
        ),
    )

    result = run_materialize(REPO_ROOT, target_root, "--decisions-file", "decisions.yml")

    assert result.returncode == 0, result.stdout + result.stderr
    marker = as_mapping(
        load_yaml(target_root / ".template-sync" / "marker.yml"),
        "marker must be a mapping",
    )
    template_sync = as_mapping(marker["template_sync"], "template_sync must be a mapping")
    protected_decisions = as_list(
        template_sync.get("protected_file_decisions"),
        "protected decisions must be a list",
    )
    protected_guide_waivers = as_list(
        template_sync.get("protected_guide_contract_waivers"),
        "protected guide waivers must be a list",
    )
    protected_decision_paths = {
        as_mapping(record, "protected decision must be a mapping")["path"]
        for record in protected_decisions
    }
    protected_guide_waiver_keys = {
        as_mapping(record, "protected guide waiver must be a mapping")["contract_key"]
        for record in protected_guide_waivers
    }
    assert protected_decision_paths == set(
        retained_protected_paths_for_modules(GITHUB_POWERSHELL_PROFILE_MODULES)
    )
    assert protected_guide_waiver_keys >= {
        "agents-azure-devops-pr-review-protocol",
        "agents-azure-devops-support-guide-path",
        "copilot-instructions-azure-devops-support-guide-path",
    }
    assert not (target_root / "docs" / "azure-devops-support.md").exists()
    subprocess.run(
        ["git", "init", "-q"],
        cwd=target_root,
        check=True,
        capture_output=True,
        text=True,
    )
    validation_result = run_downstream_adoption_validator(target_root)

    assert validation_result.returncode == 0, validation_result.stdout + validation_result.stderr
    assert "Protected guide contract waiver:" in validation_result.stdout

    # The fixture owner now authorizes removal of this excluded host protocol.
    # Other protected guide waivers still govern the other retained entry points.
    agents_path = target_root / "AGENTS.md"
    agents = read_file(agents_path)
    start = agents.index("## Azure DevOps PR Review Protocol\n")
    end = agents.index("## PR Review Workflow (Codex-adapted)\n", start)
    write_file(agents_path, agents[:start] + agents[end:])
    remaining_waivers = [
        record
        for record in protected_guide_waivers
        if as_mapping(record, "waiver must be a mapping")["path"] != "AGENTS.md"
    ]
    template_sync["protected_guide_contract_waivers"] = remaining_waivers
    write_yaml(target_root / ".template-sync/marker.yml", marker)
    direct_command = [
        sys.executable,
        str(target_root / ".template-sync/scripts/validate_instruction_contracts.py"),
        "--mode",
        "downstream",
        "--require-marker",
        "--repo-root",
        str(target_root),
    ]
    direct = subprocess.run(direct_command, check=False, capture_output=True, text=True)
    aggregate = run_downstream_adoption_validator(target_root)
    assert direct.returncode == 0, direct.stdout + direct.stderr
    assert aggregate.returncode == 0, aggregate.stdout + aggregate.stderr


@pytest.mark.slow
@pytest.mark.parametrize(
    "relative_path",
    [
        pytest.param(path, marks=pytest.mark.upstream_template_only)
        if path == "README.md"
        else path
        for path in ISSUE_693_BASELINE_DOCS
    ],
)
def test_materialized_partial_adoption_strips_shared_baseline_doc_stale_references(
    tmp_path: Path,
    relative_path: str,
) -> None:
    """Partial materialization strips excluded-stack prose from each shared baseline doc."""
    target_root = tmp_path / "partial-docs"
    target_root.mkdir()
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            list(ISSUE_693_PARTIAL_DOC_MODULES),
            protected_file_decisions=protected_take_decisions_for_modules(
                ISSUE_693_PARTIAL_DOC_MODULES
            ),
        ),
    )
    module_args = [
        argument
        for module_name in ISSUE_693_PARTIAL_DOC_MODULES
        for argument in ("--included-module", module_name)
    ]

    result = run_materialize(
        REPO_ROOT,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--last-reviewed-template-commit",
        FULL_SHA,
        "--repository",
        "octocat/hello-world",
        "--security-contact",
        "security@example.com",
        "--decisions-file",
        "decisions.yml",
        *module_args,
    )

    assert result.returncode == 0, result.stderr
    generated_path = target_root / relative_path
    assert generated_path.is_file(), result.stdout
    generated_text = read_file(generated_path)

    for retained_token in ISSUE_693_RETAINED_DOC_REFERENCES[relative_path]:
        assert retained_token in generated_text, f"{relative_path}: {retained_token}"
    for excluded_token in ISSUE_693_EXCLUDED_DOC_REFERENCES[relative_path]:
        assert excluded_token not in generated_text, f"{relative_path}: {excluded_token}"

    subprocess.run(
        ["git", "init", "-q"],
        cwd=target_root,
        check=True,
        capture_output=True,
        text=True,
    )
    report_result = run_excluded_module_report(target_root, ISSUE_693_PARTIAL_DOC_MODULES)

    assert report_result.returncode == 0, report_result.stderr
    matched_report_row = False
    for report_line in report_result.stdout.splitlines():
        if f"| {relative_path}" not in report_line:
            continue
        matched_report_row = True
        assert "required_cleanup" not in report_line, report_line
        assert "markdown-link.excluded-target" not in report_line, report_line
    assert matched_report_row, (
        f"excluded-module report produced no row mentioning {relative_path}; "
        "the report format may have changed, leaving the cleanup assertions "
        "above vacuous"
    )


@pytest.mark.upstream_template_only
@pytest.mark.slow
@pytest.mark.parametrize("template_sync_support_included", [False, True])
def test_materialized_readme_template_sync_support_reference_block(
    tmp_path: Path,
    template_sync_support_included: bool,
) -> None:
    """README template-sync surface rows materialize only when support is adopted."""
    target_root = tmp_path / "readme-template-sync"
    target_root.mkdir()
    included_modules: tuple[str, ...] = NO_DATA_NO_TEMPLATE_SYNC_MODULES
    if template_sync_support_included:
        included_modules = (*included_modules, "template-sync-support")
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            list(included_modules),
            protected_file_decisions=protected_take_decisions_for_modules(included_modules),
        ),
    )
    module_args = [
        argument
        for module_name in included_modules
        for argument in ("--included-module", module_name)
    ]

    result = run_materialize(
        REPO_ROOT,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--last-reviewed-template-commit",
        FULL_SHA,
        "--repository",
        "octocat/hello-world",
        "--security-contact",
        "security@example.com",
        "--decisions-file",
        "decisions.yml",
        *module_args,
    )

    assert result.returncode == 0, result.stderr
    generated_path = target_root / "README.md"
    assert generated_path.is_file(), result.stdout
    generated_text = read_file(generated_path)

    for reference in TEMPLATE_SYNC_SUPPORT_README_REFERENCES:
        if template_sync_support_included:
            assert reference in generated_text, reference
        else:
            assert reference not in generated_text, reference


@pytest.mark.slow
@pytest.mark.parametrize("template_sync_support_included", [False, True])
def test_materialized_contributing_template_sync_support_reference_block(
    tmp_path: Path,
    template_sync_support_included: bool,
) -> None:
    """CONTRIBUTING template-sync surface materializes only when support is adopted."""
    target_root = tmp_path / "contributing-template-sync"
    target_root.mkdir()
    included_modules: tuple[str, ...] = NO_DATA_NO_TEMPLATE_SYNC_MODULES
    if template_sync_support_included:
        included_modules = (*included_modules, "template-sync-support")
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            list(included_modules),
            protected_file_decisions=protected_take_decisions_for_modules(included_modules),
        ),
    )
    module_args = [
        argument
        for module_name in included_modules
        for argument in ("--included-module", module_name)
    ]

    result = run_materialize(
        REPO_ROOT,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--last-reviewed-template-commit",
        FULL_SHA,
        "--repository",
        "octocat/hello-world",
        "--security-contact",
        "security@example.com",
        "--decisions-file",
        "decisions.yml",
        *module_args,
    )

    assert result.returncode == 0, result.stderr
    generated_path = target_root / "CONTRIBUTING.md"
    assert generated_path.is_file(), result.stdout
    generated_text = read_file(generated_path)

    if template_sync_support_included:
        assert "validate_downstream_adoption.py" in generated_text
    else:
        assert "validate_downstream_adoption.py" not in generated_text


@pytest.mark.parametrize(
    ("included_modules", "expect_data_ci_row"),
    [
        pytest.param(
            NO_DATA_NO_TEMPLATE_SYNC_MODULES,
            True,
            id="baseline-and-github-actions",
        ),
        pytest.param(
            tuple(module for module in NO_DATA_NO_TEMPLATE_SYNC_MODULES if module != "baseline")
            + ("json",),
            False,
            id="data-and-github-actions-without-baseline",
        ),
        pytest.param(
            tuple(
                module for module in NO_DATA_NO_TEMPLATE_SYNC_MODULES if module != "github-actions"
            )
            + ("json",),
            False,
            id="baseline-and-data-without-github-actions",
        ),
    ],
)
def test_materialized_contributing_data_ci_reference_block(
    included_modules: tuple[str, ...],
    expect_data_ci_row: bool,
) -> None:
    """The GitHub Data CI row requires both baseline and GitHub Actions."""
    generated_text = materializer.remove_inline_blocks_for_modules(
        read_file(REPO_ROOT / "CONTRIBUTING.md"),
        included_modules,
        relative_path="CONTRIBUTING.md",
    )

    if expect_data_ci_row:
        assert "**Data CI**" in generated_text
        assert ".github/workflows/data-ci.yml" in generated_text
    else:
        assert "**Data CI**" not in generated_text
        assert ".github/workflows/data-ci.yml" not in generated_text


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("module_name", sorted(AZURE_TEMPLATE_MODULES))
def test_materialized_azure_support_guide_retained_by_any_azure_module(
    tmp_path: Path,
    module_name: str,
) -> None:
    """Any Azure host module retains the durable Azure DevOps guide."""
    target_root = materialize_module_fixture(tmp_path, ("baseline", module_name))

    guide_path = target_root / "docs" / "azure-devops-support.md"
    assert guide_path.is_file()
    assert "Azure DevOps Services Support Guide" in read_file(guide_path)


@pytest.mark.upstream_template_only
def test_materialized_all_azure_modules_retain_guide_and_reference_links(
    tmp_path: Path,
) -> None:
    """A full Azure host selection keeps the guide and guarded shared-doc links."""
    target_root = materialize_module_fixture(
        tmp_path, ("baseline", *sorted(AZURE_TEMPLATE_MODULES))
    )

    assert (target_root / "docs" / "azure-devops-support.md").is_file()
    readme_text = read_file(target_root / "README.md")
    contributing_text = read_file(target_root / "CONTRIBUTING.md")

    assert "docs/azure-devops-support.md" in readme_text
    assert "docs/azure-devops-support.md" in contributing_text


@pytest.mark.slow
def test_excluded_module_report_accepts_github_data_ci_and_marker(
    tmp_path: Path,
) -> None:
    """The data-CI reference marker is valid with baseline and GitHub Actions."""
    target_root = tmp_path / "github-data-ci-report"
    target_root.mkdir()
    included_modules = (*NO_DATA_NO_TEMPLATE_SYNC_MODULES, "template-sync-support")
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            list(included_modules),
            protected_file_decisions=protected_take_decisions_for_modules(included_modules),
        ),
    )
    module_args = [
        argument
        for module_name in included_modules
        for argument in ("--included-module", module_name)
    ]

    result = run_materialize(
        REPO_ROOT,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--last-reviewed-template-commit",
        FULL_SHA,
        "--repository",
        "octocat/hello-world",
        "--security-contact",
        "security@example.com",
        "--decisions-file",
        "decisions.yml",
        *module_args,
    )
    assert result.returncode == 0, result.stderr

    subprocess.run(
        ["git", "init", "-q"],
        cwd=target_root,
        check=True,
        capture_output=True,
        text=True,
    )
    report_result = run_excluded_module_report(target_root, included_modules)
    assert report_result.returncode == 0, report_result.stderr
    assert (target_root / ".github" / "workflows" / "data-ci.yml").is_file()
    assert "github-data-ci-reference-only" not in report_result.stdout, report_result.stdout


def test_materialization_preview_reports_pruned_live_marker_families(
    tmp_path: Path,
) -> None:
    """Pruning preview reports live marker families and ignores fenced examples."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {"pattern": "README.md", "requires_all": ["baseline"]},
        ],
    )
    write_file(
        template_root / "README.md",
        (
            "# Fixture\n\n"
            "```markdown\n"
            "<!-- template-sync: begin python-reference-only -->\n"
            "Example marker only.\n"
            "<!-- template-sync: end python-reference-only -->\n"
            "```\n\n"
            "<!-- template-sync: begin python-reference-only -->\n"
            "Live Python reference.\n"
            "<!-- template-sync: end python-reference-only -->\n"
        ),
    )

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-modules",
        "baseline,markdown",
    )

    assert result.returncode == 0, result.stderr
    generated_text = read_file(target_root / "README.md")
    assert "Example marker only." in generated_text
    assert "Live Python reference." not in generated_text

    pruning_preview = section_entries(result.stdout, "Pruning preview")
    assert pruning_preview == {"README.md: python-reference-only"}
    assert "Computed marker preview:" in result.stdout


def test_computed_marker_preview_uses_canonical_marker_yaml(tmp_path: Path) -> None:
    """Preview-only computed marker output is the canonical marker YAML with prefixes."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {"pattern": "README.md", "requires_all": ["baseline"]},
        ],
    )
    write_file(template_root / "README.md", "template readme\n")

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 0, result.stderr
    expected_marker = materializer.marker_yaml(
        {"template_sync": {"source_repo": SOURCE_REPO, "included_modules": ["baseline"]}}
    )
    expected_preview = "".join(f"  {line}\n" for line in expected_marker.rstrip("\n").splitlines())
    assert result.stdout.split("Computed marker preview:\n", maxsplit=1)[1] == expected_preview
    assert not (target_root / ".template-sync" / "marker.yml").exists()


def materialized_marker_fixture(tmp_path: Path) -> Path:
    """Materialize a minimal fixture that writes a computed marker."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {"pattern": ".template-sync/marker.yml", "requires_all": ["template-sync-support"]},
        ],
    )
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            ["template-sync-support"],
            local_overrides=[
                {
                    "path": "README.md",
                    "reason": "Review local README changes later.",
                    "default_decision": "DEFER",
                }
            ],
        ),
    )

    result = run_materialize(template_root, target_root, "--decisions-file", "decisions.yml")

    assert result.returncode == 0, result.stderr
    return target_root / ".template-sync" / "marker.yml"


def test_materialized_marker_yaml_is_canonical(tmp_path: Path) -> None:
    """Generated marker writes use canonical YAML before reaching disk."""
    marker_path = materialized_marker_fixture(tmp_path)
    marker_text = read_file(marker_path)
    marker_document_data = as_mapping(load_yaml(marker_path), "marker must be a mapping")
    assert marker_text == materializer.marker_yaml(marker_document_data)
    assert "  included_modules:\n    - template-sync-support\n" in marker_text
    assert "  local_overrides:\n    - path: README.md\n" in marker_text
    assert not marker_text.startswith("---")


def test_materialized_marker_yaml_is_yamllint_clean(tmp_path: Path) -> None:
    """Generated marker writes pass yamllint under the repository config."""
    if yamllint_command() is None:
        pytest.skip("yamllint executable or importable module is required for this assertion")
    marker_path = materialized_marker_fixture(tmp_path)

    assert_yamllint_clean(marker_path)


@pytest.mark.slow
def test_materialized_baseline_without_python_retains_runner_dependabot_ecosystems(
    tmp_path: Path,
) -> None:
    """A baseline-only runner keeps pip and hook updates without Python metadata."""
    target_root = tmp_path / "no-python"
    target_root.mkdir()
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            list(ISSUE_692_NO_PYTHON_MODULES),
            protected_file_decisions=protected_take_decisions_for_modules(
                ISSUE_692_NO_PYTHON_MODULES
            ),
        ),
    )
    module_args = [
        argument
        for module_name in ISSUE_692_NO_PYTHON_MODULES
        for argument in ("--included-module", module_name)
    ]

    result = run_materialize(
        REPO_ROOT,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--last-reviewed-template-commit",
        FULL_SHA,
        "--decisions-file",
        "decisions.yml",
        *module_args,
    )

    assert result.returncode == 0, result.stderr
    dependabot_path = target_root / ".github" / "dependabot.yml"
    assert dependabot_path.is_file(), result.stdout

    dependabot_text = read_file(dependabot_path)
    assert "Baseline pre-commit runner and retained Python dependencies" in dependabot_text
    assert 'package-ecosystem: "pip"' in dependabot_text
    assert "pip-minor-patch" in dependabot_text
    assert "npm (package.json) - Markdown tooling dependencies" in dependabot_text
    assert "GitHub Actions (workflows) - Action version updates" in dependabot_text
    assert "pre-commit (.pre-commit-config.yaml) - Pre-commit hook updates" in dependabot_text
    assert dependabot_update_ecosystems(dependabot_path) == DEPENDABOT_BASELINE_NO_PYTHON_ECOSYSTEMS
    assert (target_root / "requirements-pre-commit.txt").is_file()
    assert (target_root / ".pre-commit-config.yaml").is_file()
    assert not (target_root / "pyproject.toml").exists()
    validate_dependabot_vendor_schema(dependabot_path)

    subprocess.run(
        ["git", "init"],
        cwd=target_root,
        check=True,
        capture_output=True,
        text=True,
    )
    report_result = run_excluded_module_report(target_root, ISSUE_692_NO_PYTHON_MODULES)

    assert report_result.returncode == 0, report_result.stderr
    assert "dependabot-ecosystem.stale | required_cleanup | python |" not in report_result.stdout


@pytest.mark.slow
@pytest.mark.parametrize(
    "included_modules",
    [
        pytest.param(
            FULL_TEMPLATE_MODULES,
            id="full-upstream-module-set",
            marks=pytest.mark.upstream_template_only,
        ),
        pytest.param(DOWNSTREAM_PYTEST_MODULES, id="complete-downstream-pytest-module-set"),
    ],
)
def test_materialized_full_adoption_keeps_all_dependabot_ecosystems(
    tmp_path: Path,
    included_modules: tuple[str, ...],
) -> None:
    """Complete upstream and portable downstream profiles keep every ecosystem."""
    target_root = tmp_path / "full"
    target_root.mkdir()
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            list(included_modules),
            protected_file_decisions=protected_take_decisions_for_modules(included_modules),
        ),
    )
    module_args = [
        argument
        for module_name in included_modules
        for argument in ("--included-module", module_name)
    ]

    result = run_materialize(
        REPO_ROOT,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--last-reviewed-template-commit",
        FULL_SHA,
        "--decisions-file",
        "decisions.yml",
        *module_args,
    )

    assert result.returncode == 0, result.stderr
    marker = as_mapping(load_yaml(target_root / ".template-sync/marker.yml"), "adopted marker")
    template_sync = as_mapping(marker["template_sync"], "adopted template_sync")
    recorded_modules = as_string_list(template_sync["included_modules"], "adopted modules")
    assert set(recorded_modules) == set(included_modules)
    dependabot_path = target_root / ".github" / "dependabot.yml"
    assert dependabot_path.is_file(), result.stdout
    assert dependabot_update_ecosystems(dependabot_path) == DEPENDABOT_FULL_ECOSYSTEMS
    validate_dependabot_vendor_schema(dependabot_path)


def test_unrecorded_conflict_exits_two_and_allow_conflicts_does_not_advance_marker(
    tmp_path: Path,
) -> None:
    """Unrecorded conflicts are nonzero by default and preview-only with support."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {"pattern": ".template-sync/marker.yml", "requires_all": ["template-sync-support"]},
            {"pattern": "README.md", "requires_all": ["baseline"]},
            {"pattern": "CONTRIBUTING.md", "requires_all": ["baseline"]},
        ],
    )
    write_file(template_root / "README.md", "template readme\n")
    write_file(template_root / "CONTRIBUTING.md", "template contributing\n")
    write_file(target_root / "README.md", "downstream readme\n")

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-modules",
        "baseline,template-sync-support",
    )

    assert result.returncode == 2, result.stdout
    assert read_file(target_root / "README.md") == "downstream readme\n"
    assert read_file(target_root / "CONTRIBUTING.md") == "template contributing\n"
    assert "README.md: unrecorded" in result.stdout
    assert "preview-only: unrecorded conflicts remain" in result.stdout
    assert not (target_root / ".template-sync" / "marker.yml").exists()

    allow_result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-modules",
        "baseline,template-sync-support",
        "--allow-conflicts",
    )

    assert allow_result.returncode == 0, allow_result.stderr
    assert "preview-only: unrecorded conflicts remain" in allow_result.stdout
    assert not (target_root / ".template-sync" / "marker.yml").exists()


def test_decisions_file_take_skip_and_cli_overlap_conflict(tmp_path: Path) -> None:
    """TAKE and SKIP are applied, while conflicting CLI scalar inputs are rejected."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {"pattern": "README.md", "requires_all": ["baseline"]},
            {"pattern": "CONTRIBUTING.md", "requires_all": ["baseline"]},
        ],
    )
    write_file(template_root / "README.md", "template readme\n")
    write_file(template_root / "CONTRIBUTING.md", "template contributing\n")
    write_file(target_root / "README.md", "downstream readme\n")
    write_file(target_root / "CONTRIBUTING.md", "downstream contributing\n")
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            ["baseline"],
            local_overrides=[
                {
                    "path": "README.md",
                    "default_decision": "TAKE",
                    "reason": "Adopt template README.",
                },
                {
                    "path": "CONTRIBUTING.md",
                    "default_decision": "SKIP",
                    "reason": "Keep downstream contributing guide.",
                },
            ],
        ),
    )

    result = run_materialize(
        template_root,
        target_root,
        "--decisions-file",
        "decisions.yml",
    )

    assert result.returncode == 0, result.stderr
    assert read_file(target_root / "README.md") == "template readme\n"
    assert read_file(target_root / "CONTRIBUTING.md") == "downstream contributing\n"
    assert "README.md" in result.stdout
    assert "CONTRIBUTING.md" in result.stdout

    conflict_result = run_materialize(
        template_root,
        target_root,
        "--decisions-file",
        "decisions.yml",
        "--source-repo",
        "https://example.com/other.git",
    )

    assert conflict_result.returncode == 1
    assert "Conflicting source repo" in conflict_result.stderr


def test_recorded_deferral_exits_zero_and_marker_updates_then_stays_unchanged(
    tmp_path: Path,
) -> None:
    """Recorded unresolved decisions do not block marker advancement."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {"pattern": ".template-sync/marker.yml", "requires_all": ["template-sync-support"]},
            {"pattern": "README.md", "requires_all": ["baseline"]},
        ],
    )
    write_file(template_root / "README.md", "template readme\n")
    write_file(target_root / "README.md", "downstream readme\n")
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            ["baseline", "template-sync-support"],
            local_overrides=[
                {
                    "path": "README.md",
                    "default_decision": "DEFER",
                    "reason": "Review README later.",
                }
            ],
        ),
    )

    result = run_materialize(
        template_root,
        target_root,
        "--decisions-file",
        "decisions.yml",
    )

    assert result.returncode == 0, result.stderr
    assert "Recorded unresolved decisions remain:" in result.stdout
    assert "updated: computed marker written" in result.stdout
    assert read_file(target_root / "README.md") == "downstream readme\n"
    marker_text = read_file(target_root / ".template-sync" / "marker.yml")
    assert "default_decision: DEFER" in marker_text

    unchanged_result = run_materialize(
        template_root,
        target_root,
        "--decisions-file",
        "decisions.yml",
    )

    assert unchanged_result.returncode == 0, unchanged_result.stderr
    assert "unchanged: existing marker already equals computed marker" in unchanged_result.stdout


def test_remove_local_is_recorded_but_not_applied(tmp_path: Path) -> None:
    """REMOVE-LOCAL records an unresolved removal without deleting the target file."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {"pattern": ".template-sync/marker.yml", "requires_all": ["template-sync-support"]},
            {"pattern": "README.md", "requires_all": ["baseline"]},
        ],
    )
    write_file(template_root / "README.md", "template readme\n")
    write_file(target_root / "README.md", "downstream readme\n")
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            ["baseline", "template-sync-support"],
            local_overrides=[
                {
                    "path": "README.md",
                    "default_decision": "REMOVE-LOCAL",
                    "reason": "Remove after owner review.",
                }
            ],
        ),
    )

    result = run_materialize(template_root, target_root, "--decisions-file", "decisions.yml")

    assert result.returncode == 0, result.stderr
    assert read_file(target_root / "README.md") == "downstream readme\n"
    assert "Recorded but not applied removals:" in result.stdout
    assert "README.md" in result.stdout


def test_materializer_json_args_file_supplies_package_metadata(
    tmp_path: Path,
) -> None:
    """The materializer accepts shell-safe JSON args and updates package identity."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {"pattern": "package.json", "requires_all": ["baseline"]},
            {"pattern": "package-lock.json", "requires_all": ["baseline"]},
        ],
        include_placeholder_helper=True,
    )
    write_json(
        template_root / "package.json",
        {
            "name": "copilot-repo-template",
            "version": "1.0.0",
            "description": "Template repository with Copilot instructions",
            "private": True,
            "author": "Frank Lesniak",
        },
    )
    write_json(
        template_root / "package-lock.json",
        {
            "name": "copilot-repo-template",
            "version": "1.0.0",
            "lockfileVersion": 3,
            "packages": {
                "": {
                    "name": "copilot-repo-template",
                    "version": "1.0.0",
                },
                "node_modules/example": {"version": "9.9.9"},
            },
        },
    )
    args_file = tmp_path / "materialize.args.json"
    write_json(
        args_file,
        {
            "target_root": str(target_root),
            "source_repo": SOURCE_REPO,
            "included_modules": ["baseline"],
            "package_name": "downstream-markdown-tools",
            "package_description": "Markdown tooling for downstream docs, with $literal text",
            "package_author": "Example Org",
            "package_version": "2.0.0",
        },
    )

    result = subprocess.run(
        [
            sys.executable,
            str(SCRIPT_PATH),
            "--template-root",
            str(template_root),
            "--args-file",
            str(args_file),
        ],
        check=False,
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stdout + result.stderr
    package_json = json.loads(read_file(target_root / "package.json"))
    package_lock = json.loads(read_file(target_root / "package-lock.json"))
    assert package_json["name"] == "downstream-markdown-tools"
    assert package_json["description"] == "Markdown tooling for downstream docs, with $literal text"
    assert package_json["author"] == "Example Org"
    assert package_json["version"] == "2.0.0"
    assert package_lock["name"] == "downstream-markdown-tools"
    assert package_lock["version"] == "2.0.0"
    assert package_lock["packages"][""]["name"] == "downstream-markdown-tools"
    assert package_lock["packages"][""]["version"] == "2.0.0"
    assert package_lock["packages"]["node_modules/example"]["version"] == "9.9.9"


def test_materializer_args_file_cli_values_take_precedence(tmp_path: Path) -> None:
    """Direct CLI values override lower-priority args-file values."""
    args_file = tmp_path / "materialize.args.json"
    write_json(
        args_file,
        {
            "target_root": str(tmp_path / "from-file"),
            "source_repo": "https://example.com/from-file.git",
            "included_modules": ["python"],
            "package_name": "from-file",
        },
    )

    args = materializer.parse_args(
        [
            "--args-file",
            str(args_file),
            "--target-root",
            str(tmp_path / "from-cli"),
            "--source-repo",
            SOURCE_REPO,
            "--included-module",
            "baseline",
            "--package-name",
            "from-cli",
        ]
    )

    assert args.target_root == str(tmp_path / "from-cli")
    assert args.source_repo == SOURCE_REPO
    assert args.included_modules == ["baseline"]
    assert args.package_name == "from-cli"


def test_materializer_args_file_with_utf8_bom_is_parsed(tmp_path: Path) -> None:
    """A UTF-8 BOM (e.g. PowerShell Set-Content -Encoding UTF8) in an args file is tolerated."""
    args_file = tmp_path / "materialize.args.json"
    args_file.write_bytes(
        b"\xef\xbb\xbf"
        + json.dumps({"target_root": str(tmp_path / "target"), "source_repo": SOURCE_REPO}).encode(
            "utf-8"
        )
    )

    mapping = materializer.load_args_file_mapping(str(args_file), None)

    assert mapping["target_root"] == str(tmp_path / "target")
    assert mapping["source_repo"] == SOURCE_REPO


def test_materializer_args_file_cli_selectors_override_family(tmp_path: Path) -> None:
    """A CLI source/module selector overrides its whole args-file family, not just the same flag."""
    args_file = tmp_path / "materialize.args.json"
    write_json(
        args_file,
        {"template_ref": "v1.2.3", "included_modules_csv": "python,yaml"},
    )

    args = materializer.parse_args(
        [
            "--args-file",
            str(args_file),
            "--template-root",
            str(tmp_path / "tmpl"),
            "--target-root",
            str(tmp_path / "tgt"),
            "--source-repo",
            SOURCE_REPO,
            "--included-module",
            "baseline",
        ]
    )

    # CLI --template-root wins; the args-file source selector is skipped (no conflict error).
    assert args.template_root == str(tmp_path / "tmpl")
    assert args.template_ref is None
    # CLI --included-module wins; the args-file module selector is skipped (no merge).
    assert args.included_modules == ["baseline"]
    assert args.included_modules_csv is None


def test_materializer_args_file_decisions_path_traversal_is_rejected(
    tmp_path: Path,
) -> None:
    """Repo-relative path values supplied through args files cannot escape target root."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "README.md", "requires_all": ["baseline"]}])
    write_file(template_root / "README.md", "template readme\n")
    args_file = tmp_path / "materialize.args.json"
    write_json(
        args_file,
        {
            "target_root": str(target_root),
            "source_repo": SOURCE_REPO,
            "included_modules": ["baseline"],
            "decisions_file": "../outside.yml",
        },
    )

    result = subprocess.run(
        [
            sys.executable,
            str(SCRIPT_PATH),
            "--template-root",
            str(template_root),
            "--args-file",
            str(args_file),
        ],
        check=False,
        capture_output=True,
        text=True,
    )

    assert result.returncode == 1
    assert "--decisions-file must not contain traversal segments" in result.stderr


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("markdown", [True, False])
def test_claude_bootstrap_tracks_markdown_without_baseline(tmp_path: Path, markdown: bool) -> None:
    """Materialized agent hooks retain Markdown setup independently of baseline setup."""
    modules: tuple[str, ...] = ("agent-instructions", "agent-claude")
    if markdown:
        modules += ("markdown",)
    target = materialize_module_fixture(tmp_path, modules, authorize_protected_files=True)
    hook = read_file(target / ".claude/hooks/session-start.sh")
    assert ("npm ci --ignore-scripts" in hook) is markdown
    assert ("markdown_repository_root" in hook) is markdown
    assert "ensure_pre_commit" not in hook


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    ("modules", "retained"),
    [
        (("agent-instructions", "agent-claude", "github-actions"), True),
        (("agent-instructions", "github-actions"), False),
        (("agent-instructions", "agent-claude"), False),
        (("github-actions",), False),
        (("agent-instructions", "agent-claude", "azure-pipelines"), False),
        (("agent-instructions", "agent-claude", "github-actions", "azure-pipelines"), True),
    ],
)
def test_claude_review_command_selection(
    tmp_path: Path, modules: tuple[str, ...], retained: bool
) -> None:
    """The command and its discoverability text follow the selected agent and host."""
    target = materialize_module_fixture(tmp_path, modules, authorize_protected_files=True)
    command = target / ".claude/commands/review-loop.md"
    assert command.is_file() is retained
    prompts = target / "docs/PR_REVIEW_PROMPTS.md"
    if prompts.is_file():
        text = read_file(prompts)
        assert ("## Local Claude Review Command" in text) is retained
        assert ("## Requesting Copilot Review and Recording Effort" in text) is (
            "github-actions" in modules
        )
    if retained:
        assert (target / "CLAUDE.md").is_file()
        assert (target / ".github/copilot-instructions.md").is_file()
        assert not (target / ".template-sync/scripts").exists()
    before = snapshot_fixture_files(target)
    repeated = run_materialize(
        REPO_ROOT,
        target,
        "--decisions-file",
        "decisions.yml",
        *azure_provider_cli_args_for_modules(modules),
    )
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert snapshot_fixture_files(target) == before


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("removed_module", ["agent-claude", "github-actions"])
def test_claude_review_command_removal_requires_reviewed_cleanup(
    tmp_path: Path, removed_module: str
) -> None:
    """Later selection changes prune references and leave excluded files for reviewed cleanup."""
    retained = ("agent-instructions", "agent-claude", "github-actions", "template-sync-support")
    target = materialize_module_fixture(tmp_path, retained, authorize_protected_files=True)
    command = target / ".claude/commands/review-loop.md"
    original = command.read_bytes()
    omitted = tuple(module for module in retained if module != removed_module)
    protected_decisions = protected_take_decisions_for_modules(omitted)
    if removed_module == "agent-claude":
        protected_decisions.append(
            {
                "path": "CLAUDE.md",
                "decision": "REMOVE-LOCAL",
                "reason": "The fixture owner retires the existing Claude entry point.",
                "authorization_basis": "Fixture owner explicitly authorizes Claude removal.",
                "authorized_scope": "CLAUDE.md only.",
            }
        )
    write_yaml(
        target / "decisions.yml",
        marker_document(
            list(omitted),
            protected_file_decisions=protected_decisions,
            local_overrides=[
                {
                    "path": relative_path,
                    "default_decision": "TAKE",
                    "reason": "Fixture owner authorizes the reviewed command-reference pruning.",
                }
                for relative_path in (
                    "docs/PR_REVIEW_PROMPTS.md",
                    "TEMPLATE_UPDATE_PROCEDURE.md",
                    "tests/test_contract_wiring.py",
                )
            ],
        ),
    )
    result = run_materialize(REPO_ROOT, target, "--decisions-file", "decisions.yml")
    assert result.returncode == 0, result.stdout + result.stderr
    assert command.read_bytes() == original
    assert "## Local Claude Review Command" not in read_file(target / "docs/PR_REVIEW_PROMPTS.md")
    assert ".claude/commands/review-loop.md" in result.stdout
    # Normal module removal requires deliberate cleanup of the reported excluded file.
    command.unlink()
    repeated = run_materialize(REPO_ROOT, target, "--decisions-file", "decisions.yml")
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert not command.exists()


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("markdown", [True, False])
def test_copilot_recipe_contract_without_claude(tmp_path: Path, markdown: bool) -> None:
    """The materialized recipe skips its absent optional successor without Markdown tooling."""
    modules: tuple[str, ...] = (
        "agent-instructions",
        "agent-codex",
        "github-actions",
        "template-sync-support",
    )
    if markdown:
        modules += ("markdown",)
    target = materialize_module_fixture(tmp_path, modules, authorize_protected_files=True)
    record_github_agent_azure_protocol_section_waivers(target)
    text = read_file(target / "docs/PR_REVIEW_PROMPTS.md")
    assert "## Requesting Copilot Review and Recording Effort" in text
    assert "## Local Claude Review Command" not in text
    assert not (target / ".claude/commands/review-loop.md").exists()
    command = [
        sys.executable,
        str(target / ".template-sync/scripts/validate_instruction_contracts.py"),
        "--mode",
        "downstream",
        "--require-marker",
        "--repo-root",
        str(target),
    ]
    result = subprocess.run(command, check=False, capture_output=True, text=True)
    assert result.returncode == 0, result.stdout + result.stderr
    assert "validation passed" in result.stdout


def test_placeholder_replacement_ignores_selected_helper(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Placeholder replacement uses trusted code even if selected code is hostile or absent."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [{"pattern": "SECURITY.md", "requires_all": ["baseline"]}],
        include_placeholder_helper=True,
    )
    write_file(
        template_root / "SECURITY.md",
        "Report at https://github.com/OWNER/REPO/security\nEmail [security contact email]\n",
    )
    write_file(
        template_root / ".github/scripts/replace-template-placeholders.py",
        "from pathlib import Path\n"
        "Path(__file__).parents[2].joinpath('candidate-executed.txt').write_text('executed')\n"
        "raise RuntimeError('candidate placeholder helper executed')\n",
    )

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--repository",
        "octo/widget",
        "--security-contact",
        "security@example.com",
    )

    assert result.returncode == 0, result.stderr
    assert not (template_root / "candidate-executed.txt").exists()
    security_text = read_file(target_root / "SECURITY.md")
    assert "https://github.com/octo/widget/security" in security_text
    assert "security@example.com" in security_text
    assert "SECURITY.md" in result.stdout

    # Restoring the selected source as the execution anchor must trip the sentinel.
    mutant_target = tmp_path / "mutant-target"
    mutant_target.mkdir()
    with monkeypatch.context() as patch:
        patch.setattr(materializer, "TRUSTED_TOOL_ROOT", template_root)
        failed = materializer.main(
            [
                "--template-root",
                str(template_root),
                "--target-root",
                str(mutant_target),
                "--source-repo",
                SOURCE_REPO,
                "--included-module",
                "baseline",
                "--repository",
                "octo/widget",
                "--security-contact",
                "security@example.com",
            ]
        )
    assert failed == 1
    assert read_file(template_root / "candidate-executed.txt") == "executed"

    missing_helper_template = tmp_path / "missing-helper-template"
    missing_helper_target = tmp_path / "missing-helper-target"
    missing_helper_target.mkdir()
    prepare_template(
        missing_helper_template,
        [{"pattern": "SECURITY.md", "requires_all": ["baseline"]}],
        include_placeholder_helper=True,
    )
    (missing_helper_template / ".github/scripts/replace-template-placeholders.py").unlink()
    write_file(missing_helper_template / "SECURITY.md", "Email [security contact email]\n")

    missing_result = run_materialize(
        missing_helper_template,
        missing_helper_target,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--repository",
        "octo/widget",
        "--security-contact",
        "security@example.com",
    )

    assert missing_result.returncode == 0, missing_result.stdout + missing_result.stderr
    assert "security@example.com" in read_file(missing_helper_target / "SECURITY.md")


@pytest.mark.upstream_template_only
def test_materializer_replays_azure_provider_fields_from_marker(
    tmp_path: Path,
) -> None:
    """Marker-recorded Azure provider fields drive placeholder materialization."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {"pattern": "SECURITY.md", "requires_all": ["baseline"]},
            {"pattern": "CONTRIBUTING.md", "requires_all": ["baseline"]},
            {"pattern": "CODE_OF_CONDUCT.md", "requires_all": ["baseline"]},
            {
                "pattern": ".azuredevops/platform/**",
                "requires_all": ["azure-devops-platform"],
            },
            {
                "pattern": ".azuredevops/pull_request_template.md",
                "requires_all": ["azure-devops-collaboration"],
            },
        ],
        include_placeholder_helper=True,
    )
    copy_template_file(template_root, "SECURITY.md")
    copy_template_file(template_root, "CONTRIBUTING.md")
    copy_template_file(template_root, "CODE_OF_CONDUCT.md")
    copy_template_file(template_root, ".azuredevops/platform/adoption-guidance.md")
    copy_template_file(template_root, ".azuredevops/pull_request_template.md")
    write_yaml(
        target_root / "decisions.yml",
        {
            "template_sync": {
                "source_repo": SOURCE_REPO,
                "included_modules": [
                    "baseline",
                    "template-sync-support",
                    "azure-devops-platform",
                    "azure-devops-collaboration",
                ],
                "host_provider": "azure-devops-services",
                "azure_devops_organization": "contoso",
                "azure_devops_project": "Microsoft 365",
                "azure_devops_repository": "downstream-template",
                "azure_boards_policy": "work-items",
                "azure_repos_pr_template_policy": "materialize",
                "azure_branch_policy_reviewer_guidance": "manual-follow-up",
                "azure_security_intake_policy": "manual-follow-up",
                "azure_security_product_enablement": "github-secret-protection",
                "azure_dependency_update_policy": "manual-follow-up",
            }
        },
    )

    result = run_materialize(
        template_root,
        target_root,
        "--decisions-file",
        "decisions.yml",
        "--security-contact",
        "security@example.com",
    )

    assert result.returncode == 0, result.stderr
    guidance_text = read_file(target_root / ".azuredevops" / "platform" / "adoption-guidance.md")
    assert "https://dev.azure.com/contoso/Microsoft%20365" in guidance_text
    assert "https://dev.azure.com/contoso/Microsoft%20365/_git/downstream-template" in guidance_text
    assert "AZURE_DEVOPS_PROJECT_URL" not in guidance_text
    pr_template_text = read_file(target_root / ".azuredevops" / "pull_request_template.md")
    assert "[Microsoft 365](https://dev.azure.com/contoso/Microsoft%20365)" in pr_template_text
    assert (
        "[downstream-template]"
        "(https://dev.azure.com/contoso/Microsoft%20365/_git/downstream-template)"
    ) in pr_template_text
    assert "Default branch used for PR template discovery: `main`." in pr_template_text
    assert "Azure Boards intake policy: `work-items`." in pr_template_text
    assert "not by this Markdown template" in pr_template_text
    assert "AZURE_DEVOPS_PROJECT_URL" not in pr_template_text
    security_text = read_file(target_root / "SECURITY.md")
    assert "Azure DevOps Services project" in security_text
    assert "private vulnerability reporting" not in security_text
    contributing_text = read_file(target_root / "CONTRIBUTING.md")
    assert "OWNER/REPO" not in contributing_text
    assert "https://dev.azure.com/contoso/Microsoft%20365/_git/downstream-template" in (
        contributing_text
    )
    marker = as_mapping(
        load_yaml(target_root / ".template-sync" / "marker.yml"),
        "marker must be a mapping",
    )
    template_sync = as_mapping(marker["template_sync"], "template_sync must be a mapping")
    assert template_sync["host_provider"] == "azure-devops-services"
    assert template_sync["azure_devops_project"] == "Microsoft 365"
    assert template_sync["azure_security_product_enablement"] == "github-secret-protection"


@pytest.mark.parametrize(
    ("mode_args", "expected_mode", "expected_url"),
    [
        pytest.param(
            ["--security-contact", "security@example.com"],
            "both",
            "https://github.com/octo/widget/blob/HEAD/SECURITY.md",
            id="omitted-mode-backward-compatible-both",
        ),
        pytest.param(
            ["--security-reporting-mode", "both", "--security-contact", "security@example.com"],
            "both",
            "https://github.com/octo/widget/blob/HEAD/SECURITY.md",
            id="explicit-both",
        ),
        pytest.param(
            [
                "--security-reporting-mode",
                "contact-only",
                "--security-contact",
                "security@example.com",
            ],
            "contact-only",
            "https://github.com/octo/widget/blob/HEAD/SECURITY.md",
            id="contact-only",
        ),
        pytest.param(
            ["--security-reporting-mode", "github-private-only"],
            "github-private-only",
            "https://github.com/octo/widget/blob/HEAD/SECURITY.md",
            id="github-private-only",
        ),
    ],
)
@pytest.mark.upstream_template_only
def test_materializes_security_reporting_modes(
    tmp_path: Path,
    mode_args: list[str],
    expected_mode: str,
    expected_url: str,
) -> None:
    """Materialization passes supported security reporting modes to the helper."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_security_reporting_template(template_root)

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--repository",
        "octo/widget",
        "--conduct-contact",
        "conduct@example.com",
        *mode_args,
    )

    assert result.returncode == 0, result.stderr
    assert "Placeholder scan passed" in result.stdout
    assert "conduct@example.com" in read_file(target_root / "CODE_OF_CONDUCT.md")
    assert "[INSERT CONTACT METHOD]" not in read_file(target_root / "CODE_OF_CONDUCT.md")

    config = load_yaml(target_root / ".github" / "ISSUE_TEMPLATE" / "config.yml")
    contact_link = security_contact_link(config)
    assert contact_link["url"] == expected_url
    assert_issue_form_shape(
        load_yaml(target_root / ".github" / "ISSUE_TEMPLATE" / "bug_report.yml")
    )

    security_text = read_file(target_root / "SECURITY.md")
    combined_security_surfaces = "\n".join(
        (
            security_text,
            read_file(target_root / ".github" / "ISSUE_TEMPLATE" / "config.yml"),
            read_file(target_root / ".github" / "ISSUE_TEMPLATE" / "bug_report.yml"),
        )
    )
    assert "[security contact email]" not in security_text
    assert "TODO: Replace" not in security_text

    if expected_mode == "contact-only":
        assert "security@example.com" in security_text
        assert "/security/advisories/new" not in combined_security_surfaces
        assert "GitHub Security Advisories" not in combined_security_surfaces
        assert "private vulnerability reporting" not in combined_security_surfaces
    elif expected_mode == "github-private-only":
        assert "security@example.com" not in security_text
        assert "Security Contact" not in security_text
        assert "GitHub Security Advisories" in security_text
        assert "private vulnerability reporting" in combined_security_surfaces
    else:
        assert "security@example.com" in security_text
        assert "GitHub Security Advisories" in security_text
        assert "private vulnerability reporting" in combined_security_surfaces


@pytest.mark.upstream_template_only
def test_materialized_security_reporting_mode_preserves_github_host(
    tmp_path: Path,
) -> None:
    """Materialized security-mode URLs honor --github-host for GHES adoption."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_security_reporting_template(template_root)

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--repository",
        "octo/widget",
        "--security-reporting-mode",
        "github-private-only",
        "--conduct-contact",
        "conduct@example.com",
        "--github-host",
        "github.company.com",
    )

    assert result.returncode == 0, result.stderr
    config = load_yaml(target_root / ".github" / "ISSUE_TEMPLATE" / "config.yml")
    contact_link = security_contact_link(config)
    assert contact_link["url"] == "https://github.company.com/octo/widget/blob/HEAD/SECURITY.md"
    for relative_path in (
        "SECURITY.md",
        ".github/ISSUE_TEMPLATE/config.yml",
        ".github/ISSUE_TEMPLATE/bug_report.yml",
    ):
        text = read_file(target_root / relative_path)
        assert "https://github.com/octo/widget" not in text


@pytest.mark.upstream_template_only
def test_materializes_collaboration_policy_cli_inputs(
    tmp_path: Path,
) -> None:
    """Materialization passes collaboration policy CLI inputs to the helper."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_security_reporting_template(template_root)

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--repository",
        "octo/widget",
        "--security-contact",
        "security@example.com",
        "--conduct-contact",
        "conduct@example.com",
        "--issue-label-policy",
        "custom",
        "--issue-label",
        "type: bug",
        "--issue-label",
        "needs triage",
        "--discussions-policy",
        "enabled",
    )

    assert result.returncode == 0, result.stderr
    bug_report = load_yaml(target_root / ".github" / "ISSUE_TEMPLATE" / "bug_report.yml")
    assert isinstance(bug_report, dict)
    assert bug_report["labels"] == ["type: bug", "needs triage"]
    config = load_yaml(target_root / ".github" / "ISSUE_TEMPLATE" / "config.yml")
    discussions_link = contact_link_by_name(config, "Questions & Discussions")
    assert discussions_link["url"] == "https://github.com/octo/widget/discussions"


@pytest.mark.upstream_template_only
def test_materializes_collaboration_policy_from_decisions_file(
    tmp_path: Path,
) -> None:
    """Marker-recorded collaboration policy fields feed placeholder rendering."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_security_reporting_template(template_root)
    follow_up_status = "Discussions enablement is still open in _TODO-repo-init.md."
    write_yaml(
        target_root / "decisions.yml",
        {
            "template_sync": {
                "source_repo": SOURCE_REPO,
                "included_modules": ["baseline"],
                "issue_label_policy": "create-manual-follow-up",
                "discussions_policy": "deferred-not-rendered",
                "collaboration_policy_follow_up_status": follow_up_status,
            }
        },
    )

    result = run_materialize(
        template_root,
        target_root,
        "--decisions-file",
        "decisions.yml",
        "--repository",
        "octo/widget",
        "--security-contact",
        "security@example.com",
        "--conduct-contact",
        "conduct@example.com",
    )

    assert result.returncode == 0, result.stderr
    bug_report = load_yaml(target_root / ".github" / "ISSUE_TEMPLATE" / "bug_report.yml")
    assert isinstance(bug_report, dict)
    assert bug_report["labels"] == ["bug", "triage"]
    config_text = read_file(target_root / ".github" / "ISSUE_TEMPLATE" / "config.yml")
    assert "/discussions" not in config_text
    assert "_TODO-repo-init.md dependent-file status remains open" in config_text
    assert follow_up_status in config_text
    assert "issue_label_policy: create-manual-follow-up" in result.stdout
    assert "discussions_policy: deferred-not-rendered" in result.stdout


@pytest.mark.upstream_template_only
def test_materializer_rejects_missing_security_mode_and_contact_when_placeholders_run(
    tmp_path: Path,
) -> None:
    """Supplying repository replacement without mode or contact fails clearly."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_security_reporting_template(template_root)

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--repository",
        "octo/widget",
    )

    assert result.returncode == 1
    assert "Either --security-reporting-mode or --security-contact is required" in result.stderr


def test_byte_only_files_are_not_sent_to_placeholder_helper(tmp_path: Path) -> None:
    """Byte-only retained files stay byte-for-byte even when placeholders run."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [{"pattern": "SECURITY.md", "requires_all": ["baseline"]}],
        include_placeholder_helper=True,
    )
    byte_content = b"\xff\x00OWNER/REPO"
    security_path = template_root / "SECURITY.md"
    security_path.parent.mkdir(parents=True, exist_ok=True)
    security_path.write_bytes(byte_content)

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
        "--repository",
        "octo/widget",
        "--security-contact",
        "security@example.com",
    )

    assert result.returncode == 0, result.stderr
    assert (target_root / "SECURITY.md").read_bytes() == byte_content
    assert "Byte-only paths:" in result.stdout
    assert "SECURITY.md" in result.stdout


def test_protected_files_require_concrete_decisions_before_write(tmp_path: Path) -> None:
    """Protected files are conflicts until a path-scoped TAKE decision exists."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [{"pattern": "AGENTS.md", "requires_all": ["agent-instructions"]}],
    )
    write_file(template_root / "AGENTS.md", "agent instructions\n")

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "agent-instructions",
    )

    assert result.returncode == 2, result.stdout
    assert not (target_root / "AGENTS.md").exists()
    assert "unrecorded protected-file decision required" in result.stdout

    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            ["agent-instructions"],
            protected_file_decisions=[
                {
                    "path": "AGENTS.md",
                    "decision": "TAKE",
                    "adoption_mode": "minimal-preservation",
                    "authorization_basis": "Owner authorized AGENTS.md adoption.",
                    "authorized_scope": "AGENTS.md only.",
                }
            ],
        ),
    )

    take_result = run_materialize(template_root, target_root, "--decisions-file", "decisions.yml")

    assert take_result.returncode == 0, take_result.stderr
    assert read_file(target_root / "AGENTS.md") == "agent instructions\n"


@pytest.mark.parametrize(
    "existing_catalog_bytes",
    [
        pytest.param(None, id="initial-creation"),
        pytest.param(b"downstream catalog must survive\r\n", id="replacement"),
    ],
)
def test_instruction_contract_catalog_protection_guard_requires_take_before_write(
    tmp_path: Path,
    existing_catalog_bytes: bytes | None,
) -> None:
    """Creating or replacing the catalog requires a path-scoped decision."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {
                "pattern": INSTRUCTION_CONTRACTS_PATH,
                "requires_all": ["template-sync-support"],
            }
        ],
    )
    catalog_bytes = b"instruction_contracts:\r\n  version: 1\r\n  contracts: []\r\n"
    source_catalog = template_root / INSTRUCTION_CONTRACTS_PATH
    source_catalog.parent.mkdir(parents=True, exist_ok=True)
    source_catalog.write_bytes(catalog_bytes)
    target_catalog = target_root / INSTRUCTION_CONTRACTS_PATH
    if existing_catalog_bytes is not None:
        target_catalog.parent.mkdir(parents=True, exist_ok=True)
        target_catalog.write_bytes(existing_catalog_bytes)

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "template-sync-support",
    )

    assert result.returncode == 2, result.stdout + result.stderr
    assert INSTRUCTION_CONTRACTS_PATH in result.stdout
    assert "unrecorded protected-file decision required" in result.stdout
    if existing_catalog_bytes is None:
        assert not target_catalog.exists()
    else:
        assert target_catalog.read_bytes() == existing_catalog_bytes

    if existing_catalog_bytes is not None:
        # The separate ordinary-overwrite guard still rejects a replacement.
        # Initial creation isolates the catalog's protected-classification guard.
        return
    mutant_dir = copy_materializer_runtime_for_mutation(tmp_path / "mutant")
    helper = mutant_dir / "template_sync_materialization_helpers.py"
    code = helper.read_text(encoding="utf-8")
    protected_catalog = '{".template-sync/instruction-contracts.yml"}'
    assert code.count(protected_catalog) == 1
    helper.write_text(code.replace(protected_catalog, "set()"), encoding="utf-8")
    mutant = subprocess.run(
        [
            sys.executable,
            str(mutant_dir / SCRIPT_PATH.name),
            "--template-root",
            str(template_root),
            "--target-root",
            str(target_root),
            "--source-repo",
            SOURCE_REPO,
            "--included-module",
            "template-sync-support",
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    assert mutant.returncode == 0, mutant.stdout + mutant.stderr
    assert target_catalog.read_bytes() == catalog_bytes


@pytest.mark.parametrize(
    "existing_catalog_bytes",
    [
        pytest.param(None, id="initial-creation"),
        pytest.param(b"replace this downstream catalog\n", id="replacement"),
    ],
)
def test_instruction_contract_catalog_take_is_byte_identical(
    tmp_path: Path,
    existing_catalog_bytes: bytes | None,
) -> None:
    """An explicit catalog TAKE copies the reviewed source bytes exactly."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {
                "pattern": INSTRUCTION_CONTRACTS_PATH,
                "requires_all": ["template-sync-support"],
            }
        ],
    )
    catalog_bytes = b"instruction_contracts:\r\n  version: 1\r\n  contracts: []\r\n"
    source_catalog = template_root / INSTRUCTION_CONTRACTS_PATH
    source_catalog.parent.mkdir(parents=True, exist_ok=True)
    source_catalog.write_bytes(catalog_bytes)
    target_catalog = target_root / INSTRUCTION_CONTRACTS_PATH
    if existing_catalog_bytes is not None:
        target_catalog.parent.mkdir(parents=True, exist_ok=True)
        target_catalog.write_bytes(existing_catalog_bytes)
    write_yaml(
        target_root / "decisions.yml",
        marker_document(
            ["template-sync-support"],
            protected_file_decisions=[
                {
                    "path": INSTRUCTION_CONTRACTS_PATH,
                    "decision": "TAKE",
                    "adoption_mode": "minimal-preservation",
                    "authorization_basis": "Fixture owner authorizes the reviewed catalog.",
                    "authorized_scope": f"{INSTRUCTION_CONTRACTS_PATH} only.",
                }
            ],
        ),
    )

    result = run_materialize(template_root, target_root, "--decisions-file", "decisions.yml")

    assert result.returncode == 0, result.stdout + result.stderr
    assert target_catalog.read_bytes() == catalog_bytes


def test_instruction_contract_catalog_is_omitted_when_template_sync_support_is_excluded(
    tmp_path: Path,
) -> None:
    """Excluding template-sync support omits the catalog without a decision."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(
        template_root,
        [
            {
                "pattern": INSTRUCTION_CONTRACTS_PATH,
                "requires_all": ["template-sync-support"],
            }
        ],
    )
    write_file(template_root / INSTRUCTION_CONTRACTS_PATH, "instruction_contracts: {}\n")

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 0, result.stdout + result.stderr
    assert not (target_root / INSTRUCTION_CONTRACTS_PATH).exists()
    assert "unrecorded protected-file decision required" not in result.stdout


def test_decisions_file_path_traversal_is_rejected(tmp_path: Path) -> None:
    """The decisions file must be repository-relative and stay inside the target."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "README.md", "requires_all": ["baseline"]}])
    write_file(template_root / "README.md", "template readme\n")

    result = run_materialize(
        template_root,
        target_root,
        "--decisions-file",
        "../outside.yml",
    )

    assert result.returncode == 1
    assert "--decisions-file must not contain traversal segments" in result.stderr


def test_format_cli_error_summarizes_oserror_without_path() -> None:
    """OSError output is summarized via os_error_summary and omits the path."""
    secret_path = "/home/secret-user/private/secret.key"
    error = FileNotFoundError(2, "No such file or directory", secret_path)

    message = materializer.format_cli_error(error)

    assert message == "FileNotFoundError: No such file or directory"
    assert secret_path not in message


def test_format_cli_error_summarizes_shutil_error_without_paths() -> None:
    """shutil.Error (an OSError subclass) is summarized without its path tuples."""
    secret_path = "/home/secret-user/private"
    error = shutil.Error([(f"{secret_path}/a", f"{secret_path}/b", "boom")])

    message = materializer.format_cli_error(error)

    assert secret_path not in message
    assert message == "Error: I/O error"


def test_format_cli_error_preserves_domain_error_message() -> None:
    """Domain errors already carry path-safe messages and are returned verbatim."""
    error = materializer.MaterializationError("safe domain message")

    assert materializer.format_cli_error(error) == "safe domain message"


def run_workflow_render_error_control(
    read_path: str, error_kind: str, *, remove_formatter: bool = False
) -> subprocess.CompletedProcess[str]:
    """Exercise actual render reads and the real CLI boundary in a fresh interpreter."""
    script = r"""
import argparse
import importlib.util
import sys
from pathlib import Path
from unittest.mock import patch

source_path = Path(sys.argv[1])
sys.path.insert(0, str(source_path.parent))
spec = importlib.util.spec_from_file_location("render_error_control", source_path)
module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = module
source = source_path.read_text(encoding="utf-8")
if sys.argv[4] == "mutant":
    anchor = 'f"Workflow contract rendering failed: {format_cli_error(error)}; "'
    assert source.count(anchor) == 1
    source = source.replace(anchor, 'f"Workflow contract rendering failed: {error}; "')
exec(compile(source, str(source_path), "exec"), module.__dict__)
root = source_path.parents[2]
private = "PRIVATE_RENDER_FILENAME"
errors = {
    "permission": PermissionError(13, "Permission denied", private + "/source"),
    "missing": FileNotFoundError(2, "No such file or directory", private + "/source"),
    "two-names": OSError(5, "Input/output error", private + "/source", None, private + "/target"),
    "no-strerror": OSError(None, None, private + "/source"),
    "domain": ValueError("fixed domain diagnostic"),
    "unexpected": RuntimeError("fixed unexpected diagnostic"),
}
injected_error = errors[sys.argv[3]]
original_open = Path.open
def injected_open(path, *args, **kwargs):
    if path.as_posix().endswith(sys.argv[2]):
        raise injected_error
    return original_open(path, *args, **kwargs)

def render(args):
    try:
        return module.render_workflow_contract(
            root, (), {"github-actions"}, template_paths=[],
            expected_workflows={".github/workflows/markdownlint.yml"},
        )
    except module.MaterializationError as error:
        assert error.__cause__ is injected_error
        raise

module.parse_args = lambda argv: argparse.Namespace()
module.materialize = render
with patch.object(Path, "open", injected_open):
    sys.exit(module.main([]))
"""
    return subprocess.run(
        [
            sys.executable,
            "-B",
            "-c",
            script,
            str(SCRIPT_PATH),
            read_path,
            error_kind,
            "mutant" if remove_formatter else "fixed",
        ],
        check=False,
        capture_output=True,
        text=True,
        timeout=30,
    )


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    "read_path",
    [
        ".github/workflows/markdownlint.yml",
        ".github/instructions/yaml.instructions.md",
        "schemas/workflow-security-contract.schema.json",
    ],
)
@pytest.mark.parametrize(
    ("error_kind", "summary"),
    [
        ("permission", "PermissionError: Permission denied"),
        ("missing", "FileNotFoundError: No such file or directory"),
        ("two-names", "OSError: Input/output error"),
        ("no-strerror", "OSError: I/O error"),
    ],
)
def test_workflow_render_oserror_preserves_failure_without_filenames(
    read_path: str, error_kind: str, summary: str
) -> None:
    """Workflow, example and schema read errors retain safe context and native failure."""
    result = run_workflow_render_error_control(read_path, error_kind)
    assert result.returncode == 1, result.stdout + result.stderr
    assert f"ERROR: Workflow contract rendering failed: {summary};" in result.stderr
    assert "review and update the installed materializer tool bundle" in result.stderr
    assert "PRIVATE_RENDER_FILENAME" not in result.stderr
    assert "Traceback" not in result.stderr


@pytest.mark.upstream_template_only
def test_workflow_render_formatter_removal_breaks_fixed_privacy_oracle() -> None:
    """Restoring only the old wrapper leaks filenames while preserving native failure."""
    result = run_workflow_render_error_control(
        ".github/workflows/markdownlint.yml", "two-names", remove_formatter=True
    )
    assert result.returncode == 1, result.stdout + result.stderr
    assert "Workflow contract rendering failed:" in result.stderr
    assert "PRIVATE_RENDER_FILENAME/source" in result.stderr
    assert "PRIVATE_RENDER_FILENAME/target" in result.stderr
    assert "Traceback" not in result.stderr
    with pytest.raises(AssertionError):
        assert "PRIVATE_RENDER_FILENAME" not in result.stderr


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("error_kind", ["domain", "unexpected"])
def test_workflow_render_other_error_semantics_remain_visible(error_kind: str) -> None:
    """Safe domain diagnostics survive, and unrelated runtime errors remain uncaught."""
    result = run_workflow_render_error_control(".github/workflows/markdownlint.yml", error_kind)
    assert result.returncode == 1, result.stdout + result.stderr
    if error_kind == "domain":
        assert "Workflow contract rendering failed: fixed domain diagnostic;" in result.stderr
        assert "Traceback" not in result.stderr
    else:
        assert "RuntimeError: fixed unexpected diagnostic" in result.stderr
        assert "Traceback" in result.stderr
        assert "Workflow contract rendering failed:" not in result.stderr


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("enforcement", [False, True])
def test_workflow_render_preserves_named_optional_checks(enforcement: bool) -> None:
    """Real successful rendering retains fixed check identities and optional omission."""
    _, _, mappings = materializer.load_validated_manifest_context(REPO_ROOT)
    modules = {"github-actions"}
    if enforcement:
        modules.update({"agent-instructions", "instruction-enforcement", "agent-codex"})
    rendered = yaml.safe_load(materializer.render_workflow_contract(REPO_ROOT, mappings, modules))
    workflows = rendered["workflows"]
    assert (
        workflows[".github/workflows/workflow-security.yml"]["jobs"]["validate"]["controls"]["name"]
        == "Workflow Security"
    )
    instruction = ".github/workflows/instruction-contracts.yml"
    if enforcement:
        assert workflows[instruction]["jobs"]["validate"]["controls"]["name"] == (
            "Instruction Contracts"
        )
    else:
        assert instruction not in workflows


def test_summarize_helper_failure_includes_exit_code_and_output() -> None:
    """The failure summary surfaces the exit code and the helper's findings."""
    summary = materializer.summarize_helper_failure(
        returncode=1,
        stdout="Placeholder scan found issues:\n  - README.md:12: forbidden: x (bad)",
        stderr="ERROR: boom",
    )

    assert "exit code 1" in summary
    assert "README.md:12: forbidden: x (bad)" in summary
    assert "ERROR: boom" in summary


def test_summarize_helper_failure_bounds_output_lines() -> None:
    """Long helper streams are truncated to the most recent lines."""
    stdout = "\n".join(f"line {index}" for index in range(50))

    summary = materializer.summarize_helper_failure(
        returncode=2,
        stdout=stdout,
        stderr="",
        line_limit=10,
    )

    assert "line 49" in summary
    assert "line 39" not in summary
    assert "showing last 10 of 50 lines" in summary


def test_summarize_helper_failure_handles_empty_output() -> None:
    """With no captured output, only the exit-code line is returned."""
    assert (
        materializer.summarize_helper_failure(returncode=3, stdout="", stderr="")
        == "Placeholder helper failed with exit code 3."
    )


def test_non_regular_target_path_is_reported_with_repo_relative_path(tmp_path: Path) -> None:
    """A directory where a staged file would land aborts with an actionable, path-named error."""
    template_root = tmp_path / "template"
    target_root = tmp_path / "target"
    target_root.mkdir()
    prepare_template(template_root, [{"pattern": "README.md", "requires_all": ["baseline"]}])
    write_file(template_root / "README.md", "template readme\n")
    # The downstream repository has a directory where the staged file would land.
    (target_root / "README.md").mkdir()

    result = run_materialize(
        template_root,
        target_root,
        "--source-repo",
        SOURCE_REPO,
        "--included-module",
        "baseline",
    )

    assert result.returncode == 1, result.stdout
    assert "README.md" in result.stderr
    assert "not a regular file" in result.stderr


def test_validate_computed_marker_error_names_marker_document_not_schema() -> None:
    """A computed-marker schema failure is reported against the marker document path."""
    with pytest.raises(materializer.TemplateSyncMaterializationError) as excinfo:
        materializer.validate_computed_marker({"template_sync": {}}, template_root=REPO_ROOT)

    message = str(excinfo.value)
    assert ".template-sync/marker.yml" in message
    assert "schema.json" not in message


def test_ensure_regular_target_rejects_non_regular_path_with_repo_relative_path(
    tmp_path: Path,
) -> None:
    """A directory where a regular file is expected raises naming the repo-relative path."""
    conflict = tmp_path / ".template-sync" / "marker.yml"
    conflict.mkdir(parents=True)

    with pytest.raises(materializer.MaterializationError) as excinfo:
        materializer.ensure_regular_target(conflict, ".template-sync/marker.yml")

    message = str(excinfo.value)
    assert ".template-sync/marker.yml" in message
    assert "not a regular file" in message


def test_ensure_regular_target_allows_regular_file_and_missing_path(tmp_path: Path) -> None:
    """Regular files and missing paths pass the guard without raising."""
    regular = tmp_path / "file.txt"
    regular.write_text("content", encoding="utf-8")

    materializer.ensure_regular_target(regular, "file.txt")
    materializer.ensure_regular_target(tmp_path / "missing.txt", "missing.txt")


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    ("profile", "modules"),
    [
        pytest.param(
            "no-baseline",
            ("markdown", "github-actions"),
            id="no-baseline",
        ),
        pytest.param(
            "neither",
            ("baseline", "markdown", "github-actions"),
            id="neither-python-nor-support",
        ),
        pytest.param(
            "support-only",
            ("baseline", "markdown", "github-actions", "template-sync-support"),
            id="support-only",
        ),
        pytest.param(
            "python-only",
            ("baseline", "markdown", "github-actions", "python"),
            id="python-only",
        ),
        pytest.param(
            "support-python",
            (
                "baseline",
                "markdown",
                "github-actions",
                "template-sync-support",
                "python",
            ),
            id="support-plus-python",
        ),
    ],
)
def test_materialized_markdown_ci_closes_source_regression_prerequisites(
    tmp_path: Path,
    profile: str,
    modules: tuple[str, ...],
) -> None:
    """Generated Markdown CI keeps core checks and only applicable source regressions."""
    support_retained = "template-sync-support" in modules
    python_retained = "python" in modules
    target = materialize_module_fixture(
        tmp_path,
        modules,
        authorize_protected_files=support_retained,
    )
    workflow_path = target / ".github/workflows/markdownlint.yml"
    workflow_text = read_file(workflow_path)
    workflow = yaml.safe_load(workflow_text)
    steps = workflow["jobs"]["markdownlint"]["steps"]
    by_name = {step["name"]: step for step in steps}

    for core_step in (
        "Run markdownlint on outer files",
        "Run markdownlint on nested Markdown code fences",
        "Validate Markdown links",
        "Run nested Markdown regression tests",
    ):
        assert core_step in by_name, f"{profile}: {core_step}"

    support_steps = (
        "Setup Python",
        "Install Python dependencies",
        "Install pinned pre-commit runner",
        "Lint materialized generated output (nested Markdown)",
        "Run actual nested Markdown hook regression",
        "Check source-template regression results",
    )
    assert all((name in by_name) is support_retained for name in support_steps)
    assert (
        "test_nested_markdown_hook_actual_invocation_and_mutation" in workflow_text
    ) is support_retained
    assert ("steps.test-nested-hook.outcome" in workflow_text) is support_retained
    assert (
        "# template-sync: begin template-sync-support-only" in workflow_text
    ) is support_retained

    assert (target / "pyproject.toml").is_file() is python_retained
    assert (target / "tests/test_materialize_downstream_adoption.py").is_file() is support_retained
    if support_retained:
        setup_condition = " ".join(by_name["Setup Python"]["if"].split())
        generated_condition = " ".join(
            by_name["Lint materialized generated output (nested Markdown)"]["if"].split()
        )
        for condition in (setup_condition, generated_condition):
            assert "hashFiles('pyproject.toml') != ''" in condition
            assert "hashFiles('tests/test_materialize_downstream_adoption.py') != ''" in condition
        assert (python_retained and support_retained) is (
            (target / "pyproject.toml").is_file()
            and (target / "tests/test_materialize_downstream_adoption.py").is_file()
        )
        assert by_name["Install pinned pre-commit runner"]["if"] == (
            "${{ github.repository == 'franklesniak/copilot-repo-template' }}"
        )
        assert by_name["Run actual nested Markdown hook regression"]["if"] == (
            "${{ github.repository == 'franklesniak/copilot-repo-template' }}"
        )


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    "profile", ["full", "github-only", "no-support", "no-policy-engine", "no-github"]
)
def test_materialized_workflows_keep_all_checkout_credentials_disabled(
    tmp_path: Path, profile: str
) -> None:
    """Generated workflows preserve checkout intent with and without policy enforcement."""
    excluded = {
        "full": set(),
        "github-only": set(AZURE_TEMPLATE_MODULES),
        "no-support": {"template-sync-support"},
        "no-policy-engine": {"template-sync-support", "instruction-enforcement"},
        "no-github": set(GITHUB_HOST_TEMPLATE_MODULES),
    }[profile]
    modules = tuple(module for module in FULL_TEMPLATE_MODULES if module not in excluded)
    target = materialize_module_fixture(tmp_path, modules, authorize_protected_files=True)
    paths = sorted((target / ".github/workflows").glob("*.yml"))
    if profile == "no-github":
        assert not paths
        return
    expected_workflows = {
        "auto-fix-precommit.yml",
        "check-placeholders.yml",
        "copilot-setup-steps.yml",
        "data-ci.yml",
        "markdownlint.yml",
        "powershell-ci.yml",
        "precommit-ci.yml",
        "python-ci.yml",
        "terraform-ci.yml",
        "toolchain-eol.yml",
        "workflow-security.yml",
    }
    if "instruction-enforcement" in modules:
        expected_workflows.add("instruction-contracts.yml")
    assert {path.name for path in paths} == expected_workflows
    checkout_count = 0
    for path in paths:
        workflow = yaml.safe_load(read_file(path))
        for job in workflow["jobs"].values():
            for step in job.get("steps", []):
                if step.get("uses", "").startswith("actions/checkout@"):
                    checkout_count += 1
                    assert job.get("permissions", workflow.get("permissions")) == {
                        "contents": "read"
                    }
                    value = step.get("with", {}).get("persist-credentials")
                    assert value is False or (
                        isinstance(value, str) and value.strip(" \t\r\n").lower() == "false"
                    ), str(path)
    assert checkout_count >= 15 + ("instruction-enforcement" in modules)
    if "template-sync-support" not in modules:
        assert not (target / "tests/test_precommit_runner.py").exists()


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("modules", [("baseline",), ("baseline", "markdown"), ("markdown",)])
def test_materialized_contributor_node_setup_matches_retained_tooling(
    tmp_path: Path, modules: tuple[str, ...]
) -> None:
    """Complete setup and lint sections disappear with Markdown, without requiring baseline."""
    target = materialize_module_fixture(tmp_path, modules)
    for relative_path in (
        "package.json",
        "package-lock.json",
        ".github/scripts/lint-nested-markdown.js",
    ):
        assert (target / relative_path).is_file() == ("markdown" in modules)
    contributor = target / "CONTRIBUTING.md"
    assert contributor.is_file() == ("baseline" in modules)
    if not contributor.is_file():
        return
    text = read_file(contributor)
    if "markdown" in modules:
        assert "### 2. Install Node.js Dependencies" in text
        assert "npm ci --ignore-scripts" in text
        assert "npm run lint:md:nested" in text
    else:
        assert "Install Node.js Dependencies" not in text
        assert "npm " not in text
    assert "Git hooks are managed by pre-commit." in text


@pytest.mark.upstream_template_only
def test_onboarding_locked_setup_matches_selected_ci() -> None:
    """Ordinary setup preserves the root lock; optional Node guidance is conditional."""
    sections = (
        ("GETTING_STARTED_NEW_REPO.md", "### Confirm Local Validation Prerequisites"),
        ("GETTING_STARTED_EXISTING_REPO.md", "### Local Validation Prerequisites"),
    )
    for relative, heading in sections:
        text = read_file(REPO_ROOT / relative)
        section = text.split(heading, 1)[1].split("\n### ", 1)[0]
        assert "npm ci --ignore-scripts" in section
        assert "root lockfile are retained" in section
        assert "Profiles that omit Markdown do not need Node" in section
        assert "`npm install` or `npm ci`" not in section
    optional = read_file(REPO_ROOT / "OPTIONAL_CONFIGURATIONS.md")
    guidance = optional.split("> **Node availability requirement.**", 1)[1].split("\n\n", 1)[0]
    assert "When Markdown is selected" in guidance
    assert "GitHub and Azure pre-commit CI routes already install Node" in guidance
    assert "Profiles that omit Markdown do not retain those setup steps" in guidance
    assert "npm ci --ignore-scripts" in guidance
    assert "default CI workflows do not install Node" not in guidance


@pytest.mark.upstream_template_only
@pytest.mark.parametrize("keep_markdown", [True, False])
def test_materialized_aggregate_hooks_match_markdown_toolchain(
    tmp_path: Path, keep_markdown: bool
) -> None:
    """Actual host materializations retain or remove nested hook and its complete setup."""
    modules = tuple(
        module for module in FULL_TEMPLATE_MODULES if keep_markdown or module != "markdown"
    )
    target = materialize_module_fixture(tmp_path, modules, authorize_protected_files=True)
    config = yaml.safe_load(read_file(target / ".pre-commit-config.yaml"))
    hooks = [hook for repo in config["repos"] for hook in repo["hooks"]]
    assert any(hook["id"] == "lint-nested-markdown" for hook in hooks) == keep_markdown
    assert (target / ".github/scripts/lint_nested_markdown_hook.py").is_file() == keep_markdown
    for relative in (
        ".github/workflows/precommit-ci.yml",
        ".github/workflows/auto-fix-precommit.yml",
        ".azuredevops/pipelines/precommit.yml",
    ):
        text = read_file(target / relative)
        assert ("npm ci --ignore-scripts" in text) == keep_markdown
        assert ("Node" in text) == keep_markdown
        document = yaml.safe_load(text)
        steps = document.get("steps") or next(iter(document["jobs"].values()))["steps"]
        if keep_markdown:
            install = next(
                index
                for index, step in enumerate(steps)
                if step.get("run", step.get("bash")) == "npm ci --ignore-scripts"
            )
            check = next(
                index
                for index, step in enumerate(steps)
                if "pre-commit run --all-files" in step.get("run", step.get("bash", ""))
            )
            assert install < check


@pytest.mark.upstream_template_only
@pytest.mark.parametrize(
    "modules",
    [
        ("agent-instructions", "yaml"),
        ("baseline", "agent-instructions", "yaml"),
        ("agent-instructions",),
    ],
)
def test_materialized_yaml_policy_does_not_depend_on_baseline_or_markdown(
    tmp_path: Path, modules: tuple[str, ...]
) -> None:
    """Portable encoding and live URL rules survive actual optional-module pruning."""
    target = materialize_module_fixture(tmp_path, modules, authorize_protected_files=True)
    guide = target / ".github/instructions/yaml.instructions.md"
    assert guide.is_file() == ("yaml" in modules)
    if not guide.is_file():
        return
    text = read_file(guide)
    assert "YAML files **MUST** use UTF-8 without a byte-order mark (BOM)." in text
    assert "URLs **MUST** use literal `OWNER/REPO`" in text
    assert "notation **MAY** appear in explicitly labeled schematic prose" in text
    assert "Line-ending, BOM" not in text
    assert (
        "gitattributes.instructions.md" in text
        if "baseline" in modules
        else "gitattributes.instructions.md" not in text
    )
    assert not (target / ".github/instructions/docs.instructions.md").exists()


@pytest.mark.upstream_template_only
@pytest.mark.slow
@pytest.mark.parametrize(
    ("modules", "expected_languages"),
    [
        pytest.param(
            ("agent-instructions", "powershell"),
            frozenset({"powershell"}),
            id="powershell-only-without-markdown-or-support",
        ),
        pytest.param(
            ("agent-instructions", "terraform", "markdown", "template-sync-support"),
            frozenset({"terraform"}),
            id="terraform-only-with-markdown-and-support",
        ),
        pytest.param(
            ("agent-instructions", "powershell", "terraform", "markdown"),
            frozenset({"powershell", "terraform"}),
            id="both-languages-with-markdown-without-support",
        ),
        pytest.param(
            ("agent-instructions",),
            frozenset(),
            id="neither-language",
        ),
        pytest.param(
            ("powershell", "terraform", "template-sync-support"),
            frozenset(),
            id="languages-without-agent-instructions",
        ),
    ],
)
def test_materialized_upstream_style_guide_provenance_profiles(
    tmp_path: Path,
    modules: tuple[str, ...],
    expected_languages: frozenset[str],
) -> None:
    """Real outputs retain only provenance for imported guides in scope."""
    target = materialize_module_fixture(tmp_path, modules, authorize_protected_files=True)
    provenance = target / UPSTREAM_STYLE_GUIDES_PATH
    assert provenance.is_file() is bool(expected_languages)
    assert not (target / "TEMPLATE_MAINTENANCE.md").exists()
    assert (target / ".template-sync/marker.yml").is_file() is ("template-sync-support" in modules)
    assert (target / ".github/instructions/docs.instructions.md").is_file() is (
        "agent-instructions" in modules and "markdown" in modules
    )

    if provenance.is_file():
        text = read_file(provenance)
        language_tokens = {
            "powershell": (
                "franklesniak/PSStyleGuide",
                "534762988c0634d34c01059c9acb310e14c24371",
                "powershell-reference-only",
            ),
            "terraform": (
                "franklesniak/TerraformStyleGuide",
                "e81f68e38b49eebdb9669eb406c918f64f0e35fd",
                "terraform-reference-only",
            ),
        }
        for language, tokens in language_tokens.items():
            for token in tokens:
                assert (token in text) is (language in expected_languages)

    before = snapshot_fixture_files(target)
    repeated = run_materialize(REPO_ROOT, target, "--decisions-file", "decisions.yml")
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert snapshot_fixture_files(target) == before


@pytest.mark.upstream_template_only
@pytest.mark.slow
def test_upstream_style_guide_provenance_update_is_owner_selective(
    tmp_path: Path,
) -> None:
    """A/B review preserves custom bytes until the owner selects each path."""
    source = tmp_path / "source"
    source.mkdir()
    copy_tracked_worktree(source)
    copy_template_file(source, UPSTREAM_STYLE_GUIDES_PATH)
    revision_a = commit_fixture_template(source)
    target = tmp_path / "target"
    target.mkdir()
    modules = ("agent-instructions", "powershell", "template-sync-support")
    protected_take = protected_take_decisions_for_modules(modules)
    write_yaml(
        target / "decisions.yml",
        marker_document(
            list(modules),
            last_reviewed_template_commit=revision_a,
            protected_file_decisions=protected_take,
        ),
    )
    adopted = run_materialize(source, target, "--decisions-file", "decisions.yml")
    assert adopted.returncode == 0, adopted.stdout + adopted.stderr

    marker_path = target / ".template-sync/marker.yml"
    marker_a = marker_path.read_bytes()
    downstream_guide = target / POWERSHELL_GUIDE_PATH
    downstream_provenance = target / UPSTREAM_STYLE_GUIDES_PATH
    write_file(downstream_guide, read_file(downstream_guide) + "\n<!-- downstream guide note -->\n")
    write_file(
        downstream_provenance,
        read_file(downstream_provenance) + "\n<!-- downstream provenance note -->\n",
    )
    customized_guide = downstream_guide.read_bytes()
    customized_provenance = downstream_provenance.read_bytes()

    source_guide = source / POWERSHELL_GUIDE_PATH
    source_provenance = source / UPSTREAM_STYLE_GUIDES_PATH
    write_file(source_guide, read_file(source_guide) + "\n<!-- reviewed source revision B -->\n")
    write_file(
        source_provenance,
        read_file(source_provenance) + "\n<!-- reviewed provenance revision B -->\n",
    )
    revision_b = commit_fixture_template(source)
    assert revision_b != revision_a

    write_yaml(
        target / "decisions.yml",
        marker_document(list(modules), last_reviewed_template_commit=revision_b),
    )
    blocked = run_materialize(source, target, "--decisions-file", "decisions.yml")
    assert blocked.returncode == 2, blocked.stdout + blocked.stderr
    assert marker_path.read_bytes() == marker_a
    assert downstream_guide.read_bytes() == customized_guide
    assert downstream_provenance.read_bytes() == customized_provenance

    protected_skip = [
        (
            {
                "path": POWERSHELL_GUIDE_PATH,
                "decision": "SKIP",
                "reason": "Fixture owner keeps the reviewed local guide customization.",
            }
            if decision["path"] == POWERSHELL_GUIDE_PATH
            else decision
        )
        for decision in protected_take
    ]
    write_yaml(
        target / "decisions.yml",
        marker_document(
            list(modules),
            last_reviewed_template_commit=revision_b,
            protected_file_decisions=protected_skip,
            local_overrides=[
                {
                    "path": UPSTREAM_STYLE_GUIDES_PATH,
                    "default_decision": "SKIP",
                    "reason": "Fixture owner keeps the reviewed local provenance note.",
                }
            ],
        ),
    )
    skipped = run_materialize(source, target, "--decisions-file", "decisions.yml")
    assert skipped.returncode == 0, skipped.stdout + skipped.stderr
    assert downstream_guide.read_bytes() == customized_guide
    assert downstream_provenance.read_bytes() == customized_provenance
    marker = as_mapping(load_yaml(marker_path), "marker must be a mapping")
    template_sync = as_mapping(marker["template_sync"], "template_sync must be a mapping")
    assert template_sync["last_reviewed_template_commit"] == revision_b

    write_yaml(
        target / "decisions.yml",
        marker_document(
            list(modules),
            last_reviewed_template_commit=revision_b,
            protected_file_decisions=protected_take,
            local_overrides=[
                {
                    "path": UPSTREAM_STYLE_GUIDES_PATH,
                    "default_decision": "TAKE",
                    "reason": "Fixture owner accepts the reviewed provenance update.",
                }
            ],
        ),
    )
    accepted = run_materialize(source, target, "--decisions-file", "decisions.yml")
    assert accepted.returncode == 0, accepted.stdout + accepted.stderr

    clean_b = tmp_path / "clean-b"
    clean_b.mkdir()
    write_yaml(
        clean_b / "decisions.yml",
        marker_document(
            list(modules),
            last_reviewed_template_commit=revision_b,
            protected_file_decisions=protected_take,
        ),
    )
    clean_b_result = run_materialize(source, clean_b, "--decisions-file", "decisions.yml")
    assert clean_b_result.returncode == 0, clean_b_result.stdout + clean_b_result.stderr
    assert downstream_guide.read_bytes() == (clean_b / POWERSHELL_GUIDE_PATH).read_bytes()
    assert downstream_provenance.read_bytes() == (clean_b / UPSTREAM_STYLE_GUIDES_PATH).read_bytes()
    assert "franklesniak/PSStyleGuide" in read_file(downstream_provenance)
    assert "franklesniak/TerraformStyleGuide" not in read_file(downstream_provenance)
    assert "reviewed source revision B" in read_file(downstream_guide)
    assert "reviewed provenance revision B" in read_file(downstream_provenance)
    assert "downstream guide note" not in read_file(downstream_guide)
    assert "downstream provenance note" not in read_file(downstream_provenance)
    before_noop = snapshot_fixture_files(target)
    repeated = run_materialize(source, target, "--decisions-file", "decisions.yml")
    assert repeated.returncode == 0, repeated.stdout + repeated.stderr
    assert snapshot_fixture_files(target) == before_noop


@pytest.mark.upstream_template_only
@pytest.mark.slow
def test_upstream_style_guide_language_removal_requires_explicit_cleanup(
    tmp_path: Path,
) -> None:
    """Removing the final imported languages never silently deletes old files."""
    source = tmp_path / "source"
    source.mkdir()
    copy_tracked_worktree(source)
    copy_template_file(source, UPSTREAM_STYLE_GUIDES_PATH)
    revision = commit_fixture_template(source)
    retained = ("agent-instructions", "powershell", "terraform", "template-sync-support")
    omitted = ("agent-instructions", "template-sync-support")

    target = tmp_path / "target"
    target.mkdir()
    write_yaml(
        target / "decisions.yml",
        marker_document(
            list(retained),
            last_reviewed_template_commit=revision,
            protected_file_decisions=protected_take_decisions_for_modules(retained),
        ),
    )
    first = run_materialize(source, target, "--decisions-file", "decisions.yml")
    assert first.returncode == 0, first.stdout + first.stderr
    original = snapshot_fixture_files(target)

    _manifest, _module_order, mappings = materializer.load_validated_manifest_context(source)
    overrides: list[dict[str, str]] = []
    removals: list[str] = []
    for relative_path in original:
        relation = materializer.selected_relation_for_path(relative_path, mappings)
        if relation is None:
            continue
        if relation.is_retained_by(omitted):
            overrides.append(
                {
                    "path": relative_path,
                    "default_decision": "TAKE",
                    "reason": "Fixture owner approved the retained-file pruning update.",
                }
            )
        else:
            removals.append(relative_path)
    assert {UPSTREAM_STYLE_GUIDES_PATH, POWERSHELL_GUIDE_PATH, TERRAFORM_GUIDE_PATH}.issubset(
        removals
    )
    write_yaml(
        target / "decisions.yml",
        marker_document(
            list(omitted),
            last_reviewed_template_commit=revision,
            local_overrides=overrides,
            protected_file_decisions=protected_take_decisions_for_modules(omitted),
            protected_guide_contract_waivers=github_agent_azure_protocol_section_waivers(),
        ),
    )
    removal_review = run_materialize(source, target, "--decisions-file", "decisions.yml")
    assert removal_review.returncode == 0, removal_review.stdout + removal_review.stderr
    for relative_path in (
        UPSTREAM_STYLE_GUIDES_PATH,
        POWERSHELL_GUIDE_PATH,
        TERRAFORM_GUIDE_PATH,
    ):
        assert (target / relative_path).read_bytes() == original[relative_path]

    run_git(target, "init", "-q")
    run_git(target, "add", ".")
    before_cleanup = run_downstream_adoption_validator(target)
    assert before_cleanup.returncode == 1, before_cleanup.stdout + before_cleanup.stderr
    for relative_path in (
        UPSTREAM_STYLE_GUIDES_PATH,
        POWERSHELL_GUIDE_PATH,
        TERRAFORM_GUIDE_PATH,
    ):
        assert relative_path in before_cleanup.stdout

    for relative_path in removals:
        (target / relative_path).unlink()
    run_git(target, "add", "-A")
    after_cleanup = run_downstream_adoption_validator(target)
    assert after_cleanup.returncode == 0, after_cleanup.stdout + after_cleanup.stderr

    clean = tmp_path / "clean"
    clean.mkdir()
    write_yaml(
        clean / "decisions.yml",
        marker_document(
            list(omitted),
            last_reviewed_template_commit=revision,
            protected_file_decisions=protected_take_decisions_for_modules(omitted),
            protected_guide_contract_waivers=github_agent_azure_protocol_section_waivers(),
        ),
    )
    clean_result = run_materialize(source, clean, "--decisions-file", "decisions.yml")
    assert clean_result.returncode == 0, clean_result.stdout + clean_result.stderr
    ignored = {"decisions.yml", ".template-sync/marker.yml"}
    assert {
        path: content
        for path, content in snapshot_fixture_files(target).items()
        if path not in ignored
    } == {
        path: content
        for path, content in snapshot_fixture_files(clean).items()
        if path not in ignored
    }
