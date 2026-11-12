"""Exercise first-adoption quality-debt report helpers."""

from __future__ import annotations

import importlib
import io
import json
import shutil
import subprocess
import sys
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, cast

from tests._pytest_compat import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_PATH = REPO_ROOT / ".template-sync" / "scripts" / "first_adoption_quality_reports.py"
SCRIPT_DIR = SCRIPT_PATH.parent

if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

quality_reports = cast(Any, importlib.import_module("first_adoption_quality_reports"))


def _run_git(repo_root: Path, *args: str) -> str:
    """Run a Git command in a fixture repository and return stdout."""
    result = subprocess.run(
        ["git", *args],
        cwd=repo_root,
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stderr
    return result.stdout.strip()


def _write_text(repo_root: Path, relative_path: str, text: str = "content\n") -> None:
    """Write a UTF-8 fixture file."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _write_bytes(repo_root: Path, relative_path: str, data: bytes) -> None:
    """Write a binary fixture file."""
    path = repo_root / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def _init_repo(repo_root: Path) -> None:
    """Initialize a fixture Git repository."""
    _run_git(repo_root, "init")
    _run_git(repo_root, "config", "user.email", "test@example.invalid")
    _run_git(repo_root, "config", "user.name", "Test User")


def _write_powershell_report_prerequisites(repo_root: Path) -> None:
    """Write the helper files required before the report invokes the runner."""
    _write_text(repo_root, ".github/linting/PSScriptAnalyzerSettings.psd1", "@{}\n")
    _write_text(
        repo_root,
        "src/tools/Resolve-PSScriptAnalyzerCandidate.ps1",
        "function Resolve-PSScriptAnalyzerCandidate {}\n",
    )
    _write_text(
        repo_root,
        "src/tools/Resolve-PSScriptAnalyzerGate.ps1",
        "function Resolve-PSScriptAnalyzerGate {}\n",
    )


def _candidate_summary(
    repo_root: Path,
    *,
    selected: Sequence[Mapping[str, object]] | None = None,
    policy_excluded: Sequence[Mapping[str, object]] | None = None,
    unsafe: Sequence[Mapping[str, object]] | None = None,
) -> dict[str, object]:
    """Build a composite candidate summary payload."""
    selected_records = list(selected or [])
    policy_excluded_records = list(policy_excluded or [])
    unsafe_records = list(unsafe or [])
    if selected is None and not policy_excluded_records and not unsafe_records:
        selected_records = [
            {
                "CandidateFullName": str(repo_root / "src" / "sample.ps1"),
                "RepositoryRelativePath": "src/sample.ps1",
                "EscapedAnalyzerPath": str(repo_root / "src" / "sample.ps1"),
                "OutcomeCategory": "selected",
                "ReasonCode": "Selected",
            }
        ]
    return {
        "Selected": selected_records,
        "PolicyExcluded": policy_excluded_records,
        "Unsafe": unsafe_records,
        "SummaryCounts": {
            "Selected": len(selected_records),
            "PolicyExcluded": len(policy_excluded_records),
            "Unsafe": len(unsafe_records),
        },
    }


def _gate_payload(
    *,
    findings: Sequence[Mapping[str, object]] = (),
    should_fail: bool = False,
    recommended_mode: str = "strict",
) -> dict[str, object]:
    """Build a valid structured Gate payload."""
    return {
        "Status": "ok",
        "Mode": "first-adoption",
        "ShouldFail": should_fail,
        "RecommendedMode": recommended_mode,
        "SummaryLines": [
            "PSScriptAnalyzer gate mode: first-adoption.",
            "Result: pass.",
        ],
        "IssueReadyMarkdown": [
            "## PSScriptAnalyzer First-Adoption Debt Cleanup",
            "",
        ],
        "Findings": list(findings),
    }


def _write_retained_powershell_ci_surfaces(repo_root: Path) -> tuple[str, str]:
    """Write retained GitHub Actions and Azure Pipelines analyzer surfaces."""
    github_workflow = (
        "name: PowerShell CI\n"
        "on:\n"
        "  push:\n"
        "permissions:\n"
        "  contents: read\n"
        "jobs:\n"
        "  powershell-lint:\n"
        "    runs-on: ubuntu-latest\n"
        "    env:\n"
        "      PSSCRIPTANALYZER_GATE_MODE: strict\n"
        "    steps:\n"
        "      - name: Run PSScriptAnalyzer\n"
        "        shell: pwsh\n"
        "        run: |\n"
        "          $gateResult = Resolve-PSScriptAnalyzerGate `\n"
        "            -Mode $env:PSSCRIPTANALYZER_GATE_MODE `\n"
        "            -AnalyzerFinding @()\n"
    )
    azure_pipeline = (
        "parameters:\n"
        "  - name: gateMode\n"
        "    type: string\n"
        '    default: "strict"\n'
        "jobs:\n"
        "  - job: lint\n"
        "    variables:\n"
        '      PSSCRIPTANALYZER_GATE_MODE: "${{ parameters.gateMode }}"\n'
        "    steps:\n"
        "      - task: PowerShell@2\n"
        "        inputs:\n"
        '          targetType: "inline"\n'
        "          script: |\n"
        "            $gateResult = Resolve-PSScriptAnalyzerGate `\n"
        "              -Mode $env:PSSCRIPTANALYZER_GATE_MODE `\n"
        "              -AnalyzerFinding @()\n"
    )
    _write_text(repo_root, ".github/workflows/powershell-ci.yml", github_workflow)
    _write_text(repo_root, ".azuredevops/pipelines/powershell-ci.yml", azure_pipeline)
    return github_workflow, azure_pipeline


def test_quality_file_discovery_supports_tracked_only_and_ignored_controls(
    tmp_path: Path,
) -> None:
    """Report discovery defaults to tracked plus untracked non-ignored files."""
    _init_repo(tmp_path)
    _write_text(tmp_path, ".gitignore", "ignored.txt\n")
    _write_text(tmp_path, "tracked.txt")
    _write_text(tmp_path, "untracked.txt")
    _write_text(tmp_path, "ignored.txt")
    _run_git(tmp_path, "add", ".gitignore", "tracked.txt")

    default_collection = quality_reports.collect_git_files(tmp_path)
    tracked_collection = quality_reports.collect_git_files(tmp_path, tracked_only=True)
    ignored_collection = quality_reports.collect_git_files(tmp_path, include_ignored=True)

    assert default_collection.files == (".gitignore", "tracked.txt", "untracked.txt")
    assert tracked_collection.files == (".gitignore", "tracked.txt")
    assert ignored_collection.files == (
        ".gitignore",
        "ignored.txt",
        "tracked.txt",
        "untracked.txt",
    )


def test_line_ending_report_inventories_endings_and_normalization_risk(
    tmp_path: Path,
) -> None:
    """Line-ending reports detect CRLF/LF/mixed files and LF-normalization risk."""
    _init_repo(tmp_path)
    _write_bytes(tmp_path, ".gitattributes", b"*.md text eol=lf\n")
    _write_bytes(tmp_path, "tracked-crlf.md", b"one\r\ntwo\r\n")
    _write_bytes(tmp_path, "untracked-lf.md", b"one\ntwo\n")
    _write_bytes(tmp_path, "mixed.md", b"one\r\ntwo\n")
    _run_git(tmp_path, "add", ".gitattributes", "tracked-crlf.md")

    report = quality_reports.build_line_ending_report(tmp_path)

    assert report.counts["crlf"] == 1
    assert report.counts["lf"] == 2
    assert report.counts["mixed"] == 1
    assert report.normalization_risk_paths == ("mixed.md", "tracked-crlf.md")


def test_path_reference_report_catches_case_mismatch(tmp_path: Path) -> None:
    """Path-reference reports catch mismatches hidden on case-insensitive file systems."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "CSV/data.csv")
    _write_text(tmp_path, "README.md", "See [data](Csv/data.csv).\n")

    report = quality_reports.build_path_reference_report(tmp_path)

    assert len(report.findings) == 1
    finding = report.findings[0]
    assert finding.rule_id == quality_reports.PATH_REFERENCE_CASE_MISMATCH_RULE
    assert finding.literal == "Csv/data.csv"
    assert finding.matched_path == "CSV/data.csv"


def test_path_reference_report_scans_only_documented_surfaces(tmp_path: Path) -> None:
    """Files outside the documented surface set are not scanned."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "CSV/data.csv")
    _write_text(tmp_path, "notes.txt", "See Csv/data.csv.\n")

    report = quality_reports.build_path_reference_report(tmp_path)

    assert report.scanned_paths == ()
    assert report.findings == ()


def test_path_reference_suppression_can_match_rule_path_glob_and_literal_pattern(
    tmp_path: Path,
) -> None:
    """Suppressions can target a rule, source-path glob, and literal regex."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "CSV/data.csv")
    _write_text(tmp_path, "README.md", "See [data](Csv/data.csv).\n")
    _write_text(
        tmp_path,
        ".template-sync/first-adoption/quality-suppressions.json",
        json.dumps(
            {
                "path-reference": {
                    "suppressions": [
                        {
                            "ruleId": "path-reference.case-mismatch",
                            "pathGlob": "*.md",
                            "literalPattern": "^Csv/",
                            "reason": "Fixture intentionally exercises suppression matching.",
                        }
                    ]
                }
            },
            indent=2,
        )
        + "\n",
    )

    report = quality_reports.build_path_reference_report(tmp_path)

    assert report.findings == ()
    assert report.suppressed_count == 1


def test_quality_suppressions_tolerate_leading_utf8_bom(tmp_path: Path) -> None:
    """Downstream-authored suppression files load even with a leading UTF-8 BOM."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "CSV/data.csv")
    _write_text(tmp_path, "README.md", "See [data](Csv/data.csv).\n")
    suppression_document = json.dumps(
        {
            "path-reference": {
                "suppressions": [
                    {
                        "ruleId": "path-reference.case-mismatch",
                        "pathGlob": "*.md",
                        "literalPattern": "^Csv/",
                        "reason": "Fixture exercises BOM-tolerant suppression loading.",
                    }
                ]
            }
        },
        indent=2,
    )
    _write_bytes(
        tmp_path,
        ".template-sync/first-adoption/quality-suppressions.json",
        b"\xef\xbb\xbf" + (suppression_document + "\n").encode("utf-8"),
    )

    report = quality_reports.build_path_reference_report(tmp_path)

    assert report.findings == ()
    assert report.suppressed_count == 1


def test_path_reference_suppression_rejects_unknown_rule_identifier(tmp_path: Path) -> None:
    """Malformed suppression files fail with actionable errors."""
    _init_repo(tmp_path)
    _write_text(
        tmp_path,
        ".template-sync/first-adoption/quality-suppressions.json",
        json.dumps(
            {
                "path-reference": {
                    "suppressions": [
                        {
                            "ruleId": "path-reference.not-real",
                            "reason": "Unknown rule must fail.",
                        }
                    ]
                }
            }
        )
        + "\n",
    )

    with pytest.raises(quality_reports.FirstAdoptionQualityError) as excinfo:
        quality_reports.load_quality_suppressions(
            tmp_path,
            ".template-sync/first-adoption/quality-suppressions.json",
        )

    assert "unknown ruleId" in str(excinfo.value)


def test_quality_suppressions_require_path_reference_section(tmp_path: Path) -> None:
    """Suppression files follow the same required sections as their schema."""
    _init_repo(tmp_path)
    _write_text(
        tmp_path,
        ".template-sync/first-adoption/quality-suppressions.json",
        "{}\n",
    )

    with pytest.raises(quality_reports.FirstAdoptionQualityError) as excinfo:
        quality_reports.load_quality_suppressions(
            tmp_path,
            ".template-sync/first-adoption/quality-suppressions.json",
        )

    assert "path-reference" in str(excinfo.value)


def test_path_reference_suppression_requires_a_selector(tmp_path: Path) -> None:
    """Suppressions cannot accidentally match every path-reference finding."""
    _init_repo(tmp_path)
    _write_text(
        tmp_path,
        ".template-sync/first-adoption/quality-suppressions.json",
        json.dumps(
            {
                "path-reference": {
                    "suppressions": [
                        {
                            "reason": "Reason-only suppressions are too broad.",
                        }
                    ]
                }
            }
        )
        + "\n",
    )

    with pytest.raises(quality_reports.FirstAdoptionQualityError) as excinfo:
        quality_reports.load_quality_suppressions(
            tmp_path,
            ".template-sync/first-adoption/quality-suppressions.json",
        )

    assert "at least one selector" in str(excinfo.value)


def test_markdownlint_report_reports_missing_npm_without_silent_skip(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Missing optional Markdown tooling produces an unavailable state."""
    _write_text(tmp_path, "package.json", '{"scripts":{"lint:md":"markdownlint-cli2 ."}}\n')

    def missing_executable(_name: str) -> None:
        return None

    monkeypatch.setattr(quality_reports.shutil, "which", missing_executable)

    report = quality_reports.build_markdownlint_report(tmp_path)

    assert report.available is False
    assert "npm was not found" in report.message


def test_markdownlint_output_parser_records_rule_and_file_counts() -> None:
    """Markdownlint text output is parsed into reportable findings."""
    output = "\n".join(
        [
            "README.md:10:81 MD013/line-length Line length [Expected: 80]",
            "docs/guide.md:3 MD032/blanks-around-lists Lists should be surrounded by blank lines",
            "",
        ]
    )

    findings = quality_reports.parse_markdownlint_findings(output)

    assert [finding.rule for finding in findings] == [
        "MD013/line-length",
        "MD032/blanks-around-lists",
    ]
    assert [finding.path for finding in findings] == ["README.md", "docs/guide.md"]


def test_markdownlint_fixer_reports_changed_files(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Markdown fixer mode reports files changed after the fixer command."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "package.json", '{"scripts":{"lint:md":"markdownlint-cli2 ."}}\n')
    _write_text(tmp_path, "README.md", "# Title\n")
    _run_git(tmp_path, "add", "package.json", "README.md")
    _run_git(tmp_path, "commit", "-m", "initial")
    monkeypatch.setattr(quality_reports, "npm_executable", lambda: "npm")

    def fake_runner(
        command: Sequence[str],
        repo_root: Path,
        *,
        env: Mapping[str, str] | None = None,
    ) -> subprocess.CompletedProcess[str]:
        del env
        (repo_root / "README.md").write_text("# Title\n\nBody\n", encoding="utf-8")
        return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

    report = quality_reports.build_markdownlint_report(
        tmp_path,
        fix=True,
        runner=fake_runner,
    )

    assert report.available is True
    assert report.changed_files == ("README.md",)


def test_decode_git_nul_paths_preserves_crlf_inside_record() -> None:
    """Raw Git path decoding preserves valid CR/LF bytes inside path records."""
    output = b"src/line\r\nbreak.ps1\0"

    paths = quality_reports.decode_git_nul_paths(output)

    assert paths == ("src/line\r\nbreak.ps1",)


def test_decode_git_nul_paths_rejects_undecodable_record() -> None:
    """Undecodable Git path bytes fail with a manual-review diagnostic."""
    with pytest.raises(quality_reports.FirstAdoptionQualityError, match="not valid UTF-8"):
        quality_reports.decode_git_nul_paths(b"src/bad-\xff.ps1\0")


def test_decode_git_nul_paths_rejects_empty_record() -> None:
    """Empty NUL-delimited records fail closed instead of being silently skipped."""
    with pytest.raises(quality_reports.FirstAdoptionQualityError, match="empty NUL-delimited"):
        quality_reports.decode_git_nul_paths(b"src/a.ps1\0\0src/b.ps1\0")


def test_collect_powershell_candidate_paths_uses_analyzer_suffixes_case_insensitively(
    tmp_path: Path,
) -> None:
    """PowerShell report discovery uses the same analyzer input suffix family."""
    _init_repo(tmp_path)
    _write_text(tmp_path, ".github/linting/PSScriptAnalyzerSettings.psd1")
    _write_text(tmp_path, "src/Script.PS1")
    _write_text(tmp_path, "src/Module.PSM1")
    _write_text(tmp_path, "src/Data.PSD1")
    _write_text(tmp_path, "src/readme.txt")

    candidates = quality_reports.collect_powershell_candidate_paths(tmp_path)

    assert set(candidates) == {
        ".github/linting/PSScriptAnalyzerSettings.psd1",
        "src/Script.PS1",
        "src/Module.PSM1",
        "src/Data.PSD1",
    }


@pytest.mark.upstream_template_only
def test_report_suffix_policy_matches_shared_helper() -> None:
    """The Python report and PowerShell helper expose the same analyzer suffix set."""
    executable = quality_reports.powershell_executable()
    if not isinstance(executable, str):
        pytest.skip("PowerShell is not available")

    result = subprocess.run(
        [
            executable,
            "-NoProfile",
            "-Command",
            (
                "$ErrorActionPreference = 'Stop'; "
                ". './src/tools/Resolve-PSScriptAnalyzerCandidate.ps1'; "
                "Get-PSScriptAnalyzerCandidateAllowedExtension | ConvertTo-Json -Compress"
            ),
        ],
        cwd=REPO_ROOT,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert result.returncode == 0, result.stderr

    helper_suffixes = json.loads(result.stdout)
    assert set(helper_suffixes) == quality_reports.PSSCRIPTANALYZER_CANDIDATE_SUFFIXES


def test_candidate_summary_lines_escape_human_facing_crlf() -> None:
    """Candidate summaries preserve structure but render paths on one line."""
    lines = quality_reports.candidate_summary_lines(
        {
            "Selected": [],
            "PolicyExcluded": [
                {
                    "CandidateFullName": "/repo/.github/linting/PSScriptAnalyzerSettings.psd1",
                    "RepositoryRelativePath": ".github/linting/PSScriptAnalyzerSettings.psd1",
                    "OutcomeCategory": "policy-excluded",
                    "ReasonCode": "AnalyzerSettingsFile",
                },
                {
                    "CandidateFullName": "/repo/node_modules/.bin/tool.ps1",
                    "RepositoryRelativePath": "node_modules/.bin/tool.ps1",
                    "OutcomeCategory": "policy-excluded",
                    "ReasonCode": "NodeModulesSegment",
                },
            ],
            "Unsafe": [
                {
                    "CandidateFullName": "/repo/src/line\r\nbreak.ps1",
                    "RepositoryRelativePath": "src/line\r\nbreak.ps1",
                    "OutcomeCategory": "unsafe",
                    "ReasonCode": "MissingTarget",
                }
            ],
            "SummaryCounts": {
                "Selected": 0,
                "PolicyExcluded": 2,
                "Unsafe": 1,
            },
        }
    )

    assert "node_modules/.bin/tool.ps1" in "\n".join(lines)
    assert (
        "Policy-excluded reason counts: AnalyzerSettingsFile (1); NodeModulesSegment (1)." in lines
    )
    assert "src/line\\r\\nbreak.ps1" in "\n".join(lines)
    assert all("\r" not in line and "\n" not in line for line in lines)


def test_powershell_report_parses_injected_runner_output(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """The PowerShell report sends stdin candidates and parses composite JSON."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "src/café.ps1", "Write-Output 'sample'\n")
    _write_powershell_report_prerequisites(tmp_path)
    monkeypatch.setattr(quality_reports, "powershell_executable", lambda: "pwsh")

    captured_input = ""

    def fake_runner(
        command: Sequence[str],
        repo_root: Path,
        *,
        env: Mapping[str, str] | None = None,
        input: str | None = None,
    ) -> subprocess.CompletedProcess[str]:
        nonlocal captured_input
        assert env is None
        captured_input = input or ""
        payload = json.dumps(
            {
                "Gate": _gate_payload(),
                "Candidates": {
                    "Selected": [
                        {
                            "CandidateFullName": str(repo_root / "src" / "café.ps1"),
                            "RepositoryRelativePath": "src/café.ps1",
                            "EscapedAnalyzerPath": str(repo_root / "src" / "café.ps1"),
                            "OutcomeCategory": "selected",
                            "ReasonCode": "Selected",
                        }
                    ],
                    "PolicyExcluded": [],
                    "Unsafe": [],
                    "SummaryCounts": {
                        "Selected": 1,
                        "PolicyExcluded": 0,
                        "Unsafe": 0,
                    },
                },
            }
        )
        return subprocess.CompletedProcess(command, 0, stdout=payload, stderr="")

    report = quality_reports.build_powershell_report(tmp_path, runner=fake_runner)

    assert report.available is True
    assert report.summary_lines == (
        "PSScriptAnalyzer gate mode: first-adoption.",
        "Result: pass.",
    )
    assert report.issue_ready_markdown == (
        "## PSScriptAnalyzer First-Adoption Debt Cleanup",
        "",
    )
    assert report.candidate_summary_lines == (
        "PSScriptAnalyzer candidates: 1 selected; 0 policy-excluded; 0 unsafe.",
    )
    assert report.analyzer_debt_records == ()
    assert report.opt_in_guidance_lines == ()
    assert "café" in captured_input
    assert "\\u00e9" not in captured_input
    assert {"RepositoryRelativePath": "src/café.ps1"} in json.loads(captured_input)


def test_powershell_report_surfaces_settings_policy_exclusion(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """The report sends the settings file to the helper and renders its exclusion reason."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "src/module.psm1", "function Get-Sample {}\n")
    _write_powershell_report_prerequisites(tmp_path)
    monkeypatch.setattr(quality_reports, "powershell_executable", lambda: "pwsh")

    captured_input = ""

    def fake_runner(
        command: Sequence[str],
        repo_root: Path,
        *,
        env: Mapping[str, str] | None = None,
        input: str | None = None,
    ) -> subprocess.CompletedProcess[str]:
        nonlocal captured_input
        del env
        captured_input = input or ""
        payload = json.dumps(
            {
                "Gate": _gate_payload(),
                "Candidates": _candidate_summary(
                    repo_root,
                    selected=[
                        {
                            "CandidateFullName": str(repo_root / "src" / "module.psm1"),
                            "RepositoryRelativePath": "src/module.psm1",
                            "EscapedAnalyzerPath": str(repo_root / "src" / "module.psm1"),
                            "OutcomeCategory": "selected",
                            "ReasonCode": "Selected",
                        }
                    ],
                    policy_excluded=[
                        {
                            "CandidateFullName": str(
                                repo_root / ".github" / "linting" / "PSScriptAnalyzerSettings.psd1"
                            ),
                            "RepositoryRelativePath": (
                                ".github/linting/PSScriptAnalyzerSettings.psd1"
                            ),
                            "OutcomeCategory": "policy-excluded",
                            "ReasonCode": "AnalyzerSettingsFile",
                        }
                    ],
                ),
            }
        )
        return subprocess.CompletedProcess(command, 0, stdout=payload, stderr="")

    report = quality_reports.build_powershell_report(tmp_path, runner=fake_runner)

    assert {"RepositoryRelativePath": "src/module.psm1"} in json.loads(captured_input)
    assert {
        "RepositoryRelativePath": ".github/linting/PSScriptAnalyzerSettings.psd1"
    } in json.loads(captured_input)
    assert report.candidate_summary_lines == (
        "PSScriptAnalyzer candidates: 1 selected; 1 policy-excluded; 0 unsafe.",
        "Policy-excluded reason counts: AnalyzerSettingsFile (1).",
        (
            "Policy-excluded candidates: 1 (examples: "
            ".github/linting/PSScriptAnalyzerSettings.psd1)."
        ),
    )


def test_powershell_report_emits_debt_records_and_opt_in_guidance(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Warning-only first adoption findings produce debt records and manual guidance."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "src/sample.ps1", "Write-Output 'sample'\n")
    _write_powershell_report_prerequisites(tmp_path)
    github_workflow, azure_pipeline = _write_retained_powershell_ci_surfaces(tmp_path)
    monkeypatch.setattr(quality_reports, "powershell_executable", lambda: "pwsh")

    finding = {
        "Severity": "Warning",
        "RuleName": "PSAvoidUsingWriteHost",
        "Message": "Avoid Write-Host.",
        "ScriptPath": "src/sample.ps1",
        "Line": 7,
        "Column": 5,
        "TrackedDebt": True,
    }

    def fake_runner(
        command: Sequence[str],
        repo_root: Path,
        *,
        env: Mapping[str, str] | None = None,
        input: str | None = None,
    ) -> subprocess.CompletedProcess[str]:
        del env, input
        payload = json.dumps(
            {
                "Gate": _gate_payload(
                    findings=[finding],
                    recommended_mode="first-adoption",
                ),
                "Candidates": _candidate_summary(repo_root),
            }
        )
        return subprocess.CompletedProcess(command, 0, stdout=payload, stderr="")

    report = quality_reports.build_powershell_report(tmp_path, runner=fake_runner)
    stdout = io.StringIO()
    quality_reports.print_powershell_report(report, stdout=stdout)

    output = stdout.getvalue()
    assert report.analyzer_debt_records == (
        quality_reports.PowerShellAnalyzerDebtRecord(
            rule_name="PSAvoidUsingWriteHost",
            severity="Warning",
            normalized_path="src/sample.ps1",
            line=7,
            column=5,
            message="Avoid Write-Host.",
        ),
    )
    assert "Analyzer debt records for manual tracking" in output
    assert "Owner: <owner>; expected removal date: <YYYY-MM-DD>" in output
    assert "src/sample.ps1:7:5 [Warning] PSAvoidUsingWriteHost" in output
    assert "Non-mutating first-adoption opt-in guidance" in output
    assert ".github/workflows/powershell-ci.yml `jobs.powershell-lint.env" in output
    assert ".azuredevops/pipelines/powershell-ci.yml `parameters[gateMode].default`" in output
    assert "manually set .github/workflows/powershell-ci.yml" in output
    assert "manually set .azuredevops/pipelines/powershell-ci.yml" in output
    assert "runtime overrides" in output
    assert "queued-run parameters" in output
    assert "Resolve-PSScriptAnalyzerGate -Mode $env:PSSCRIPTANALYZER_GATE_MODE" in output
    assert (tmp_path / ".github/workflows/powershell-ci.yml").read_text(
        encoding="utf-8"
    ) == github_workflow
    assert (tmp_path / ".azuredevops/pipelines/powershell-ci.yml").read_text(
        encoding="utf-8"
    ) == azure_pipeline


def test_powershell_report_keeps_debt_records_separate_from_blocking_errors(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Error findings remain blocking while tracked Warning debt is still recorded."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "src/sample.ps1", "Write-Output 'sample'\n")
    _write_powershell_report_prerequisites(tmp_path)
    _write_retained_powershell_ci_surfaces(tmp_path)
    monkeypatch.setattr(quality_reports, "powershell_executable", lambda: "pwsh")

    findings = [
        {
            "Severity": "Error",
            "RuleName": "PSParserError",
            "Message": "A parser error remains blocking.",
            "ScriptPath": "src/sample.ps1",
            "Line": 1,
            "Column": 1,
            "TrackedDebt": False,
        },
        {
            "Severity": "Warning",
            "RuleName": "PSAvoidUsingWriteHost",
            "Message": "Avoid Write-Host.",
            "ScriptPath": "src/sample.ps1",
            "Line": 8,
            "Column": 3,
            "TrackedDebt": True,
        },
    ]

    def fake_runner(
        command: Sequence[str],
        repo_root: Path,
        *,
        env: Mapping[str, str] | None = None,
        input: str | None = None,
    ) -> subprocess.CompletedProcess[str]:
        del env, input
        payload = json.dumps(
            {
                "Gate": _gate_payload(
                    findings=findings,
                    should_fail=True,
                    recommended_mode="strict",
                ),
                "Candidates": _candidate_summary(repo_root),
            }
        )
        return subprocess.CompletedProcess(command, 0, stdout=payload, stderr="")

    report = quality_reports.build_powershell_report(tmp_path, runner=fake_runner)

    assert [record.rule_name for record in report.analyzer_debt_records] == [
        "PSAvoidUsingWriteHost"
    ]
    assert "PSParserError" not in "\n".join(
        quality_reports.format_analyzer_debt_record(record)
        for record in report.analyzer_debt_records
    )
    assert any("blocking findings are present" in line for line in report.opt_in_guidance_lines)


def test_powershell_report_unavailable_gate_preserves_candidate_summary(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Unavailable analyzer state emits manual-review output without debt records."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "src/sample.ps1", "Write-Output 'sample'\n")
    _write_powershell_report_prerequisites(tmp_path)
    monkeypatch.setattr(quality_reports, "powershell_executable", lambda: "pwsh")

    def fake_runner(
        command: Sequence[str],
        repo_root: Path,
        *,
        env: Mapping[str, str] | None = None,
        input: str | None = None,
    ) -> subprocess.CompletedProcess[str]:
        del env, input
        payload = json.dumps(
            {
                "Gate": {
                    "Status": "unavailable",
                    "Message": "PSScriptAnalyzer report unavailable.",
                    "SummaryLines": [],
                    "IssueReadyMarkdown": [],
                },
                "Candidates": _candidate_summary(repo_root),
            }
        )
        return subprocess.CompletedProcess(command, 0, stdout=payload, stderr="")

    report = quality_reports.build_powershell_report(tmp_path, runner=fake_runner)

    assert report.available is False
    assert report.candidate_summary_lines == (
        "PSScriptAnalyzer candidates: 1 selected; 0 policy-excluded; 0 unsafe.",
    )
    assert report.analyzer_debt_records == ()
    assert report.opt_in_guidance_lines == ()


def test_powershell_report_malformed_gate_status_is_manual_review(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Missing Gate.Status is non-fatal and emits no debt or opt-in instructions."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "src/sample.ps1", "Write-Output 'sample'\n")
    _write_powershell_report_prerequisites(tmp_path)
    monkeypatch.setattr(quality_reports, "powershell_executable", lambda: "pwsh")

    def fake_runner(
        command: Sequence[str],
        repo_root: Path,
        *,
        env: Mapping[str, str] | None = None,
        input: str | None = None,
    ) -> subprocess.CompletedProcess[str]:
        del env, input
        payload = json.dumps(
            {
                "Gate": {
                    "SummaryLines": ["PSScriptAnalyzer gate mode: first-adoption."],
                    "IssueReadyMarkdown": [],
                    "Findings": [],
                },
                "Candidates": _candidate_summary(repo_root),
            }
        )
        return subprocess.CompletedProcess(command, 0, stdout=payload, stderr="")

    report = quality_reports.build_powershell_report(tmp_path, runner=fake_runner)

    assert report.available is False
    assert "Gate.Status" in report.message
    assert "manual review" in "\n".join(report.summary_lines).lower()
    assert report.analyzer_debt_records == ()
    assert report.opt_in_guidance_lines == ()


def test_powershell_report_malformed_gate_findings_are_manual_review(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Malformed structured findings avoid debt output instead of scraping text."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "src/sample.ps1", "Write-Output 'sample'\n")
    _write_powershell_report_prerequisites(tmp_path)
    monkeypatch.setattr(quality_reports, "powershell_executable", lambda: "pwsh")

    def fake_runner(
        command: Sequence[str],
        repo_root: Path,
        *,
        env: Mapping[str, str] | None = None,
        input: str | None = None,
    ) -> subprocess.CompletedProcess[str]:
        del env, input
        payload = json.dumps(
            {
                "Gate": _gate_payload(
                    findings=[
                        {
                            "Severity": "Warning",
                            "RuleName": "PSRule",
                            "Message": "Malformed tracked flag.",
                            "ScriptPath": "src/sample.ps1",
                            "Line": 1,
                            "Column": 1,
                            "TrackedDebt": "true",
                        }
                    ],
                    recommended_mode="first-adoption",
                ),
                "Candidates": _candidate_summary(repo_root),
            }
        )
        return subprocess.CompletedProcess(command, 0, stdout=payload, stderr="")

    report = quality_reports.build_powershell_report(tmp_path, runner=fake_runner)

    assert "required Gate fields" in report.message
    assert report.analyzer_debt_records == ()
    assert report.opt_in_guidance_lines == ()


def test_powershell_report_unsafe_candidates_take_precedence_over_unavailable_gate(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Unsafe candidates keep their reserved exit precedence with valid JSON."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "src/sample.ps1", "Write-Output 'sample'\n")
    _write_powershell_report_prerequisites(tmp_path)
    monkeypatch.setattr(quality_reports, "powershell_executable", lambda: "pwsh")

    unsafe_candidate = {
        "CandidateFullName": str(tmp_path / "src" / "sample.ps1"),
        "RepositoryRelativePath": "src/sample.ps1",
        "OutcomeCategory": "unsafe",
        "ReasonCode": "TargetOutsideRepository",
    }

    def fake_runner(
        command: Sequence[str],
        repo_root: Path,
        *,
        env: Mapping[str, str] | None = None,
        input: str | None = None,
    ) -> subprocess.CompletedProcess[str]:
        del env, input
        payload = json.dumps(
            {
                "Gate": {
                    "Status": "unavailable",
                    "Message": "Analyzer was not run because unsafe candidates were found.",
                },
                "Candidates": _candidate_summary(
                    repo_root,
                    selected=[],
                    unsafe=[unsafe_candidate],
                ),
            }
        )
        return subprocess.CompletedProcess(command, 0, stdout=payload, stderr="")

    report = quality_reports.build_powershell_report(tmp_path, runner=fake_runner)
    stdout = io.StringIO()
    monkeypatch.setattr(quality_reports, "build_powershell_report", lambda *args, **kwargs: report)
    args = quality_reports.parse_args(["--repo-root", str(tmp_path), "powershell"])

    result = quality_reports.run_report(args, stdout=stdout)

    assert result == quality_reports.UNSAFE_CANDIDATE_EXIT_CODE
    assert report.unsafe_candidate_count == 1
    assert "Unsafe candidate: src/sample.ps1" in stdout.getvalue()


def test_powershell_report_non_json_stdout_is_fatal(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Non-JSON runner stdout keeps the existing fatal report error."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "src/sample.ps1", "Write-Output 'sample'\n")
    _write_powershell_report_prerequisites(tmp_path)
    monkeypatch.setattr(quality_reports, "powershell_executable", lambda: "pwsh")

    def fake_runner(
        command: Sequence[str],
        repo_root: Path,
        *,
        env: Mapping[str, str] | None = None,
        input: str | None = None,
    ) -> subprocess.CompletedProcess[str]:
        del env, input, repo_root
        return subprocess.CompletedProcess(command, 0, stdout="not json", stderr="")

    with pytest.raises(quality_reports.FirstAdoptionQualityError, match="non-JSON"):
        quality_reports.build_powershell_report(tmp_path, runner=fake_runner)


def test_gate_mode_static_value_reports_yaml_boolean_scalars_as_manual_review() -> None:
    """PyYAML 1.1 truthy scalars are not treated as gate-mode strings."""
    for scalar in ("on", "off", "yes", "no", "true", "false", "True", "False"):
        loaded = quality_reports.yaml.safe_load(f"value: {scalar}\n")["value"]

        normalized = quality_reports.normalize_gate_mode_static_value(loaded)

        assert isinstance(loaded, bool)
        assert normalized.manual_review is True
        assert normalized.recognized_mode is None
        assert "YAML boolean" in normalized.display_value


def test_github_actions_gate_mode_detection_covers_env_specificity() -> None:
    """Workflow, job, and step env values are detected with specificity labels."""
    document = quality_reports.yaml.safe_load(
        "env:\n"
        "  PSSCRIPTANALYZER_GATE_MODE: strict\n"
        "jobs:\n"
        "  powershell-lint:\n"
        "    env:\n"
        "      PSSCRIPTANALYZER_GATE_MODE: first-adoption\n"
        "    steps:\n"
        "      - name: Run PSScriptAnalyzer\n"
        "        env:\n"
        "          PSSCRIPTANALYZER_GATE_MODE: on\n"
        "        run: |\n"
        "          Resolve-PSScriptAnalyzerGate `\n"
        "            -Mode $env:PSSCRIPTANALYZER_GATE_MODE\n"
    )

    settings, notes, retained = quality_reports.github_actions_gate_mode_settings(
        document,
        path=".github/workflows/powershell-ci.yml",
    )

    assert notes == ()
    assert retained is True
    assert [setting.specificity for setting in settings] == [
        "workflow env",
        "job env",
        "step env",
    ]
    assert settings[2].value.manual_review is True


def test_github_actions_gate_mode_settings_preserves_notes_when_jobs_missing() -> None:
    """A workflow-level env note survives the missing/non-mapping ``jobs`` early return."""
    document = quality_reports.yaml.safe_load(
        "env: malformed-workflow-env\n"
        "jobs: |\n"
        "  Resolve-PSScriptAnalyzerGate `\n"
        "    -Mode $env:PSSCRIPTANALYZER_GATE_MODE\n"
    )

    settings, notes, retained = quality_reports.github_actions_gate_mode_settings(
        document,
        path=".github/workflows/powershell-ci.yml",
    )

    assert retained is True
    assert settings == ()
    assert notes == (
        (
            "GitHub Actions: .github/workflows/powershell-ci.yml `env` env is not a mapping; "
            "manual review is required."
        ),
        "GitHub Actions: jobs is missing or not a mapping; manual review is required.",
    )


def test_azure_pipelines_gate_mode_detection_resolves_mapping_variables() -> None:
    """Azure mapping variables resolve the parameter expression to its default."""
    document = quality_reports.yaml.safe_load(
        "parameters:\n"
        "  - name: gateMode\n"
        "    type: string\n"
        '    default: "strict"\n'
        "jobs:\n"
        "  - job: lint\n"
        "    variables:\n"
        '      PSSCRIPTANALYZER_GATE_MODE: "${{ parameters.gateMode }}"\n'
        "    steps:\n"
        "      - task: PowerShell@2\n"
        "        inputs:\n"
        "          script: |\n"
        "            Resolve-PSScriptAnalyzerGate `\n"
        "              -Mode $env:PSSCRIPTANALYZER_GATE_MODE\n"
    )

    parameter, variables, notes, retained = quality_reports.azure_pipelines_gate_mode_settings(
        document,
        path=".azuredevops/pipelines/powershell-ci.yml",
    )

    assert notes == ()
    assert retained is True
    assert parameter is not None
    assert parameter.value.recognized_mode == "strict"
    assert variables[0].value.display_value == "${{ parameters.gateMode }}"
    assert variables[0].effective_value is not None
    assert variables[0].effective_value.recognized_mode == "strict"


def test_azure_pipelines_gate_mode_detection_supports_sequence_variables() -> None:
    """Azure sequence variables are detected for the gate-mode variable."""
    document = quality_reports.yaml.safe_load(
        "parameters:\n"
        "  gateMode:\n"
        '    default: "first-adoption"\n'
        "jobs:\n"
        "  lint:\n"
        "    variables:\n"
        "      - name: PSSCRIPTANALYZER_GATE_MODE\n"
        '        value: "${{ parameters.gateMode }}"\n'
        "    steps:\n"
        "      - powershell: |\n"
        "          Resolve-PSScriptAnalyzerGate `\n"
        "            -Mode $env:PSSCRIPTANALYZER_GATE_MODE\n"
    )

    parameter, variables, notes, retained = quality_reports.azure_pipelines_gate_mode_settings(
        document,
        path=".azuredevops/pipelines/powershell-ci.yml",
    )

    assert notes == ()
    assert retained is True
    assert parameter is not None
    assert parameter.value.recognized_mode == "first-adoption"
    assert variables[0].location == "jobs[lint].variables[PSSCRIPTANALYZER_GATE_MODE].value"
    assert variables[0].effective_value is not None
    assert variables[0].effective_value.recognized_mode == "first-adoption"


def test_azure_parameter_default_flags_missing_default_in_sequence_form() -> None:
    """A sequence-form gateMode parameter without a default requires manual review."""
    document = quality_reports.yaml.safe_load("parameters:\n  - name: gateMode\n    type: string\n")

    setting, notes = quality_reports.azure_parameter_default_setting(
        document,
        path=".azuredevops/pipelines/powershell-ci.yml",
    )

    assert setting is None
    assert notes == (
        (
            "Azure Pipelines: .azuredevops/pipelines/powershell-ci.yml "
            "`parameters[gateMode].default` has no static default; the value must be supplied "
            "when the pipeline runs, so manual review is required."
        ),
    )


def test_azure_parameter_default_flags_missing_default_in_mapping_form() -> None:
    """A mapping-form gateMode parameter without a default requires manual review."""
    document = quality_reports.yaml.safe_load("parameters:\n  gateMode:\n    type: string\n")

    setting, notes = quality_reports.azure_parameter_default_setting(
        document,
        path=".azuredevops/pipelines/powershell-ci.yml",
    )

    assert setting is None
    assert notes == (
        (
            "Azure Pipelines: .azuredevops/pipelines/powershell-ci.yml "
            "`parameters.gateMode.default` has no static default; the value must be supplied "
            "when the pipeline runs, so manual review is required."
        ),
    )


@pytest.mark.upstream_template_only
def test_real_powershell_ci_surfaces_retain_gate_mode_and_force_visible_discovery() -> None:
    """Checked-in PowerShell CI surfaces keep gate-mode and candidate discovery shape.

    This test reads the upstream template's PowerShell CI surfaces for both
    hosts (``.github/workflows/powershell-ci.yml`` and
    ``.azuredevops/pipelines/powershell-ci.yml``) and asserts their
    template-authored content. Downstream materializations that exclude the
    ``powershell``, ``github-actions``, or ``azure-pipelines`` module do not
    retain those files, so the downstream pytest gate
    (``-m "not upstream_template_only"``) must not select this test.
    """
    github_document, github_load_notes = quality_reports.load_ci_yaml_mapping(
        REPO_ROOT,
        quality_reports.GITHUB_ACTIONS_POWERSHELL_CI,
        platform_name="GitHub Actions",
    )
    azure_document, azure_load_notes = quality_reports.load_ci_yaml_mapping(
        REPO_ROOT,
        quality_reports.AZURE_PIPELINES_POWERSHELL_CI,
        platform_name="Azure Pipelines",
    )
    assert github_load_notes == ()
    assert azure_load_notes == ()
    assert github_document is not None
    assert azure_document is not None

    github_settings, github_gate_notes, github_retained = (
        quality_reports.github_actions_gate_mode_settings(
            github_document,
            path=quality_reports.GITHUB_ACTIONS_POWERSHELL_CI,
        )
    )
    azure_parameter, azure_variables, azure_gate_notes, azure_retained = (
        quality_reports.azure_pipelines_gate_mode_settings(
            azure_document,
            path=quality_reports.AZURE_PIPELINES_POWERSHELL_CI,
        )
    )

    assert github_gate_notes == ()
    assert azure_gate_notes == ()
    assert github_retained is True
    assert azure_retained is True
    assert any(setting.value.recognized_mode == "strict" for setting in github_settings)
    assert azure_parameter is not None
    assert azure_parameter.value.recognized_mode == "strict"
    assert azure_variables[0].effective_value is not None
    assert azure_variables[0].effective_value.recognized_mode == "strict"

    github_text = (REPO_ROOT / quality_reports.GITHUB_ACTIONS_POWERSHELL_CI).read_text(
        encoding="utf-8"
    )
    azure_text = (REPO_ROOT / quality_reports.AZURE_PIPELINES_POWERSHELL_CI).read_text(
        encoding="utf-8"
    )
    assert (
        "Get-PSScriptAnalyzerCandidate -RepositoryRoot $env:GITHUB_WORKSPACE "
        "-DirectoryVisibility All"
    ) in " ".join(github_text.split())
    assert (
        "Get-PSScriptAnalyzerCandidate -RepositoryRoot $repositoryRoot -DirectoryVisibility All"
    ) in " ".join(azure_text.split())
    assert "Resolve-PSScriptAnalyzerGate ` -Mode $env:PSSCRIPTANALYZER_GATE_MODE" in " ".join(
        github_text.split()
    )
    assert "Resolve-PSScriptAnalyzerGate ` -Mode $env:PSSCRIPTANALYZER_GATE_MODE" in " ".join(
        azure_text.split()
    )


def test_powershell_report_unsafe_candidates_use_reserved_exit_code(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Unsafe analyzer candidates propagate the reserved report exit code."""
    report = quality_reports.PowerShellAnalyzerReport(
        available=True,
        message="Unsafe candidate path(s) were found.",
        candidate_summary_lines=(
            "Unsafe candidate: src/link.ps1; reason: TargetOutsideRepository",
        ),
        unsafe_candidate_count=1,
        summary_lines=(),
        issue_ready_markdown=(),
        analyzer_debt_records=(),
        opt_in_guidance_lines=(),
    )
    monkeypatch.setattr(quality_reports, "build_powershell_report", lambda *args, **kwargs: report)
    stdout = io.StringIO()
    args = quality_reports.parse_args(
        [
            "--repo-root",
            str(tmp_path),
            "powershell",
        ]
    )

    result = quality_reports.run_report(args, stdout=stdout)

    assert result == quality_reports.UNSAFE_CANDIDATE_EXIT_CODE
    assert "Unsafe candidate" in stdout.getvalue()


def test_host_setup_report_distinguishes_azure_service_tasks(tmp_path: Path) -> None:
    """Azure host setup reporting separates service tasks from file materialization."""
    _write_text(
        tmp_path,
        ".template-sync/marker.yml",
        (
            "template_sync:\n"
            "  source_repo: https://github.com/franklesniak/copilot-repo-template.git\n"
            "  included_modules:\n"
            "    - baseline\n"
            "    - azure-devops-platform\n"
            "  host_provider: azure-devops-services\n"
            "  azure_boards_policy: work-items\n"
            "  azure_repos_pr_template_policy: materialize\n"
            "  azure_branch_policy_reviewer_guidance: manual-follow-up\n"
            "  azure_security_intake_policy: manual-follow-up\n"
            "  azure_security_product_enablement: github-advanced-security\n"
            "  azure_dependency_update_policy: manual-follow-up\n"
        ),
    )
    stdout = io.StringIO()

    report = quality_reports.build_host_setup_report(tmp_path)
    quality_reports.print_host_setup_report(report, stdout=stdout)

    output = stdout.getvalue()
    assert report.azure_modules_retained is True
    assert "Azure DevOps Services service setup tasks" in output
    assert "Azure Boards intake policy: work-items" in output
    assert "Azure Repos pull request template policy: materialize" in output
    assert "Branch policy reviewer guidance: manual-follow-up" in output
    assert "not local file materialization failures" in output
    assert "not local file materialization failures or GitHub issue-template findings" in output


def test_load_marker_rejects_symlinked_marker(tmp_path: Path) -> None:
    """A symlinked marker (even a broken one) is rejected, not treated as absent."""
    marker_path = tmp_path / ".template-sync" / "marker.yml"
    marker_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        marker_path.symlink_to(tmp_path / "nonexistent-target.yml")
    except (OSError, NotImplementedError):
        pytest.skip("Filesystem does not support symlink creation")

    with pytest.raises(quality_reports.FirstAdoptionQualityError, match="Expected a regular file"):
        quality_reports.load_marker_template_sync(tmp_path)


def test_load_ci_yaml_mapping_flags_broken_symlink(tmp_path: Path) -> None:
    """A broken symlink at a CI YAML path is flagged for manual review, not ignored."""
    relative_path = ".github/workflows/example.yml"
    ci_path = tmp_path / relative_path
    ci_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        ci_path.symlink_to(tmp_path / "nonexistent-target.yml")
    except (OSError, NotImplementedError):
        pytest.skip("Filesystem does not support symlink creation")

    document, notes = quality_reports.load_ci_yaml_mapping(
        tmp_path,
        relative_path,
        platform_name="GitHub Actions",
    )

    assert document is None
    assert notes == (
        (
            "GitHub Actions: .github/workflows/example.yml is not a regular YAML file; "
            "manual review is required."
        ),
    )


def test_load_ci_yaml_mapping_flags_non_utf8_file(tmp_path: Path) -> None:
    """A present-but-non-UTF-8 CI YAML file degrades to a manual-review note, not a crash."""
    relative_path = ".github/workflows/example.yml"
    ci_path = tmp_path / relative_path
    ci_path.parent.mkdir(parents=True, exist_ok=True)
    ci_path.write_bytes(b"\xff\xfe not valid utf-8")

    document, notes = quality_reports.load_ci_yaml_mapping(
        tmp_path,
        relative_path,
        platform_name="GitHub Actions",
    )

    assert document is None
    assert notes == (
        (
            "GitHub Actions: .github/workflows/example.yml is not valid UTF-8; "
            "manual review is required."
        ),
    )


def test_path_reference_cli_can_fail_on_findings(tmp_path: Path) -> None:
    """The CLI offers an explicit non-zero path-reference gate when selected."""
    _init_repo(tmp_path)
    _write_text(tmp_path, "CSV/data.csv")
    _write_text(tmp_path, "README.md", "See [data](Csv/data.csv).\n")
    stdout = io.StringIO()
    args = quality_reports.parse_args(
        [
            "--repo-root",
            str(tmp_path),
            "path-references",
            "--fail-on-findings",
        ]
    )

    result = quality_reports.run_report(args, stdout=stdout)

    assert result == 1
    assert "path-reference.case-mismatch" in stdout.getvalue()


def _copy_installed_host_report(root: Path) -> Path:
    """Copy the support-owned report and its actual installed import closure."""
    for relative in (
        ".template-sync/scripts/first_adoption_quality_reports.py",
        ".template-sync/scripts/template_sync_materialization_helpers.py",
        ".github/scripts/instruction_contract_support.py",
    ):
        destination = root / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(REPO_ROOT / relative, destination)
    assert not (root / ".github/scripts/instruction_contract_core.py").exists()
    assert not (root / "pyproject.toml").exists()
    return root / ".template-sync/scripts/first_adoption_quality_reports.py"


@pytest.mark.parametrize(
    ("marker", "exit_code", "message"),
    [
        (None, 0, "not found; no host setup marker decisions"),
        ("template_sync: {included_modules: [github-actions]}\n", 0, "No Azure DevOps"),
        (
            (
                "template_sync:\n  included_modules: [azure-pipelines]\n"
                "  azure_dependency_update_policy: manual-follow-up\n"
            ),
            0,
            "manual-follow-up",
        ),
        (
            (
                "\ufeffbase: &base {included_modules: [github-actions]}\n"
                "template_sync:\n  <<: *base\n  included_modules: [azure-pipelines]\n"
                "  azure_dependency_update_policy: inherited-override\n"
            ),
            0,
            "inherited-override",
        ),
        (
            (
                "template_sync:\n  included_modules: [azure-pipelines]\n"
                "  included_modules: [github-actions]\n"
            ),
            1,
            "duplicate explicit mapping key",
        ),
        (
            (
                "template_sync: {included_modules: [azure-pipelines]}\n"
                "template_sync: {included_modules: [github-actions]}\n"
            ),
            1,
            "duplicate explicit mapping key",
        ),
        ("template_sync: [\n", 1, "Invalid YAML in .template-sync/marker.yml:"),
        ("[]\n", 1, "must contain a YAML mapping"),
        ("hello\n", 1, "must contain a YAML mapping"),
        ("null\n", 1, "must contain a YAML mapping"),
        ("template_sync: []\n", 1, "must contain template_sync mapping"),
    ],
)
def test_installed_host_report_uses_unambiguous_marker(
    tmp_path: Path, marker: str | None, exit_code: int, message: str
) -> None:
    """Direct installed CLI reports follow-ups only from unambiguous input."""
    script = _copy_installed_host_report(tmp_path / "installed")
    target = tmp_path / "target"
    target.mkdir()
    if marker is not None:
        _write_text(target, ".template-sync/marker.yml", marker)
    result = subprocess.run(
        [sys.executable, "-B", str(script), "--repo-root", str(target), "host-setup"],
        capture_output=True,
        text=True,
        check=False,
        timeout=15,
    )
    assert result.returncode == exit_code, result.stderr
    assert message in result.stdout + result.stderr
    if exit_code:
        assert result.stdout == ""
        assert result.stderr.count(".template-sync/marker.yml") == 1
        if message.startswith("must contain"):
            assert result.stderr == f"ERROR: .template-sync/marker.yml {message}.\n"
        else:
            assert result.stderr.startswith("ERROR: Invalid YAML in .template-sync/marker.yml:")
        assert "is not valid YAML" not in result.stderr


@pytest.mark.parametrize("marker", ["[]\n", "template_sync: [\n", "key: 1\nkey: 2\n"])
def test_host_report_preserves_marker_parser_cause(tmp_path: Path, marker: str) -> None:
    """Domain translation retains the parser's precise message and original cause."""
    _write_text(tmp_path, ".template-sync/marker.yml", marker)

    with pytest.raises(quality_reports.FirstAdoptionQualityError) as caught:
        quality_reports.load_marker_template_sync(tmp_path)

    cause = caught.value.__cause__
    assert isinstance(cause, quality_reports.TemplateSyncMaterializationError)
    assert str(caught.value) == str(cause)


def test_host_report_old_wrapper_fails_exact_shape_diagnostic(tmp_path: Path) -> None:
    """A fixed diagnostic oracle catches restoration of the misleading wrapper."""
    script = _copy_installed_host_report(tmp_path / "installed")
    source = script.read_text(encoding="utf-8")
    translation = "raise FirstAdoptionQualityError(str(error)) from error"
    assert source.count(translation) == 1
    script.write_text(
        source.replace(
            translation,
            'raise FirstAdoptionQualityError(f"{MARKER_PATH} is not valid YAML: {error}") from error',
        ),
        encoding="utf-8",
    )
    target = tmp_path / "target"
    _write_text(target, ".template-sync/marker.yml", "[]\n")
    result = subprocess.run(
        [sys.executable, "-B", str(script), "--repo-root", str(target), "host-setup"],
        capture_output=True,
        text=True,
        check=False,
        timeout=15,
    )
    assert result.returncode == 1, result.stderr
    assert result.stdout == ""
    assert "is not valid YAML" in result.stderr
    with pytest.raises(AssertionError):
        assert result.stderr == "ERROR: .template-sync/marker.yml must contain a YAML mapping.\n"


def test_host_report_duplicate_guard_removal_restores_false_no_tasks(
    tmp_path: Path,
) -> None:
    """An independent fixed expected-error oracle detects parser guard removal."""
    installed = tmp_path / "installed"
    script = _copy_installed_host_report(installed)
    support = installed / ".github/scripts/instruction_contract_support.py"
    source = support.read_text(encoding="utf-8")
    guard = "yaml.load(text, Loader=UniqueKeySafeLoader)"
    assert source.count(guard) == 1
    support.write_text(source.replace(guard, "yaml.safe_load(text)"), encoding="utf-8")
    target = tmp_path / "target"
    _write_text(
        target,
        ".template-sync/marker.yml",
        "template_sync:\n  included_modules: [azure-pipelines]\n"
        "  included_modules: [github-actions]\n"
        "  azure_dependency_update_policy: manual-follow-up\n",
    )
    result = subprocess.run(
        [sys.executable, "-B", str(script), "--repo-root", str(target), "host-setup"],
        capture_output=True,
        text=True,
        check=False,
        timeout=15,
    )
    assert result.returncode == 0, result.stderr
    assert "No Azure DevOps Services host setup tasks" in result.stdout
    with pytest.raises(AssertionError):
        assert result.returncode == 1


@pytest.mark.parametrize("escape", [False, True])
def test_host_report_preserves_marker_symlink_rejection(tmp_path: Path, escape: bool) -> None:
    """Shared parsing does not weaken the reader's lexical or containment checks."""
    target = tmp_path / "target"
    (target / ".template-sync").mkdir(parents=True)
    destination = (tmp_path if escape else target) / "actual.yml"
    destination.write_text("template_sync: {}\n", encoding="utf-8")
    (target / ".template-sync/marker.yml").symlink_to(destination)
    with pytest.raises(quality_reports.FirstAdoptionQualityError, match="escapes|regular file"):
        quality_reports.build_host_setup_report(target)
