<!-- markdownlint-disable MD013 -->
# ROMSorter Template Adoption Summary

## Outcome and Review Boundary

**Finalization mode:** completed local branch with working-tree changes for inspection. All adoption changes remain uncommitted and unstaged on `demo/template-adoption-2026-11-part-3`. No GitHub Issue, pull request, branch rename, push or hosted settings change was performed. The default branch remains `master`.

First adoption used **full reconciliation** of all 303 tracked upstream paths at `2f6aa17caa063eaa0b39afb3afb1d27757b4eec3`. There is no upstream range base because no prior adoption marker existed. Downstream comparison HEAD is `dcb2e4a42eabccc3ce3f4dc118781f570ccc65a4`, tree `610221c03d0d62b34160f4cd70108772bb36b687`. The obsolete remote adoption ref was not used as lineage or policy evidence. The pinned template clone was verified before materialization.

The [adoption ledger](docs/adoption/ledger.md) was generated with `--ledger-only` and reviewed before protected files were materialized. The [303 per-file decisions](docs/adoption/file-decisions.md) and [machine-readable source inventory](docs/adoption/source-inventory.json) are the complete reconciliation record. The external working notes retain the full HEAD-to-template comparison and original byte backups at `C:\Users\flesniak\AppData\Local\Temp\romsorter-adoption-record`.

**Review marker:** `template_sync.last_reviewed_template_commit` now records the full pinned SHA after reconciliation and validation. It remains an uncommitted working-tree record and is not durable until the owner commits it.

## Scope and Explicit Decisions

Adopted modules: baseline, powershell, markdown, json, yaml, github-platform, github-actions, github-templates, template-sync-support, agent-instructions, agent-codex, agent-copilot, instruction-enforcement.

Source-path decisions: **213 TAKE, 33 MERGE, 57 SKIP**. The linked table lists every unchanged adoption, manual merge and skipped path individually.

Excluded activity was reviewed and intentionally skipped: Python example-product packaging and guides; Terraform examples/hooks/workflows; schema example-product tooling; Azure platform, pipeline and collaboration assets; Claude, Cursor, Gemini and Hermes entry points; Git LFS opt-in attributes; and template-onboarding instructions. Shared inline blocks were pruned by module. Two Python import helpers remain explicitly mapped to template-sync-support because its retained tests require them. No unknown modules or unmapped paths remain. No existing downstream file was removed because it was absent from upstream; no protected file was removed or deferred.

All 13 protected paths and the template-derived governance, community, process, workflow and collaboration material use **minimal-preservation**. The owner explicitly authorized those concrete paths before creation; authorization and scope are recorded in the marker. Every retained instruction anchor passes validation and the sync catalog is byte-identical to pinned upstream. No instruction-contract waiver was introduced. No commit-isolation waiver was needed because no commits were created; future commits must respect the applicable isolation rules.

Reviewed local merges include repository identity and support routing; sole-maintainer ownership; a PVR-only policy; the PowerShell host matrix and temporary analyzer gate; Pester 5.9.1 pins; selected workflow contract fingerprints; Node.js 24 development metadata; exact CSV protection; downstream pytest configuration and upstream-only fixture classification. Every path and rationale is in the decision table and marker. Action identities, full release pins and workflow trust boundaries remain enforced.

## Structure and Preservation

Required structural changes were implemented: `.github/`, `.template-sync/`, schemas, validation tooling and tests; the existing MIT `LICENSE.txt` was renamed to conventional `LICENSE` with exact text and author bytes. Strongly recommended development/support guidance and offline syntax coverage were added. Behavioral characterization and helper consolidation are post-adoption tasks. A wholesale move into `src/` or `scripts/` is not recommended during adoption: all 25 product script paths and four public CSV paths remain unchanged.

The initial `.vscode/settings.json` change to `window.title: ROMSorter-BF-Part3` is preserved byte-for-byte, not absorbed or reverted. All four CSVs are byte-identical. All script executable content is preserved. Two scripts lose spaces from one blank line each, outside strings: `Find-MAMEAndMAME2003PlusMatchUsingCRC.ps1` and `Find-MAMEAndMAME2010MatchUsingCRC.ps1`, line 448. No product script was executed or dot-sourced. [Preservation hashes](docs/adoption/preservation.json) cover these checks.

README retains ROMSorter identity and examples, with an actual-filename case correction and narrow development/support/security/compatibility guidance. Mechanical Markdown list/emphasis/fence fixes were verified by identical rendered HTML in README, the arcade background document and Copilot prompts.

EOL normalization was a separate reviewed byte-only phase: **16 original files, 10,402 CRLF pairs removed**. [The exact file list](docs/adoption/eol-review.md) and [before/after byte counts and SHA-256 values](docs/adoption/eol-accounting.json) account for the phase. No separate commit was made because all changes must remain uncommitted. README, LICENSE and VS Code settings were already LF; an intermediate README serialization round-trip contributes no net normalization. New template files follow their adopted policy. No CSV or generated product data was normalized.

## Validation and Findings

[Exact results and commands](docs/adoption/validation-results.json) distinguish initial failures, fixes, retries and remaining limitations. Full first-adoption checks passed all ten groups, including tracked and untracked files, pre-commit, data/schema checks, placeholders, line endings, marker checks and Markdown/link/nested-fence checks. The final full run completed in 132.161 seconds. Post-stamp actionlint, workflow, marker and instruction hooks passed. One status-table Markdown spacing finding was corrected and its focused retry passed; final repository-wide link and nested-Markdown checks passed. Doctor, plan, bootstrap review, five instruction contracts and ten workflow security contracts (with release verification) passed.

Node regression tests: **122 passed, two skipped, zero failed**. Pester 5.9.1 on both PowerShell 7.6.9 and Windows PowerShell 5.1: **67 passed, two skipped, zero failed per host**. Windows skips require symlink privileges. Local product tests parse scripts only. Hosted macOS/Linux PowerShell execution and emitted GitHub check names remain post-push verification.

The exhaustive initial downstream Windows pytest run selected 2,362 tests: **2,331 passed, 20 skipped, 11 failed**. Two failures were downstream fixture assumptions about `pyproject.toml`; the fixture now supports retained `pytest.ini` and both pass on Windows. Seven failures required full upstream README/Azure assets; eight such cases are now explicitly upstream-only, including one formerly passing case. All nine non-symlink failures passed unchanged against pinned upstream. The two Windows symlink privilege failures also reproduce upstream; both pass in an isolated Ubuntu snapshot, together with both corrected fixture cases (**four passed, seven deselected**). Final collection is 2,354 selected and 2,923 deselected. The exhaustive suite was not rerun after these narrow corrections; its initial result is not represented as an all-green Windows run. No applicable validator or security assertion was disabled.

Finding dispositions:

- **Downstream-local fixes:** identity/support links, real filename case, Markdown formatting, host matrix shell placement, explicit UTF-8 discovery output for Windows PowerShell 5.1, Pester major pin, selected-module test fixtures, CSV formatter exclusions and Node support floor. These are implemented and covered by focused/final checks.
- **Upstream-template compatibility findings:** copied support tests assumed full upstream product files or packaging; narrow downstream scope/configuration adaptations are recorded per file. Runtime validators remain intact. The same Windows symlink limitation was reproduced in upstream.
- **Deferred existing debt:** PSScriptAnalyzer reports **249 warnings, zero errors, zero unknown severities**, all in existing product scripts. The approved rubric selected temporary `first-adoption` gating without disabling rules; strict gating and behavioral fixtures are tracked in TODO.
- **Deferred inherited tooling debt:** pinned npm development dependencies report **15 advisories (3 low, 11 moderate, 1 high)**; production audit reports zero. Deliberate dependency maintenance is tracked in TODO.
- **Environment/hosted validation:** native Windows symlink cases require privilege; isolated Ubuntu covers the two failing pytest cases. Hosted jobs/settings have not been exercised or enabled by this local task.

## Owner Handoff

No unresolved policy question remains. The preflight answers are recorded in [_TODO-repo-init.md](_TODO-repo-init.md), together with weighted ruleset/analyzer/Node decisions. The owner, **@franklesniak**, must immediately enable PVR and Discussions, create/verify `bug`, `enhancement`, `documentation`, `triage`, and activate the specified zero-approval default-branch ruleset. Select the exact required check names after successful hosted runs; verify issue/security/support routing and Copilot setup. No additional review-flow label requirement was found. The adopted content assumes those settings, but their hosted state remains unverified until the owner records evidence in TODO.

Support, usage help and general troubleshooting go to Discussions. Issues are reserved for reproducible bugs, enhancements, documentation defects and actionable repository work. The PVR-only 48-hour acknowledgment/seven-day fuller-response/latest-version commitments and generic maintainer GitHub conduct contact are retained.

The complete copy-ready future `master` to `main` Issue draft is fenced in TODO, including raw-content URL migration, external consumers, compatibility, rulesets, open PRs, clones, acceptance criteria and validation. **No Issue URL exists because the owner instructed that no Issue be created.**

Before any later commit or PR, inspect the full working-tree diff, preserve the pre-existing VS Code change deliberately, resolve or acknowledge the recorded debt and complete the owner's hosted activation steps. This handoff is local inspection, not a claim that hosted CI or repository settings are complete. The [final file/status inventory](docs/adoption/git-status.md) lists every added, modified and removed path, including the unstaged license rename pair.
