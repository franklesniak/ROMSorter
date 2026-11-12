<!-- markdownlint-disable MD013 -->
# ROMSorter Repository Initialization

## Metadata

- **Status:** Active
- **Owner:** @franklesniak
- **Last Updated:** 2026-11-11
- **Scope:** First-adoption decisions, immediate manual GitHub settings, and post-adoption work. This is downstream-owned state.

## Discoverable Repository State

- Repository: [franklesniak/ROMSorter](https://github.com/franklesniak/ROMSorter), public, default branch `master`.
- First adoption: no prior marker or adoption note existed. Reconcile the complete pinned template at `2f6aa17caa063eaa0b39afb3afb1d27757b4eec3`; the deleted remote adoption branch is not lineage evidence.
- Preserve all root product script and CSV paths, project README identity, and existing `.vscode/settings.json` content including `window.title: ROMSorter-BF-Part3`.
- Keep `master`; the future rename is a separate task below. Leave this adoption uncommitted with an empty index.

## Maintainer Policy Decisions

- Selected modules: baseline, powershell, markdown, json, yaml, github-platform, github-actions, github-templates, template-sync-support, agent-instructions, agent-codex, agent-copilot, instruction-enforcement.
- Adoption mode: `minimal-preservation` for protected and template-derived files. No broad protected rewrite or anchor waiver is authorized.
- Security: GitHub private vulnerability reporting only; acknowledgment within 48 hours, fuller response within seven days, latest-version support.
- Conduct: contact the maintainer using contact information on their GitHub profile. CODEOWNERS: `@franklesniak` only. GHES override: none.
- Support and troubleshooting: GitHub Discussions. Issues: reproducible bugs, enhancements, documentation defects and actionable repository work.
- Compatibility: Windows PowerShell 5.1 and Microsoft-supported PowerShell 7 release lines on Windows, macOS and Linux.
- Protected files: all 13 concrete selected paths are authorized in `.template-sync/marker.yml` before creation. Copy the sync instruction catalog intact. No protected removals. Commit isolation remains applicable to any later owner commit; this task creates no commits.
- Required layout: `.github/`, `.template-sync/`, schemas, tooling and tests. Rename `LICENSE.txt` to `LICENSE`, preserving MIT text and author. Normalize policy-covered text in a separate byte-reviewed phase. Keep CSV bytes unchanged.

## Immediate Manual GitHub Settings

Each item has deferral state **asked and deferred**, owner **@franklesniak**, and trigger **immediately after adoption**. The owner explicitly authorized content to assume these settings are enabled. Actual hosted activation remains unverified until evidence is recorded here.

- [ ] Enable private vulnerability reporting under repository Settings > Security. Verify [private report intake](https://github.com/franklesniak/ROMSorter/security/advisories/new). Dependency: `SECURITY.md` and issue security links. Dependent-file status: owner-approved future-state content; settings activation pending.
- [ ] Enable Discussions under Settings > General > Features. Verify [Discussions](https://github.com/franklesniak/ROMSorter/discussions) and a help/Q&A category. Dependency: issue chooser support link and CONTRIBUTING support guidance. Dependent-file status: owner-approved future-state content; settings activation pending.
- [ ] Create or verify labels `bug`, `enhancement`, `documentation`, and `triage`. Use `triage` color `d4c5f9`, description `Needs triage`; retain the standard meanings of the other labels. Review-flow label requirements, if any, are listed in the final adoption summary. Dependency: issue-form labels. Dependent-file status: owner-approved future-state content; labels activation pending.
- [ ] Create an active branch ruleset named `Default branch quality` targeting the default branch (currently `master`). Require a pull request, **zero approvals**, resolved conversations, current passing required checks, and an up-to-date branch. Block force pushes and deletions. No routine bypass actors. Do not require code-owner approval, last-push approval, signed commits, linear history, deployments or merge queue. These extra requirements have no established contributor/release setup here. CODEOWNERS still routes ownership. Dependency: the retained CI and review policy; content assumes this policy, hosted activation pending.
- [ ] After the first successful hosted runs, select the exact check names in the adoption summary, with GitHub Actions as expected source where available. Verify an ordinary pull request cannot merge with a failing required check. Do not require scheduled EOL or bot-only preview checks on every pull request.
- [ ] Verify the issue chooser, issue forms, PR template, security links and Discussions routing on GitHub after the files reach the default branch. Verify Copilot setup activation on the default branch separately from ordinary workflow runs.

Required check candidates, as declared by the adopted workflows: `Pre-commit`, `markdownlint`,
`Data file linting`, `Check for OWNER/REPO Placeholders`, `Instruction Contracts`, `Workflow Security`,
`Lint (PSScriptAnalyzer)`, `PowerShell Tests (ubuntu-latest, pwsh)`,
`PowerShell Tests (windows-latest, pwsh)`, `PowerShell Tests (macos-latest, pwsh)`, and
`PowerShell Tests (windows-latest, powershell)`. Confirm their exact emitted names from successful hosted
runs before selecting them. `copilot-setup-steps` is agent setup, and `Node.js EOL` is scheduled maintenance;
neither is a required ordinary pull-request check.

Collaboration follow-up status: Owner-approved future-state content; @franklesniak completes settings immediately after adoption.

References: [private reporting](https://docs.github.com/en/code-security/how-tos/report-and-fix-vulnerabilities/configure-vulnerability-reporting/configuring-private-vulnerability-reporting-for-a-repository), [Discussions](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/enabling-features-for-your-repository/enabling-or-disabling-github-discussions-for-a-repository), [labels](https://docs.github.com/en/issues/using-labels-and-milestones-to-track-work/managing-labels), and [ruleset controls](https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-rulesets/available-rules-for-rulesets).

## Decision: Sole-Maintainer Ruleset

The AGENTS shared decision process applies. Options considered before scoring: leave protection unchanged; require one/code-owner approval; require pull requests with zero approvals and required checks; allow direct pushes with checks. A classic protection rule with equivalent controls is grouped with the ruleset option because its enforcement intent is equivalent; the ruleset offers clearer independent management.

Scores are judgments from 1 (poor) to 5 (strong). Weights prioritize integrity (35%), sole-maintainer usability (30%), auditability (20%), and maintenance (15%). Integrity has highest weight because a convenient workflow must still enforce validation. Independent human approval is unavailable for self-authored changes.

| Option | Integrity 35% | Usability 30% | Auditability 20% | Maintenance 15% | Weighted total |
| --- | ---: | ---: | ---: | ---: | ---: |
| Unchanged settings | 1 | 5 | 1 | 5 | 2.80 |
| One/code-owner approval | 5 | 1 | 5 | 2 | 3.35 |
| Pull request, zero approvals, checks | 4 | 5 | 5 | 4 | 4.50 |
| Direct pushes with checks | 3 | 4 | 2 | 4 | 3.25 |

Select the zero-approval ruleset above. It enforces CI and preserves a review record without an impossible self-approval requirement. No routine bypass preserves enforcement. Emergency settings changes remain explicit owner actions. GitHub documents that a required pull request need not require approval; no independent human-review guarantee is claimed.

## Decision: Initial Analyzer Gate

Options considered before scoring: fix every legacy warning during adoption; keep strict gating and leave adoption blocked by legacy debt; use the template's temporary `first-adoption` mode with visible debt and a strict-gate exit task; disable rules or skip legacy scripts. Narrow suppressions are reserved for findings proved inapplicable, not blanket debt clearance.

Scores use 1–5. Weights: behavioral preservation (30%), failure truth and security (30%), maintainability (25%), scoped delivery (15%). Existing scripts lack behavioral regression coverage, so broad warning-driven changes risk altering ROM handling.

| Option | Preservation 30% | Failure truth 30% | Maintainability 25% | Scope 15% | Weighted total |
| --- | ---: | ---: | ---: | ---: | ---: |
| Fix all warnings now | 2 | 5 | 4 | 1 | 3.25 |
| Strict, blocked by legacy debt | 5 | 5 | 2 | 1 | 3.65 |
| Temporary first-adoption, visible debt | 5 | 4 | 4 | 5 | 4.45 |
| Disable rules or skip scripts | 5 | 1 | 1 | 5 | 2.80 |

The reviewed baseline is [249 warnings across existing product scripts](docs/adoption/analyzer-debt.json), with zero errors or unknown severities. New validation tooling contributes no findings.

Select `first-adoption` temporarily in PowerShell CI. Preserve all analyzer rules and all product candidate paths. Errors and unknown severities fail; warnings remain annotated and reported. This mode does not prevent new warnings automatically. Review warning changes against the recorded baseline until strict mode returns. The owner delegated this choice through the template decision process.

## Post-Adoption Tasks

### Decision: Node.js Development Support Floor

The retained EOL check reports that the inherited Node.js 22 floor is within its 180-day warning window
(end of support 2027-04-30). All retained hosted Node jobs already use Node.js 24. Options were considered
before scoring: retain 22 and defer the failing maintenance check; raise the floor to 24; raise it to 26;
or weaken the warning window. The rubric prioritizes supported-tool policy (35%), CI consistency (30%),
developer compatibility (20%), and ongoing maintenance (15%). Scores use 1–5.

| Option | Policy 35% | CI 30% | Compatibility 20% | Maintenance 15% | Total |
| --- | ---: | ---: | ---: | ---: | ---: |
| Retain 22 and defer | 2 | 4 | 5 | 2 | 3.20 |
| Raise floor to 24 | 5 | 5 | 4 | 5 | 4.80 |
| Raise floor to 26 | 5 | 2 | 2 | 3 | 3.20 |
| Weaken warning window | 1 | 4 | 5 | 1 | 2.70 |

Select a Node.js 24 development floor in package metadata and the matching lockfile root. This aligns
with existing CI, preserves the warning policy, and does not affect PowerShell product compatibility.
The scheduled EOL check remains enabled. Node.js 26 remains compatible with this minimum.

### Decision: Downstream Support-Test Profile

The first full pytest run exposed fixtures that assumed upstream `pyproject.toml`, a full template README,
and excluded Azure assets. Options considered before scoring were to import unselected source assets,
skip the two entire test files, or adapt configuration discovery and mark only source-dependent cases.
Weights prioritize retained coverage (45%), adopted scope (35%) and maintenance (20%); scores use 1–5.

| Option | Coverage 45% | Scope 35% | Maintenance 20% | Total |
| --- | ---: | ---: | ---: | ---: |
| Import unselected product assets | 5 | 1 | 2 | 3.00 |
| Skip both test files | 1 | 5 | 4 | 2.90 |
| Narrow fixture and case changes | 5 | 5 | 4 | 4.80 |

Select narrow changes: prefer retained `pytest.ini` with upstream `pyproject.toml` fallback; retain
strict marker assertions; classify eight full-source README/Azure cases as `upstream_template_only`.
All nine non-symlink initial failures passed against pinned upstream. The two corrected configuration
cases pass on Windows and Ubuntu. No runtime validator, applicable security assertion or complete test
file is disabled. The original full-run result and focused retries remain visible in the adoption summary.

### Follow-Ups

- [ ] Characterize product behavior with offline Pester fixtures before refactoring. Cover parser success, malformed inputs, output schemas and rename/copy collision handling. Mock network and external filesystem effects. Preserve every public path and command. Use Windows PowerShell 5.1 and supported PowerShell 7 on all three operating systems.
- [ ] Triage the recorded analyzer debt by rule and script family. Fix safe findings with behavioral tests; use justified narrow suppressions only for intentional compatibility. Restore strict gating when warnings are resolved. Do not weaken shared analyzer settings. Owner: @franklesniak; begin immediately after adoption, complete before declaring the analyzer migration finished.
- [ ] Assess duplicated helper extraction only after characterization tests pass. Keep root scripts and CSVs as compatibility entry points. Do not move public files without a separate compatibility plan; a wholesale `src/` or `scripts/` migration is not recommended during adoption.
- [ ] For each later task, preserve inputs, outputs, filenames and supported hosts; update touched `Last Updated`/`Version` metadata; run applicable Pester/pre-commit/Markdown gates. Protected instruction edits require separate explicit authority.
- [ ] Review the development-tool advisories inherited from the pinned template lockfile. The 2026-11-11 `npm audit` result is 15 findings (3 low, 11 moderate, 1 high); `npm audit --omit=dev` reports zero. Evaluate the affected Markdown tool chain and transitive packages together, update deliberately without `npm audit fix --force`, and rerun the Markdown and Node regression tests. Record justified exceptions if no compatible fix exists. Owner: @franklesniak; prioritize before processing untrusted Markdown in local tooling.
- [ ] Maintain the supported PowerShell matrix against Microsoft's lifecycle table. On 2026-11-11, 7.6 is the supported PowerShell 7 release line; 7.4 and 7.5 reached end of support on 2026-11-10. Keep hosted runner patches current, add new supported release lines, and retain the Windows PowerShell 5.1 job. Hosted macOS/Linux execution is a post-push validation step, not a locally verified result.

## Copy-Ready GitHub Issue Draft: Rename the Default Branch

No GitHub Issue is created during this adoption, as directed by the owner.

````markdown
Title: Plan and migrate the default branch from master to main without breaking consumers

## Context

Template adoption is complete. ROMSorter still uses `master`. This issue handles a separately authorized future migration to `main`.

## Scope

- Confirm current default branch and owner intent before changing remote settings or branch names.
- Inventory repository links, script download URLs, CI references, rulesets, open pull-request bases, release automation and local-clone guidance.
- Audit every `raw.githubusercontent.com/franklesniak/ROMSorter/master/` consumer. GitHub branch web redirects do not guarantee raw-content URL compatibility.
- Identify known external scripts, documentation, bookmarks and automation consuming the old raw URLs. Record unknown consumers as a compatibility risk.
- Choose and document a migration/transition plan, including whether a maintained compatibility `master` ref is needed and its retirement criteria.
- Update repository-owned URLs and migrate known external consumers before retiring compatibility paths.
- Retarget rulesets and open pull requests without weakening required checks; give local clone update instructions.
- Protected instruction files are out of scope. Obtain explicit authority if a required reference resides in one.

## Non-Goals

- Do not rename branches as part of template adoption.
- Do not change script behavior, public script/CSV paths, license, history or unrelated tooling.
- Do not assume web redirects preserve raw downloads or external automation.

## Acceptance Criteria

- Record the owner-authorized change window, current/target branch and rollback plan.
- Record the complete in-repository old-branch reference inventory and external-consumer audit results.
- Validate representative raw script/CSV downloads from both intended new URLs and any compatibility URLs.
- Keep required-check and force-push/deletion protections effective throughout transition.
- Retarget open pull requests deliberately; publish local-clone instructions and compatibility retirement criteria.
- Verify GitHub default branch is `main` only after explicit owner execution authority.
- Update touched files' `Last Updated` and synchronized `Version` metadata where present.

## Validation

- Search tracked files for `master`, repository URLs and `raw.githubusercontent.com` references; classify intentional historical references.
- Run retained pre-commit, Markdown, workflow-security and instruction checks, and affected offline Pester tests.
- Verify branch metadata, rulesets, required-check names and open pull-request bases through GitHub.
- Check download status and expected content/hash for representative public assets without running downloaded scripts.
- Verify normal clone/fetch and documented existing-clone migration steps in an isolated checkout.

## Preservation Notes

Preserve commit history, releases, contributor access and external consumers. The issue is limited to this repository and needs no other checkout. Any external change requires the relevant owner's authorization.
````
