<!-- markdownlint-disable MD013 -->
# Downstream Template Update Procedure

**Version:** 1.3.20260929.0

## Metadata

- **Status:** Active
- **Owner:** Repository Maintainers
- **Last Updated:** 2026-09-29
- **Scope:** Defines the selective review procedure for downstream repositories that were created from, or adopted files from, this template repository. Covers manual and agent-assisted syncs from later upstream template changes, first-adoption preflight state, the first-adoption bootstrap command, the read-only first-adoption preflight/questionnaire mode, raw first-adoption state reporting, first-adoption quality-debt reports and suppressions, the adoption difficulties journal, one-shot first-adoption materialization, shell-safe first-adoption args files, package identity and collaboration-policy materialization, first-adoption structural convention assessment, first-adoption working-tree validation and doctor diagnostics, downstream local path ownership records, the human-readable view of the template sync manifest, required/recommended/deferred structural-change classification, protected-file decision records, the marker-aware retained-state validation helper command, the excluded-module cleanup report, the sync candidate table generator, post-adoption issue drafting, the generated adoption ledger review artifact, and the concise adoption summary for PR descriptions. Does not define an automated ongoing upstream sync tool.
- **Related:** [Optional Configurations](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/OPTIONAL_CONFIGURATIONS.md), [Getting Started for New Repositories](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/GETTING_STARTED_NEW_REPO.md), [Getting Started for Existing Repositories](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/GETTING_STARTED_EXISTING_REPO.md), [Repository Copilot Instructions](.github/copilot-instructions.md)

## Purpose

Downstream template updates are a selective review process, not a raw merge. A downstream repository may have removed optional modules, customized project identity, changed workflows, or retained only part of the language/tooling guidance from this template. The goal of this procedure is to review later upstream template changes deterministically, adopt only relevant improvements, and preserve downstream ownership decisions.

Use this procedure when a downstream repository wants to review new changes from `franklesniak/copilot-repo-template` after initial adoption. The procedure is suitable for a repository owner or for a coding agent operating under owner direction.

For the first import into an existing repository, the separate one-shot materialization command can copy retained template-managed files into a target working tree. A non-empty target is expected to surface conflicts for pre-existing files such as `README.md`, `LICENSE`, and `.gitignore`. Treat those as per-path adoption decisions recorded through `local_overrides`, `protected_file_decisions`, or deferred protected candidates, not as runtime/tool failures. When the downstream license text already exists under an alternate root filename such as `LICENSE.txt` or `LICENSE.md`, the first-adoption materializer can normalize that owner-approved text to root `LICENSE` while suppressing the template license and leaving the original source path for manual cleanup.

## Terminology

- **Module:** A unit in the taxonomy defined by `.template-sync/manifest.yml` and rendered in this procedure, such as `markdown`, `powershell`, or `terraform`. Procedure logic operates on modules.
- **Stack:** Informal shorthand for a related grouping of modules. For example, a "PowerShell stack" may mean `powershell`, `markdown`, `yaml`, and `agent-instructions`, depending on what the downstream repository adopted. Stack is acceptable in prose, but sync decisions MUST be recorded in module terms.
- **Downstream sync marker:** The `.template-sync/marker.yml` file in the downstream repository. Under `template_sync`, it records the upstream template repository, the newest upstream template commit that has been reviewed, the modules the downstream repository has adopted, local override rules, downstream local path ownership records, protected-file decision records, deferred protected-file candidates, and instruction-contract waivers.
- **Instruction contract:** A machine-readable required-anchor contract in `.template-sync/instruction-contracts.yml`. Each entry names a protected instruction file, the modules that make the contract relevant downstream, and the required headings or phrases that MUST remain present unless a marker waiver or authorized protected-file removal applies.
- **First-adoption preflight checklist:** A root `_TODO-repo-init.md` file, or an equivalent committed adoption note named by this procedure, that records manual GitHub settings and maintainer policy decisions that cannot be inferred from repository files during first-time template adoption.
- **First-adoption state:** The resolved answers from the first-adoption preflight checklist, `.template-sync/marker.yml`, or an equivalent committed adoption note. Examples include conduct and security reporting channels, private vulnerability reporting, Discussions, expected labels, CODEOWNERS owner/team identity, default-branch protection policy, adoption mode, and any GHES host override.
- **First-adoption bootstrap command:** The `.template-sync/scripts/bootstrap_first_adoption.py` helper used before pruning to compose the existing preflight, adoption journal, raw-state, marker validation, and working-tree validation-plan helpers under one command. Its default mode is read-only and prints proposed `_TODO-repo-init.md`, `_ADOPTION-DIFFICULTIES.md`, and draft marker content. Its `--write` mode creates missing downstream-owned bootstrap files idempotently, preserves existing checklist and journal content, updates only the bootstrap-owned structured state block when `--update-existing-todo-state` is selected, and writes a missing draft marker only when `--write-draft-marker` is selected.
- **Structural convention finding:** A first-adoption observation about repository layout, path conventions, retained workflow or validator roots, template-module assumptions, or modernization opportunities discovered through the structural convention assessment in [Getting Started for Existing Repositories](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/GETTING_STARTED_EXISTING_REPO.md#structural-convention-assessment).
- **Required structural alignment:** A path, filename, directory, or command shape that must exist, or must be explicitly remapped, for a selected template module, repository platform feature, validator, or retained instruction contract to work.
- **Post-adoption issue draft:** A self-contained GitHub Issue description for deferred modernization or cleanup after template adoption is complete. Each draft is scoped to one repository and includes scope, non-goals, acceptance criteria, validation, and preservation notes.
- **First-adoption materialization command:** The `.template-sync/scripts/materialize_downstream_adoption.py` one-shot helper used during first adoption to stage retained manifest-owned template files, prune inline blocks for excluded modules, optionally run approved placeholder replacement against the staging tree, optionally normalize owner-approved downstream license text from an alternate source path to root `LICENSE`, and reconcile non-conflicting staged files into a target working tree. It can read from an explicit `--template-root`, or from a locally available upstream `--template-ref` / `--template-revision` through a private temporary detached worktree created from `--template-repo`. A complete, clean `--template-root` can be reported as observed source identity, but it becomes reviewed marker state only when the operator explicitly supplies `--stamp-resolved-source-as-reviewed` or `--last-reviewed-template-commit FULL_SHA` after review. It is a write-once adoption helper, not an automated ongoing upstream sync tool.
- **First-adoption working-tree validation runner:** The `.template-sync/scripts/run_first_adoption_checks.py` helper used before the first adoption commit. It collects tracked and untracked non-ignored regular files using `git ls-files --cached --others --exclude-standard`, prints a deterministic numbered command plan, runs read-only line-ending, path-reference, Markdownlint, and PSScriptAnalyzer debt reports when the quality report helper is present, runs chunked `pre-commit run --files ...` commands against that file list, and runs the placeholder scan, marker validation, supported Markdown package scripts, and the downstream-safe pytest gate when those checks are applicable. Its default `--check` mode makes validation intent explicit, its `--fix` mode is the opt-in surface for mutating hooks or fixers, and its `--plan-only` mode prints discovery results, notes, and the command plan without running validation commands or pytest availability probes. Its `--doctor` mode runs bounded, non-mutating Git, Python, pre-commit, pytest, Node/npm, PowerShell, PSScriptAnalyzer, and yamllint availability probes without running validation hooks or fixers. Normal validation execution prints UTC start and end timestamps, elapsed time, group label, index, and exit status for each planned command, prints a deterministic Git changed-file summary with aggregate before/after status and, when applicable, command-boundary status transitions observed after executed planned commands, then prints total elapsed time for the run.
- **First-adoption quality suppressions:** The optional downstream-created `.template-sync/first-adoption/quality-suppressions.json` file records durable suppressions for first-adoption quality reports. The current schema defines the report-scoped `path-reference` section only, with selector fields for rule/category identifier, source path or glob, and literal or pattern. The file is retained downstream state like `.template-sync/marker.yml`; the upstream template classifies and validates it when present but does not ship a concrete suppression file.
- **Adoption ledger:** A generated Markdown review artifact emitted by `.template-sync/scripts/generate_sync_candidates.py`. It summarizes manifest module assignments, marker local overrides, local path ownership records, protected-file flags and decisions, adoption-mode posture, `_TODO-repo-init.md` checklist links, and affected validation commands. It is not authoritative state; `.template-sync/manifest.yml` and `.template-sync/marker.yml` remain the machine-readable sources of truth.
- **Adoption summary:** A concise Markdown review artifact emitted by `.template-sync/scripts/generate_sync_candidates.py --summary` for PR descriptions. It summarizes included and excluded modules, protected-file decision records, local overrides, local path ownership records, unresolved maintainer decisions, machine-interpretable manual TODO items, and retained-module validation commands. It is not the detailed path-level review artifact; use the adoption ledger for protected-file and per-path review details.
- **Adoption difficulties journal:** An optional root `_ADOPTION-DIFFICULTIES.md` file created from `templates/adoption/_TEMPLATE-ADOPTION-DIFFICULTIES.md`. It records process and difficulty evidence from first adoption. It is not `_TODO-repo-init.md` decision state and is not `.template-sync/marker.yml` sync state.
- **Adoption mode:** The preservation posture applied before editing protected files and template-derived governance, community, process, workflow, or collaboration files. Valid named modes are `minimal-preservation` and `tailored`.
- **`minimal-preservation`:** The default adoption mode for protected files and template-derived governance, community, process, workflow, and collaboration files. Keep upstream wording and structure; limit edits to placeholder substitution, removing complete delimited sections owned by unadopted manifest modules, fixing broken links, and adding required local overrides that are recorded in `.template-sync/marker.yml`.
- **`tailored`:** A maintainer-selected adoption mode for a specific file or file set that allows broader downstream rewriting. The maintainer MUST select `tailored` explicitly before the broader rewrite starts. For protected files, the selection MUST be recorded in `template_sync.protected_file_decisions` with `tailored_authorization_basis`; for non-protected files, record the selection in `_TODO-repo-init.md`, `.template-sync/marker.yml` local overrides, the sync working notes, or the final sync summary.
- **Reviewed range:** The upstream template commit range inspected during a delta sync. It is recorded as `RANGE_BASE_SHA..RANGE_HEAD_SHA`, where both endpoints are resolved upstream template commit SHAs. Full reconciliation does not use a delta range; it compares a committed downstream snapshot against the resolved upstream range head.
- **Included module:** A module listed in the downstream sync marker under `template_sync.included_modules`.
- **Unadopted-module activity:** Upstream activity in a known taxonomy module that is not listed in `included_modules`.
- **Unknown module:** A module name introduced by a newer upstream manifest or procedure that the downstream marker does not recognize. Unknown modules MUST be surfaced for explicit owner decision.
- **Local path ownership:** A marker record under `template_sync.local_path_ownership` that documents downstream-owned exact paths or directory-prefix path families, such as `docs/local.md` or `docs/`. Local ownership records are not manifest rows and do not make downstream project files template-managed.
- **Protected file:** A governance or instruction file that requires explicit owner authorization before editing.
- **Sync working notes:** Temporary notes maintained while applying this procedure. They MAY be a scratch document, a draft PR body, or another local checklist, but they are not the final sync summary. They MUST capture the first-adoption preflight disposition and any unresolved `_TODO-repo-init.md` manual GitHub settings when applicable, structural convention findings and their required/strongly recommended/post-adoption/not-recommended classifications when first adoption applies, required structural changes, the range mode, range endpoints or reconciliation command, range-base rationale, saved Step 6 candidate table or table citation, saved adoption ledger location or ledger citation, unmapped paths, per-file decisions, protected-file dispositions, line-ending normalization actions, validation results, validation issue classifications, post-adoption issue drafts, finalization mode, and open questions as those facts are discovered. Step 14 turns these working notes into the final sync summary.
- **Sync summary:** The final owner-facing record created in Step 14 from the sync working notes. Depending on the finalization mode, it MAY be a PR description, a committed summary artifact, a local handoff note, or a dry-run report. It is the durable review artifact for modes that commit a branch or open a PR; working-tree inspection and dry-run modes MUST still present it clearly before stopping.

## Safety Rules

- Do not run `git pull template main`, `git merge template/main`, or `git rebase template/main` as the update mechanism. Fetch first, inspect the range, and make explicit per-file decisions.
- Do not overwrite downstream project identity, repository URLs, issue templates, PR templates, workflow runner choices, validation commands, README content, package metadata, or local policy without a recorded decision.
- Do not invent contact emails, reporting channels, branch protection policy, default-branch ruleset settings, CODEOWNERS identities, label existence, Discussions state, private vulnerability reporting state, GHES hosts, adoption modes, or template-preservation policy.
- Do not edit protected files unless the owner gives explicit, path-scoped authorization in the current task.
- Do not edit or remove protected files until `.template-sync/marker.yml` records a matching `template_sync.protected_file_decisions` entry for the protected path and decision.
- Do not treat `minimal-preservation` as protected-file authorization. Adoption mode limits what an authorized edit may do; it does not grant permission to edit protected files.
- Do not weaken existing security, validation, or pre-commit expectations to make a sync easier.
- Do not silently include or exclude unknown modules.
- Do not re-ask first-adoption preflight questions that are already recorded as resolved in `_TODO-repo-init.md`, `.template-sync/marker.yml`, or an equivalent committed adoption note named by this procedure.

## Procedure Overview

At the start of the procedure, determine whether the first-adoption preflight gate below applies. If it applies, generate or update the checklist before any content edits whose correctness depends on unresolved adoption answers, and run the structural convention assessment from [Getting Started for Existing Repositories](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/GETTING_STARTED_EXISTING_REPO.md#structural-convention-assessment). Before editing protected files or template-derived governance, community, process, workflow, or collaboration files, create or review the adoption ledger and record the applicable adoption mode. Use `minimal-preservation` when no explicit maintainer selection exists.

1. Create a dedicated sync branch.
2. Add this template repository as a `template` remote if it is not already present.
3. Fetch upstream template changes without merging.
4. Identify the upstream commit range under review with rename detection.
5. Initialize or update `.template-sync/marker.yml`.
6. Filter upstream changes through the authoritative module taxonomy and create or refresh the adoption ledger.
7. Review every candidate file with an explicit adoption mode and decision.
8. Perform manual merges using an ignored scratch location when needed, and normalize line endings after `.gitattributes` changes.
9. Handle protected files through authorization or deferral.
10. Preserve local customizations and downstream project identity.
11. Re-substitute template placeholders such as `OWNER/REPO`.
12. Run whitespace checks and validation for the adopted modules.
13. Record the latest reviewed template commit.
14. Finalize the sync using a declared mode and clear sync summary.

Independent substeps, such as inspecting unrelated candidate files or collecting validation commands for separate modules, MAY be performed in parallel. The final recorded decisions and marker update MUST remain deterministic.

## First-Adoption Preflight Gate

Run this gate only when the downstream repository is receiving template contents for the first time, or when the repository is missing recorded first-adoption state that affects the files under review. Normal initialized delta syncs MUST NOT re-trigger this gate when `_TODO-repo-init.md`, `.template-sync/marker.yml`, or an equivalent committed adoption note already records the relevant answers.

If the gate applies, generate or update root `_TODO-repo-init.md` before editing files whose content depends on unresolved adoption answers. Discovery, manifest inspection, range selection, and non-content setup may still happen before the checklist is complete. The checklist is downstream-owned state, not an upstream template file, and is excluded from upstream sync candidate review.

When `.template-sync/scripts/bootstrap_first_adoption.py`, `.template-sync/scripts/generate_sync_candidates.py`, `.template-sync/scripts/run_first_adoption_checks.py`, `.template-sync/scripts/validate_marker.py`, `.template-sync/manifest.yml`, and the template-sync schemas are available, run the bootstrap command before pruning, protected-file edits, or policy-dependent file materialization:

```bash
python .template-sync/scripts/bootstrap_first_adoption.py
```

The default bootstrap mode is read-only. It prints proposed `_TODO-repo-init.md`, `_ADOPTION-DIFFICULTIES.md`, and optional draft `.template-sync/marker.yml` content, reuses the existing preflight protected-file authorization questions, reports raw first-adoption state, names intended operations and protocol-preservation constraints, and surfaces a runnable validation plan from `.template-sync/scripts/run_first_adoption_checks.py`.

After reviewing the report, use write mode only to create missing bootstrap-owned files:

```bash
python .template-sync/scripts/bootstrap_first_adoption.py --write
```

Write mode is idempotent. It creates `_TODO-repo-init.md` when absent, creates `_ADOPTION-DIFFICULTIES.md` when absent, and preserves existing checklist, journal, and marker content. Use `--update-existing-todo-state` only to insert or refresh the bootstrap-owned structured state block in an existing checklist while preserving free-form Markdown notes outside that block. Use `--write-draft-marker` only to create a missing validated draft marker; existing marker content is never overwritten.

Use `--include-github-metadata` only after the maintainer explicitly opts in to read-only GitHub metadata discovery through the `gh` CLI. Without that opt-in, GitHub-only settings remain manual-review items in the generated report.

The lower-level preflight/questionnaire mode remains available when a read-only report without bootstrap write surfaces is needed:

```bash
python .template-sync/scripts/generate_sync_candidates.py --preflight
```

The preflight report includes a raw first-adoption state section that separates initialized sync state from local repository clutter. It reports marker evidence, adoption notes, tracked files, untracked files, ignored files, physical empty directory trees, missing state files, and high-signal path presence. Default output is bounded to deterministic samples with explicit remainder counts. Use `--full-state` with `--preflight` when the full raw inventory is needed:

```bash
python .template-sync/scripts/generate_sync_candidates.py --preflight --full-state
```

The bootstrap command also prints the journal scaffold and can create the optional journal with `--write`. If the adoption work exposes process difficulties worth preserving and the bootstrap command is unavailable, create the optional journal directly before continuing:

```bash
python .template-sync/scripts/initialize_adoption_journal.py
```

Use `--journal-path` when the downstream repository keeps adoption notes somewhere other than the root journal path. The helper never overwrites an existing journal.

After context loss, interruption, or compaction, reread `_ADOPTION-DIFFICULTIES.md` when present, `_TODO-repo-init.md`, `.template-sync/marker.yml` when present, and the current raw repository state before continuing.

When the gate applies, run the structural convention assessment in [Getting Started for Existing Repositories](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/GETTING_STARTED_EXISTING_REPO.md#structural-convention-assessment) and carry the findings into the first-adoption working notes. Each finding MUST be classified as required for selected template modules, strongly recommended during adoption, post-adoption follow-up, or intentionally not recommended. Required structural changes MUST be implemented or explicitly remapped before adoption is finalized. Post-adoption modernization MUST be drafted as issue follow-up and MUST NOT be bundled into template adoption unless the owner explicitly authorizes it.

When governance, security, community, CODEOWNERS, issue-template, PR-template, or other template-derived file content depends on unknown owner policy, the agent MUST ask the owner a concrete question before finalizing that file. The gate is a decision point and evidence record, not a substitute for asking. Generated or adopted files MUST NOT ship unresolved placeholders or misleading policy defaults.

Record an unresolved dependent-file item only after identifying the concrete owner question and the dependent file/status impact. Each unresolved item MUST use exactly one deferral state:

- `not yet asked`: use when the concrete owner question is identified but has not been asked.
- `asked and deferred`: use when the owner was asked and intentionally deferred the answer, with dependent-file status recorded.
- `unavailable through current safe tooling / manual review required`: use as the could-not-determine state when the answer cannot be discovered through current safe tooling and requires owner or manual reviewer action.

The checklist MUST separate:

- **Discoverable repository state:** owner/name, default branch, existing files, existing marker state, and prior committed adoption notes.
- **Manual GitHub settings:** private vulnerability reporting, Discussions, expected labels such as `triage`, and default-branch protection or rulesets.
- **Maintainer policy decisions:** Code of Conduct reporting contact method, security reporting channel, CODEOWNERS owner/team identity, adoption mode for protected and template-derived files, explicit `tailored` opt-ins, and any GHES host override.
- **Protected-file adoption decisions:** protected instruction-file identification, edit authorization, removal authorization, and marker updates.
- **Unresolved settings:** concrete owner questions, exactly one deferral state, and dependent file/status impact.
- **Resolution evidence:** proof of resolved settings, intentionally deferred items, and placeholder/default safety.

Use this copyable starting point when the downstream repository does not already have an equivalent adoption note:

```markdown
# Repository Initialization Checklist

This file records first-adoption decisions for this downstream repository. It is downstream-owned state, not an upstream template file, and is excluded from upstream template sync candidate review.

Keep unresolved items in this file until the maintainer completes them through the GitHub UI, an authorized GitHub connector, a safe API call, or a later maintainer action. Do not create this file in the upstream template repository unless the template manifest and sync procedure are deliberately changed to handle it.

When a template-derived file depends on unknown governance, security, community, CODEOWNERS, issue-template, PR-template, or other owner policy, ask the owner a concrete question before finalizing that file. Record unresolved dependent-file items only after the question and the dependent file/status impact are identified.

## Discoverable Repository State

- [ ] Repository owner/name recorded from the GitHub URL or Git remote: `OWNER/REPO`
- [ ] Repository visibility recorded: `[public, private, or internal]`
- [ ] Repository default branch recorded neutrally from GitHub or `git remote show origin`: `[recorded default branch]`
- [ ] Existing `SECURITY.md`, `CODE_OF_CONDUCT.md`, `.github/CODEOWNERS`, issue templates, PR template, and `.github/dependabot.yml` reviewed before replacement.
- [ ] Existing `.template-sync/marker.yml`, `_TODO-repo-init.md`, or equivalent adoption note checked before asking repeated questions.
- [ ] Existing labels reviewed in GitHub under **Issues** > **Labels**.
- [ ] Existing Discussions state reviewed in GitHub under **Settings** > **General** > **Features**.
- [ ] Existing repository rulesets or branch protection reviewed in GitHub under **Settings** > **Rules**.

## Manual GitHub Settings

These settings may be completed through the GitHub UI even when `gh` is unavailable. If a GitHub connector or approved API call is used instead, record the action and evidence in the item.

- [ ] Private vulnerability reporting decision: `[enable, leave disabled, or not available]`
  - UI path: open **Settings** > **Code security and analysis**, then review **Private vulnerability reporting**.
  - After enabling, verify the **Security** tab and `SECURITY.md` private-reporting links.
  - Docs: [Configuring private vulnerability reporting for a repository](https://docs.github.com/en/code-security/how-tos/report-and-fix-vulnerabilities/configure-vulnerability-reporting/configuring-private-vulnerability-reporting-for-a-repository). If this link breaks, search `private vulnerability reporting repository` on `docs.github.com`.
  - Evidence: `[date, actor, screenshot/link, connector result, or maintainer note]`
- [ ] GitHub Discussions decision: `[enable or leave disabled]`
  - UI path: open **Settings** > **General** > **Features**, then review **Discussions**.
  - If enabled, verify the repository **Discussions** tab appears and decide whether issue templates should point users there.
  - Docs: [Enabling or disabling GitHub Discussions for a repository](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/enabling-features-for-your-repository/enabling-or-disabling-github-discussions-for-a-repository). If this link breaks, search `enable GitHub Discussions repository` on `docs.github.com`.
  - Evidence: `[date, actor, screenshot/link, connector result, or maintainer note]`
- [ ] Labels decision, including `triage`: `[create, skip, or already present]`
  - UI path: open **Issues** > **Labels**, then create or edit labels required by issue templates and local triage policy.
  - Minimum template label to review: `triage` with description `Needs triage` and color `d4c5f9`, unless the maintainer chooses a different triage label.
  - Docs: [Managing labels](https://docs.github.com/en/issues/using-labels-and-milestones-to-track-work/managing-labels). If this link breaks, search `manage labels GitHub issues` on `docs.github.com`.
  - Evidence: `[date, actor, labels created/skipped, connector result, or maintainer note]`
- [ ] Default branch decision: `[keep recorded default branch, rename, or defer]`
  - UI path: open **Settings** > **Branches**, then review the default branch selector before changing it.
  - If the recorded default branch is not `main`, surface that as repository state and ask whether to keep or rename it without implying it is wrong.
  - If renaming, update local clones, open PR bases, branch protection or rulesets, documentation, and CI references.
  - Docs: [Changing the default branch](https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-branches-in-your-repository/changing-the-default-branch). If this link breaks, search `change default branch GitHub repository` on `docs.github.com`.
  - Evidence: `[date, actor, old branch, new branch, connector result, or maintainer note]`
- [ ] Repository ruleset or branch-protection decision for the default branch: `[create ruleset, use classic branch protection, leave unchanged, or defer]`
  - UI path: open **Settings** > **Rules** > **Rulesets**, then create or review rules for branch names such as `main` or the recorded default branch.
  - Review required status checks, pull request review requirements, force-push restrictions, deletion restrictions, bypass roles, and enforcement mode.
  - Docs: [Creating rulesets for a repository](https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-rulesets/creating-rulesets-for-a-repository). If this link breaks, search `create repository ruleset GitHub` on `docs.github.com`.
  - Evidence: `[date, actor, ruleset name/link, connector result, or maintainer note]`

## Maintainer Policy Decisions

- [ ] Code of Conduct reporting contact method: `[explicit contact email/address, repository-maintainer profile contact, another maintainer-approved contact path, or intentionally deferred (dependent file status recorded)]`
- [ ] Security vulnerability reporting channel: `[private vulnerability reporting, monitored email, both, or intentionally deferred (dependent file status recorded)]`
- [ ] CODEOWNERS owner/team identity: `[@user or @org/team]`
- [ ] Label-dependent issue-template behavior: `[use template labels, choose alternate labels, remove label metadata, or intentionally deferred (dependent file status recorded)]`
- [ ] Adoption mode for protected and template-derived governance, community, process, workflow, and collaboration files: `minimal-preservation` by default; list any specific `tailored` opt-ins.
- [ ] GHES host override: `[none or github.company.com]`

## Protected-File Adoption Decisions

- [ ] Protected instruction files identified before editing: `.github/copilot-instructions.md`, `.github/instructions/*.instructions.md`, `.cursor/rules/*.mdc`, `.hermes.md`, `AGENTS.md`, `CLAUDE.md`, and `GEMINI.md`.
- [ ] When `template-sync-support` is retained, `.template-sync/instruction-contracts.yml` identified as a separate protected governance path, including in profiles with no agent platform.
- [ ] Protected-file edits authorized by maintainer: `[none, path-scoped authorization, or deferred]`
- [ ] Protected-file removals authorized by maintainer: `[none, path-scoped authorization, or deferred]`
- [ ] `.template-sync/marker.yml` protected-file decisions updated if template sync support is retained.

## Unresolved Settings

Tag each unresolved item with exactly one deferral state:

- `not yet asked`: the concrete owner question is identified but has not been asked.
- `asked and deferred`: the owner was asked and intentionally deferred the answer; dependent-file status is recorded.
- `unavailable through current safe tooling / manual review required`: the answer could not be determined through current safe tooling and requires owner or manual reviewer action.

- [ ] Items that must be completed later:
  - Question: `[concrete owner question]`; state: `[not yet asked, asked and deferred, or unavailable through current safe tooling / manual review required]`; owner: `[person/team]`; target date or trigger: `[date/event]`; dependency: `[file, workflow, or GitHub setting]`; dependent-file status: `[not finalized, blocked, placeholder removed, local default withheld, or other concrete status]`
- [ ] Dependent files that must not be considered finalized until the unresolved items are complete:
  - `[path]` depends on `[concrete owner question]`; state: `[same deferral state used above]`; status: `[dependent-file status]`

## Resolution Evidence

- [ ] Every unresolved dependent-file item was recorded only after the concrete owner question and dependent file/status impact were identified.
- [ ] Dependent files are finalized only after their policy questions are resolved, or after the owner intentionally chooses `asked and deferred` and the dependent-file status avoids placeholders and misleading defaults.
- [ ] No generated or adopted file ships unresolved placeholders or misleading policy defaults.
- [ ] Evidence captured for every resolved GitHub setting or policy decision: `[commit, PR, screenshot/link, connector result, API response note, or maintainer note]`
- [ ] Deferred items copied into the adoption PR description, sync summary, issue, or other maintainer-owned follow-up with their deferral state.
```

Downstream work may assume a checklist item is complete only after it is recorded as resolved, or after the owner intentionally chooses `asked and deferred` with dependent-file status, in `_TODO-repo-init.md`, `.template-sync/marker.yml`, or the equivalent committed adoption note named by this procedure. Items tagged `not yet asked` or `unavailable through current safe tooling / manual review required` keep dependent files unfinalized. If the owner prefers a different committed adoption note instead of `_TODO-repo-init.md`, name that note in the sync working notes and final sync summary.

## Step 1: Create a Sync Branch

Start from the downstream repository's normal integration branch, then create a branch dedicated to the sync.

```bash
git checkout main
git pull --ff-only origin main
git checkout -b template-sync/YYYYMMDD
```

Use the downstream repository's actual default branch name if it is not `main`.

## Step 2: Add the Template Remote

Add this template repository as a remote named `template` if it is not already configured.

```bash
git remote add template https://github.com/franklesniak/copilot-repo-template.git
```

If the remote already exists, verify it points to the template repository.

```bash
git remote get-url template
```

## Step 3: Fetch Without Merging

Fetch upstream template refs without changing the working tree.

```bash
git fetch template
```

Do not merge, rebase, or pull from `template/main` at this step.

## Step 4: Identify the Reviewed Range

This step decides which upstream template changes need review during this sync. It does **not** decide which changes will be adopted. Later steps filter the changed paths by module, review each file, and record which upstream changes were taken, merged, skipped, or deferred.

### Terms Used in This Step

- **Marker path:** Downstream repositories use `.template-sync/marker.yml` as the sync marker path. The marker lives inside `.template-sync/` so the directory can hold committed template-sync support files: the sync marker and the sync manifest are committed today, while additional items such as review artifacts are only added if a later issue explicitly defines them as committed outputs.
- **Marker:** Short name for `.template-sync/marker.yml`.
- **Marker authority:** The marker is authoritative for marker-backed adoption and template sync while `template-sync-support` is retained. Standalone instruction enforcement uses the protected `.github/instruction-profile.yml` instead and rejects a conflicting marker. Before removing sync support, deliberately migrate relevant declarations and review the protected profile candidate; see the [upstream instruction-enforcement guide](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/docs/instruction-enforcement.md). A local exception declaration does not independently prove owner authorization.
- **`template_sync.last_reviewed_template_commit`:** The marker field that stores the newest upstream template commit already reviewed in a prior sync. The durable marker value MUST be a resolved upstream template commit SHA, not a branch name, tag name, or other moving ref. Always store the full 40-character SHA; short SHAs are ambiguous and are not durable marker values. See the [Step 5 example marker](#step-5-initialize-or-update-the-sync-marker) for the field in context.
- **Range mode:** The review path for this sync: normal delta sync, first sync from known lineage, timestamp-proxy delta sync, or full reconciliation.
- **Range base SHA:** The older endpoint of a delta reviewed range. Changes after this commit are candidates for review.
- **Marker-supplied range base:** The present, non-empty `template_sync.last_reviewed_template_commit` value from the marker.
- **Initial range base:** The upstream template commit where a first sync review starts. Upstream changes after that commit are candidates for review.
- **Exact initial range base:** An initial range base that is known to be the upstream template commit used to create or copy the downstream template content.
- **Timestamp-proxy delta sync:** A first-sync delta mode that uses the latest upstream template commit at or before a known copy or import timestamp as an approximate initial range base. It is useful for over-reviewing likely upstream changes, but it is not proof of lineage.
- **Range head ref:** The upstream branch, tag, or commit the owner chooses as the newer endpoint. By default, this is `template/main`.
- **Range head SHA:** The resolved upstream template commit SHA printed from the range head ref.
- **Full reconciliation:** A first-sync mode that compares the committed downstream snapshot against the resolved upstream range head instead of trusting a delta range. It does not require shared Git history between the downstream repository and the template repository.

### Discover the Marker

Before reading any marker contents, choosing a range mode, or filtering by module, check for `.template-sync/marker.yml`. Also check for `_TODO-repo-init.md` or an equivalent committed adoption note when this appears to be first-time adoption or missing first-adoption state.

- If `.template-sync/marker.yml` is present, read it as the downstream sync marker.
- If `.template-sync/marker.yml` is absent, proceed without a marker. The maintainer then chooses one of the first-sync paths below.

### Choose the Range Mode

| Mode | Definition | When to use | Example |
| --- | --- | --- | --- |
| Normal delta sync | Review upstream changes after the marker-supplied range base through the resolved range head. | `.template-sync/marker.yml` exists and `template_sync.last_reviewed_template_commit` is present and non-empty. | Review `1111111111111111111111111111111111111111..2222222222222222222222222222222222222222`. |
| First sync from known lineage | Review upstream changes after the exact upstream commit used to create or copy the downstream template content. | The marker does not supply a range base, but a trustworthy upstream origin commit is known. | Review `dddddddddddddddddddddddddddddddddddddddd..2222222222222222222222222222222222222222`. |
| Timestamp-proxy delta sync | Review upstream changes after an approximate date-based upstream commit chosen from the downstream copy or import timestamp. | The marker does not supply a range base, the exact source commit is unknown, and the owner accepts over-review risk from an approximate base. | Review `aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa..2222222222222222222222222222222222222222` and record the timestamp rationale. |
| Full reconciliation | Compare the committed downstream snapshot against the resolved upstream range head instead of trusting a delta range. | The marker does not supply a range base and no trustworthy exact or timestamp-proxy base should be used. | Compare `HEAD` to `2222222222222222222222222222222222222222`. |

Read `.template-sync/marker.yml`, when present, before running any diff. A marker supplies a range base only when `template_sync.last_reviewed_template_commit` is present and non-empty. If the marker is absent, or if the field is missing or empty, the marker does not supply a range base; the maintainer chooses one of the first-sync paths.

When an agent runs this procedure, the repository owner confirms any initial range base or timestamp-proxy choice before the agent trusts it.

Range mode does not, by itself, authorize re-asking first-adoption questions. First sync from known lineage, timestamp-proxy delta sync, and full reconciliation all use the first-adoption preflight gate only when the required adoption answers are absent or unresolved. Normal initialized delta sync uses the recorded first-adoption state and MUST NOT re-prompt for resolved answers.

| Situation | Mode | What to do |
| --- | --- | --- |
| Marker supplies a range base | Normal delta sync | Use the marker value as the marker-supplied range base, resolve the range head, run the reachability check, then diff the reviewed range. |
| No range base from the marker, but a trustworthy upstream origin commit is known | First sync from known lineage | Use that upstream template commit as the exact initial range base, resolve the range head, run the reachability check, then diff the reviewed range. |
| No range base from the marker, no exact upstream origin commit is known, and the owner accepts a date-based proxy | Timestamp-proxy delta sync | Use the latest upstream template commit at or before the known copy or import timestamp as the initial range base, resolve the range head, run the reachability check, then diff the reviewed range. |
| No range base from the marker, and no trustworthy upstream origin commit is known | Full reconciliation | Compare the downstream tree against the resolved upstream range head, apply the pre-filter below, adjudicate candidate files through Steps 6-7, then set the marker to the resolved upstream range head only after review is complete. |

A present, non-empty marker value that later fails the reachability check is handled by **Check That the Base Is Reachable** below. Do not run a separate routing-time reachability check, and do not silently replace the marker value with the range head.

Full reconciliation is the recommended path for downstream repositories that adopted this template by manual copy, ZIP download, GitHub web UI copy, cherry-picking selected files, heavily editing template files, or using the copy-based paths in [Getting Started for Existing Repositories](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/GETTING_STARTED_EXISTING_REPO.md).

### Find an Initial Range Base

Use this recipe only when the marker does not supply a range base and the owner wants a delta range instead of full reconciliation:

1. Check adoption records first: the PR or commit that created the downstream repository, local sync notes, release tags, or any recorded upstream template commit.
2. If the exact upstream source commit is known, resolve that commit and record it as an exact initial range base.
3. If the exact source commit is unknown but the template was copied as a snapshot on a known date, the owner MAY choose a timestamp-proxy delta sync: the latest upstream template commit at or before the copy or import timestamp.
4. Verify any proxy by spot-diffing several copied template files against the candidate upstream commit.
5. If the proxy remains uncertain, choose an older commit or use full reconciliation. Over-review is safer than under-review because under-review can silently skip upstream changes.

Record the range base type, the rationale, and any uncertainty in the sync working notes. A stale local checkout, partial copy, cherry-pick, or later manual edit can make a timestamp-proxy delta sync wrong. Record the timestamp-proxy base as date-based and approximate, not as a known source commit.

### Choose the Range Head

Use `template/main` as the range head ref unless the owner explicitly chooses a different upstream branch, tag, or commit. Resolve the ref to a commit SHA and record the printed SHA in the sync working notes as the range head SHA.

Get the range head SHA explicitly:

```bash
git rev-parse 'RANGE_HEAD_REF^{commit}'
```

For example, if the owner is using the default upstream branch, replace `RANGE_HEAD_REF` with `template/main`:

```bash
git rev-parse 'template/main^{commit}'
```

`template/main` can move after each `git fetch`, while the resolved SHA records exactly what was reviewed and is the value Step 13 later writes to `template_sync.last_reviewed_template_commit`. Typing `template/main` directly is acceptable shorthand during interactive inspection when upstream is not changing under the maintainer, but the sync working notes, reachability check, diff command, and marker update use the resolved SHA.

The `^{commit}` suffix peels the reference to its underlying commit. This matters when the range head is an annotated tag: a plain `git rev-parse` on an annotated tag prints the tag object's SHA instead of the commit SHA. For branches and lightweight tags the suffix is harmless and still prints the commit SHA. The single quotes keep the `^{commit}` suffix as literal text, so the shell does not interpret `^`, `{`, or `}`.

If the chosen range base is a ref or tag rather than a commit SHA, resolve it the same way before recording or using it:

```bash
git rev-parse 'RANGE_BASE_REF^{commit}'
```

`RANGE_BASE_REF`, `RANGE_HEAD_REF`, `RANGE_BASE_SHA`, and `RANGE_HEAD_SHA` are placeholders for values recorded in the sync working notes. They are not shell variables that the snippets set.

### Check That the Base Is Reachable

Run this check only for delta range modes. After choosing the range base SHA and range head SHA, confirm that the range base is still in the history that leads to the range head. In beginner terms, "reachable from the range head" means Git can start at `RANGE_HEAD_SHA`, walk backward through parent commits, and eventually find `RANGE_BASE_SHA`. If Git cannot find the base this way, the two endpoints do not describe a normal forward span of upstream template history.

Check reachability with the recorded SHAs:

```bash
git merge-base --is-ancestor RANGE_BASE_SHA RANGE_HEAD_SHA
```

For example:

```bash
git merge-base --is-ancestor 1111111111111111111111111111111111111111 2222222222222222222222222222222222222222
```

Git may print no output when this command succeeds. A successful exit, exit code `0`, means the range base is reachable from the range head and the reviewed range is coherent. A non-zero exit means the base is not reachable from the head. In that case, stop and ask the owner to choose a new range base before running or trusting the diff.

Reasonable replacement bases include the exact upstream origin commit, a conservative timestamp-proxy base, or an older upstream commit chosen after inspecting repository history. Do not silently reset the marker to the range head.

### Run the Delta Diff

After choosing both endpoints and confirming the base is reachable, replace `RANGE_BASE_SHA` and `RANGE_HEAD_SHA` with the recorded values and list the upstream paths changed in that range:

```bash
git diff --name-status -M RANGE_BASE_SHA..RANGE_HEAD_SHA --
```

The `-M` flag is required so upstream renames appear as renames, such as `R100 old/path new/path`, instead of unrelated add/delete pairs.

Example where the marker says the last reviewed upstream commit was `1111111111111111111111111111111111111111` and the resolved range head SHA is `2222222222222222222222222222222222222222`:

```bash
git diff --name-status -M 1111111111111111111111111111111111111111..2222222222222222222222222222222222222222 --
```

Example first-time sync where upstream commit `dddddddddddddddddddddddddddddddddddddddd` is the exact initial range base and the resolved range head SHA is `2222222222222222222222222222222222222222`:

```bash
git diff --name-status -M dddddddddddddddddddddddddddddddddddddddd..2222222222222222222222222222222222222222 --
```

Do not use the range head as an initial range base just to make the diff empty. That would mark upstream changes as reviewed without reviewing them and would erase the distinction between reviewed upstream changes and adopted upstream changes.

### Run Full Reconciliation

Use this path when the marker does not supply a range base and there is no trustworthy upstream origin commit. Full reconciliation works even when the downstream repository and the template repository share no Git history, because `git diff` can compare two trees directly.

Before running the comparison, confirm that the downstream working tree is clean or that any local edits that should be part of the reconciliation have been committed. In this command, `HEAD` means the committed downstream snapshot being reconciled. If a different local commit or branch is the downstream snapshot under review, use that ref instead of `HEAD`.

```bash
git diff --name-status -M HEAD RANGE_HEAD_SHA
```

Concrete example after substituting the resolved range head SHA:

```bash
git diff --name-status -M HEAD 2222222222222222222222222222222222222222
```

Hypothetical full-reconciliation output:

```text
A       .github/workflows/markdownlint.yml
D       src/app/legacy_service.py
M       README.md
R075    docs/old-template-guide.md   docs/template-sync-guide.md
```

When the downstream repository and the template repository share no Git history, read the status letters this way:

| Status | Meaning |
| --- | --- |
| `A` | Present in the upstream template only. |
| `D` | Present in the downstream repository only. |
| `M` | Present in both trees, but different. |
| `R` | Pairs a downstream-only path, shown first, with a similar upstream-only path, shown second. With no shared Git history, this is a content-similarity match, not a tracked rename; review each path on its merits. |

A cross-tree `R` row is advisory unless shared lineage is confirmed. For example, `R075 docs/old-template-guide.md docs/template-sync-guide.md` might mean the downstream file was renamed from an earlier template guide, or it might only mean the two files happen to share similar content. Without confirmed shared lineage, treat `docs/old-template-guide.md` as a `D` candidate and `docs/template-sync-guide.md` as an `A` candidate, then decide each through the same taxonomy and per-file review process.

Apply a Step 4 pre-filter before per-file review: when a downstream-only path matches no template-managed path or module mapping and is not template-derived, summarize it as local-only noise and exclude it from Steps 6-7. This keeps the downstream repository's own project files out of per-file adjudication.

Always treat `_TODO-repo-init.md` as downstream-owned first-adoption state. If it appears as a downstream-only path during full reconciliation, exclude it from upstream sync candidate review and cite it only as first-adoption state in the sync working notes or summary.

The pre-filter does not require listing every downstream-only project file. A count plus one-line categorization is enough, such as `247 application source files under src/app/** excluded as local-only noise`.

Do not exclude paths that appear template-derived or require owner attention. Template-derived paths include a copied template file that was later moved, renamed, or locally edited; a workflow copied from the template and renamed for the downstream project; a Markdown guide that still carries template placeholder conventions; or a protected instruction file whose content is based on this template. Send those paths through the Steps 6-7 taxonomy and per-file decision process unchanged.

### Step 4 Completion Checklist

Before moving to Step 5, the sync working notes MUST contain:

- Range mode: normal delta sync, first sync from known lineage, full reconciliation, or timestamp-proxy delta sync when that optional sub-path was used
- Range base SHA, when using a delta range, plus the ref or tag it resolved from if the base was not already a commit SHA
- Range base type: marker-supplied range base, exact initial range base, or timestamp-proxy base
- Range base rationale
- Range head ref
- Range head SHA
- Reachability check result, when using a delta range
- Diff command or full-reconciliation enumeration command used
- First-adoption preflight disposition: not applicable, already resolved with record path, generated or updated with unresolved items, or blocked pending owner answers
- Structural convention findings and their required/strongly recommended/post-adoption/not-recommended classifications, when first adoption applies
- Required structural changes, including whether each change was implemented or remapped in retained tooling, plus any intentional `.template-sync/marker.yml` local override recorded in addition for a deliberate downstream deviation
- Post-adoption issue drafts or draft locations for deferred structural modernization, when first adoption applies
- Local-only noise excluded by the full-reconciliation pre-filter, when applicable
- Candidate inline module blocks discovered in changed files, when applicable. Retain or strip decisions for these blocks belong to Step 6 once path mapping has run; record them per file in the Step 7 decision notes and the sync summary.
- Any uncertainty that should be carried into the sync summary

Example sync working-notes block:

```markdown
## Template Sync Working Notes

- Marker discovery: `.template-sync/marker.yml` present
- Range mode: normal delta sync
- Range base SHA: `1111111111111111111111111111111111111111`
- Range base type: marker-supplied range base
- Range base rationale: read from `template_sync.last_reviewed_template_commit`
- Range head ref: `template/main`
- Range head SHA: `2222222222222222222222222222222222222222`
- Reachability check: passed
- Enumeration command: `git diff --name-status -M 1111111111111111111111111111111111111111..2222222222222222222222222222222222222222 --`
- First-adoption preflight: not applicable; first-adoption state was already recorded
- Structural convention findings: not applicable for this initialized delta sync
- Required structural changes: none
- Post-adoption issue drafts: none
- Local-only noise: not applicable
- Candidate inline module blocks: none discovered in changed files
- Uncertainty: none
```

Step 5 initializes or updates `.template-sync/marker.yml` after the range mode is chosen. Step 13 later advances `template_sync.last_reviewed_template_commit` to the resolved upstream range head SHA only after the relevant range or reconciliation has actually been reviewed.

## Step 5: Initialize or Update the Sync Marker

Downstream repositories SHOULD keep the sync marker at `.template-sync/marker.yml`, matching the marker path rule in Step 4. The marker distinguishes reviewed upstream changes from adopted upstream changes. Selective syncs may intentionally skip upstream files, so the preferred field under `template_sync` is `last_reviewed_template_commit`, not `last_adopted_template_commit`.

Marker contents are schema-backed by [`schemas/template-sync-marker.schema.json`](schemas/template-sync-marker.schema.json). The `validate-template-sync-marker` pre-commit hook validates `.template-sync/marker.yml` when that file is present; repositories without a marker are unaffected because no file matches the hook's anchored pattern. Marker changes MUST be rejected when they fail the schema. The schema's `included_modules` enum mirrors `.template-sync/manifest.yml`, and [`tests/test_template_manifest.py`](tests/test_template_manifest.py) fails if the schema enum drifts from the manifest module list.

The marker may record sync-specific first-adoption state, such as adopted modules, path-specific local overrides, downstream local path ownership, protected-file decisions, explicit `tailored` opt-ins for protected paths, deferred protected-file candidates, and instruction-contract waivers. It is not a general replacement for `_TODO-repo-init.md` when manual GitHub settings or maintainer policy decisions still need explicit resolution.

For example, suppose upstream changed `README.md` and `.github/workflows/terraform-ci.yml`, and the downstream repository reviewed both but adopted neither because `README.md` is locally owned and Terraform is not adopted. The sync still advances `last_reviewed_template_commit` to the resolved range head after review, because those upstream changes were inspected and intentionally skipped. A `last_adopted_template_commit` field would incorrectly imply that skipped-but-reviewed changes need to be reviewed again during the next sync.

If Step 4 used a first-sync delta range because the marker was missing or incomplete, initialize `.template-sync/marker.yml` in this step. Set `template_sync.last_reviewed_template_commit` to the resolved Step 4 range base SHA until Step 13 advances it to the resolved upstream range head SHA. Carry the range-base rationale from the sync working notes into the final sync summary.

If Step 4 selected full reconciliation, the marker still has no reviewed upstream commit at this step. You MAY initialize or update other marker fields, such as `source_repo`, `included_modules`, and local overrides chosen by the owner, but do not set `template_sync.last_reviewed_template_commit` until Step 13 records the resolved upstream range head SHA after review is complete.

`local_overrides`, `local_path_ownership`, `protected_file_decisions`, `deferred_protected_candidates`, and `instruction_contract_waivers` are explained immediately after the example.

Example marker:

```yaml
template_sync:
  source_repo: https://github.com/franklesniak/copilot-repo-template.git
  last_reviewed_template_commit: "1111111111111111111111111111111111111111"
  included_modules:
    - baseline
    - agent-instructions
    - github-actions
    - github-platform
    - github-templates
    - markdown
    - powershell
    - json
    - yaml
    - template-sync-support
  local_overrides:
    - path: README.md
      reason: "Project-specific; use template only as reference."
      default_decision: SKIP
    - path: .github/ISSUE_TEMPLATE/config.yml
      reason: "Repository-specific contact links."
      default_decision: MERGE
  local_path_ownership:
    - path: docs/
      reason: "Project documentation is downstream-owned."
    - path: corpus/
      reason: "Evaluation corpus is downstream application data."
    - path: schemas/local-project.schema.json
      reason: "Project-specific schema lives beside retained template schemas."
      overlap_exception_reason: "This exact local schema is not an upstream template file."
  deferred_protected_candidates:
    - path: .github/copilot-instructions.md
      source_commit: "dddddddddddddddddddddddddddddddddddddddd"
      reason: "Adds a stack-selection clause; pending owner authorization."
  protected_file_decisions:
    - path: CLAUDE.md
      adoption_mode: minimal-preservation
      decision: MERGE
      authorization_basis: "Owner authorized protected instruction file edits on 2026-05-27."
      authorized_scope: "Remove unadopted stack references and fix broken links only."
    - path: GEMINI.md
      decision: REMOVE-LOCAL
      authorization_basis: "Owner explicitly authorized removing GEMINI.md in issue 123."
      authorized_scope: "GEMINI.md only; no other protected files."
      reason: "Gemini agent not used by this downstream repository."
  instruction_contract_waivers:
    - path: CLAUDE.md
      anchor: "## Handling Code Review Comments"
      reason: "Downstream owner replaced this protocol with an equivalent local process."
      authorization_basis: "Owner approved this CLAUDE.md anchor waiver on 2026-05-27."
```

### Marker Semantics

- `source_repo` is the upstream template repository under review.
- `last_reviewed_template_commit` is the resolved upstream template commit SHA whose changes were most recently reviewed, regardless of how many were adopted. It MUST NOT be a branch name, tag name, or other moving ref.
- `included_modules` is the adoption state. Anything not listed is not adopted.
- `local_overrides` changes the starting recommendation for a path, but it MUST NOT hide upstream activity from the sync.
- `local_path_ownership` records downstream-owned exact paths or directory-prefix path families without modifying upstream-owned manifest rows.
- `protected_file_decisions` records the current path-scoped protected-file authorization and decision. Content adoption decisions (`TAKE` and `MERGE`) require `adoption_mode`, `authorization_basis`, and `authorized_scope`. Broad rewrites require `adoption_mode: tailored` plus `tailored_authorization_basis`. Protected-file removals require `decision: REMOVE-LOCAL`, explicit removal authorization, `authorized_scope`, and a substantive `reason`.
- `deferred_protected_candidates` records protected-file updates that were reviewed but not applied because path-scoped owner authorization was absent.
- `instruction_contract_waivers` records explicit owner-approved waivers for missing required instruction-contract anchors. Each waiver MUST include `path`, `anchor`, `reason`, and `authorization_basis`.

### Instruction Contract Waivers

Use `template_sync.instruction_contract_waivers` only when a retained protected instruction file intentionally omits a required heading or phrase from `.template-sync/instruction-contracts.yml`.

Each waiver entry MUST include:

- `path`: the protected instruction file path.
- `anchor`: the exact required heading or phrase being waived.
- `reason`: why the retained downstream file intentionally omits that anchor.
- `authorization_basis`: the owner authorization for the waiver, named consistently with `protected_file_decisions[].authorization_basis`.

A valid waiver MAY allow `python .template-sync/scripts/validate_instruction_contracts.py --mode downstream` to exit successfully, but the validator MUST report `passed with waivers` and list each applied waiver. A waiver is not protected-file edit or removal authorization; edits and removals still require the matching `protected_file_decisions` records described below.

### Local Overrides

When a changed upstream path appears in `local_overrides` and every mapped module is in `included_modules`, start the per-file decision at the override's `default_decision`. Each applied override MUST still appear in the sync summary with a brief description of the upstream change.

The agent or maintainer does not decide that an upstream change is too minor to mention under an override. Listing every applied override is the mechanism that lets the owner notice stale overrides, security-sensitive changes, validation changes, or governance changes that should override the override.

Worked local-overrides mini scenario:

1. The marker has a `README.md` override with `default_decision: SKIP` because the downstream README is project-specific.
2. The reviewed upstream range changes `README.md` to add a security reporting note.
3. The per-file row starts with `SKIP` from the override, but the maintainer upgrades the decision to `MERGE` because the security note is relevant.
4. The sync summary still lists the applied override and the upstream change, for example: `README.md` defaulted to `SKIP`; upstream added security reporting guidance; final decision `MERGE`.

### Local Path Ownership

Use `template_sync.local_path_ownership` for downstream project paths that upstream cannot know in advance, such as `docs/`, `corpus/`, `.devcontainer/`, application source directories, or implementation-state directories. Local ownership lives only in the marker. Do not add downstream-local paths to `.template-sync/manifest.yml` and do not add rows to the generated module-definition or path-mapping tables for them.

Each record MUST include:

- `path`: the downstream-owned exact path or directory-prefix path family.
- `reason`: a concise ownership explanation that is durable without private context.
- `overlap_exception_reason`: optional; use only when the validator reports broad manifest-area proximity and the path is still downstream-owned.

Path syntax is intentionally narrow:

- `docs` means the exact path `docs`.
- `docs/` means the `docs/` directory family. Directory-family ownership covers both the directory path and descendants.
- `docs/**` is not valid local ownership syntax. Manifest-style glob semantics remain reserved for `.template-sync/manifest.yml`.

Local ownership records may name future paths whose final segment does not exist yet. Validators still reject unsafe lexical forms such as absolute paths, backslashes, `..` traversal, duplicate slashes, empty segments, and glob characters. Validators also reject existing symlink ancestors and surface Git-visible symlink paths that would resolve outside the repository root.

Duplicate normalized local ownership paths MUST be rejected. Parent and child records, such as `docs/` plus `docs/api/`, MAY coexist when a subtree needs a more specific reason. Matching uses the most-specific record for diagnostics and grouping.

Exact collisions with concrete upstream-owned manifest paths MUST fail validation and cannot be overridden with `overlap_exception_reason`. A local path inside a broad manifest-owned area, such as `schemas/local-project.schema.json` near `schemas/**`, MUST include `overlap_exception_reason`. That exception reason is permitted only for broad-area proximity, not for exact upstream-file collisions and not for unrelated local paths.

Local path ownership affects only unrecorded local-path diagnostics. It MUST NOT silence retained-template-file missing checks, excluded-module leftover checks, protected-file decision checks, downstream instruction-contract checks, or inline-block consistency checks. Use:

- `local_path_ownership` when the path family is downstream-owned project state.
- `local_overrides` when an upstream-changed template path needs a default sync decision such as `SKIP`, `MERGE`, or `DEFER`.
- `protected_file_decisions` or `deferred_protected_candidates` for protected instruction/governance files.

The marker validator discovers unrecorded local paths with `git ls-files --cached --others --exclude-standard`, so ignored build and cache files do not create ownership suggestions. Unrecorded local paths are warnings by default and include bounded, deterministic, copy-ready suggested records. Run `python .template-sync/scripts/validate_marker.py --require-marker --strict-local-path-ownership` only when the downstream repository intentionally wants unrecorded local paths to fail validation.

### Adoption Mode Records

Before editing protected files or template-derived governance, community, process, workflow, or collaboration files, record the mode for the affected file or file set:

- Record the default `minimal-preservation` mode in `_TODO-repo-init.md` when that checklist exists.
- For protected paths, record path-specific decisions in `.template-sync/marker.yml` as `protected_file_decisions` before editing or removing the file.
- For non-protected paths, record path-specific `tailored` opt-ins or required local ownership exceptions in `.template-sync/marker.yml` as `local_overrides`, with a reason that names the selected mode and the local policy being preserved.
- Record the mode in the sync working notes and final sync summary when the sync does not modify `_TODO-repo-init.md` or `.template-sync/marker.yml`.

Do not add ad hoc, non-schema fields to `.template-sync/marker.yml`. Use the schema-backed marker fields above, or record the mode in the checklist, working notes, or sync summary.

### Protected File Decisions

Use `template_sync.protected_file_decisions` for every protected file that is edited, merged, taken, skipped, removed, deferred, or sent through protected review during first adoption or later sync. Valid `decision` values are `TAKE`, `MERGE`, `SKIP`, `REMOVE-LOCAL`, `DEFER`, and `PROTECTED-REVIEW`.

The retained `.template-sync/instruction-contracts.yml` catalog requires its own explicit decision before initial creation or later replacement. An authorized `TAKE` copies the pinned reviewed catalog intact; a `MERGE` requires authorization for the specific obligation changes. Do not infer catalog authority from an agent-platform bundle, coordinated instruction edits, or validation repair. Keep the catalog across support-retaining profiles and let `requires_modules` determine applicability; do not prune obligations for excluded agents. A candidate-local validator is a drift check, not independent evidence of human authorization. When `template-sync-support` is excluded, no catalog decision or catalog output is required.

- `TAKE` and `MERGE` MUST include `adoption_mode`, `authorization_basis`, and `authorized_scope`.
- `adoption_mode: tailored` MUST include `tailored_authorization_basis`; use it only for authorized broad rewrites of the named protected path or path set.
- `SKIP`, `DEFER`, and `PROTECTED-REVIEW` MUST include `reason` and MUST NOT imply edit authorization.
- `REMOVE-LOCAL` MUST include `authorization_basis`, `authorized_scope`, and `reason`. The `authorization_basis` MUST explicitly name the removed protected file and mention removal, deletion, or equivalent removal-specific wording. Generic wording such as "protected file edits authorized" is insufficient for removal.

Path overlap with `local_overrides` is allowed only when the two records are compatible. The marker validator reports compatible overlaps side by side and fails contradictory same-path decisions. A protected path MUST NOT appear in both `protected_file_decisions` and `deferred_protected_candidates` at the same time, because one record asserts a current decision while the other asserts the change is awaiting authorization.

### Downstream adoption: Dependabot schema regression surface

Downstream repositories that adopt both `github-platform` and `schema` SHOULD adopt [`tests/test_dependabot_schema.py`](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/tests/test_dependabot_schema.py), [`tests/fixtures/dependabot/auto-assignment.yml`](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/tests/fixtures/dependabot/auto-assignment.yml), and the related Dependabot validation hooks (`validate-dependabot-config` and `validate-dependabot-config-valid-examples`) when their live `.github/dependabot.yml` stays within the pinned `vendor.dependabot` schema surface. This keeps the documented Dependabot auto-assignment guidance under test at the hook's pinned `check-jsonschema` `rev` and under pytest; the pytest `check-jsonschema` dev dependency and the hook `rev` are not required to match.

A downstream repository MAY skip this regression surface when its live `.github/dependabot.yml` intentionally uses GitHub-supported fields rejected by the pinned built-in `vendor.dependabot` schema, or when the repository does not adopt Dependabot platform configuration. Record that skip explicitly: prefer path-scoped `local_overrides` entries in `.template-sync/marker.yml` for the fixture, test, and related Dependabot validation paths, each with the skip rationale; alternatively, record an explicit validation-policy decision in the sync summary that names the skipped paths and explains why the repository chose not to retain the pinned vendor-schema regression surface.

### Deferred Protected Candidates

Protected-file candidates remain in `deferred_protected_candidates` until the owner applies them through an authorized PR or explicitly dismisses them. Each entry includes:

- `path`
- `source_commit`
- `reason`

If the same protected path changes again upstream before the candidate is resolved, update the existing entry to the latest relevant source commit and preserve the prior rationale. Add a short refresh note to `reason` when the newer upstream change materially changes what is deferred. Add a separate entry only when the same path has distinct deferred candidates that cannot be represented clearly by one entry.

To explicitly dismiss a deferred protected candidate:

1. Remove that candidate from `.template-sync/marker.yml`.
2. Record the dismissed candidate's `path`, `source_commit`, and dismissal rationale in the sync PR description or a linked owner comment.
3. Allow later upstream changes to the same protected path to surface a new candidate during a future sync.

## Step 6: Use the Authoritative Module Taxonomy

The machine-readable `.template-sync/manifest.yml` file is authoritative for module definitions, path mappings, and filtering semantics. The tables in this section are a rendered human-readable view of that manifest, and [`tests/test_template_manifest.py`](tests/test_template_manifest.py) checks that the tables do not drift from the manifest.

When changing the taxonomy, update `.template-sync/manifest.yml` first, then update or regenerate the rendered tables below.

Select `agent-instructions` for shared guidance and each desired agent explicitly through `agent-copilot`, `agent-codex`, `agent-claude`, `agent-cursor`, `agent-gemini`, or `agent-hermes`. Agent modules require `agent-instructions`. A legacy selection must explicitly identify retained agents before reviewed cleanup; the new taxonomy does not authorize silent removal of existing agent files.

`instruction-enforcement` additionally requires `agent-instructions` and at least one execution route: `baseline`, `github-actions`, or `azure-pipelines`. Baseline retains the local pre-commit hook; each selected host retains its dedicated instruction-validation route, including when baseline is omitted. The Python project module is not required. Omitting enforcement is an explicit policy-only choice. The shared validation core remains available to the marker adapter when sync support is retained; this does not implicitly select standalone enforcement. See the [upstream instruction-enforcement guide](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/docs/instruction-enforcement.md) for mode selection, exact exceptions and migration.

The Copilot setup workflow requires `agent-instructions`, `agent-copilot` and `github-actions`; the local Claude review command requires `agent-instructions`, `agent-claude` and `github-actions`. Removing an agent or host uses ordinary reviewed cleanup for the file and its conditioned references. Setup completion does not replace the validation gates, and a branch test does not establish default-branch Copilot activation.

### Module Definitions

| Module | Scope |
| --- | --- |
| `instruction-enforcement` | Explicit optional instruction-enforcement selection. |
| `agent-copilot` | Explicit optional agent-copilot selection. |
| `agent-codex` | Explicit optional agent-codex selection. |
| `agent-claude` | Explicit optional agent-claude selection. |
| `agent-cursor` | Explicit optional agent-cursor selection. |
| `agent-gemini` | Explicit optional agent-gemini selection. |
| `agent-hermes` | Explicit optional agent-hermes selection. |
| `baseline` | Core repository scaffolding, community files, starter identity files, and repository-level configuration not owned by a narrower module. |
| `git-lfs` | Opt-in Git LFS-managed attributes for selected opaque authoring and project formats in the baseline root `.gitattributes`. |
| `agent-instructions` | Agent entry points, Copilot instructions, Cursor rules, reusable prompt guidance, and modular instruction docs. |
| `github-platform` | Repository-level GitHub platform configuration that is not itself an issue template, PR template, CODEOWNERS file, or workflow. Includes Dependabot configuration, repository funding metadata, security-advisory configuration, code-scanning configuration outside of workflows, and similar repository-scope GitHub-only settings. Current example: `.github/dependabot.yml`. Likely future inhabitants include `.github/FUNDING.yml`, `.github/security/*`, and other repository-scope GitHub configuration files. |
| `azure-devops-platform` | Azure DevOps Services project, security, and service configuration guidance that is not itself an Azure Pipelines file or Azure Repos collaboration template. |
| `github-actions` | GitHub Actions workflow files under `.github/workflows/**`. |
| `azure-pipelines` | Azure Pipelines CI assets and setup guidance under `.azuredevops/pipelines/**`. |
| `github-templates` | GitHub issue templates, PR templates, CODEOWNERS, and GitHub collaboration surfaces. |
| `azure-devops-collaboration` | Azure Repos PR templates, Azure Boards intake guidance, reviewer policies, and Azure DevOps Services security intake guidance. |
| `template-onboarding` | Template adoption and template maintainer guidance that downstream repositories typically remove after adoption. |
| `template-sync-support` | Committed files used to perform future template syncs, such as the sync procedure, sync marker, sync manifest, sync validation helper scripts, and future sync validation docs. |
| `markdown` | Markdown linting, Markdown templates, docs guidance, and Markdown-only documentation assets. |
| `powershell` | PowerShell scripts, Pester tests, PSScriptAnalyzer configuration, and PowerShell CI. |
| `json` | JSON and JSONC guidance, examples, validation commands, and JSON-oriented template files. |
| `yaml` | YAML guidance, YAML template files, and YAML validation not owned by a narrower module. |
| `schema` | JSON Schema contracts, schema examples, schema validation docs, and schema-specific tests. |
| `python` | Python package scaffolding, Python tests, Python CI, and Python tooling configuration. |
| `terraform` | Terraform modules, tests, linting, documentation, and template files. |

### Path Mapping

Apply the most specific matching row. The most-specific match wins: when an exact path or a narrower glob row covers a path, broader catch-all rows do not contribute additional modules to that path. If two rows match at the same specificity level, use the de-duplicated union of their relation modules, preserving the relation kind.

Manifest version 1 rows use `requires_all` only: the path is included in the per-file review only when every listed module appears in `included_modules`.

Manifest version 2 and version 3 rows MAY also use `requires_any`: the path is included only when every `requires_all` module appears in `included_modules` and, when `requires_any` is present, at least one `requires_any` module also appears in `included_modules`. A row with `requires_all` plus `requires_any` therefore means "all of these modules, plus at least one of these alternatives." Manifest version 3 additionally records compatibility groups for related host-family modules.

| Path pattern | Module(s) |
| --- | --- |
| `.claude/commands/review-loop.md` | `agent-instructions`, `agent-claude`, `github-actions` |
| `.github/workflows/copilot-setup-steps.yml`, `docs/copilot-setup.md` | `agent-instructions`, `agent-copilot`, `github-actions` |
| `tests/test_copilot_setup.py` | `template-sync-support`, `github-actions` |
| `.github/scripts/instruction_contract_core.py` | one of `instruction-enforcement`, `template-sync-support` |
| `.github/scripts/instruction_contract_support.py` | one of `instruction-enforcement`, `template-sync-support`, `github-actions` |
| `.github/scripts/validate_instruction_profile.py`, `.github/instruction-profile.yml`, `.github/instruction-contracts.yml` | `agent-instructions`, `instruction-enforcement` |
| `schemas/instruction-profile.schema.json`, `schemas/instruction-contracts.schema.json` | one of `instruction-enforcement`, `template-sync-support` |
| `schemas/examples/instruction-contracts/**`, `schemas/examples/instruction-profile/**`, `docs/instruction-enforcement.md` | `agent-instructions`, `instruction-enforcement` |
| `tests/test_instruction_profile.py`, `tests/test_instruction_profile_output.py`, `tests/test_instruction_profile_references.py`, `tests/test_instruction_yaml_duplicates.py` | `template-sync-support` |
| `.github/workflows/instruction-contracts.yml` | `agent-instructions`, `instruction-enforcement`, `github-actions` |
| `.azuredevops/pipelines/instruction-contracts.yml` | `agent-instructions`, `instruction-enforcement`, `azure-pipelines` |
| `.template-sync/marker.yml`, `.template-sync/manifest.yml`, `.template-sync/instruction-contracts.yml`, `.template-sync/first-adoption/**`, `.template-sync/scripts/**`, `templates/adoption/**` | `template-sync-support` |
| `.github/copilot-instructions.md` | `agent-instructions` |
| `.github/instructions/docs.instructions.md` | `markdown`, `agent-instructions` |
| `.github/instructions/gitattributes.instructions.md` | `baseline`, `agent-instructions` |
| `.github/instructions/json.instructions.md` | `json`, `agent-instructions` |
| `.github/instructions/powershell.instructions.md` | `powershell`, `agent-instructions` |
| `.github/instructions/python.instructions.md` | `python`, `agent-instructions` |
| `.github/instructions/terraform.instructions.md` | `terraform`, `agent-instructions` |
| `.github/instructions/yaml.instructions.md` | `yaml`, `agent-instructions` |
| `.github/instructions/*.instructions.md` | `agent-instructions` |
| `.cursor/rules/**` | `agent-instructions`, `agent-cursor` |
| `.hermes.md` | `agent-instructions`, `agent-hermes` |
| `AGENTS.md` | `agent-instructions`, `agent-codex` |
| `CLAUDE.md` | `agent-instructions`, `agent-claude` |
| `GEMINI.md` | `agent-instructions`, `agent-gemini` |
| `COPILOT_CHAT_PROMPTS.md`, `docs/ISSUE_EVALUATION_PROMPT.md`, `docs/PR_REVIEW_PROMPTS.md` | `agent-instructions` |
| `.codex/**` | `agent-instructions`, `agent-codex` |
| `.claude/**` | `agent-instructions`, `agent-claude` |
| `.github/ISSUE_TEMPLATE/**`, `.github/pull_request_template.md`, `.github/CODEOWNERS` | `github-templates` |
| `.azuredevops/pull_request_template.*`, `.azuredevops/pull_request_template/**` | `azure-devops-collaboration` |
| `.github/dependabot.yml` | `github-platform` |
| `.azuredevops/platform/**` | `azure-devops-platform` |
| `docs/azure-devops-support.md` | one of `azure-devops-platform`, `azure-pipelines`, `azure-devops-collaboration` |
| `tests/test_dependabot_schema.py`, `tests/fixtures/dependabot/auto-assignment.yml` | `github-platform`, `schema` |
| `docs/workflow-security.md` | `github-actions` |
| `docs/upstream-style-guides.md` | `agent-instructions` plus one of `powershell`, `terraform` |
| `.github/workflow-security-contract.yml` | `github-actions` |
| `.github/scripts/validate_workflow_security.py`, `schemas/workflow-security-contract.schema.json` | one of `github-actions`, `template-sync-support` |
| `.github/workflows/workflow-security.yml` | `github-actions` |
| `schemas/examples/workflow-security-contract/**`, `tests/test_workflow_security_contract.py` | `github-actions`, `template-sync-support` |
| `tests/test_contract_wiring.py` | `template-sync-support` |
| `tests/test_workflow_security_lifecycle.py` | `github-actions`, `template-sync-support` |
| `.github/workflows/markdownlint.yml`, `.github/workflows/toolchain-eol.yml` | `markdown`, `github-actions` |
| `.github/workflows/powershell-ci.yml` | `powershell`, `github-actions` |
| `.github/workflows/python-ci.yml` | `python`, `github-actions` |
| `.github/workflows/precommit-ci.yml` | `baseline`, `github-actions` |
| `.github/workflows/terraform-ci.yml` | `terraform`, `github-actions` |
| `.github/workflows/data-ci.yml` | `baseline`, `github-actions` |
| `.github/template-placeholders.json` | one of `baseline`, `template-sync-support` |
| `.github/workflows/check-placeholders.yml` | `baseline`, `github-actions` |
| `.github/scripts/replace-template-placeholders.py` | one of `baseline`, `template-sync-support` |
| `.github/scripts/validate-placeholder-schema-examples.py` | `baseline` |
| `.github/workflows/auto-fix-precommit.yml` | `baseline`, `github-actions` |
| `.azuredevops/pipelines/README.md` | `azure-pipelines` |
| `.azuredevops/pipelines/precommit.yml`, `.azuredevops/pipelines/check-placeholders.yml` | `baseline`, `azure-pipelines` |
| `.azuredevops/pipelines/markdownlint.yml` | `markdown`, `azure-pipelines` |
| `.azuredevops/pipelines/powershell-ci.yml` | `powershell`, `azure-pipelines` |
| `.azuredevops/pipelines/python-ci.yml` | `python`, `azure-pipelines` |
| `.azuredevops/pipelines/terraform-ci.yml` | `terraform`, `azure-pipelines` |
| `.azuredevops/pipelines/data-ci.yml` | `baseline`, `azure-pipelines` |
| `.azuredevops/pipelines/**` | `azure-pipelines` |
| `.yamllint.yml` | `yaml` |
| `.pre-commit-config.yaml`, `requirements-pre-commit.txt` | `baseline` |
| `.markdownlint.jsonc`, `.remarkignore`, `.remarkrc.mjs`, `package.json`, `package-lock.json`, `.github/scripts/lint-nested-markdown.js`, `.github/scripts/lint-nested-markdown.test.js` | `markdown` |
| `.github/scripts/lint_nested_markdown_hook.py` | `baseline`, `markdown` |
| `.github/scripts/check-toolchain-eol.js`, `.github/scripts/check-prohibited-placeholders.py` | `markdown` |
| `tests/test_replace_template_placeholders.py` | `baseline` |
| `tests/test_check_prohibited_placeholders.py`, `tests/toolchain-eol/check-toolchain-eol.test.js`, `tests/toolchain-eol/fixtures/**`, `templates/markdown/**` | `markdown` |
| `templates/powershell/**`, `tests/PowerShell/**`, `.github/linting/PSScriptAnalyzerSettings.psd1`, `src/tools/*.ps1` | `powershell` |
| `templates/json/**` | `json` |
| `templates/yaml/**` | `yaml` |
| `schemas/template-sync-manifest.schema.json`, `schemas/template-sync-marker.schema.json`, `schemas/template-sync-instruction-contracts.schema.json`, `schemas/first-adoption-quality-suppressions.schema.json` | `template-sync-support` |
| `schemas/template-placeholders.schema.json` | one of `baseline`, `template-sync-support` |
| `schemas/examples/template-sync-marker/**`, `schemas/examples/template-sync-instruction-contracts/**`, `schemas/examples/first-adoption-quality-suppressions/**` | `template-sync-support` |
| `schemas/examples/template-placeholders/**` | `baseline` |
| `schemas/**` | `schema` |
| `tests/test_schema_examples.py` | one of `schema`, `template-sync-support` |
| `tests/test_generate_sync_candidates.py`, `tests/test_bootstrap_first_adoption.py`, `tests/test_first_adoption_state.*`, `tests/test_initialize_adoption_journal.*`, `tests/test_report_excluded_module_references.py`, `tests/test_materialize_downstream_adoption.*`, `tests/test_run_first_adoption_checks.*`, `tests/test_first_adoption_quality_reports.*`, `tests/test_template_manifest.py`, `tests/test_template_sync_materialization_helpers.py`, `tests/test_validate_marker.py`, `tests/test_validate_downstream_adoption.py`, `tests/test_validate_instruction_contracts.py`, `tests/test_instruction_input_limits.py`, `tests/test_precommit_runner.py` | `template-sync-support` |
| `.github/scripts/terraform_hooks.py`, `tests/test_terraform_hooks.py` | `terraform` |
| `templates/python/**`, `pyproject.toml`, `pyrightconfig.json`, `src/copilot_repo_template/**`, `tests/*.py`, `tests/**/*.py` | `python` |
| `templates/terraform/**`, `docs/terraform/**`, `modules/**`, `tests/**/*.tftest.hcl`, `.tflint.hcl`, `*.tf`, `*.tfvars`, `*.tftpl`, `*.tfbackend` | `terraform` |
| `README.md` | `baseline` |
| `TEMPLATE_UPDATE_PROCEDURE.md` | `template-sync-support` |
| `GETTING_STARTED_NEW_REPO.md`, `GETTING_STARTED_EXISTING_REPO.md`, `OPTIONAL_CONFIGURATIONS.md`, `TEMPLATE_MAINTENANCE.md`, `.github/TEMPLATE_DESIGN_DECISIONS.md` | `template-onboarding` |
| `CONTRIBUTING.md`, `SECURITY.md`, `CODE_OF_CONDUCT.md`, `LICENSE`, `.gitignore`, `.gitattributes`, `.vscode/**` | `baseline` |

### Manifest Version Migration

Version 1 manifests remain valid for downstream repositories that have not adopted version 2 or version 3. Treat every version 1 mapping as a `requires_all` mapping and ignore `requires_any`, because version 1 does not allow that field. Version 3 keeps the version 2 path-mapping semantics and adds `compatibility_groups`.

To migrate a manifest from version 1 to version 2:

1. Change `template_manifest.version` from `1` to `2`.
2. Leave existing `requires_all` rows unchanged unless a path truly has cross-module alternatives.
3. Add `requires_any` only for rows where a containing file should be retained when at least one alternative module is adopted. For example, `tests/test_schema_examples.py` uses `requires_any: [schema, template-sync-support]`. Data CI instead requires both `baseline` and its matching host module through `requires_all`.
4. Add `filtering.requires_any_semantics: OR`.
5. Remove any `known_limitations` entry whose only purpose was to record a cross-module relation that is now expressed by `requires_any`.
6. Validate the manifest with `pre-commit run validate-template-sync-manifest --all-files` and run [`tests/test_template_manifest.py`](tests/test_template_manifest.py).

If a changed upstream path does not match the table, classify it as `UNMAPPED` in the sync working notes and ask the owner to assign a module before deciding whether to include it.

Retained Markdown documents, meaning `.md` or `.mdc` paths whose most-specific manifest mapping does not include `template-onboarding`, MUST NOT use relative Markdown links to `template-onboarding` files. Use an absolute upstream URL of the form `https://github.com/franklesniak/copilot-repo-template/blob/HEAD/<path>` when the reference remains useful after downstream onboarding files are removed, or use neutral wording with no link when the retained document does not need to navigate to the onboarding file.

### Inline Module Blocks

Some retained files contain module-owned blocks delimited by YAML comments or Markdown-safe HTML comments. The `*-only` family is for configuration or workflow blocks owned by an optional module. The `*-reference-only` family is for textual references to an optional module inside a retained protected or shared document. Both families use the marker's registered AND or OR retention relation, described below. If `included_modules` does not satisfy that relation, remove the complete block including the begin and end marker lines.

The current `*-only` marker forms are:

```yaml
# template-sync: begin terraform-only
# ...
# template-sync: end terraform-only

# template-sync: begin python-only
# ...
# template-sync: end python-only

# template-sync: begin markdown-only
# ...
# template-sync: end markdown-only

# template-sync: begin yaml-only
# ...
# template-sync: end yaml-only

# template-sync: begin schema-only
# ...
# template-sync: end schema-only

# template-sync: begin template-sync-support-only
# ...
# template-sync: end template-sync-support-only

# template-sync: begin github-actions-only
# ...
# template-sync: end github-actions-only

# template-sync: begin github-platform-only
# ...
# template-sync: end github-platform-only

# template-sync: begin schema-template-sync-support-only
# ...
# template-sync: end schema-template-sync-support-only

# template-sync: begin git-lfs-only
# ...
# template-sync: end git-lfs-only
```

The current `*-reference-only` marker forms are Markdown-safe HTML comments:

```markdown
<!-- template-sync: begin markdown-reference-only -->
...
<!-- template-sync: end markdown-reference-only -->

<!-- template-sync: begin powershell-reference-only -->
...
<!-- template-sync: end powershell-reference-only -->

<!-- template-sync: begin python-reference-only -->
...
<!-- template-sync: end python-reference-only -->

<!-- template-sync: begin terraform-reference-only -->
...
<!-- template-sync: end terraform-reference-only -->

<!-- template-sync: begin json-reference-only -->
...
<!-- template-sync: end json-reference-only -->

<!-- template-sync: begin yaml-reference-only -->
...
<!-- template-sync: end yaml-reference-only -->

<!-- template-sync: begin schema-reference-only -->
...
<!-- template-sync: end schema-reference-only -->

<!-- template-sync: begin template-sync-support-reference-only -->
...
<!-- template-sync: end template-sync-support-reference-only -->

<!-- template-sync: begin azure-devops-guide-reference-only -->
...
<!-- template-sync: end azure-devops-guide-reference-only -->

<!-- template-sync: begin github-actions-reference-only -->
...
<!-- template-sync: end github-actions-reference-only -->

<!-- template-sync: begin github-platform-reference-only -->
...
<!-- template-sync: end github-platform-reference-only -->
```

Most markers in both families use AND-retention; a single-module marker requires that module. The `schema-template-sync-support-only` marker requires both `schema` and `template-sync-support`. The `github-data-ci-reference-only` marker requires both `baseline` and `github-actions`, matching the data workflow's dependencies. Its delimiters are `<!-- template-sync: begin github-data-ci-reference-only -->` and `<!-- template-sync: end github-data-ci-reference-only -->`. The `baseline-reference-only` and `baseline-only` markers require `baseline`; the reference delimiters are `<!-- template-sync: begin baseline-reference-only -->` and `<!-- template-sync: end baseline-reference-only -->`.

The `instruction-enforcement-only` and `instruction-enforcement-reference-only` markers require both `agent-instructions` and `instruction-enforcement`. The former guards the local `validate-instruction-profile` hook inside the baseline-owned pre-commit configuration; the latter guards optional enforcement guidance in retained documents. The `copilot-setup-reference-only` marker requires `agent-instructions`, `agent-copilot` and `github-actions`. The `claude-review-command-reference-only` marker requires `agent-instructions`, `agent-claude` and `github-actions` and guards the command section in `docs/PR_REVIEW_PROMPTS.md`. All use AND-retention.

Two marker families use OR-retention: `azure-devops-guide-reference-only` requires any of `azure-devops-platform`, `azure-pipelines`, or `azure-devops-collaboration`; `pip-dependencies-only` requires `baseline` or `python`. Each is retained when at least one named module is adopted and stripped only when all are excluded.

These inline blocks let a downstream repository keep the containing baseline or cross-module file while removing toolchain assumptions for a module it did not adopt. During Step 6, after path mapping decides whether the containing file itself is in scope, apply these rules:

1. For an AND-retention marker, retain its blocks unchanged when every named module is present in `included_modules`, unless the per-file review records a separate `MERGE` decision. Otherwise remove each complete block, including its begin and end marker lines, before accepting or merging the containing file.
2. For an OR-retention marker, retain its blocks unchanged when at least one named module is present in `included_modules`, unless the per-file review records a separate `MERGE` decision. If none is present, remove each complete block, including its begin and end marker lines, before accepting or merging the containing file.
3. Treat unmatched, nested, or unknown inline-block markers as an explicit sync question for the owner; do not silently keep or drop the affected block.

For a single-module AND marker such as `python-reference-only`, retain the block when its module is present and strip it when that module is absent. For multi-module markers, use the registered AND or OR relation described above.

When an inline module block from either the `*-only` or `*-reference-only` family splits or participates in a retained sequentially numbered Markdown structure, authors **MUST NOT** leave a numbering gap in the post-strip Markdown source or rendered output. Running numbers **MUST** stay contiguous in the full template and after the block is stripped during partial adoption. When multiple inline blocks appear near the same numbered sequence, every supported retain/strip combination **MUST** leave that sequence contiguous. This rule applies to numbered Markdown headings, such as `### 3. ...`, and ordered-list items, such as `1.`, `2.`, and `3.`. A post-strip gap in an ordered list also fails markdownlint MD029 because `.markdownlint.jsonc` configures MD029 with `"style": "ordered"` when the stripped output is linted; a numbered-heading gap is a readability defect rather than a markdownlint numbering failure. A complete independent numbered sequence wholly inside an inline block is acceptable when stripping the block leaves no retained numbering gap.

Recommended pattern: keep optional, stack-specific steps unnumbered, such as `### Optional: Install {module} tooling`, and reserve running numbers for steps that are always retained.

**MUST NOT** wrap a running numbered step in an optional inline block:

```markdown
### 1. Clone the repository

### 2. Install Node.js dependencies

<!-- template-sync: begin {module}-reference-only -->

### 3. Install {module} tooling

<!-- template-sync: end {module}-reference-only -->

### 4. Install pre-commit
```

After the `{module}` block is stripped, the retained headings read 1, 2, 4, leaving a numbering gap.

**MUST** keep the optional step unnumbered and keep the retained sequence contiguous:

```markdown
### 1. Clone the repository

### 2. Install Node.js dependencies

<!-- template-sync: begin {module}-reference-only -->

### Optional: Install {module} tooling

<!-- template-sync: end {module}-reference-only -->

### 3. Install pre-commit
```

The numbered headings are always-retained steps, and the optional `{module}` step is unnumbered, so stripping the block leaves the sequence contiguous. The `{module}-reference-only` marker family in these examples is inert because it does not match the template-sync marker parser.

The current `markdown-reference-only`, `powershell-reference-only`, `python-reference-only`, `terraform-reference-only`, `json-reference-only`, `yaml-reference-only`, and `schema-reference-only` inline blocks live in:

- `.github/copilot-instructions.md` for removable optional-stack references in protected canonical guidance, including instruction-file rows, validation examples, module-owned validator prose, and workflow tool-version examples.
- `.cursor/rules/repository-instructions.mdc`, `.hermes.md`, `AGENTS.md`, `CLAUDE.md`, and `GEMINI.md` for removable optional-stack references in protected agent entry-point summaries.
- `README.md` and `CONTRIBUTING.md` for removable optional-stack references in shared baseline contributor-facing documentation.
- `.github/pull_request_template.md` for removable Python, PowerShell, and schema checklist sections in the retained PR template.
- `.github/instructions/json.instructions.md`, `.github/instructions/yaml.instructions.md`, `schemas/README.md`, `GETTING_STARTED_NEW_REPO.md`, and `GETTING_STARTED_EXISTING_REPO.md` for optional companion-guide links.
- `OPTIONAL_CONFIGURATIONS.md` for optional Python guidance and YAML starter links.
- `docs/upstream-style-guides.md` for the provenance record of each retained imported PowerShell or Terraform instruction guide.
- `docs/PR_REVIEW_PROMPTS.md` for the PowerShell version-support prompt.

The current `github-actions-reference-only` inline blocks live in:

- `README.md` for GitHub Actions-specific `actionlint` validation prose.
- `AGENTS.md`, `CLAUDE.md`, and `docs/PR_REVIEW_PROMPTS.md` for the shared Copilot review recipe and its entry-point references.
- `.github/pull_request_template.md` for the GitHub Actions checklist section in the retained PR template.

The current `github-platform-reference-only` inline blocks live in:

- `README.md` for the GitHub Dependabot key-file rows.
- `OPTIONAL_CONFIGURATIONS.md` for GitHub Dependabot optional configuration guidance.
- `schemas/README.md` for GitHub Dependabot built-in schema validation guidance.

The current `template-sync-support-reference-only` inline blocks live in:

- `README.md` for the optional `.template-sync/` and `schemas/template-sync-*.schema.json` surface rows, which are removed when `template-sync-support` is excluded.
- `CONTRIBUTING.md` for the Template-Sync Validation section, which is removed when `template-sync-support` is excluded.

The current `github-data-ci-reference-only` inline blocks live in:

- `CONTRIBUTING.md` for the Data CI workflow row and `README.md` for its key-file bullet.
- `.github/copilot-instructions.md`, `AGENTS.md`, `CLAUDE.md`, `.hermes.md`, `GEMINI.md`, and `.cursor/rules/repository-instructions.mdc` for the data-CI summary.
- `GETTING_STARTED_NEW_REPO.md` for the data-CI validation bullet.

All these blocks require baseline and GitHub Actions. The schema inventory, JSON starter, and maintenance guide use upstream-template links for durable references to data CI that remain useful when the downstream workflow is absent.

The `baseline-reference-only` blocks in `.github/pull_request_template.md` guard the generic pre-commit verification and data-file checklists. Blocks in `.github/copilot-instructions.md` and all five platform entry points guard baseline configuration references and checklist commands. Blocks in the JSON and YAML guides, schema inventory, `docs/terraform/TERRAFORM_LINTING_GUIDE.md`, new-repository onboarding guide, and optional-configuration guide guard baseline-owned companion links, hook execution, and placeholder assets. Portable authoring requirements, including the JSON single-newline requirement, remain outside those blocks. The `baseline-only` blocks guard the Claude session hook's pre-commit provisioning and the Dependabot pre-commit ecosystem header/update block. The `pip-dependencies-only` blocks guard the Dependabot pip header/update block, which scans baseline runner requirements and any retained Python project metadata.

The current `azure-devops-guide-reference-only` inline blocks live in:

- `README.md`, `CONTRIBUTING.md`, `OPTIONAL_CONFIGURATIONS.md`, `COPILOT_CHAT_PROMPTS.md`, `docs/PR_REVIEW_PROMPTS.md`, and `schemas/README.md` for optional links to `docs/azure-devops-support.md`, retained when any of `azure-devops-platform`, `azure-pipelines`, or `azure-devops-collaboration` is adopted and removed only when all three are excluded.

The current `python-only` inline blocks live in:

- `.pre-commit-config.yaml` for the `black` and `ruff-check` Python project hooks.
- `.github/workflows/copilot-setup-steps.yml` for Python project prerequisites.

The current `markdown-only` inline blocks live in:

- `.pre-commit-config.yaml` for outer and nested Markdown lint and placeholder checks.
- `.github/workflows/precommit-ci.yml`, `.github/workflows/auto-fix-precommit.yml`, and `.azuredevops/pipelines/precommit.yml` for root Node dependency setup required by nested lint.
- `.claude/hooks/session-start.sh` for the locked root Node dependency installation in Claude Code web sessions.

The current `yaml-only` inline blocks live in:

- `.pre-commit-config.yaml` for the `yamllint` hook that depends on `.yamllint.yml`.
- `.github/workflows/data-ci.yml` for the `yamllint` hook-list documentation and the dedicated `Run yamllint` step.
- `.azuredevops/pipelines/data-ci.yml` for the `yamllint` hook-list documentation and the dedicated `Run yamllint` step.

The current `schema-only` inline blocks live in:

- `.pre-commit-config.yaml` for worked-example schema validators and the `check-metaschema` schema self-validation hook.
- `.github/workflows/data-ci.yml` for worked-example schema validation hook-list documentation and dedicated schema validation alias steps.
- `.azuredevops/pipelines/data-ci.yml` for worked-example schema validation hook-list documentation and dedicated schema validation alias steps.

The current `template-sync-support-only` inline blocks live in:

- `.pre-commit-config.yaml` for template sync schema example validators, first-adoption quality suppression example validators, runtime schema self-validation hooks, the `validate-template-sync-manifest`, `validate-template-sync-marker`, `validate-template-sync-instruction-contracts`, `validate-first-adoption-quality-suppressions`, `validate-instruction-contracts-upstream`, and `validate-instruction-contracts-downstream` hooks.
- `.github/workflows/markdownlint.yml` for source-template Python setup, generated-output nested-Markdown regression testing, actual nested-Markdown hook invocation testing, and their aggregate failure predicates.
- `.github/workflows/data-ci.yml` for template sync schema example validation, first-adoption quality suppression validation, runtime schema self-validation, live template sync validation hook-list documentation, and the dedicated template sync validation alias steps.
- `.azuredevops/pipelines/data-ci.yml` for template sync schema example validation, first-adoption quality suppression validation, runtime schema self-validation, live template sync validation hook-list documentation, and the dedicated template sync validation alias steps.

The current `github-platform-only` inline blocks live in:

- `.pre-commit-config.yaml` for the `validate-dependabot-config` and `validate-dependabot-config-valid-examples` hooks.
- `.github/workflows/data-ci.yml` for Dependabot validation hook-list documentation and the dedicated `Run validate-dependabot-config` and `Run validate-dependabot-config-valid-examples` steps.
- `.azuredevops/pipelines/data-ci.yml` for Dependabot validation hook-list documentation and the dedicated `Run validate-dependabot-config` and `Run validate-dependabot-config-valid-examples` steps.

The current `github-actions-only` inline block lives in:

- `.pre-commit-config.yaml` for the `actionlint` hook.

The `schema-template-sync-support-only` inline marker family is registered for future blocks that require both `schema` and `template-sync-support`; no checked-in block currently uses it.

The current `git-lfs-only` inline block lives in:

- `.gitattributes` for optional Git LFS-managed opaque authoring and project format patterns.

The Copilot setup workflow also uses `baseline-only`, `markdown-only`, `terraform-only`, and `powershell-only` blocks to install only retained prerequisites. Its baseline block reads the direct pin from `requirements-pre-commit.txt`, and its Markdown block uses the committed root lockfile with lifecycle scripts disabled.

The current `terraform-only` inline blocks live in:

- `.pre-commit-config.yaml` for the `terraform-fmt`, `terraform-validate`, and `terraform-tflint` repo-local hooks.
- `.github/workflows/precommit-ci.yml` for the Terraform and TFLint setup steps required only when those hooks are retained.
- `.github/workflows/auto-fix-precommit.yml` for the Terraform and TFLint setup steps required only when those hooks are retained.
- `.azuredevops/pipelines/precommit.yml` for the Terraform and TFLint setup steps required only when those hooks are retained.
- `.claude/hooks/session-start.sh` for the Terraform installation in Claude Code web sessions. This block stays last in the hook because its idempotency check exits the hook early.

After stripping `python-only` blocks, a downstream repository that keeps baseline but excludes `python` can run `pre-commit run --all-files` without Python project formatters or linters such as Black and Ruff. The tooling runtime and `requirements-pre-commit.txt` remain. If GitHub platform support is retained, the pip updater remains for that runner requirement; it disappears only when both baseline and Python are excluded.

The baseline module owns both `.pre-commit-config.yaml` and the exact runner requirement. Aggregate and data-file pre-commit CI require baseline plus the matching host. Selecting data modules without baseline keeps their content but supplies no template-managed pre-commit configuration or data CI. A no-baseline agent profile removes the pre-commit provisioning block from the Claude hook, even if the downstream repository owns a separate local config. The hook's Terraform installation follows its own `terraform-only` block.

After stripping `markdown-only` blocks, a downstream repository that excludes `markdown` should be able to run `pre-commit run --all-files` without installing Node.js or markdownlint.

After stripping `yaml-only` blocks, a downstream repository that excludes `yaml` should be able to run `pre-commit run --all-files` and the retained data-file workflow without retaining `.yamllint.yml` or invoking a missing `yamllint` hook.

After stripping `schema-only` blocks, a downstream repository that excludes `schema` should be able to run `pre-commit run --all-files` and the retained data-file workflow without retaining worked-example schema validators or the worked-example `check-metaschema` hook.

After stripping `template-sync-support-only` blocks, a downstream repository that excludes `template-sync-support` should be able to run `pre-commit run --all-files` and the retained data-file workflow without invoking template sync schema example validators, first-adoption quality suppression validators, runtime schema self-validation hooks, manifest validators, marker validators, or marker-backed instruction-contract validators. Independently selected `instruction-enforcement` retains its standalone profile check, local hook when baseline is selected, and dedicated selected-host routes; removing sync support does not remove that enforcement.

After stripping `github-platform-only` blocks, a downstream repository that excludes `github-platform` should be able to run `pre-commit run --all-files` and the retained data-file workflow without retaining Dependabot validation hooks or invoking a missing `validate-dependabot-config` hook.

After stripping `github-actions-only` blocks, a downstream repository that excludes `github-actions` should be able to run `pre-commit run --all-files` without retaining GitHub Actions-only `actionlint` hook installation or execution. Azure Pipelines validation is intentionally separate from this hook.

After stripping `git-lfs-only` blocks, a downstream repository that excludes `git-lfs` should receive the baseline root `.gitattributes` without any `filter=lfs`, `diff=lfs`, or `merge=lfs` attributes.

After stripping `terraform-only` blocks, a downstream repository that excludes `terraform` should be able to run `pre-commit run --all-files` and the retained non-Terraform workflows without installing HashiCorp Terraform or TFLint. The Claude session hook also stops downloading Terraform.

Record the resulting retain or strip decisions per affected path in the Step 7 per-file decision notes, and summarize them in the sync summary so the audit trail captures both the discovery from Step 4 and the resolution from Step 6.

### Filtering Rules

After Step 3 has fetched the template remote and Step 4 has resolved the range endpoints and confirmed reachability, and after Step 5 has a schema-valid marker with the intended `included_modules`, `local_overrides`, `protected_file_decisions`, and `deferred_protected_candidates`, run the candidate table generator from the downstream repository root:

```bash
python .template-sync/scripts/generate_sync_candidates.py --range-head RANGE_HEAD_SHA
```

The helper defaults `--range-base` to `template_sync.last_reviewed_template_commit` from `.template-sync/marker.yml`. For a first-sync delta review where the marker has no reviewed commit yet, pass the owner-approved initial range base explicitly:

```bash
python .template-sync/scripts/generate_sync_candidates.py --range-base RANGE_BASE_SHA --range-head RANGE_HEAD_SHA
```

If `--range-head` is omitted, the helper uses the local `template/main` ref when that ref is present. It does **not** fetch automatically; run `git fetch template` in Step 3 before relying on the default head, or pass a resolved `RANGE_HEAD_SHA`.

The generator validates `.template-sync/marker.yml` and `.template-sync/manifest.yml` against the checked-in schemas before producing output. Its report header prints the exact delta command it models, `git diff --name-status -M RANGE_BASE_SHA..RANGE_HEAD_SHA --`, so maintainers can cross-check the path list directly. It also compares local `TEMPLATE_UPDATE_PROCEDURE.md` with the upstream copy at the resolved range head; when they differ, the report prints a warning that local procedure text may be stale and includes `git show RANGE_HEAD_SHA:TEMPLATE_UPDATE_PROCEDURE.md` for reviewing the current upstream procedure. It prints a Markdown table with one row per changed upstream path and columns for the matched module relation, retained/excluded status, local override status, deferred protected candidate status, protected-file decision status, protected instruction/governance-file status, and notes. The notes explicitly surface unmapped paths, unknown modules, cross-module manifest relations, manifest inline-block notes, protected-file handling, protected decision records, and renames. The output is a decision aid only: it does not update the marker, apply file changes, strip inline blocks, or make final per-file decisions. The manual review process in this procedure remains authoritative.

Before Step 13 advances `template_sync.last_reviewed_template_commit`, save the exact Step 6 candidate table into the sync working notes or into an owner-chosen local file cited by those notes. To write only the rendered candidate table to a repository-contained file while preserving the full normal report on stdout, pass `--write-candidates`:

```bash
python .template-sync/scripts/generate_sync_candidates.py --range-head RANGE_HEAD_SHA --write-candidates SYNC_CANDIDATE_TABLE_PATH
```

The `SYNC_CANDIDATE_TABLE_PATH` value MUST remain inside the repository root; the helper rejects paths that escape it. Record both `RANGE_BASE_SHA` and `RANGE_HEAD_SHA` with the saved table. After Step 13 advances the marker, rerunning the helper without an explicit old `--range-base` correctly uses the new marker value and can produce an empty range. To rerun the helper for the same reviewed range after marker advancement, pass the recorded range explicitly, even if `template/main` has moved:

```bash
python .template-sync/scripts/generate_sync_candidates.py --range-base RANGE_BASE_SHA --range-head RANGE_HEAD_SHA
```

The saved Step 6 candidate table is the source of truth for the reviewed candidate set. A rerun for the same range uses the current `.template-sync/marker.yml` values for `included_modules`, `local_overrides`, `protected_file_decisions`, and `deferred_protected_candidates`, so its output can diverge from the saved table when Step 13 changed any of those fields even though the commit range is unchanged.

Create or refresh the adoption ledger during first sync and full-reconciliation syncs, and review it before editing protected files. To include the ledger in the normal delta report, pass `--ledger`; to write a generated snapshot that can be committed to the repository, pass `--write-ledger`:

```bash
python .template-sync/scripts/generate_sync_candidates.py --range-head RANGE_HEAD_SHA --ledger --write-ledger ADOPTION_LEDGER_PATH
```

For first-adoption or full-reconciliation work that needs the ledger before a delta range is available, emit only the ledger and skip git range inspection:

```bash
python .template-sync/scripts/generate_sync_candidates.py --ledger-only --write-ledger ADOPTION_LEDGER_PATH
```

For a PR-description-ready summary after `.template-sync/marker.yml` and any `_TODO-repo-init.md` state are current, emit the concise adoption summary and skip git range inspection:

```bash
python .template-sync/scripts/generate_sync_candidates.py --summary
```

Use `--summary` for the final adoption PR body or owner handoff when maintainers need the high-level module set, protected-file decisions, local overrides, unresolved decisions, manual TODOs, and retained-module validation commands in a compact form. Use `--ledger` or `--ledger-only` for the detailed path-level review artifact, especially before editing protected files, auditing local overrides, or reviewing `_TODO-repo-init.md` checklist links. The summary is a standalone mode and cannot be combined with `--ledger`, `--ledger-only`, or `--preflight`. The summary is deterministic and pasteable, but it is not authoritative state and it does not replace the full ledger when protected-file review details are needed.

The `ADOPTION_LEDGER_PATH` value MUST remain inside the repository root; the helper rejects paths that escape it. The ledger is a generated review artifact. It MUST be regenerated when `.template-sync/manifest.yml`, `.template-sync/marker.yml`, `_TODO-repo-init.md`, the selected adoption mode, or the included module set changes. Do not treat a stale committed ledger as authoritative: `.template-sync/manifest.yml` and `.template-sync/marker.yml` remain the machine-readable source of truth. Rows marked `manual TODO` link to `_TODO-repo-init.md` checklist lines when that file exists. Rows marked `local override` MUST show a reason from `.template-sync/marker.yml`; if a local override has no durable reason, fix the marker before relying on the ledger. Rows marked `local ownership` summarize `template_sync.local_path_ownership` records without treating those paths as manifest-owned template files. Protected rows with matching `protected_file_decisions` show the protected decision state, distinguish default `minimal-preservation` from authorized minimal edits and authorized tailored rewrites, and include a distinct `REMOVE-LOCAL` authorization section when protected removals are recorded.

`_TODO-repo-init.md` link targets in the ledger depend on the rendering destination. When `--write-ledger` is used, the helper writes Markdown link targets relative to the saved file's directory so the links resolve correctly when the saved ledger is committed and rendered on GitHub. When the ledger is emitted only to stdout (no `--write-ledger`), the helper writes repo-root-relative link targets, which are informative when reading the ledger in a terminal but **do not** render as clickable links to repository files when pasted into a GitHub pull request or issue body — GitHub resolves relative Markdown link targets in PR/issue bodies relative to the PR/issue URL, not the repository root. For clickable rendered links, run the helper with `--write-ledger ADOPTION_LEDGER_PATH`, commit the saved file, and review it directly on GitHub.

The generated candidate table models upstream path changes. It MUST NOT be expanded to review `_TODO-repo-init.md` as an upstream candidate; that file is downstream-owned first-adoption state. During full reconciliation, carry forward the Step 4 pre-filter decision that excludes `_TODO-repo-init.md` from candidate review.

For each path from `git diff --name-status -M`:

1. Map the path to its manifest relation.
2. Include the path in the per-file review table only if every `requires_all` module is present in `included_modules` and, when `requires_any` is present, at least one `requires_any` module is present.
3. Exclude the path from the per-file review table if any `requires_all` module is absent, or if `requires_any` is present and none of its modules are included.
4. Summarize excluded paths as unadopted-module activity by module in the sync summary.
5. Surface unknown modules and unmapped paths for explicit owner review before completing the sync.

Unadopted-module activity and unknown modules are different cases:

- **Unadopted-module activity** uses known modules from this taxonomy, such as `terraform`, that the downstream marker intentionally omits. Summarize it by module and path count or path list.
- **Unknown modules** are not known to the downstream marker or procedure. The owner MUST decide whether the module should be added to `included_modules`, mapped to an existing module, or deferred.

If summarized unadopted-module activity appears relevant during review, the owner MAY opt into that module before completing the sync by adding it to `included_modules`, or MAY defer opt-in to a later PR. Record either choice in the sync summary.

## Step 7: Review Each Candidate File

Every included candidate requires a row in a per-file decision table. When the Step 6 generator is used, treat its `Retained` rows as the starting candidate set and its `Excluded`, `Unmapped`, local override, deferred protected candidate, protected-file decision, protected-file, rename, and inline-block notes as prompts for manual review and sync-summary entries. When a Step 6 adoption ledger exists, review it before editing any protected file and carry forward its protected-file flag, protected-file decision state, local-override reason, adoption-mode value, `_TODO-repo-init.md` link, and validation-command hints into the per-file decision table or sync working notes.

For protected files and template-derived governance, community, process, workflow, or collaboration files, assign an adoption mode before choosing the file decision:

- Use `minimal-preservation` by default. Do not prompt the maintainer repeatedly when this default applies.
- Use `tailored` only when the maintainer explicitly selects it for the specific file or file set.
- For protected files, record the selected mode in `template_sync.protected_file_decisions` before editing the affected file.
- For non-protected files, record the selected mode in `_TODO-repo-init.md`, `.template-sync/marker.yml` local overrides, the sync working notes, or the final sync summary before editing the affected file.

In `minimal-preservation` mode, permitted edits are limited to:

- substituting repository placeholders such as `OWNER/REPO`, confirmed contacts, confirmed owner/team names, and GHES hosts
- removing complete delimited sections owned by unadopted manifest modules, such as a `# template-sync: begin terraform-only` block or an `<!-- template-sync: begin python-reference-only -->` block when the corresponding module is excluded
- fixing links that would otherwise be broken after placeholder substitution, module removal, or path deletion
- adding required local overrides or local ownership records in `.template-sync/marker.yml`

In `minimal-preservation` mode, do not shorten or restructure instruction files for style, paraphrase protected prose, rewrite community or workflow files merely to sound more project-specific, or remove rules that still apply to adopted modules.

For example, when a downstream repository excludes Python but retains Codex instructions, a minimal-preservation edit MAY remove complete `python-reference-only` blocks from `AGENTS.md` while keeping `## GitHub Plugin Usage` and `## PR Review Workflow (Codex-adapted)` unchanged. The same edit MUST NOT collapse `AGENTS.md` into a short generic summary or remove the retained Codex platform protocol merely because the file is described as a thin entry point.

For example, when a downstream repository excludes Terraform but retains Claude instructions, a minimal-preservation edit MAY remove complete `terraform-reference-only` blocks from `CLAUDE.md` while keeping `## Handling Code Review Comments` and `## Automated Review Loop` unchanged. The same edit MUST NOT delete or rewrite the Claude review loop unless a protected-file decision records an explicit owner waiver for that protocol.

Adoption mode applies in addition to protected-file authorization. Selecting `minimal-preservation` does not by itself authorize editing `.github/copilot-instructions.md`, `.github/instructions/*.instructions.md`, root agent files, or other protected instruction files.

| Decision | Meaning |
| --- | --- |
| `TAKE` | Adopt the upstream version as-is. |
| `MERGE` | Manually merge relevant upstream changes with local customizations. |
| `SKIP` | Do not adopt because the change is irrelevant, not adopted, or locally superseded. |
| `REMOVE-LOCAL` | Remove a downstream file because the module or feature is intentionally excluded. |
| `DEFER` | Leave unresolved pending owner decision. |
| `PROTECTED-REVIEW` | Protected-file change requiring explicit owner authorization before it can be applied. |

Suggested table:

```markdown
| Path | Module(s) | Adoption Mode | Template Change | Local Customization | Decision | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| `.github/instructions/docs.instructions.md` | `markdown`, `agent-instructions` | `minimal-preservation` | Updated style rules | None | `PROTECTED-REVIEW` | Requires owner authorization; mode does not grant it. |
| `.github/workflows/powershell-ci.yml` | `powershell`, `github-actions` | `minimal-preservation` | Updated validation steps | Local runner change | `MERGE` | Preserve local runner while retaining upstream workflow structure. |
| `README.md` | `baseline` | `tailored` | Updated setup prose | Project-specific README | `SKIP` | In scope per relation filtering; local override defaulted to `SKIP`; recorded in the sync summary. |
```

Upstream deletions MUST be surfaced for owner decision rather than applied automatically. Valid decisions for deletion rows include `TAKE` when the downstream owner agrees to delete the local file, `SKIP` when the downstream file is intentionally retained, and `MERGE` when only part of the deletion rationale applies.

Upstream renames MUST preserve both old and new paths in the table. If the rename is adopted, apply it as a rename locally rather than recreating the file without history where practical.

## Step 8: Perform Manual Merges

Use `MERGE` when the downstream file contains local customizations that should survive. A merge MUST inspect both the upstream candidate and the local file.

### Scratch Location

This template ignores `.cache/` in `.gitignore`, so `.cache/template-sync/` is the preferred scratch location when the downstream repository still has that ignore rule.

Verify the location is ignored:

```bash
git check-ignore .cache/template-sync/
```

If the command does not report the path as ignored, either add an equivalent downstream-specific ignore rule before using the location or choose another already ignored scratch location. Do not commit scratch copies.

Create the scratch directory:

```bash
mkdir -p .cache/template-sync
```

PowerShell equivalent:

```powershell
New-Item -ItemType Directory -Force -Path .cache/template-sync
```

### Inspect Upstream Content

Inspect an upstream file directly:

```bash
git show TEMPLATE_SHA:.github/workflows/powershell-ci.yml
```

Or restore an upstream copy into scratch for side-by-side reconciliation:

```bash
git show TEMPLATE_SHA:.github/workflows/powershell-ci.yml > .cache/template-sync/powershell-ci.upstream.yml
```

Then manually reconcile the scratch copy with the downstream file. Preserve local customizations unless the per-file decision explicitly replaces them.

### Normalize Line Endings After `.gitattributes` Changes

When the sync adopts, updates, or reintroduces `.gitattributes`, run a reviewable line-ending normalization pass after the intended `.gitattributes` content is in the working tree and before validation. This is required when new `eol=lf` pins may change the clean-filtered content of tracked files, such as adding template-managed Markdown, PowerShell, JSON, TOML, JavaScript, or Python defaults. Exclude unrelated local edits first so the normalization review is limited to the sync.

`git add --renormalize .` stages every tracked path whose clean-filtered content has changed, including substantive sync edits that happen to be unstaged, so before running it stage and commit (or stash) any substantive sync edits already in the working tree. When isolating substantive edits is impractical, restrict the renormalization pass to specific paths (for example, `git add --renormalize <path>...` for the files affected by the `.gitattributes` change) so the staged result contains only line-ending normalization.

The preferred Git normalization pass is:

```bash
git add --renormalize .
```

Then inspect the staged normalization before committing it:

```bash
git diff --check
git diff --cached --check
git diff --cached --name-status
```

If the downstream workflow does not stage changes during review, use an equivalent reviewable pass that shows which tracked files changed because of `.gitattributes`. Record the command or method in the sync working notes.

The repository-wide commit-hygiene rule in [`.github/copilot-instructions.md`](.github/copilot-instructions.md) still forbids routine standalone formatting-only or lint-only commits. A separate normalization commit is allowed in this procedure when `.gitattributes` adoption causes broad mechanical line-ending churn and separating that churn materially improves reviewability. That commit MUST contain only the reviewed normalization effects caused by the `.gitattributes` change, MUST stay adjacent to the related sync change, and MUST be recorded in the sync summary. If the normalization is small, include it in the same substantive sync commit instead.

## Step 9: Handle Protected Files

Protected governance and instruction files require explicit owner authorization before editing.

Protected files include:

- `.github/copilot-instructions.md`
- `.github/instructions/*.instructions.md`
- `.cursor/rules/*`
- `.hermes.md`
- `AGENTS.md`
- `CLAUDE.md`
- `GEMINI.md`

The default decision for protected-file changes is `PROTECTED-REVIEW`.

The default adoption mode for protected-file candidates is `minimal-preservation`, but that default only constrains the edit once authorization exists. It does not replace the explicit owner authorization requirement. A maintainer MAY select `tailored` for a protected file, but that selection still requires the same path-scoped protected-file authorization before editing.

### Same-Sync Protected Edits

Protected-file changes MAY be included in the same sync change set, whether it finalizes as a PR-ready branch or a completed local branch, only when the owner gives explicit, current-task authorization for the specific protected path or paths being changed.

Before applying the change, record a matching `template_sync.protected_file_decisions` entry:

- Use `decision: TAKE` or `decision: MERGE` for protected content adoption. Include `adoption_mode`, `authorization_basis`, and `authorized_scope`.
- Use `adoption_mode: tailored` only when the owner explicitly authorized a broad rewrite, and include `tailored_authorization_basis`.
- Use `decision: REMOVE-LOCAL` only when the owner explicitly authorized protected-file removal. Include `authorization_basis`, `authorized_scope`, and `reason`.

The sync summary MUST record the authorization basis in a reviewer-verifiable form, such as:

- a linked owner comment
- a linked authorization issue
- a quoted owner instruction from the current task

The sync summary MUST list each protected path, the authorization basis, and the validation performed.

For `REMOVE-LOCAL`, `authorization_basis` MUST explicitly name the removed file and mention removal, deletion, or equivalent removal-specific wording; generic "protected file edits authorized" wording is not enough. Reviewers MUST verify that `authorization_basis`, `authorized_scope`, and `reason` support removal before approving the PR. The validator's optional `--strict-remove-local-phrasing` flag can automate the brittle substring check when maintainers want it:

```bash
python .template-sync/scripts/validate_marker.py --require-marker --strict-remove-local-phrasing
```

Use `--remove-local-authorization-token TOKEN` one or more times, or `--remove-local-authorization-tokens remov,delet`, to tune the case-insensitive substring list. Normal upstream CI does not require this flag unless maintainers deliberately opt in.

Protected-file edits MUST be placed in a separate commit from non-protected changes unless the owner's current-task authorization explicitly waives commit isolation. If commit isolation is waived, record the waiver and authorization basis in the sync summary.

### Deferred Protected Edits

If authorization is absent, leave the protected file unchanged and record the candidate in `.template-sync/marker.yml` under `deferred_protected_candidates`. Do not also add a same-path `protected_file_decisions` entry until the owner makes a current decision for that protected path.

Example:

```yaml
template_sync:
  deferred_protected_candidates:
    - path: .github/copilot-instructions.md
      source_commit: "2222222222222222222222222222222222222222"
      reason: "Updates stack-selection guidance; awaiting owner authorization."
```

Leave the PR thread or sync question open when the owner still needs to act on the protected-file candidate.

## Step 10: Preserve Local Identity

Before taking or merging a file, check whether it contains local project identity or downstream policy. Examples include:

- repository owner/name
- package names
- workflow runner labels
- branch names
- support or security contact details
- issue labels, projects, assignees, and issue types
- PR checklist items
- README product description
- license policy
- validation commands that reflect adopted modules

Use `MERGE`, not `TAKE`, when upstream content is useful but the downstream identity must remain.

## Step 11: Re-Substitute Template Placeholders

After applying upstream content, search for unresolved template placeholders and replace them with downstream values or remove the placeholder-bearing content.

```bash
git grep "OWNER/REPO"
```

Common placeholder shapes include:

- `OWNER/REPO`
- `https://github.com/OWNER/REPO`
- `https://github.com/OWNER/REPO.git`
- `https://github.com/OWNER/REPO/blob/HEAD/PATH`

References to the upstream template repository itself (for example, `franklesniak/copilot-repo-template` in `.template-sync/marker.yml` under `source_repo`, in the Step 2 `git remote add template` example, or in adopted documentation that intentionally links to the upstream template) are not template placeholders and are out of scope for Step 11. Review and decide on those references under Step 10 (Preserve Local Identity) so that legitimate retentions are not rewritten and any deliberate rebranding remains an explicit decision.

Do not replace didactic examples that intentionally explain the placeholder convention unless the downstream repository has removed the associated placeholder-check workflow and no longer wants the convention documented.

## Step 12: Validate by Adopted Module

Run validation appropriate to the included modules and files changed. Full template-like validation remains the safest default when the downstream repository keeps the relevant tooling.

During first adoption, run the working-tree validation runner before the first adoption commit when the helper is present:

```bash
python .template-sync/scripts/run_first_adoption_checks.py --doctor
```

```bash
python .template-sync/scripts/run_first_adoption_checks.py --plan-only
```

```bash
python .template-sync/scripts/run_first_adoption_checks.py --check
```

If check mode reports final Git status changes after otherwise passing planned commands, inspect the changed-file summary before deciding whether to keep those changes. To intentionally let mutating hooks or fixers update files, run:

```bash
python .template-sync/scripts/run_first_adoption_checks.py --fix
```

After fix mode or any unexpected final Git status change, review `git status`, use `git diff` and `git diff --cached` for tracked-content diffs, keep or discard the edits intentionally, then rerun check mode:

```bash
python .template-sync/scripts/run_first_adoption_checks.py --check
```

Use `--doctor` first when validation tooling availability is unclear. Doctor mode runs small non-mutating probes, captures bounded stdout/stderr diagnostics, records per-probe timeout state, reports both `pwsh` and `powershell` independently when present, checks PSScriptAnalyzer availability inside each PowerShell host, and distinguishes missing shims or optional tools from failing tests and hooks. It also reports pre-commit hook IDs from `.pre-commit-config.yaml` as pre-commit-managed surfaces instead of direct-probing every hook executable. The pytest gate-ready verdict is tied to the Python interpreter running the helper, so a working bare `pytest` command or a different Python launcher is only supplemental evidence when that exact interpreter cannot run pytest.

This helper is different from `pre-commit run --all-files`: it uses `git ls-files --cached --others --exclude-standard` to include newly copied, untracked, non-ignored files before they are committed, then invokes `pre-commit run --files ...` against that explicit file list. Use `--plan-only` to inspect the deterministic file collection result, explanatory notes, and numbered command plan without running pre-commit, placeholder scans, marker validation, npm scripts, pytest, or other validation commands; planned commands printed in this mode are not executed validation-command boundaries, so status changes in plan-only mode cannot receive command-boundary attribution. In plan-only mode, an applicable pytest gate is shown as a planned command, but pytest tool availability is not evaluated; use `--doctor` for environment and tool diagnostics. In normal execution, the helper prints the full command plan before the first validation command starts, warns that cold pre-commit hook environment bootstrapping may be slow, chunks large pre-commit file lists to avoid command-line length limits, and prints each command's group label, index, UTC start and end timestamps, elapsed time, and exit status. It continues through the whole plan, reports multiple failures together, prints a deterministic before/after Git status summary for files changed during the invocation, supplements that aggregate summary with command-boundary status transitions observed after executed planned validation commands when transitions occur, prints total elapsed time, and also runs the placeholder scan, marker validation, supported Markdown npm scripts, and the downstream pytest gate when those checks are applicable. Nonzero planned commands still fail the run. In check mode, when planned commands otherwise pass, a final aggregate Git status change exits with the distinct changed-file code after status-scoped follow-up guidance. In fix mode, final aggregate Git status changes are tolerated only when planned commands otherwise pass. Transient command-boundary status transitions are reported but do not change the exit code by themselves. If a planned command also fails, resolve the validation failure separately from any changed-file keep/revert decision. If the `pre-commit` console script is not present or does not execute successfully, generated validation plans prefer an available `python -m pre_commit run --files ...` form. Continue to run `pre-commit run --all-files` before committing after the adopted files are tracked.

The runner plans the downstream pytest gate only when pytest support is retained. Once `.template-sync/marker.yml` exists, the marker's `template_sync.included_modules` controls retention. Before the marker exists, pass `--retained-module MODULE` once for each selected module when you need the plan to reflect explicit pre-marker decisions. If no marker and no explicit retained-module decision is available, the runner falls back to file presence. A planned pytest gate requires both the retained pytest configuration (`pyproject.toml`) and at least one retained downstream pytest candidate path; otherwise `--plan-only` and normal planning print a skip reason that distinguishes pruned pytest configuration from retained configuration with no downstream pytest candidates. When the gate runs, it uses pytest's module-form invocation with the official `not upstream_template_only` marker selection.

When `.template-sync/scripts/first_adoption_quality_reports.py` is present, the plan also inventories line endings, path-reference casing, optional PowerShell analyzer debt, and Markdownlint debt before broad validation or fixers run. In `--fix` mode, the Markdown report command switches to an explicit Markdown fixer command and reports files changed afterward while the runner's normal changed-file summary remains authoritative. Missing optional tools such as npm, PowerShell, or PSScriptAnalyzer are reported as unavailable states; the helper does not install tools automatically. Intentional path-reference findings MAY be suppressed in the downstream-created `.template-sync/first-adoption/quality-suppressions.json` file, whose current schema supports the report-scoped `path-reference` section only.

When first-adoption materialization is used, inspect the command surface first:

```bash
python .template-sync/scripts/materialize_downstream_adoption.py --help
```

The Python interpreter running the materializer must provide PyYAML for parsing and jsonschema for manifest and marker validation. Help output requires PyYAML but does not load jsonschema. These requirements apply when only template-sync support is retained.

The running materializer loads executable helpers only from its own reviewed installation. It validates selected workflow contract data against the workflow security schema from that installation. Selected workflow validators, workflow schemas, and placeholder scripts remain inert staged files during the current run, including sources obtained by ref or revision. Keep the installed materializer, shared helpers, and their data dependencies together; review and upgrade that bundle when a newer source contract is incompatible. Template-sync support retains the shared helpers and workflow schema even when Actions or baseline is omitted.

Use `--template-root` when a local source checkout already exists. When that root is the top level of a complete, clean Git worktree, the helper reports the observed source commit and source worktree root. That observed commit is not automatically reviewed state. After completing review, rerun with `--stamp-resolved-source-as-reviewed`, or pass `--last-reviewed-template-commit FULL_SHA`. The stamp flag asserts both review completion and intended upstream lineage; a complete, clean local `HEAD` proves local checkout state, not source lineage by itself.

Local-root stampability covers tracked content and rejects dirty tracked files plus non-ignored untracked files; ignored untracked files are outside this gate. A `--template-root` that points at a subdirectory of a larger Git worktree is outside stampability and is not silently promoted to the worktree root. Sparse checkout, skip-worktree, and assume-unchanged states are outside stampability because tracked content can be absent or hidden while ordinary status output remains clean. On Git older than 2.36, or when the Git version cannot be parsed, configured `core.fsmonitor` makes a `--template-root` source not stampable because the helper requires Git detected as 2.36 or newer before it relies on `-c core.fsmonitor=false`; Git 2.35.1 and earlier treat `false` as a hook pathname rather than boolean false, and Git older than 2.26 can silently report incomplete status results. Use a Git executable detected as 2.36 or newer, or pass `--last-reviewed-template-commit FULL_SHA` after review. Template sources that record submodule gitlinks are not stampable, and are fatal for `--template-ref` / `--template-revision`, because the materializer neither initializes nor verifies submodules. Restore a full checkout of the template source before materializing from it when these gates fail.

To materialize from a locally fetched upstream ref or a full upstream commit SHA without manually creating a source checkout, pass `--template-ref REF` or `--template-revision FULL_SHA`. The helper resolves that value against `--template-repo PATH` without fetching, pulling, merging, or rebasing; when `--template-repo` is omitted, it defaults to `--target-root` so locally fetched template refs in the downstream repository can be used. The helper creates a private detached worktree outside `--target-root`, forces that temporary checkout out of sparse-checkout mode, verifies the temporary checkout before trusting its contents, reports the supplied source ref or revision, resolved source SHA, source repository, diagnostic temporary checkout path, cleanup status, and target root, then removes the temporary worktree. Do not set `template_sync.last_reviewed_template_commit` merely because a ref was materialized; record it only after the review procedure is complete. If a reviewed SHA is supplied and it differs from the resolved source SHA, the helper rejects the run.

For shell-sensitive identity values, pass `--args-file PATH` instead of relying on shell quoting. Both JSON and `.yaml`/`.yml` args files are supported, with `.yaml`/`.yml` parsed through the retained YAML parser path. This materializer already requires the retained YAML parser (PyYAML) for manifest and marker processing, so it must be available regardless of args-file format; converting the args file to JSON does not remove that requirement. `--args-format json` or `--args-format yaml` overrides the extension and is required for extensionless or unknown-extension files. CLI flags take precedence over args-file values, and the merged source/module values must still agree with any `--decisions-file` values. Repository-relative path values inside args or decisions, such as `decisions_file` and `license_source_path`, must stay inside `--target-root`.

The materializer and placeholder helper reject repeated YAML mapping keys and JSON object names in argument files, including nested objects. For example, two `included_modules` entries fail instead of allowing the second selection to replace the first. Marker-based placeholder classification and the host-setup report also reject repeated YAML keys before interpreting module selections. Safe YAML merge overrides remain supported; an explicit value may override an inherited value. The host-setup report remains advisory and does not establish hosted validation.

When the `markdown` module is retained, the args file may include `package_name`, `package_description`, `package_author`, `package_keywords`, and, only when intentionally changing the package release version, `package_version`. Package identity edits update `package.json` and the root identity fields in `package-lock.json` deterministically without running `npm install`; lockfile version fields change only when `package_version` is explicitly supplied.

When `github-templates` is retained, the args file or marker may include `issue_label_policy`, `issue_labels`, `discussions_policy`, and `collaboration_policy_follow_up_status`. `issue_label_policy` accepts `existing`, `create-manual-follow-up`, `omit`, or `custom`; `custom` requires `issue_labels`. `discussions_policy` accepts `enabled`, `disabled`, `deferred-planned-render`, or `deferred-not-rendered`. Policies that render future-state content or defer a GitHub setting require `collaboration_policy_follow_up_status` with the matching `_TODO-repo-init.md` dependent-file status so issue and PR collaboration templates do not ship misleading defaults.

When the owner has approved preserving existing downstream license text from an alternate source path such as `LICENSE.txt` or `LICENSE.md`, pass `--preserve-existing-license --license-source-path SOURCE_PATH`. The helper validates that the source path is repository-relative, non-symlink, readable UTF-8 text, and distinct from root `LICENSE`; copies the source bytes to root `LICENSE`; suppresses the template `LICENSE`; records or previews a `local_overrides` `SKIP` decision for root `LICENSE`; and reports the original source path as residual manual cleanup. If root `LICENSE` already exists, do not use the alternate-path normalization flags. Preserve that same-path license with the existing `template_sync.local_overrides` `SKIP` decision instead. The helper never deletes the source license file and MUST NOT be used to alter license text without explicit owner authorization.

The materialization command writes non-conflicting retained files, reports conflicting paths as a separate actionable section, and uses exit code `2` for unrecorded conflicts by default. Existing downstream files such as `README.md`, `LICENSE`, and `.gitignore` commonly conflict during first adoption; resolve them with path-scoped `local_overrides`, `protected_file_decisions`, or deferred protected candidates rather than treating the run as a runtime failure. If source worktree cleanup fails after otherwise successful materialization, the helper exits with a dedicated cleanup-failure code and prints the residual path plus the manual `git worktree remove --force` command; the target tree remains intact. If materialization itself fails and cleanup also fails, the materialization error remains primary and the cleanup diagnostic names the residual worktree and manual removal command.

Before editing protected files or finalizing adoption decisions, generate review artifacts with the existing candidate generator rather than reading the source tree broadly. Use `python .template-sync/scripts/generate_sync_candidates.py --ledger` during a normal delta review, `python .template-sync/scripts/generate_sync_candidates.py --ledger-only` when first-adoption or full-reconciliation work needs the ledger before a range is available, and `python .template-sync/scripts/generate_sync_candidates.py --summary` for the concise PR-description or owner-handoff summary after marker and checklist state are current.

Run the downstream adoption validation command after file removals, retained files, `included_modules`, `local_overrides`, `local_path_ownership`, `protected_file_decisions`, and `deferred_protected_candidates` reflect the intended sync result, and before finalizing the sync summary. Use `--require-marker` once the downstream repository has committed to carrying `.template-sync/marker.yml` in CI:

```bash
python .template-sync/scripts/validate_downstream_adoption.py --require-marker
```

The aggregate command composes marker schema validation, retained-file presence checks, excluded-module leftover checks, protected-file decision checks, downstream instruction-contract validation, inline-block consistency checks, retained Markdown relative-link checks, and local path ownership warnings. Omit `--require-marker` only during initial adoption or exploratory sync work where the marker may intentionally be absent; in that mode, the helper exits zero with a clear no-marker message.

Downstream repositories that retain pytest-based template support SHOULD run the official downstream pytest gate with negative marker selection:

```bash
python -m pytest -m "not upstream_template_only"
```

The committed pytest configuration registers the `upstream_template_only`, `downstream_template_support`, and `slow` markers and enables strict marker validation, so marker typos fail during collection. The downstream gate intentionally excludes only tests marked `upstream_template_only`; a newly added unmarked test remains included by default. The shipped Python CI surfaces apply the same negative selection automatically: `.github/workflows/python-ci.yml` deselects `upstream_template_only` tests in every repository except the upstream template (keyed on `github.repository`), and `.azuredevops/pipelines/python-ci.yml` always deselects them because Azure Pipelines never run for the upstream template.

Retained downstream tests MUST derive optional-module expectations from `.template-sync/manifest.yml` and the materialized `.template-sync/marker.yml`. After modules such as `terraform` or `powershell` are excluded, retained tests may still assert absence, cleanup reporting, or documented exclusion behavior, but they MUST NOT require excluded module-owned files, commands, inline blocks, pytest markers, or validation surfaces to exist. Tests that genuinely require the complete upstream template module set MUST be marked `upstream_template_only` so the downstream gate does not select them.

Run instruction-contract validation directly when debugging protected agent instruction protocols or when validating the upstream template repository. Mode selection is explicit and has no fallback:

- In the upstream template repository and upstream template CI, use `--mode upstream-template`. This validates every contract entry against the template's own protected files, does not read `.template-sync/marker.yml`, does not apply marker-derived module gating, and never fails merely because the marker is absent. If `.template-sync/marker.yml` is present, the validator emits a non-blocking warning because the caller may be validating a downstream working tree with the upstream mode.
- In downstream repositories, use `--mode downstream`. This reads `.template-sync/marker.yml`, checks only contracts whose `requires_modules` are all present in `template_sync.included_modules`, and honors `--require-marker` with the same semantics as `validate_marker.py`: missing marker fails when `--require-marker` is set and exits zero with a clear no-marker message otherwise.

Upstream template validation command:

```bash
python .template-sync/scripts/validate_instruction_contracts.py --mode upstream-template
```

Downstream instruction-contract-only validation command once the repository carries a marker:

```bash
python .template-sync/scripts/validate_instruction_contracts.py --mode downstream --require-marker
```

In downstream mode, an absent contracted protected file is skipped only when `template_sync.protected_file_decisions` records an authorized `REMOVE-LOCAL` decision for that path and the file is absent from the working tree; the validator reports the skip as an authorized removal. Otherwise, absent files and missing anchors fail with the exact file and heading or phrase. A matching `template_sync.instruction_contract_waivers` entry can waive a missing anchor, but successful output says `passed with waivers` and lists the waiver instead of reporting ordinary success.

The manifest concrete-pattern integrity check is an upstream-maintainer audit and is excluded from the official downstream pytest gate. When run directly as a diagnostic, it has two modes. In the upstream template repository, where `.template-sync/marker.yml` is intentionally absent, `pytest tests/test_template_manifest.py -v` uses upstream-template mode: every concrete manifest path MUST resolve to a Git-tracked file unless it is explicitly allowlisted. In a downstream repository that carries `.template-sync/marker.yml`, the same direct test uses downstream-marker mode: it reads `template_sync.included_modules` and `template_sync.local_overrides`, checks only concrete paths retained by the marker's module relation, skips local overrides, treats untracked non-ignored files as present, and does not treat ignored scratch or cache files as retained-template files.

Before module-specific validators, check for whitespace errors and unresolved conflict markers:

```bash
git diff --check
```

When changes are staged or a separate normalization commit is being prepared, also check the staged diff:

```bash
git diff --cached --check
```

When `.gitattributes` changed or line-ending normalization occurred, verify the expected EOL attributes for the touched paths. A broad reviewable check is:

```bash
git ls-files --eol -- .
```

Before module-specific validators, verify structural consistency for retained modules:

- Retained workflows under `.github/workflows/` MUST invoke test roots, source roots, config files, and scripts that exist after stack selection and structural alignment.
- Retained test commands MUST point at the downstream repository's actual test roots, such as `tests/`, `tests/PowerShell/`, Terraform test directories, or schema example fixture directories.
- Retained package scripts, pre-commit hooks, schema validators, and markdown/YAML/Terraform linters MUST point at config files that still exist after excluded modules are removed.
- User-facing docs MUST match any adopted structural change that alters commands, workflow names, test paths, or validation prerequisites.

| Module | Example validation |
| --- | --- |
| `baseline` | `pre-commit run --all-files` |
| `agent-instructions` | `npm run lint:md`, `npm run lint:md:links`, `npm run lint:md:nested`, `pre-commit run check-json --all-files`, `pre-commit run check-toml --all-files`, shell-script syntax check for any session hooks (e.g., `if [ -d .claude/hooks ]; then find .claude/hooks -type f -name '*.sh' -exec bash -n {} \;; fi` — POSIX-portable; the `if [ -d ... ]` guard makes the check a clean no-op for downstream repos without `.claude/hooks/`, and `find` returns exit 0 when no `*.sh` files match), and any repo-specific instruction checks |
| `github-platform` | `pre-commit run check-yaml --all-files`, `pre-commit run yamllint --all-files`, `pre-commit run validate-dependabot-config --all-files` and `pre-commit run validate-dependabot-config-valid-examples --all-files` where configured, and repository-settings review |
| `github-actions` | `pre-commit run check-yaml --all-files`, `pre-commit run yamllint --all-files`, `pre-commit run actionlint --all-files` |
| `github-templates` | `pre-commit run check-yaml --all-files`, `pre-commit run yamllint --all-files`, `npm run lint:md`, `npm run lint:md:links`, and issue or PR template rendering review |
| `template-onboarding` | `npm run lint:md`, `npm run lint:md:links`, `npm run lint:md:nested`, and walkthrough review for kept onboarding paths |
| `template-sync-support` | `python .template-sync/scripts/run_first_adoption_checks.py --check` before the first adoption commit when validating copied files that may still be untracked, `python .template-sync/scripts/run_first_adoption_checks.py --fix` only when intentionally allowing mutating hooks or fixers, `python .template-sync/scripts/validate_downstream_adoption.py --require-marker` after marker decisions and retained files are current, `python .template-sync/scripts/validate_marker.py --require-marker` or `python .template-sync/scripts/validate_instruction_contracts.py --mode downstream --require-marker` for narrow debugging, `python -m pytest -m "not upstream_template_only"` for the downstream pytest gate when pytest support is retained, `npm run lint:md`, `npm run lint:md:links`, `npm run lint:md:nested`, `pre-commit run check-yaml --all-files`, `pre-commit run yamllint --all-files`, `pre-commit run validate-template-sync-marker-valid-examples --all-files`, `pre-commit run validate-template-sync-instruction-contracts-valid-examples --all-files`, `pre-commit run validate-first-adoption-quality-suppressions-valid-examples --all-files`, `pre-commit run validate-template-sync-manifest-schema --all-files`, `pre-commit run validate-template-sync-marker-schema --all-files`, `pre-commit run validate-template-sync-instruction-contracts-schema --all-files`, `pre-commit run validate-first-adoption-quality-suppressions-schema --all-files`, `pre-commit run validate-template-sync-manifest --all-files`, `pre-commit run validate-template-sync-marker --all-files`, `pre-commit run validate-template-sync-instruction-contracts --all-files`, `pre-commit run validate-first-adoption-quality-suppressions --all-files`, `pre-commit run validate-instruction-contracts-upstream --all-files`, `pre-commit run validate-instruction-contracts-downstream --all-files`, and `pytest tests/test_schema_examples.py -v` after template-sync schema example changes, plus a dry-run review of the sync procedure examples |
| `markdown` | `npm run lint:md`, `npm run lint:md:links`, `npm run lint:md:nested`, `pre-commit run check-json --all-files` |
| `powershell` | `Invoke-Pester -Path tests/ -Output Detailed` |
| `json` | `pre-commit run check-json --all-files` |
| `yaml` | `pre-commit run check-yaml --all-files`, `pre-commit run yamllint --all-files` |
| `schema` | `pre-commit run validate-example-config-valid-examples --all-files`, `pre-commit run validate-example-config-schema --all-files`, and `pytest tests/test_schema_examples.py -v` after worked-example schema or schema-example changes |
| `python` | `python -m pytest -m "not upstream_template_only"` for downstream template-support tests, `python -m pyright --project pyrightconfig.json`, `pytest tests/ -m "not slow" -v --cov --cov-report=term-missing` for default upstream-template coverage, `pytest tests/ -m slow -v --no-cov` for explicit slow template tests when retained, `pre-commit run check-toml --all-files` |
| `terraform` | `terraform fmt -check -recursive`, `tflint --recursive`, `terraform test -verbose`, `pytest tests/test_terraform_hooks.py -v` after terraform-hook script changes |

Run `pre-commit run --all-files` before committing when the downstream repository uses pre-commit. During first adoption, run `python .template-sync/scripts/run_first_adoption_checks.py --check` before that first commit when the helper is present, because `pre-commit run --all-files` primarily evaluates tracked files and may miss copied but untracked non-ignored adoption files. If the helper or any validator changes files, inspect those changes, keep or discard them intentionally, and rerun the relevant check command before finalizing. If a repository intentionally removed a module and its validation tooling, record that in the sync summary rather than reintroducing validation commands blindly.

### Validation Triage

For each validation failure, record the validator, path, failing condition, whether the upstream range changed that path, and one classification in the sync working notes:

| Classification | Use when | Required disposition |
| --- | --- | --- |
| upstream-template fix | The failure is caused by adopted upstream template content, or is reproducible in the upstream template at the reviewed range head. A template-owned file having changed in the reviewed range is not by itself sufficient when the failing condition comes from downstream customization in that same file. | Fix in the sync when in scope, or record the upstream defect and owner-facing follow-up. |
| downstream-local fix | The failure is in downstream-owned content or downstream customization that the sync can safely fix now. | Fix locally in the sync and list the fix in the sync summary. |
| deferred follow-up | The failure is pre-existing downstream debt or a policy decision that is outside the sync's approved scope. | Leave the validator failure visible, record the owner decision or follow-up needed, and do not describe the sync as PR-ready unless required validation is passing or explicitly deferred by the owner. |

When a newly adopted validator fails on a file that the upstream template did not modify in the reviewed range, classify the finding as pre-existing downstream debt unless the same failure is reproducible in the upstream template at the reviewed range head. Do not weaken or remove the validator to make the sync pass. Fix the downstream issue, defer it with owner acknowledgement, or choose a finalization mode that accurately reflects the remaining validation debt.

## Step 13: Record the Reviewed Commit

After all decisions are recorded and validation is complete, update `.template-sync/marker.yml`:

- Set `template_sync.last_reviewed_template_commit` to the resolved upstream range head SHA that was reviewed.
- Keep `included_modules` current.
- Add, update, or remove `local_overrides` only when the owner made that adoption decision.
- Add, update, or remove `local_path_ownership` only for downstream-owned project paths; do not use it to bypass template-managed retained-file, excluded-module, protected-file, or inline-block validation.
- Add, update, or remove `protected_file_decisions` for protected files that were taken, merged, skipped, removed locally, deferred as a decision, or sent through protected review during this sync.
- Add or refresh `deferred_protected_candidates` for unresolved protected-file changes.
- Preserve adoption-mode records: keep the default `minimal-preservation` record in `_TODO-repo-init.md` or the sync summary, represent protected path-specific decisions in `protected_file_decisions`, and represent non-protected path-specific `tailored` opt-ins or required local ownership exceptions as `local_overrides` when `.template-sync/marker.yml` is the durable record.

Do not set `template_sync.last_reviewed_template_commit` to a commit that was not actually reviewed through the taxonomy and per-file process. Do not store a branch name, tag name, short SHA, or other moving ref in this marker field; store the full 40-character resolved upstream template commit SHA that was reviewed.

If Step 14 will use dry-run analysis only, do not update `.template-sync/marker.yml`. Record the proposed marker update in the sync summary instead. If Step 14 will use a local-inspection mode with uncommitted changes, the marker update MAY remain in the working tree for owner inspection, but the sync summary MUST state that the reviewed commit is not durable until those changes are committed.

## Step 14: Finalize the Sync

The final sync summary is assembled from the sync working notes after decisions and validation are complete. The working notes can remain a scratch artifact; the sync summary is the owner-facing record for the chosen finalization mode.

Step 14 does not always open a PR. Choose and record exactly one finalization mode:

| Finalization mode | Use when | Required result |
| --- | --- | --- |
| completed local branch with committed sync summary | The owner wants a complete local branch to inspect before any PR is opened. | Commit the applied sync changes and a sync summary artifact in an owner-approved or downstream-defined path. The working tree is clean, and the summary states whether a future PR is still needed. |
| completed local branch with working-tree changes for inspection | The owner wants to inspect edits before committing, or validation debt requires local review first. | Leave the working tree in the requested inspection state and provide the sync summary as the local handoff record. The summary MUST state which changes are uncommitted, which validation passed or failed, and what remains before commit or PR. |
| PR-ready branch | The sync is ready for normal review. | Commit the intended changes, record validation results, and prepare the sync summary for the PR body. Opening the PR MAY happen immediately when remote tooling and owner direction allow it. |
| dry-run analysis only | The owner requested analysis without applying changes, or the sync must stop before changing files. | Do not advance `.template-sync/marker.yml` or commit sync changes. Provide the sync summary as a dry-run report with proposed decisions, proposed marker update, validation not run or simulated, and open questions. |

The sync summary MUST include:

- finalization mode
- first-adoption preflight record and unresolved items, when applicable
- unresolved `_TODO-repo-init.md` manual GitHub settings, each with owner, next action, and evidence or follow-up location, when applicable
- structural convention findings and their required/strongly recommended/post-adoption/not-recommended classifications, when first adoption applies
- required structural changes, including whether each was implemented or remapped, plus any intentional local override recorded in addition for a deliberate downstream deviation
- post-adoption issue drafts or their committed handoff location
- adoption mode record for protected files and template-derived governance, community, process, workflow, or collaboration files
- upstream template commit range reviewed
- included modules
- unadopted-module activity summarized by module
- unknown modules or unmapped paths surfaced for owner decision
- files adopted unchanged
- files manually merged
- files skipped
- files removed locally because a module was intentionally excluded
- protected files applied with explicit authorization, including authorization basis and any commit-isolation waiver
- protected files removed with explicit removal authorization, authorized scope, and substantive reason
- protected files deferred for separate authorization
- local overrides applied during this sync, each with a brief upstream change description
- local customizations preserved
- line-ending normalization actions and whether a separate normalization commit was used
- validation commands run and results
- saved Step 6 candidate table location or sync working-notes citation, including the recorded `RANGE_BASE_SHA` and `RANGE_HEAD_SHA`
- saved adoption ledger location or sync working-notes citation, including whether it was generated by `--ledger` or `--ledger-only`
- each issue surfaced during validation or review, classified as upstream-template fix, downstream-local fix, or deferred follow-up
- open questions for the owner

Example summary skeleton:

```markdown
## Template Sync Summary

**Finalization mode:** PR-ready branch
**First-adoption preflight:** not applicable; existing `_TODO-repo-init.md` and `.template-sync/marker.yml` recorded resolved adoption answers.
**Unresolved `_TODO-repo-init.md` settings:** none
**Structural convention findings:** not applicable for this initialized delta sync
**Required structural changes:** none
**Post-adoption issue drafts:** none
**Adoption mode:** `minimal-preservation` by default for protected files and template-derived governance, community, process, workflow, and collaboration files; `README.md` has a path-specific `tailored` local override.
**Upstream range reviewed:** `1111111111111111111111111111111111111111..2222222222222222222222222222222222222222`
**Included modules:** baseline, agent-instructions, github-actions, github-templates, markdown, powershell, template-sync-support
**Unadopted-module activity:** terraform (`.github/workflows/terraform-ci.yml`)
**Unknown modules or unmapped paths:** none
**Files adopted unchanged:** `templates/markdown/README.md`
**Files manually merged:** `.github/workflows/powershell-ci.yml`, `.github/pull_request_template.md`
**Files skipped:** none
**Files removed locally:** none
**Protected files applied:** none
**Protected files deferred:** `.github/copilot-instructions.md` at `2222222222222222222222222222222222222222`
**Local overrides applied:** `README.md` defaulted to `SKIP`; upstream changed setup prose.
**Local customizations preserved:** self-hosted runner block; project-specific PR checklist.
**Line-ending normalization:** not needed; `.gitattributes` unchanged.
**Validation:** `pre-commit run --all-files` (passed), `npm run lint:md` (passed), `npm run lint:md:nested` (passed), `Invoke-Pester -Path tests/ -Output Detailed` (passed)
**Saved candidate table:** sync working notes section "Step 6 Candidate Table" for `1111111111111111111111111111111111111111..2222222222222222222222222222222222222222`
**Saved adoption ledger:** `.cache/template-sync/adoption-ledger.md` generated with `--ledger`

## Issue Classifications

- upstream-template fix: none
- downstream-local fix: none
- deferred follow-up: none

## Open Questions

- Should the downstream repository opt into the `schema` module in a follow-up PR?
```

## Worked Example

This example is illustrative. A downstream repository adopts `baseline`, `agent-instructions`, `github-actions`, `github-templates`, `markdown`, `powershell`, and `template-sync-support`. It does not use `terraform`, `python`, `json`, `yaml`, `schema`, or `github-platform` as independent modules.

### Scenario State

- Downstream sync marker at `.template-sync/marker.yml`: `template_sync.last_reviewed_template_commit` is `1111111111111111111111111111111111111111`
- First-adoption state: `_TODO-repo-init.md` exists and records resolved conduct, security, CODEOWNERS, label, Discussions, private vulnerability reporting, default-branch protection, default `minimal-preservation` adoption mode, and GHES-host decisions
- Upstream range head ref: `template/main`
- Resolved upstream range head SHA: `2222222222222222222222222222222222222222`
- Included modules: `baseline`, `agent-instructions`, `github-actions`, `github-templates`, `markdown`, `powershell`, and `template-sync-support`
- Local customization: `.github/workflows/powershell-ci.yml` uses a self-hosted runner block.
- Local customization: `.github/pull_request_template.md` includes project-specific checklist items.
- Removed at adoption time: Terraform and Python workflows and template directories.

### Fetch and Enumerate

Marker discovery happens before range selection: `.template-sync/marker.yml` is present, so read it as the marker. The marker supplies `1111111111111111111111111111111111111111` as the range base, so this sync uses normal delta sync.

```bash
git fetch template
git rev-parse 'template/main^{commit}'
git merge-base --is-ancestor 1111111111111111111111111111111111111111 2222222222222222222222222222222222222222
git diff --name-status -M 1111111111111111111111111111111111111111..2222222222222222222222222222222222222222 --
mkdir -p .cache/template-sync
python .template-sync/scripts/generate_sync_candidates.py --range-head 2222222222222222222222222222222222222222 --ledger --write-candidates .cache/template-sync/candidates.md --write-ledger .cache/template-sync/adoption-ledger.md
```

The sync working notes start with the reviewed range endpoints:

- **Marker discovery:** `.template-sync/marker.yml` present
- **Range mode:** normal delta sync
- **Range base:** `1111111111111111111111111111111111111111`
- **Range head:** `2222222222222222222222222222222222222222`
- **Saved candidate table:** `.cache/template-sync/candidates.md` for `1111111111111111111111111111111111111111..2222222222222222222222222222222222222222`
- **Saved adoption ledger:** `.cache/template-sync/adoption-ledger.md`

Hypothetical output:

```text
M       .github/copilot-instructions.md
M       .github/workflows/powershell-ci.yml
M       .github/instructions/powershell.instructions.md
M       .github/pull_request_template.md
M       TEMPLATE_UPDATE_PROCEDURE.md
M       templates/markdown/README.md
M       .github/workflows/terraform-ci.yml
R100    templates/markdown/intro.md   templates/markdown/getting-started.md
```

### Filter by Modules

The manual module table below mirrors the candidate generator's retained/excluded classification for the hypothetical range. The generator output is still reviewed and rewritten into the Step 7 decision table before any file is adopted, merged, skipped, or deferred.

| Path | Module(s) | In scope? |
| --- | --- | --- |
| `.github/copilot-instructions.md` | `agent-instructions` | yes |
| `.github/workflows/powershell-ci.yml` | `powershell`, `github-actions` | yes |
| `.github/instructions/powershell.instructions.md` | `powershell`, `agent-instructions` | yes |
| `.github/pull_request_template.md` | `github-templates` | yes |
| `TEMPLATE_UPDATE_PROCEDURE.md` | `template-sync-support` | yes |
| `templates/markdown/README.md` | `markdown` | yes |
| `.github/workflows/terraform-ci.yml` | `terraform`, `github-actions` | no; `terraform` is a `requires_all` module and is absent, so relation matching excludes the row even though `github-actions` is present |
| `templates/markdown/intro.md` to `templates/markdown/getting-started.md` | `markdown` | yes |

There are no unknown modules or unmapped paths in this example.

### Decide Per File

| Path | Adoption Mode | Decision | Notes |
| --- | --- | --- | --- |
| `.github/copilot-instructions.md` | `minimal-preservation` | `PROTECTED-REVIEW` | Owner authorization is absent; defer and record under `deferred_protected_candidates`. |
| `.github/workflows/powershell-ci.yml` | `minimal-preservation` | `MERGE` | Preserve self-hosted runner block; adopt upstream validation step changes without restructuring the workflow. |
| `.github/instructions/powershell.instructions.md` | `minimal-preservation` | `PROTECTED-REVIEW` | Defer and record under `deferred_protected_candidates`. |
| `.github/pull_request_template.md` | `minimal-preservation` | `MERGE` | Preserve project checklist; adopt upstream checklist additions. |
| `TEMPLATE_UPDATE_PROCEDURE.md` | `minimal-preservation` | `TAKE` | Retain the sync procedure and adopt upstream clarification. |
| `templates/markdown/README.md` | not applicable | `TAKE` | No local customization. |
| `templates/markdown/intro.md` to `templates/markdown/getting-started.md` | not applicable | `TAKE` | Adopt upstream rename. |

Unadopted-module activity:

| Module | Upstream activity | Disposition |
| --- | --- | --- |
| `terraform` | `.github/workflows/terraform-ci.yml` changed | Not adopted; summarize in PR. |

### Merge With Scratch Copies

```bash
mkdir -p .cache/template-sync
git show 2222222222222222222222222222222222222222:.github/workflows/powershell-ci.yml > .cache/template-sync/powershell-ci.upstream.yml
git show 2222222222222222222222222222222222222222:.github/pull_request_template.md > .cache/template-sync/pr-template.upstream.md
```

Manually reconcile each scratch file against the downstream file. The PowerShell workflow keeps its self-hosted runner block and adopts the upstream validation step changes. The PR template keeps project-specific checklist items and adopts the upstream checklist additions.

### Update the Marker

```yaml
template_sync:
  source_repo: https://github.com/franklesniak/copilot-repo-template.git
  last_reviewed_template_commit: "2222222222222222222222222222222222222222"
  included_modules:
    - baseline
    - agent-instructions
    - github-actions
    - github-templates
    - markdown
    - powershell
    - template-sync-support
  local_overrides:
    - path: README.md
      reason: "Tailored adoption mode: project-specific; use template only as reference."
      default_decision: SKIP
  deferred_protected_candidates:
    - path: .github/copilot-instructions.md
      source_commit: "2222222222222222222222222222222222222222"
      reason: "Updated stack-selection clause; awaiting owner authorization."
    - path: .github/instructions/powershell.instructions.md
      source_commit: "2222222222222222222222222222222222222222"
      reason: "Adds parameter-validation guidance; awaiting owner authorization."
```

### Validate

```bash
pre-commit run --all-files
npm run lint:md
npm run lint:md:nested
```

PowerShell validation:

```powershell
Invoke-Pester -Path tests/ -Output Detailed
```

### Sync Summary Fragment

```markdown
**Finalization mode:** PR-ready branch
**First-adoption preflight:** not applicable; existing `_TODO-repo-init.md` and `.template-sync/marker.yml` recorded resolved adoption answers.
**Unresolved `_TODO-repo-init.md` settings:** none
**Structural convention findings:** not applicable for this initialized delta sync
**Required structural changes:** none
**Post-adoption issue drafts:** none
**Adoption mode:** `minimal-preservation` by default for protected files and template-derived governance, community, process, workflow, and collaboration files; `README.md` uses a path-specific `tailored` local override.
**Upstream range reviewed:** `1111111111111111111111111111111111111111..2222222222222222222222222222222222222222`
**Included modules:** baseline, agent-instructions, github-actions, github-templates, markdown, powershell, template-sync-support
**Unadopted-module activity:** terraform (`.github/workflows/terraform-ci.yml`)
**Unknown modules or unmapped paths:** none
**Files adopted unchanged:** `TEMPLATE_UPDATE_PROCEDURE.md`, `templates/markdown/README.md`, `templates/markdown/getting-started.md` renamed from `templates/markdown/intro.md`
**Files manually merged:** `.github/workflows/powershell-ci.yml`, `.github/pull_request_template.md`
**Protected files deferred:** `.github/copilot-instructions.md` at `2222222222222222222222222222222222222222`, `.github/instructions/powershell.instructions.md` at `2222222222222222222222222222222222222222`
**Local overrides applied:** none in scope this sync
**Local customizations preserved:** self-hosted runner block; project-specific PR checklist
**Line-ending normalization:** not needed; `.gitattributes` unchanged.
**Validation:** `pre-commit run --all-files` (passed), `npm run lint:md` (passed), `npm run lint:md:nested` (passed), `Invoke-Pester -Path tests/ -Output Detailed` (passed)
**Saved candidate table:** `.cache/template-sync/candidates.md` for `1111111111111111111111111111111111111111..2222222222222222222222222222222222222222`
**Saved adoption ledger:** `.cache/template-sync/adoption-ledger.md`

## Issue Classifications

- upstream-template fix: none
- downstream-local fix: none
- deferred follow-up: none
```

## Future Automation

Future automation MAY add:

- a pre-commit hook that checks manifest coverage for managed paths
- a helper script that regenerates the Module Definitions and Path Mapping tables from `.template-sync/manifest.yml`
- a higher-level dry-run reporter that combines the candidate table with validation planning without applying changes

`.template-sync/manifest.yml` is authoritative for the taxonomy. During marker-backed sync, `.template-sync/marker.yml` is authoritative for included modules, local overrides, protected-file decisions, and deferred protected candidates. The candidate generator and generated adoption ledger are intentionally read-only review aids; this document remains the authoritative manual procedure.

<!-- template-sync: begin github-actions-only -->

## Workflow Security Governance

The `github-actions` module retains a self-contained workflow security contract, validator, shared parsing helper, schema, and standalone CI gate. Template-sync support also retains the validator, helper, and schema as trusted dependencies for future materialization. The manifest owns selection. Before staging, materialization requires a retained contract and checks that it covers every retained manifest-owned workflow in the source inventory. It then validates the reviewed source and renders the retained controls. Baseline selections also retain the always-running pre-commit hook and Data CI invocation. Language-specific requirements disappear with their owning modules.

Authorize `.github/workflow-security-contract.yml` explicitly as a protected policy path before first adoption or replacement. Record permitted local customization and the reviewed update decision. Adopter-created workflows remain outside baseline policy unless the owner explicitly enables `--strict`; synchronization preserves their bytes. Run `python .github/scripts/validate_workflow_security.py` after adoption and updates, and add `--verify-releases` when verifying changed action pins against upstream.

Removing GitHub Actions requires reviewed cleanup of excluded standalone files and replacement of retained shared files with their pruned forms. Materialization never silently deletes those files. Validate the final downstream tree for excluded-module leftovers, including hooks, updater entries, documentation, and policy references. Remove the shared validator and schema only when neither Actions nor template-sync support remains. The parsing helper also supports instruction enforcement; remove it only when all three modules are omitted. Dependabot's Actions updater survives only when both GitHub platform automation and Actions are retained; otherwise action updates need manual review.

See the [workflow security guide](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/docs/workflow-security.md) for the baseline/strict boundary, required authorization, exact command fingerprints, intentional failure-aggregation exceptions, and the SHA-pin vulnerability-alert limitation. Copilot prerequisite setup is available when `agent-instructions`, `agent-copilot`, and `github-actions` are selected together. Action-free acquisition remains a future opt-in profile.

<!-- template-sync: end github-actions-only -->
