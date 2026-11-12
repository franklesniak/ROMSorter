<!-- markdownlint-disable MD013 -->
# Agent Instructions for OpenAI Codex CLI

**Version:** 1.5.20261111.0

## Metadata

- **Status:** Active
- **Owner:** Repository Maintainers
- **Last Updated:** 2026-11-11
- **Scope:** Agent-specific entry point for OpenAI Codex CLI and compatible AI coding agents operating in this repository. Mirrors a minimal inline summary of the highest-priority shared rules; `.github/copilot-instructions.md` remains the canonical source of truth.
<!-- template-sync: begin markdown-reference-only -->
- **Related:** [Repository Copilot Instructions](.github/copilot-instructions.md), [Documentation Writing Style](.github/instructions/docs.instructions.md)
<!-- template-sync: end markdown-reference-only -->

This file provides project-specific instructions for OpenAI Codex CLI and compatible AI coding agents operating in this repository. These instructions ensure that agents follow the same coding standards, safety rules, and workflows that apply to all contributors.

## Canonical Instructions

The authoritative source of truth for all repository rules is **`.github/copilot-instructions.md`** (the repo-wide constitution). All rules defined there apply without exception. **Read that file before making any changes.**

This file intentionally keeps only a minimal inline summary of the highest-priority shared rules so that agents receive critical guidance immediately. The full shared rule set remains in the canonical file above.

**Thin entry point classification:** A thin entry point keeps shared repository rules brief; it does not mean platform-specific or required protocol sections may be discarded. Sections explicitly labeled as platform protocol or required protocol must be preserved unless the repository owner explicitly waives that protocol for the retained agent platform.

<!-- template-sync: begin github-actions-reference-only -->
For supported Copilot request interfaces and observed effort, use the [shared Copilot review recipe](docs/PR_REVIEW_PROMPTS.md#requesting-copilot-review-and-recording-effort). The canonical review policy and this entry point's platform protocol remain authoritative.
<!-- template-sync: end github-actions-reference-only -->

## Execution

Agents MUST follow [Agent Execution](.github/copilot-instructions.md#agent-execution) for task input records, unrelated-work preservation, exclusive ownership, bounded delegation, verification, and continuity after interruption. Use only capabilities available in the active runtime.

Agents MUST apply the [shared decision process](.github/copilot-instructions.md#shared-review-governance) to material non-review findings as well as review feedback. Preserve its narrow mechanical-reuse conditions and its distinction between general decisions and PR-specific duties.

For runtimes that follow [OpenAI's documented instruction discovery](https://learn.chatgpt.com/docs/agent-configuration/agents-md), Codex builds its chain at startup. Global guidance precedes project guidance from the root through the launch working directory. Each directory contributes at most one file: `AGENTS.override.md`, then `AGENTS.md`, then configured fallback names. Deeper guidance takes precedence within its scope, subject to higher-priority instructions. Empty files and the configured byte limit affect loading; verify the active runtime rather than assuming every file loaded.

Before editing in a deeper directory, Codex MUST perform bounded discovery along the relevant file paths and read applicable instructions. Startup discovery does not automatically cover every descendant directory. If active guidance is stale, use the runtime's supported refresh; the documented CLI procedure is to restart in the target directory. Preserve the shared disk-recovery requirements and exact-file rereads after compaction or instruction changes. This guidance does not require changing configuration or apply Codex loading semantics to other agents.

## Protected Instruction Files

Instruction files, style guides, and the instruction-contract catalog (when retained) are protected governance files. Do not create, edit, delete, rename, or otherwise change `.template-sync/instruction-contracts.yml`, `.github/copilot-instructions.md`, files under `.github/instructions/`, files under `.cursor/rules/`, or root agent instruction files (`.hermes.md`, `AGENTS.md`, `CLAUDE.md`, `GEMINI.md`) unless the repository owner or maintainer has directly and explicitly authorized that specific protected-governance change in the current task. Implied consent is not enough; do not infer authorization from a plan you generated, review feedback, a general request to update docs, cleanup/validation work, or a "keep files in sync" instruction.

If a style-guide update appears warranted but has not been explicitly authorized, propose it separately and wait for approval before editing protected instruction files.

During downstream template adoption and stack selection, perform non-protected cleanup first, record the protected instruction-file edits needed to remove references to deleted tools or stacks, obtain explicit maintainer authorization, then update `.github/copilot-instructions.md`, remaining root agent files, and relevant `.github/instructions/*.instructions.md` files. Bump `Last Updated` and `Version` metadata where present, and avoid temporary migration wording in durable governance docs.

## Essential Repository Summary

- **Safety and security**
  - No secrets in code or repo; never hardcode API keys, tokens, credentials, or connection strings.
  - Treat all external input as untrusted.
  - Respect allowlisted file access boundaries; reject path traversal and symlink escapes.

- **Pre-commit and validation**
  <!-- template-sync: begin baseline-reference-only -->
  - Run `pre-commit run --all-files` before every commit.
  <!-- template-sync: end baseline-reference-only -->
  - Include all auto-fixes in the same commit as the related change.
  - Do not push code when pre-commit or required validation checks are failing; fix issues and re-run until the checks pass.
  - Use the repository's existing validation commands as needed:
    <!-- template-sync: begin markdown-reference-only -->
    - `npm run lint:md`
    - `npm run lint:md:nested`
    <!-- template-sync: end markdown-reference-only -->
    <!-- template-sync: begin powershell-reference-only -->
    - `Invoke-Pester -Path tests/ -Output Detailed`
    <!-- template-sync: end powershell-reference-only -->
  <!-- template-sync: begin baseline-reference-only -->
  - The `pre-commit run --all-files` command exercises the active hooks configured in [`.pre-commit-config.yaml`](.pre-commit-config.yaml), the authoritative list of active hooks.
  <!-- template-sync: end baseline-reference-only -->
  <!-- template-sync: begin json-reference-only -->
  - Retained JSON checks include strict JSON syntax (`check-json`).
  <!-- template-sync: end json-reference-only -->
  <!-- template-sync: begin yaml-reference-only -->
  - Retained YAML checks include YAML parsing (`check-yaml`) and style (`yamllint`).
  <!-- template-sync: end yaml-reference-only -->
  - Retained GitHub Actions checks include GitHub Actions linting (`actionlint`).
  - Retained Azure Pipelines assets require host-neutral local hooks plus Azure DevOps Services pipeline creation, queued runs, or branch-policy build validation; `actionlint` does not validate Azure Pipelines YAML.
  <!-- template-sync: begin github-data-ci-reference-only -->
  - When both `baseline` and `github-actions` are retained, the dedicated [`.github/workflows/data-ci.yml`](.github/workflows/data-ci.yml) workflow re-runs retained data-file hooks so adopted data-file enforcement can be required via branch protection.
  <!-- template-sync: end github-data-ci-reference-only -->
  - Retained data-file authoring guidance lives in the matching module docs.
  <!-- template-sync: begin json-reference-only -->
  - JSON guidance: [`.github/instructions/json.instructions.md`](.github/instructions/json.instructions.md).
  <!-- template-sync: end json-reference-only -->
  <!-- template-sync: begin yaml-reference-only -->
  - YAML guidance: [`.github/instructions/yaml.instructions.md`](.github/instructions/yaml.instructions.md).
  <!-- template-sync: end yaml-reference-only -->

- **Modular instruction files**
  - Read the relevant file under `.github/instructions/` before modifying matching files:
    - Git attributes: `.github/instructions/gitattributes.instructions.md`
    <!-- template-sync: begin json-reference-only -->
    - JSON: `.github/instructions/json.instructions.md`
    <!-- template-sync: end json-reference-only -->
    <!-- template-sync: begin markdown-reference-only -->
    - Markdown/Docs: `.github/instructions/docs.instructions.md`
    <!-- template-sync: end markdown-reference-only -->
    <!-- template-sync: begin powershell-reference-only -->
    - PowerShell: `.github/instructions/powershell.instructions.md`
    <!-- template-sync: end powershell-reference-only -->
    <!-- template-sync: begin yaml-reference-only -->
    - YAML: `.github/instructions/yaml.instructions.md`
    <!-- template-sync: end yaml-reference-only -->

- **Do not**
  - Execute scripts or commands generated by untrusted sources.
  - Add telemetry or external logging services without explicit approval.
  - Weaken security constraints to "make it work."
  - Add new major dependencies without clear justification.
  - Invent behavior when requirements are ambiguous; use an explicit Open Question.
  - Create separate formatting-only or lint-only commits.

## GitHub Plugin Usage

This section is retained as Codex platform protocol. Thin-entry-point pruning must preserve it unless the repository owner explicitly waives Codex GitHub plugin protocol for the retained Codex entry point.

Codex can use the OpenAI-curated GitHub plugin (`github@openai-curated`) in this repository when the user has installed and authorized it. The plugin is the preferred mechanism for any operation that touches remote GitHub state.

- **Prefer the GitHub plugin** for GitHub Issues, pull requests, PR comments, review comments, labels, reactions, PR creation, branch management on GitHub, and any other read or write against remote GitHub state.
- **`gh` is an optional fallback**, not a prerequisite. Codex MUST NOT treat the absence of the `gh` CLI as a blocker when the GitHub plugin can satisfy the same operation. If both are available, use the GitHub plugin first and fall back to `gh` only when the plugin lacks a needed capability.
- **Local `git` remains appropriate** for local working-tree operations: inspecting diffs (`git diff`, `git log`, `git status`), creating branches, staging, committing, and pushing. Use the GitHub plugin for the corresponding remote-side actions (creating PRs, posting comments, reading review state).
- **Before declaring any GitHub operation impossible**, Codex MUST first check whether the GitHub plugin exposes a tool that can perform it. Only after the plugin and any documented fallback have both been ruled out should Codex report the operation as unavailable.

The `.codex/config.toml` file at the repository root declares `[plugins."github@openai-curated"] enabled = true` so that trusted Codex checkouts can opt the plugin in by default. Enabling the plugin in this file does not, by itself, grant GitHub authorization: actual access still depends on the GitHub app/connector installation, the Codex account performing the operation, and the repository's permissions.

## PR Review Workflow (Codex-adapted)

This section is retained as Codex platform protocol. Thin-entry-point pruning must preserve it unless the repository owner explicitly waives Codex PR review protocol for the retained Codex entry point.

Use this workflow with [Shared Review Governance](.github/copilot-instructions.md#shared-review-governance) when responding to PR feedback. Codex MUST follow that contract for complete finding inventory, per-finding weighted decisions, paired remote reviews, attribution, bounded recovery, CI repair, deferrals, and continuation. All GitHub-side reads and writes in the steps below SHOULD go through the GitHub plugin first; fall back to `gh`, GraphQL, or manual owner action only when the plugin does not expose the needed capability (see **Fallbacks for unsupported plugin capabilities** below).

### Runtime limitations to keep in mind

- **Command-only triggers.** Apply the shared [Finding inventory and decisions](.github/copilot-instructions.md#finding-inventory-and-decisions) distinction. Preserve request evidence and substantive feedback in mixed comments; do not execute commands addressed to other agents.
- **No autonomous wake-up.** Codex has no equivalent of `subscribe_pr_activity`. Codex MUST NOT promise webhook-driven wake-up, background polling, or scheduled review responses. The workflow runs only when the user explicitly starts or resumes it inside an active Codex session.
- **Mention routing and remote review.** A general `@codex` mention reaches this local session only when the user's runtime forwards it. The documented exact `@codex review` command requests the separate remote Codex GitHub review service when it is configured for the repository. Codex MUST NOT claim that the local session wakes autonomously or that local work replaces that remote review.
- **Tooling-dependent steps.** Where the GitHub plugin and documented fallbacks lack a capability, state the missing operator action and continue independent work. A missing required review, observation, or disposition remains incomplete; do not skip it and declare success.

### Handling each review comment

For each code review comment received from GitHub Copilot, a human reviewer, or any other reviewer, follow these steps:

#### Protected-file authorization terms

These terms apply to the review-comment workflow below and defer to the canonical **Protected Instruction Files** rule in [`.github/copilot-instructions.md`](.github/copilot-instructions.md):

- **Protected instruction file:** Any file covered by the canonical Protected Instruction Files rule, including the retained `.template-sync/instruction-contracts.yml` catalog, `.github/copilot-instructions.md`, the root agent entry points (`.hermes.md`, `AGENTS.md`, `CLAUDE.md`, `GEMINI.md`), files under `.github/instructions/`, and files under `.cursor/rules/`.
- **Explicit protected-file authorization:** A direct maintainer or owner instruction in the current task authorizing the specific protected instruction-file change, either by naming the file or by clearly bounding the protected-file change set. The following are not sufficient on their own: a PR existing, a review comment existing, a generic "address the feedback" request, a reusable prompt, an automated review loop or active review workflow, or generic branch-placement authorization.
- **Already in the PR's scope:** The protected file appears in the PR's changed-files list or diff against its base branch before the review-driven edit under consideration. This is relevant context, not authorization.
- **Newly introduced protected file:** A protected file the PR did not modify before the review-driven edit. Introducing one exceeds any authorization scoped to the PR's existing changes and requires the narrow authorization question in step 7.
- **Within the already-authorized scope:** An edit that resolves the reviewer's comment without expanding the protected file's changes beyond the specific protected-file change the maintainer already explicitly authorized for this task. A larger or more structural change, or one that newly introduces a protected file, exceeds the already-authorized scope.
- **Secondary style-guide recommendation:** A step-8 recommendation to update a style guide to prevent similar issues in the future, distinct from the selected step-7 fix for the current review comment.

1. **Signal processing (conditional).** If the GitHub plugin (or a documented fallback) supports adding emoji reactions to review comments, add an `eyes` (👀) reaction when work begins on the comment and remove it when the comment is fully processed (after step 9, or after the early-exit path in step 2). The reaction's `content` value is the literal string `eyes` as used by the GitHub Reactions API, not the Markdown shortcode `:eyes:`. If reaction tooling is not available in the current runtime, skip this step silently.

2. **Validate the concern.** Determine whether the reviewer's feedback identifies a genuine gap, bug, style violation, or improvement opportunity. If the concern is not valid, post a reply explaining why through the GitHub plugin (or fallback), skip steps 3-8, and continue to step 9.

3. **List options.** Follow the shared finding process: validate and research the concern, then finish all materially distinct reasonable options and useful combinations before defining the rubric.

4. **Build an evaluation rubric.** Define and display a fresh weighted rubric with 4–6 finding-specific criteria and a 1–5 scale before scoring. Explain the weights and prioritize substantive correctness, security, portability, and maintainability over effort and churn.

5. **Score and select.** Display every option's criterion scores, weights, and checked weighted total before selecting the highest-supported eligible option. Follow the shared technical-tie rule; a small margin or a topic touching governance alone does not require escalation. Keep the selected option fixed at the protected-file checkpoint. Apply Escalation protocol below when escalation is necessary.

6. **Post the evaluation.** Publish the complete shared finding evaluation before editing, through the GitHub plugin or a documented fallback. Include the options, fresh rubric, score table, selected action, references, tests, and intended placement. After implementation, reply with the implementing PR-head SHA and validation evidence.

7. **Implement the fix.** Apply the selected option locally, commit, and push to the agent's working branch using local `git`. Apply the Protected content and placement section below before editing or placing the fix.

8. **Evaluate style guide impact.** Read the full applicable guide and follow the shared prevention and deferral rules. Implement a secondary guide change when explicit current-task authority covers it. Otherwise, post the required ready-to-file fenced prompt. For a GitHub inline finding, use the same native review thread. For a GitHub body-only finding, use a standalone PR comment that records its synthetic finding key, source review identity, reviewed commit, and location when available. Record the prompt comment's native identity, keep the guide action pending until attributable disposition, and continue independent authorized work.

9. **Resolve or leave open.** Reply with the fix or refutation evidence and close a complete finding when no guide action remains pending. For body-only findings, record an attributable disposition. Inspect the runtime's actual reaction and thread-resolution capabilities; retrieve the GraphQL thread ID through an authenticated equivalent when needed. If resolution is unavailable, state the manual owner action and leave it incomplete. Remove the temporary `eyes` reaction when supported. A resolved flag alone is not proof that the finding is complete.

### Escalation protocol

**Escalation path.** Escalate only for a decisive owner preference, new authority, or a material change to explicit scope or intended outcome that available evidence cannot resolve. Post a **standalone PR comment** (not a reply to the review thread) through the GitHub plugin containing:

- A brief summary of the reviewer's concern and which file/line it applies to
- The options and scoring tables
- The specific question the owner needs to answer
- Instructions: *"Reply to this PR comment with your chosen option or direction, then bring the reply back to your active Codex session so Codex can act on it. Posting `@codex` in the reply only routes the comment to Codex when the user's runtime is configured to forward it."*

**PAUSE** processing of this comment until the owner responds. Continue processing other independent review comments in the meantime.

### Protected content and placement

Codex MUST apply [Safe PR-head placement](.github/copilot-instructions.md#safe-pr-head-placement) whenever direct placement is authorized. The procedure also applies to an available equivalent API mechanism; it does not replace the explicit PR-specific authorization and conditions below.

Apply the shared [Continuity and recovery](.github/copilot-instructions.md#continuity-and-recovery) rule to verified task grants. Codex requires an explicit PR-specific task grant for direct placement; Claude's documented active-loop grant applies only under its own preconditions. This is an intentional platform policy difference, not a transferable grant. Neither placement rule supplies protected-content or merge authority.

**Protected-file authorization checkpoint.** Before creating, editing, deleting, renaming, or otherwise changing any protected instruction file, including a style guide under `.github/instructions/`, determine whether explicit protected-file authorization already covers that specific protected-file content change in the current task. Keep the selected option fixed while making this authorization determination; do not reopen option selection or ask the maintainer to choose among the scored options again merely because protected-file authorization is required.

- If explicit protected-file authorization already covers the change and the edit stays within the already-authorized scope, proceed with the selected option under the placement rules below.
- Otherwise, including when no explicit authorization exists, when the intended edit exceeds the already-authorized scope, or when the edit would newly introduce a protected file the PR did not previously modify, ask one narrow authorization question before editing. The question states the selected option, the protected file, the intended change, the agent's recommendation, and, when applicable, that the protected file is already in the PR's scope. Ask the question in the active Codex session when the owner is present; if the workflow is mediated through PR comments, post it as a standalone PR comment through the GitHub plugin and wait for the user to bring the maintainer's authorization back to the active Codex session.
- If authorization is declined, record the decision and resolve or leave the review thread according to step 9.

This checkpoint governs only authorization to change protected-file content. It does not expand any existing authorization for direct PR-head placement, which continues to govern only where an authorized commit lands. To make the fix visible on the PR (i.e., reachable from the PR's head ref), choose the appropriate placement strategy:

- **Default — push to the working branch only.** Cross-branch integration onto the PR head is a manual owner action. State in the step-6 reply which branch the commit will be pushed to and whether a merge or cherry-pick will be required to make it visible on the PR head.
- **Direct PR-head push (only with explicit user authorization).** Codex MAY push directly to the PR head branch only when **all** of the following hold:

1. The user has **explicitly authorized** direct PR-head pushes for this specific PR in the current task, including a verified resume under the shared continuity rule. Implied consent is not enough; do not infer authorization from a general "address the review" instruction.
2. The PR head branch is in the **same repository** as the agent's working branch (cross-fork PRs are excluded).
3. The push is **non-destructive**: no force-push and no history rewrite on the PR head branch.
4. All branch protections, required status checks, signing requirements, and CI/CD validation rules on the PR head branch continue to be satisfied.

When direct PR-head placement is used, record the resulting PR-head commit SHA(s) and post a follow-up reply confirming the placement and listing those SHA(s). If any of the four conditions above is not met, fall back to the default behavior.

### Optional user-initiated review cycle

When the PR owner explicitly asks Codex to drive multiple review rounds inside an active session (for example, *"run the review cycle on PR #N"*), Codex MAY iterate on the following loop. The loop runs only while the Codex session is active; it MUST NOT be presented as autonomous.

1. **Request the reviewer pair.** Apply the shared input, attribution, and recovery rules. Capture both native baselines, prefer supported Copilot Balanced selection, and accept Lite fallback. Use the GitHub plugin first; the documented CLI fallback is `gh pr edit PR-NUMBER --add-reviewer '@copilot'`. Confirm Copilot delivery before sending the exact remote trigger `@codex review`. Do not duplicate a pending request.
2. **Observe both services.** While the session is active, poll each pending service at least 60 seconds apart through authenticated tooling. Collect complete reviews, all-state inline threads with nested pagination, conversation results, and relevant runs. Codex MUST apply the shared ten-observation limit independently to each pending reviewer, the five-observation missing-body limit, and the terminal-failure/retry table. Repeated stale results or one completed reviewer cannot freeze the other's counter. Missing capability or exhausted recovery leaves the gate incomplete.
3. **Process every finding.** Use the per-comment and shared finding process, including review-body findings. Reuse only a still-valid evidence-backed disposition. Continue independent fixes and CI repair while either reviewer is pending. One clean service cannot complete the pair.
4. **Place fixes and reconcile input.** Re-request both services after a new head only after all intended fixes are reachable from the PR head and older in-flight requests are reconciled. If fixes exist only on the working branch, use step 7's placement rules and report the resulting PR-head SHAs. Without direct-head authority, pause for owner integration. Material description changes invalidate affected reviews; status-only updates do not justify another request.
5. **Complete the gate.** Verify both attributable clean reviews on the final unchanged input, all earlier finding dispositions, successful required CI, and any required independent checks. Report the actual remaining limitation when a gate is incomplete. Review completion grants no new merge authority.

### Safety limits for the optional review cycle

When the optional review cycle is used, retain these finite safety limits:

- **Maximum rounds:** 8 review iterations per cycle invocation. After the eighth round, PAUSE and ask the user to confirm whether to continue.
- **Wall-clock timeout:** 6 hours from cycle start. If the timeout is reached, PAUSE and ask the user to confirm whether to continue.
- **Recovery and disposition:** Track native identities, current finding text, and evidence. On explicit resume, recover both service states, authority, input, fix reachability, and retry counters before acting. Finish authorized paused work; do not reset a same-input retry budget or blindly request new reviews.

### Fallbacks for unsupported plugin capabilities

When a workflow step depends on a capability the GitHub plugin does not currently expose, use the following fallbacks and document the chosen fallback in the relevant reply or PR comment:

| Capability | Primary | Fallback |
| --- | --- | --- |
| Request a Copilot code review | GitHub plugin | `gh pr edit PR-NUMBER --add-reviewer '@copilot'`, supported UI, or owner request; confirm delivery |
| Request remote Codex review | GitHub plugin PR comment | Authenticated PR comment containing exact `@codex review`; owner action if unavailable |
| Resolve a review thread | GitHub plugin | `gh api graphql` against the `resolveReviewThread` mutation, or ask the owner to resolve the thread manually |
| Add a reaction on a review comment | GitHub plugin | `gh api -X POST /repos/{owner}/{repo}/pulls/comments/{comment_id}/reactions -f content=eyes`, or skip silently if neither path is available |
| Remove a reaction on a review comment | GitHub plugin | First list the comment's reactions via `gh api /repos/{owner}/{repo}/pulls/comments/{comment_id}/reactions` and find the reaction whose `content` matches the one to remove (e.g. `eyes`) and whose `user.login` is the agent's own identity; then delete by reaction id via `gh api -X DELETE /repos/{owner}/{repo}/pulls/comments/{comment_id}/reactions/{reaction_id}`. Skip silently if neither path is available |
| Post a reply to a review comment thread | GitHub plugin | `gh api` against `/repos/{owner}/{repo}/pulls/{pr}/comments/{comment_id}/replies`, or post a standalone PR comment that quotes the original review comment |

If the primary capability and all listed fallbacks are unavailable, record the incomplete step and needed operator action. Continue independent authorized work. Do not claim a required gate passed.

---

> This file is part of the `franklesniak/copilot-repo-template` template. Customize or remove agent instruction files for platforms you do not use. See [OPTIONAL_CONFIGURATIONS.md](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/OPTIONAL_CONFIGURATIONS.md) for details.
