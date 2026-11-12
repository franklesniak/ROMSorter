<!-- markdownlint-disable MD013 -->

# PR and Code Review Prompts

- **Status:** Active
- **Owner:** Repository Maintainers
- **Last Updated:** 2026-09-29
- **Scope:** Ready-to-use prompts for responding to PR comments, code review
  feedback, branch management, and common false positives during code review.
- **Related:** [Copilot Chat Prompts for Template Adoption](../COPILOT_CHAT_PROMPTS.md)

## Purpose

This document captures prompts used during pull request and code review
workflows. These prompts are designed to be copied directly into GitHub PR
comments or Copilot Chat conversations.

To refine a proposed issue without filing it, use the [Issue Evaluation Prompt](ISSUE_EVALUATION_PROMPT.md).

<!-- template-sync: begin github-actions-reference-only -->
## Requesting Copilot Review and Recording Effort

Use this GitHub recipe with the local canonical Shared Review Governance and the retained agent's platform protocol. Prefer the GitHub plugin for capabilities it exposes; use a documented CLI/API fallback only for a missing capability. Before requesting, reconcile the current head and material PR description with existing requests and results. An accepted pending or unchanged clean request must not be duplicated.

1. When a supported interface exposes effort selection, choose **Balanced** before requesting review. GitHub's web UI places this selector under **Reviewers**, next to Copilot. If the available interface cannot select effort, send one supported request and accept Lite. Do not resend an accepted Lite request merely to change effort. See [GitHub's effort selection instructions](https://docs.github.com/en/copilot/how-tos/use-copilot-agents/request-a-code-review/use-code-review#choosing-a-review-effort-level).
2. For an authenticated GitHub CLI fallback, use `gh pr edit PR-NUMBER --repo OWNER/REPO --add-reviewer '@copilot'`, replacing the PR number and repository. Quote `'@copilot'` in PowerShell so it remains a literal argument rather than a splatting expression. This is a reviewer request, not an assignee change. The [CLI manual](https://cli.github.com/manual/gh_pr_edit) documents this special reviewer value; it is not supported on GitHub Enterprise Server.
3. For a supported REST fallback, send `POST /repos/{owner}/{repo}/pulls/{pull_number}/requested_reviewers` with the JSON `reviewers` array containing `copilot-pull-request-reviewer[bot]`. The CLI alias and REST bot login are different interfaces. See [GitHub's supported Copilot REST identity](https://docs.github.com/en/copilot/how-tos/use-copilot-agents/request-a-code-review/use-code-review) and [review-request API](https://docs.github.com/en/rest/pulls/review-requests#request-reviewers-for-a-pull-request). The documented CLI and REST request parameters do not expose an effort selector; do not invent one or claim they requested Balanced.
4. Confirm delivery through fresh authenticated native request/review/run evidence before requesting the separate remote reviewer. A successful command or HTTP response alone does not establish a completed clean review. Reconcile uncertain delivery before retrying; apply the canonical per-service limits and complete finding inventory.
5. Read the effort shown in the returned review's overview comment and record it separately from requested effort, alongside the native review link, identity and reviewed head. If the result does not disclose effort, record **unknown**, not an inference from timing or request method. GitHub documents [where actual effort is displayed](https://docs.github.com/en/copilot/concepts/agents/code-review#review-effort-level). Verify the authenticated actor, request linkage, timing and stable input before accepting a result.

If an interface, entitlement, or authenticated readback is unavailable, record the native failure and the required operator action; keep that gate incomplete. Do not store or replay browser cookies, CSRF tokens, nonces, credentials, or private internal form fields. This recipe supplies operating details, not new authority or altered retry/round limits. Recheck the linked primary documentation when interface behavior changes.
<!-- template-sync: end github-actions-reference-only -->


## Responding to Code Review Comments

### Agree with the Reviewer

Use this when you have reviewed the comment and agree with the feedback:

```markdown
I agree with the code reviewer's comment.
```

### Validate and Agree

Use this when you want to confirm the reviewer's concern before agreeing:

```markdown
Please double-check the code reviewer's recommendation. If the gap or concern
they pointed out is valid, then I agree with the code reviewer's comment.
```

### Evaluate, Decide, and Implement — Secondary Guide Prompt Only

Use this when the selected fix should be implemented, but a secondary style-guide change should be proposed for a separate task. This explicit prompt-only restriction applies to the secondary recommendation even when an earlier grant would permit that change.

```markdown
Read and follow Shared Review Governance and Protected Instruction Files in
.github/copilot-instructions.md. Validate this finding and complete the
finding-specific options, fresh weighted rubric, displayed scores, pre-edit
evaluation, authorized implementation, tests, guide-impact assessment and
disposition required there.

Keep the selected option fixed at the protected-content checkpoint. Use
specific authority already granted for this task; ask one narrow question for
an uncovered protected-content change and continue independent work. This
prompt grants no protected-content, branch-placement or merge authority.

For a secondary style-guide recommendation, return only a ready-to-file
prompt in a Markdown code fence. Include the proposed rule, rationale, scope
and acceptance tests. Do not implement that secondary guide change in this
task, even if earlier authority would allow it.
```

### Evaluate, Decide, and Implement — Authorized Secondary Guide Changes

Use this when specific guide changes have already been authorized in the task. Name or link that grant in the request. Copying this variant does not supply missing authorization.

```markdown
Read and follow Shared Review Governance and Protected Instruction Files in
.github/copilot-instructions.md. Validate this finding and complete the
finding-specific options, fresh weighted rubric, displayed scores, pre-edit
evaluation, authorized implementation, tests, guide-impact assessment and
disposition required there.

Implement a selected fix or secondary style-guide change only when specific
current-task authority covers its content. Keep the selected option fixed.
Do not ask again for unchanged authority already granted. For an uncovered
protected-content change, provide the ready-to-file prompt and ask the narrow
authorization question required by the canonical process. Continue independent
authorized work while that change is pending.

This prompt grants no protected-content, branch-placement or merge authority.
```

### Evaluate, Decide, and Implement — No Secondary Guide Proposal

Use this when the guide itself is the selected fix, or when no separate guide proposal is requested. A guide that is part of the selected fix still needs specific protected-content authority.

```markdown
Read and follow Shared Review Governance and Protected Instruction Files in
.github/copilot-instructions.md. Validate this finding and complete the
finding-specific options, fresh weighted rubric, displayed scores, pre-edit
evaluation, authorized implementation, tests and disposition required there.

Keep the selected option fixed at the protected-content checkpoint. Use
specific authority already granted for this task; ask one narrow question for
an uncovered protected-content change and continue independent work. This
prompt grants no protected-content, branch-placement or merge authority.

Assess guide impact, but do not prepare or implement a separate secondary
guide proposal. Record any relevant remaining limitation instead of describing
an unfinished requirement as complete.
```

### Azure DevOps PR Review Protocol Check

Use this when a review thread or PR is hosted in Azure DevOps Services with
Azure Repos rather than GitHub:


```markdown
Please handle this as an Azure DevOps Services / Azure Repos pull request, not
as a GitHub pull request.

Before acting, verify current Microsoft Learn behavior for Azure Repos Copilot
code review and Azure DevOps PR REST APIs. Keep the GitHub workflow intact and
apply the repository's Azure DevOps PR review protocol instead.

Account for these constraints explicitly:

- Azure Repos Copilot code review is a limited public preview that requires
  organization, repository, and individual-user enablement plus linked Azure
  billing; Azure DevOps review usage does not draw down GitHub Copilot plan AI
  credits. Treat licensing and pricing details as Microsoft Learn-dependent,
  and do not assume GitHub-hosted Copilot review entitlements cover Azure Repos
  review usage.
- Copilot review must be requested manually unless the available Azure DevOps
  tooling explicitly verifies an API-supported request path.
- Copilot leaves Comment reviews only, does not satisfy required-reviewer
  policies, does not read replies, does not follow up, and does not
  automatically re-review new commits.
- If Azure DevOps connector/API tooling is available and safely authenticated,
  use it for reviewers, PR threads, thread comments, thread status, and PR
  statuses. If tooling is absent, identify the manual owner action instead.
- Keep authentication guidance high-level and secure. Do not embed tokens,
  credential-bearing URLs, service connections, or secret-like placeholders in
  files, commands, logs, or comments.
```

## Branch Management

### Merge Main into Branch

Use this to catch a branch up with recent changes on `main`. Replace the
placeholder with the actual commit link:

```markdown
@copilot I need to catch this branch up with recent changes made to `main`, so
please merge `main` (at **link to commit here**) into this branch.
```

### Merge Main into Branch (Scoped to One File)

Use this variant when you need to catch the branch up with `main` while
ensuring that only a specific file remains modified in the PR. Replace the
placeholder commit link and filename:

```markdown
@copilot I need to catch this branch up with recent changes made to `main`, so
please merge `main` (at **link to commit here**) into this branch. After this
operation, only `File-We-Are-Working-On.xyz` should appear as modified in the
PR.
```

### Bring Branch Up to Date with Main

Use this when the branch is not up to date with `main`, causing extra files to
appear in the PR diff. Replace the placeholder commit link and filename:

```markdown
@copilot, it seems something got a bit off the rails. I don't believe this
branch is up to date with `main`. Please fix this by merging `main` (at **link
to commit here**) into the branch, so that the PR shows only
`File-We-Are-Working-On.xyz` as a modified file.
```

## Responding to False Positives

### Hallucinated Table Formatting Issue

The GitHub Copilot code reviewer sometimes flags a false positive related to
improperly formatted tables. Use this prompt in response:

```markdown
I believe this is a hallucination. I don't see any double pipes (`||` or
`| |`) in the table.
```

### Markdownlint Compliance Dispute

Use this when the GitHub Copilot code reviewer appears to falsely flag
markdownlint compliance:

```markdown
I'm OK with leaving it as is if it's currently markdownlint-compliant without
any markdownlint rule customizations. I think MD032 doesn't apply if the
previous line is an ordered list, as long as the unordered list item is
indented. If it's not markdownlint-compliant the way it currently is, fix it.
```

<!-- template-sync: begin powershell-reference-only -->
## Version and Compatibility Clarifications

### PowerShell Version Support

Use this when a new version of PowerShell is released and the script's version
support requirements need to be updated. Adjust the version numbers and release
context as needed:

```markdown
PowerShell 7.6.x was recently released. So, the script must support
Windows PowerShell 5.1, PowerShell 7.4.x, PowerShell 7.5.x, and PowerShell
7.6.x. Please ensure this requirement is thoroughly clarified.
```

<!-- template-sync: end powershell-reference-only -->
