<!-- markdownlint-disable MD013 -->

# Copilot Validation Setup

The [Copilot setup workflow](../.github/workflows/copilot-setup-steps.yml) prepares the validation tools selected during adoption. Retain `agent-instructions`, `agent-copilot`, and `github-actions` to receive this workflow and guide. Azure-only profiles and profiles without Copilot omit both files. A mixed GitHub/Azure profile can retain them for its GitHub sessions.

## Selected prerequisites

- `github-actions`: Python and the parser/schema dependencies used by the retained Workflow Security validator. The instruction enforcement module uses the same dependencies when selected.
<!-- template-sync: begin baseline-reference-only -->
- `baseline`: Python runtime and the exact runner in `requirements-pre-commit.txt`; the installer verifies the command's reported version. Hooks provision their own environments when run.
<!-- template-sync: end baseline-reference-only -->
<!-- template-sync: begin markdown-reference-only -->
- `markdown`: The same Node.js release line as Markdown CI and root lockfile dependencies using `npm ci --ignore-scripts`.
<!-- template-sync: end markdown-reference-only -->
<!-- template-sync: begin powershell-reference-only -->
- `powershell`: PSScriptAnalyzer and Pester 5 or newer, matching PowerShell CI's existing installation policy.
<!-- template-sync: end powershell-reference-only -->

Removing a module removes its setup steps during materialization. For example, retaining baseline and Markdown without the Python project still installs pre-commit and the root Node packages; it does not install a Python project. Retaining only shared instructions and Copilot with GitHub Actions installs the workflow validator's Python prerequisites without unrelated language tools. The workflow does not introduce a second package tree or hook manager.

## Activation and validation

GitHub requires this file on the repository's default branch to activate it for Copilot. It recognizes the single job named `copilot-setup-steps`; this template uses supported job permissions, runner, steps, and a 30-minute timeout. Ordinary pull-request checks and manual dispatch help test setup, but a passing branch check does not prove Copilot integration on the default branch. See [GitHub's environment setup documentation](https://docs.github.com/en/copilot/how-tos/copilot-on-github/customize-copilot/customize-cloud-agent/customize-the-agent-environment).

A failed setup step stops later setup steps, but Copilot may still start with the partially prepared environment. Inspect the session logs and verify retained tools before claiming validation. A successful setup is not a test pass: run the retained repository checks and keep required CI gates authoritative. Setup has read-only repository permissions, does not persist checkout credentials, and does not suppress installation failures.

## Updating or removing setup

Review tool installation changes together with the workflow's entry in the [workflow security contract](../.github/workflow-security-contract.yml). Existing package and tool pin sources continue to govern updates; action identities remain full SHAs with release annotations. Setup does not make transitive Python or PowerShell dependencies fully locked.

When removing Copilot or GitHub Actions, materialize the new selection and review its cleanup report. The materializer preserves existing files until reviewed removal; it does not silently delete them. Remove the excluded setup workflow and this guide, and authorize the corresponding protected workflow-contract update. Replace retained setup references with their pruned versions. Repeat materialization and downstream validation to check that no excluded references remain. Unrelated adopter workflows retain their local ownership.
