# Contributing to This Project

<!--
TEMPLATE DESIGN DECISION: Placeholder Strategy

This file uses franklesniak/ROMSorter placeholders (not generic <your-repo> syntax) because:
- Enables bulk find-and-replace for template adopters (single operation)
- CI automation can verify all placeholders are replaced (.github/workflows/check-placeholders.yml)
- Results in working, copy-pastable commands after replacement
- Consistent with issue templates and other template files

Alternative considered: Generic angle-bracket syntax like <your-repository-clone-url>
Rejected because: Harder to replace in bulk, produces non-working commands,
inconsistent with other files that require real values (CI configs, package.json)

See README.md Template Setup Checklist for adoption instructions.
See [OPTIONAL_CONFIGURATIONS.md](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/OPTIONAL_CONFIGURATIONS.md) for detailed customization guidance.
-->

Thank you for your interest in contributing. This document describes the retained contributor workflow, validation expectations, and pull-request readiness checks for this repository.


## Development Setup

### 1. Clone the Repository

<!-- CUSTOMIZE: Replace `franklesniak/ROMSorter` with your organization and repository name -->

```bash
git clone https://github.com/franklesniak/ROMSorter.git
cd REPO
```

<!-- template-sync: begin markdown-reference-only -->
### 2. Install Node.js Dependencies

Install the Markdown tooling from the committed lockfile without running install lifecycle scripts:

```bash
npm ci --ignore-scripts
```

This keeps normal setup aligned with `package-lock.json` and reports a mismatch instead of rewriting the lock. For intentional dependency maintenance, use the appropriate `npm install --ignore-scripts` or `npm update --ignore-scripts` command, review the package and lockfile changes, and run validation. Later lint and test commands execute tools; disabling install scripts does not prevent that execution.

Nested Markdown validation limits each input to 1 MiB, 64 nested Markdown fences, 1,024 extracted blocks, and 8 MiB of cumulative extracted content. Exceeding a limit fails with a diagnostic; split unusually large examples into smaller files. Claude hosted sessions install the same locked dependencies automatically when Markdown is retained. The Claude hook first checks for npm and a stable Node.js version of 22 or newer. Missing commands or an unsupported or unreadable Node version stop setup with a prerequisite error before npm runs.
<!-- template-sync: end markdown-reference-only -->

Git hooks are managed by pre-commit.

### Install Python

`pre-commit` is installed and run through `pip`/`pipx`, and the template-sync helper scripts run on Python, so a working Python 3 interpreter is required even when your project does not adopt the Python module:

```bash
python --version
```


### 3. Install Pre-commit

From the repository root, install the version in `requirements-pre-commit.txt` using **one** of the following. With `pip` in your chosen Python environment:

```bash
python -m pip install -r requirements-pre-commit.txt
```

Or, for isolated tooling, with `pipx`:

```bash
pipx install --force "$(cat requirements-pre-commit.txt)"
```

`pre-commit` manages isolated hook environments, so it does not need to be installed as a project runtime dependency. The requirement pins the runner directly; it does not lock transitive packages. CI derives its expected version from this file and rejects a mismatched command. For PowerShell pipx setup, use `pipx install --force (Get-Content -Raw requirements-pre-commit.txt).Trim()`.

### 4. Install Git Hooks

```bash
pre-commit install
```

### 5. Run Pre-commit Manually

Run all hooks across the whole repository:

```bash
pre-commit run --all-files
```

Or run hooks against only the staged changes:

```bash
pre-commit run
```

Run the all-files form before opening a pull request.

## Git Hooks

This repository uses pre-commit for git hooks. Configured hooks include:

- **Formatting:** trailing whitespace, end-of-file, and line-ending fixes.
- **Markdown linting:** markdownlint and local Markdown link validation.
- **PowerShell linting:** PSScriptAnalyzer with repository settings.
- **Data-file validation:** `check-json` for strict `.json` files, `check-yaml`, and `actionlint` for GitHub Actions workflows.
<!-- template-sync: begin yaml-reference-only -->
- **YAML style validation:** `yamllint` configured by `.yamllint.yml`.
<!-- template-sync: end yaml-reference-only -->
- **Schema validation:** When schema-backed modules (the schema or template-sync support module) are retained, `check-jsonschema` and `check-metaschema` validate retained schema-backed configuration, including template-sync schemas.
- **Safety:** large file detection.

**`check-json` validates strict `.json` only.** It does **not** validate `.jsonc`. JSONC files are allowed only when the consuming tool supports JSONC; downstream repositories that need stricter `.jsonc` enforcement should add JSONC-aware tooling rather than retrofitting `check-json`.

<!-- template-sync: begin json-reference-only -->
See [`.github/instructions/json.instructions.md`](.github/instructions/json.instructions.md) for the full JSON/JSONC dialect policy.
<!-- template-sync: end json-reference-only -->
<!-- template-sync: begin yaml-reference-only -->
See [`.github/instructions/yaml.instructions.md`](.github/instructions/yaml.instructions.md) for YAML authoring standards.
<!-- template-sync: end yaml-reference-only -->

**`actionlint` first-run-on-restricted-networks caveat.** The `actionlint` pre-commit hook builds the `actionlint` binary from source on first install, which downloads a Go toolchain. On networks that block Go module downloads, the first-run install can fail. CI is the shared enforcement environment, so contributors who hit a network restriction locally can rely on CI to enforce this hook. The same caveat is documented inline in `.pre-commit-config.yaml` and in [`.github/TEMPLATE_DESIGN_DECISIONS.md`](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/.github/TEMPLATE_DESIGN_DECISIONS.md).


If you need to bypass hooks temporarily, which is not recommended:

```bash
git commit --no-verify -m "your message"
```

## Manual Validation

<!-- template-sync: begin markdown-reference-only -->
### Markdown Linting

```bash
npm run lint:md
npm run lint:md:links
npm run lint:md:nested
```
<!-- template-sync: end markdown-reference-only -->

### PowerShell Validation

```powershell
Invoke-ScriptAnalyzer -Path .\script.ps1 -Settings .\.github\linting\PSScriptAnalyzerSettings.psd1
Invoke-Pester -Path tests/ -Output Detailed
```


<!-- template-sync: begin template-sync-support-reference-only -->
### Template-Sync Validation

When template-sync support is retained, use the downstream validator after module pruning or marker changes:

```bash
python .template-sync/scripts/validate_downstream_adoption.py --require-marker
```

For cleanup planning, run the excluded-module reporter:

```bash
python .template-sync/scripts/report_excluded_module_references.py
```
<!-- template-sync: end template-sync-support-reference-only -->

## Code Quality Standards

### Pre-commit Discipline

**Always run pre-commit checks before committing code.**

Pre-commit hooks are not optional. They enforce:

- Formatting, end-of-file, and line-ending hygiene.
- Markdown linting and local Markdown link validation.
- PowerShell linting.
- Data-file validation for strict `.json`, YAML parsing, and GitHub Actions workflows.
<!-- template-sync: begin yaml-reference-only -->
- YAML style validation through `yamllint`.
<!-- template-sync: end yaml-reference-only -->
- Schema validation for retained schema-backed file families.

> **Network and dialect notes:**
>
> - First-run hook setup may require network access.
> - `check-json` validates strict `.json` files only and does **not** validate `.jsonc`.
> - CI runs the same hooks and is the shared enforcement environment.
> - Auto-fixes produced by hooks **must** be committed with the related change.

See `.pre-commit-config.yaml` for the complete list of configured hooks.

**Workflow:**

1. Make your code changes.
2. Run `pre-commit run --all-files`.
3. Review and commit all auto-fixes as part of your change.
4. Push to GitHub.

**If pre-commit CI fails:**

1. Pull the latest branch.
2. Run `pre-commit run --all-files` locally.
3. Integrate the auto-fixes into the commit that introduced the change.
4. Push again.

CI is a safety net, not a substitute for local checks.

### Language-Specific Guidelines

This repository includes coding standards for retained languages and data file formats:

- **Markdown/Documentation:** [`.github/instructions/docs.instructions.md`](.github/instructions/docs.instructions.md)
- **PowerShell:** [`.github/instructions/powershell.instructions.md`](.github/instructions/powershell.instructions.md)
- **Git attributes:** [`.github/instructions/gitattributes.instructions.md`](.github/instructions/gitattributes.instructions.md)
<!-- template-sync: begin json-reference-only -->
- **JSON/JSONC:** [`.github/instructions/json.instructions.md`](.github/instructions/json.instructions.md)
<!-- template-sync: end json-reference-only -->
<!-- template-sync: begin yaml-reference-only -->
- **YAML:** [`.github/instructions/yaml.instructions.md`](.github/instructions/yaml.instructions.md)
<!-- template-sync: end yaml-reference-only -->

These standards apply to all contributions in the retained file families and are enforced by the repository's pre-commit hooks and CI workflows.

### Template Taxonomy Updates

When a change adds, removes, renames, moves, or changes the primary purpose of a template-managed file, update the downstream sync path mapping in `TEMPLATE_UPDATE_PROCEDURE.md` as part of the same change. See [Reviewing the Template Sync Module Taxonomy](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/TEMPLATE_MAINTENANCE.md#reviewing-the-template-sync-module-taxonomy) for the maintainer checklist.

### Data-File and Schema Validation

Data files and GitHub Actions workflow files are first-class validation surfaces. Contributor expectations:

- Run `pre-commit run --all-files` before opening a PR.
- Strict `.json` files must pass `check-json`; `.jsonc` is intentionally not validated by that hook.
- YAML files must pass `check-yaml`.
<!-- template-sync: begin yaml-reference-only -->
- YAML files must also pass `yamllint` when the YAML module is retained.
<!-- template-sync: end yaml-reference-only -->
- GitHub Actions workflow files must pass `actionlint`.
- Schema-backed files must pass `check-jsonschema`, and retained schemas must pass `check-metaschema` where configured.

### CI Workflows

This repository includes retained GitHub Actions workflows that run automatically:

- **Pre-commit CI** (`.github/workflows/precommit-ci.yml`) - Runs the aggregate `pre-commit run --all-files` gate over every hook in `.pre-commit-config.yaml`.
- **Auto-fix Pre-commit** (`.github/workflows/auto-fix-precommit.yml`) - Produces an untrusted fix preview on matching Copilot-agent branches. The workflow wrapper does not commit or push; review or reproduce the proposed fixes locally and run the required checks before committing them with the substantive change.
- **Markdown Lint** (`.github/workflows/markdownlint.yml`) - Validates Markdown formatting and local links.
- **PowerShell CI** (`.github/workflows/powershell-ci.yml`) - Runs PSScriptAnalyzer and Pester on PowerShell files.
<!-- template-sync: begin github-data-ci-reference-only -->
- **Data CI** (`.github/workflows/data-ci.yml`) - Runs retained baseline placeholder, data-file, GitHub Actions, template-sync, and schema validation hooks.
<!-- template-sync: end github-data-ci-reference-only -->

The **Auto-fix Pre-commit** workflow is scoped specifically to the GitHub Copilot Coding Agent. Human-authored PRs and PRs on non-`copilot/**` branches are not affected; their authors must run `pre-commit run --all-files` locally and integrate the fixes themselves before pushing.

### Workflow Version Pinning

When editing files under `.github/workflows/`, follow the [**Workflow Version Pinning**](.github/copilot-instructions.md#workflow-version-pinning) rule in the repo-wide constitution. In short:

- Keep third-party action versions directly in `uses:` lines so Dependabot can update them.
- Do **not** mirror a `uses:` version into a workflow-level `env:` variable, comment, cache key, file path, or shell literal.
- For tool versions that Dependabot does not manage, a workflow-level `env:` value is a fine single source of truth across multiple steps.
- If a Dependabot-managed dependency genuinely cannot be expressed without duplication, add a narrowly scoped `.github/dependabot.yml` `ignore:` entry with a YAML comment explaining why.

## Making Changes

### 1. Create a Branch

```bash
git checkout -b your-feature-branch
```

### 2. Make Your Changes

Follow the coding standards for the language or file type you are changing.

### 3. Run Pre-commit Hooks

```bash
pre-commit run --all-files
```

Fix any issues that are reported.

### 4. Run Tests

Before submitting a pull request, ensure the retained test suites pass locally.


#### PowerShell Tests

```powershell
Install-Module -Name Pester -RequiredVersion 5.9.1 -Force -Scope CurrentUser
Import-Module Pester -RequiredVersion 5.9.1
Invoke-Pester -Path tests/ -Output Detailed
```

#### Test Requirements

- PowerShell changes should include Pester tests in `tests/PowerShell/`.
- Retained test suites must pass on the applicable CI matrix.

### 5. Commit Your Changes

```bash
git add .
git commit -m "Your descriptive commit message"
```

Pre-commit hooks will run automatically. If they make changes, review them and commit again.

### 6. Push Your Branch

```bash
git push origin your-feature-branch
```

### 7. Open a Pull Request

Open a PR on GitHub and fill out the PR template checklist.

## Pull Request Guidelines

When submitting a pull request:

- [ ] Confirm `pre-commit run --all-files` passes locally.
- [ ] Include tests for new functionality.
- [ ] Update documentation as needed.
- [ ] Ensure all CI checks pass.
- [ ] Confirm no secrets, credentials, tokens, or connection strings were committed.

## Questions or Issues?

Use [GitHub Discussions](https://github.com/franklesniak/ROMSorter/discussions) for support questions,
usage help, and general troubleshooting. Search existing discussions before starting a new one.

Use [GitHub Issues](https://github.com/franklesniak/ROMSorter/issues) for reproducible bugs,
enhancement requests, documentation defects, and other actionable repository work.
Review existing issues before opening a new report. Report vulnerabilities privately through
[SECURITY.md](SECURITY.md).

Supported hosts are Windows PowerShell 5.1 and Microsoft-supported PowerShell 7 release lines on
Windows, macOS, and Linux. CI tests PowerShell 7.6 and Windows PowerShell 5.1; update the matrix when
Microsoft adds a supported release line. The analyzer temporarily uses `first-adoption` mode:
errors and unknown severities fail, while warnings remain visible. See [_TODO-repo-init.md](_TODO-repo-init.md)
for the debt cleanup and strict-gate exit task.

Template-support regression tests run with `python -m pytest -m "not upstream_template_only"`.
Their Python dependency is development tooling; ROMSorter remains a PowerShell project.

## License

<!--
TEMPLATE ADOPTERS: Update this section if your project uses a license other than MIT.

If using a different open source license (Apache 2.0, GPL, BSD, etc.):
- Replace "MIT License" with your license name
- Ensure consistency with the LICENSE file in the repository root

If using a proprietary license:
- Replace the entire contribution agreement text below
- Consider adding Contributor License Agreement (CLA) requirements
- Consult with legal counsel for appropriate contribution terms

See the [OPTIONAL_CONFIGURATIONS.md "License Customization" section](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/OPTIONAL_CONFIGURATIONS.md#license-customization) for detailed guidance.
-->

By contributing to this project, you agree that your contributions will be licensed under the same license as the project (MIT License).
