<!--
Template users: see OPTIONAL_CONFIGURATIONS.md for guidance on tailoring this PR
template. Delete this comment once the PR template is tailored for your needs.

LINK STYLE: Markdown links to repo-internal files in this template MUST use
absolute `https://github.com/<OWNER>/<REPO>/blob/HEAD/<path>` URLs (with the
angle-bracket placeholders replaced by your actual values). Relative forms
(e.g., `../blob/HEAD/<file>`) are unreliable across non-GitHub.com
renderers, GitHub Mobile, and email notifications. See the
"Issue and PR templates" carve-out in `.github/instructions/docs.instructions.md`.

CUSTOMIZE: Replace `franklesniak/ROMSorter` below with your actual org/repo.
If you keep `.github/workflows/check-placeholders.yml` (an optional
adoption step), CI will fail until this substitution is made; if you
do not adopt that workflow or you remove it after setup, no CI
guardrail catches a missed substitution and you must verify the
replacement manually.
GHES users: Replace `github.com` with your GHES host (e.g., `github.company.com`).

KNOWN LIMITATION (template repo only): The contributing-guidelines link
below contains the literal `franklesniak/ROMSorter` placeholder in its URL rather than
substituted values, so it resolves to a non-existent target on `github.com`
when rendered on PRs against the template repository
(`franklesniak/copilot-repo-template`). The `franklesniak/ROMSorter` placeholder on
the checklist link is intentionally not substituted in the template repo —
that substitution is the responsibility of downstream adopters cloning
this template. The reason the unsubstituted placeholder is allowed to
persist here (rather than being rejected by CI) is that
`.github/workflows/check-placeholders.yml` is gated off in the template repo
itself (`if: github.repository != 'franklesniak/copilot-repo-template'`);
downstream adopters who keep `.github/workflows/check-placeholders.yml`
(an optional adoption step) do not have that gate and so have their
substitution enforced by CI before merge. Downstream adopters who do not
adopt that workflow, or who remove it after initial setup, have no CI
guardrail and must verify the substitution manually. This is an accepted
trade-off: downstream adopters who keep the placeholder workflow get a
working absolute link after running the placeholder substitution, which
is the dominant audience. Contributors to the template
repo itself can navigate to `CONTRIBUTING.md` via the file tree.

The illustrative URL in this comment uses the angle-bracket placeholder form
`<OWNER>/<REPO>` so that the only absolute `github.com` URL in this file
that contains the literal `franklesniak/ROMSorter` placeholder is the live
contributing-guidelines link below — that live link is the substitution
target described in the placeholder table in `GETTING_STARTED_NEW_REPO.md`,
and the only hit reported by section [6] of
`.github/workflows/check-placeholders.yml` (which greps for the literal
`franklesniak/ROMSorter`-bearing absolute `github.com` URL form; see the workflow for
the exact pattern). The literal `franklesniak/ROMSorter` placeholder also
appears as plain text in the `CUSTOMIZE`, `KNOWN LIMITATION`, and rationale
paragraphs above; those occurrences are not absolute `github.com` URLs and
are not flagged by section [6]. The file-wide `franklesniak/ROMSorter` substitution
scripts under *Initial Placeholder Replacement* in
`GETTING_STARTED_NEW_REPO.md` (PowerShell `.Replace`, GNU `sed -i`, and
BSD `sed -i ''`) will rewrite those plain-text references too, but that is
harmless because adopters delete this entire HTML comment block per the
*Delete Template Comment* step under *Customizing the Pull Request Template*
in the same guide. See the "Issue and PR templates" carve-out in
`.github/instructions/docs.instructions.md` for the broader rule against
shipping live placeholder URLs disguised as comments. That rule and section
[6]'s grep both target the literal `franklesniak/ROMSorter`-bearing URL form (which is
why the live checklist link below is detected and customized); the
angle-bracket documentation form used in this comment block is intentionally
outside the scope of that rule and CI check.
-->

## Description

<!-- Provide a brief description of the changes in this PR -->

## Type of Change

<!-- Check all that apply -->

- [ ] Bug fix (non-breaking change which fixes an issue)
- [ ] New feature (non-breaking change which adds functionality)
- [ ] Breaking change (fix or feature that would cause existing functionality to not work as expected)
- [ ] Documentation update
- [ ] Dependencies update
- [ ] Configuration/tooling change

## Checklist

<!-- Review and check all items before submitting -->

### General

- [ ] I have read the [contributing guidelines](https://github.com/franklesniak/ROMSorter/blob/HEAD/CONTRIBUTING.md)
- [ ] My code follows the coding standards in `.github/instructions/`
- [ ] My changes follow `.github/copilot-instructions.md` and any applicable `.github/instructions/*`
- [ ] No secrets, real PII, or production credentials appear in any committed file
- [ ] I have performed a self-review of my own code
- [ ] I have commented my code where necessary, particularly in hard-to-understand areas
- [ ] I have made corresponding changes to the documentation
- [ ] My changes generate no new warnings
- [ ] I have added or updated tests where appropriate
- [ ] New and existing tests pass locally (where applicable)

<!-- template-sync: begin baseline-reference-only -->
### Pre-commit Verification (if configured)

- [ ] If this repository uses pre-commit, I ran `pre-commit run --all-files` and all checks pass
- [ ] If pre-commit made auto-fixes, I reviewed and committed them
<!-- template-sync: end baseline-reference-only -->

<!-- template-sync: begin powershell-reference-only -->
### PowerShell-Specific (if applicable)

- [ ] PSScriptAnalyzer passes locally (`Invoke-ScriptAnalyzer -Path . -Settings .github/linting/PSScriptAnalyzerSettings.psd1`)
- [ ] Pester tests pass locally (`Invoke-Pester -Path tests/ -Output Detailed`)
- [ ] PowerShell formatting follows repository standards (OTBS, consistent line endings)

<!-- template-sync: end powershell-reference-only -->
<!-- template-sync: begin baseline-reference-only -->
### Data-File-Specific (if applicable)

- [ ] Retained data-file checks pass locally through `pre-commit run --all-files` or the matching targeted pre-commit hooks
- [ ] No secrets, real PII, or production credentials appear in any committed data file or fixture

<!-- template-sync: end baseline-reference-only -->
<!-- template-sync: begin github-actions-reference-only -->
### GitHub Actions-Specific (if applicable)

- [ ] If a GitHub Actions workflow was modified, `actionlint` passes (run via `pre-commit run actionlint --all-files`)

<!-- template-sync: end github-actions-reference-only -->

## Additional Notes

<!-- Add any additional information that reviewers should know -->

## Related Issues

<!-- Link any related issues using #issue_number -->

Closes #
