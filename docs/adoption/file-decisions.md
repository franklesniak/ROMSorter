# Per-File Adoption Decisions

Reviewed source: `2f6aa17caa063eaa0b39afb3afb1d27757b4eec3`.
All 303 upstream tracked paths were reconciled against downstream HEAD; this is first adoption, not an incremental lineage diff.
The generated [ledger](ledger.md) was reviewed before protected materialization. No downstream public path was deleted
because it was absent from upstream. `LICENSE.txt` becomes `LICENSE`; the existing license text wins.
`TAKE` means exact source bytes; `MERGE` identifies narrow retained-template adaptations; `SKIP` preserves downstream
content or excludes an unselected module. These are explicit decisions for this adoption, not blanket future acceptance.
All 13 protected decisions use minimal-preservation; the sync catalog remains byte-identical to the pinned source.
Unselected protected paths shown as potential decisions in the generated ledger do not exist downstream and are excluded here.

| Path | Decision | Rationale |
| --- | --- | --- |
| `.azuredevops/pipelines/README.md` | SKIP | Unselected module; do not create this upstream path. |
| `.azuredevops/pipelines/check-placeholders.yml` | SKIP | Unselected module; do not create this upstream path. |
| `.azuredevops/pipelines/data-ci.yml` | SKIP | Unselected module; do not create this upstream path. |
| `.azuredevops/pipelines/instruction-contracts.yml` | SKIP | Unselected module; do not create this upstream path. |
| `.azuredevops/pipelines/markdownlint.yml` | SKIP | Unselected module; do not create this upstream path. |
| `.azuredevops/pipelines/powershell-ci.yml` | SKIP | Unselected module; do not create this upstream path. |
| `.azuredevops/pipelines/precommit.yml` | SKIP | Unselected module; do not create this upstream path. |
| `.azuredevops/pipelines/python-ci.yml` | SKIP | Unselected module; do not create this upstream path. |
| `.azuredevops/pipelines/terraform-ci.yml` | SKIP | Unselected module; do not create this upstream path. |
| `.azuredevops/platform/adoption-guidance.md` | SKIP | Unselected module; do not create this upstream path. |
| `.azuredevops/pull_request_template.md` | SKIP | Unselected module; do not create this upstream path. |
| `.claude/commands/review-loop.md` | SKIP | Unselected module; do not create this upstream path. |
| `.claude/hooks/session-start.sh` | SKIP | Unselected module; do not create this upstream path. |
| `.claude/settings.json` | SKIP | Unselected module; do not create this upstream path. |
| `.codex/config.toml` | TAKE | Retain exact pinned upstream bytes. |
| `.cursor/rules/repository-instructions.mdc` | SKIP | Unselected module; do not create this upstream path. |
| `.gitattributes` | MERGE | Prune unselected Git LFS block; protect exact bytes of four public CSV inputs with -text. |
| `.github/CODEOWNERS` | MERGE | Sole owner @franklesniak; remove unused Azure-host path rule. |
| `.github/ISSUE_TEMPLATE/bug_report.yml` | MERGE | Render identity and assumed labels; use current master branch. |
| `.github/ISSUE_TEMPLATE/config.yml` | MERGE | Render confirmed repository, private-security path and Discussions support destination. |
| `.github/ISSUE_TEMPLATE/documentation_issue.yml` | MERGE | Render identity; use current master branch. |
| `.github/ISSUE_TEMPLATE/feature_request.yml` | TAKE | Retain exact pinned upstream bytes. |
| `.github/TEMPLATE_DESIGN_DECISIONS.md` | SKIP | Unselected module; do not create this upstream path. |
| `.github/copilot-instructions.md` | MERGE | Prune complete unselected module blocks and bump metadata; retain canonical anchors. |
| `.github/dependabot.yml` | TAKE | Retain exact pinned upstream bytes. |
| `.github/instruction-contracts.yml` | MERGE | Materializer retains selected contracts with all selected anchors. |
| `.github/instruction-profile.yml` | MERGE | Marker-backed downstream profile. |
| `.github/instructions/docs.instructions.md` | TAKE | Retain exact pinned upstream bytes. |
| `.github/instructions/gitattributes.instructions.md` | TAKE | Retain exact pinned upstream bytes. |
| `.github/instructions/json.instructions.md` | TAKE | Retain exact pinned upstream bytes. |
| `.github/instructions/powershell.instructions.md` | TAKE | Retain exact pinned upstream bytes. |
| `.github/instructions/python.instructions.md` | SKIP | Unselected module; do not create this upstream path. |
| `.github/instructions/terraform.instructions.md` | SKIP | Unselected module; do not create this upstream path. |
| `.github/instructions/yaml.instructions.md` | TAKE | Retain exact pinned upstream bytes. |
| `.github/linting/PSScriptAnalyzerSettings.psd1` | TAKE | Retain exact pinned upstream bytes. |
| `.github/pull_request_template.md` | MERGE | Reviewed deterministic selected-module pruning, identity substitution or companion rendering. |
| `.github/scripts/check-prohibited-placeholders.py` | TAKE | Retain exact pinned upstream bytes. |
| `.github/scripts/check-toolchain-eol.js` | TAKE | Retain exact pinned upstream bytes. |
| `.github/scripts/instruction_contract_core.py` | TAKE | Retain exact pinned upstream bytes. |
| `.github/scripts/instruction_contract_support.py` | TAKE | Retain exact pinned upstream bytes. |
| `.github/scripts/lint-nested-markdown.js` | TAKE | Retain exact pinned upstream bytes. |
| `.github/scripts/lint-nested-markdown.test.js` | TAKE | Retain exact pinned upstream bytes. |
| `.github/scripts/lint_nested_markdown_hook.py` | TAKE | Retain exact pinned upstream bytes. |
| `.github/scripts/replace-template-placeholders.py` | TAKE | Retain exact pinned upstream bytes. |
| `.github/scripts/terraform_hooks.py` | SKIP | Unselected module; do not create this upstream path. |
| `.github/scripts/validate-placeholder-schema-examples.py` | TAKE | Retain exact pinned upstream bytes. |
| `.github/scripts/validate_instruction_profile.py` | TAKE | Retain exact pinned upstream bytes. |
| `.github/scripts/validate_workflow_security.py` | TAKE | Retain exact pinned upstream bytes. |
| `.github/template-placeholders.json` | TAKE | Retain exact pinned upstream bytes. |
| `.github/workflow-security-contract.yml` | MERGE | Prune excluded workflows; update reviewed PowerShell controls and PowerShell/Copilot Pester fingerprints. |
| `.github/workflows/auto-fix-precommit.yml` | MERGE | Reviewed deterministic selected-module pruning, identity substitution or companion rendering. |
| `.github/workflows/check-placeholders.yml` | TAKE | Retain exact pinned upstream bytes. |
| `.github/workflows/copilot-setup-steps.yml` | MERGE | Prune unselected setup blocks and pin Pester 5.9.1 to the retained 5.x requirement. |
| `.github/workflows/data-ci.yml` | MERGE | Reviewed deterministic selected-module pruning, identity substitution or companion rendering. |
| `.github/workflows/instruction-contracts.yml` | TAKE | Retain exact pinned upstream bytes. |
| `.github/workflows/markdownlint.yml` | TAKE | Retain exact pinned upstream bytes. |
| `.github/workflows/powershell-ci.yml` | MERGE | Temporary analyzer gate; supported-host matrix; explicit UTF-8 workflow output for Windows PowerShell 5.1; unchanged trust boundary. |
| `.github/workflows/precommit-ci.yml` | MERGE | Reviewed deterministic selected-module pruning, identity substitution or companion rendering. |
| `.github/workflows/python-ci.yml` | SKIP | Unselected module; do not create this upstream path. |
| `.github/workflows/terraform-ci.yml` | SKIP | Unselected module; do not create this upstream path. |
| `.github/workflows/toolchain-eol.yml` | TAKE | Retain exact pinned upstream bytes. |
| `.github/workflows/workflow-security.yml` | TAKE | Retain exact pinned upstream bytes. |
| `.gitignore` | TAKE | Retain exact pinned upstream bytes. |
| `.hermes.md` | SKIP | Unselected module; do not create this upstream path. |
| `.markdownlint.jsonc` | TAKE | Retain exact pinned upstream bytes. |
| `.pre-commit-config.yaml` | MERGE | Prune unselected hooks; preserve exact four CSV inputs from whitespace/EOL formatters. |
| `.remarkignore` | TAKE | Retain exact pinned upstream bytes. |
| `.remarkrc.mjs` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/instruction-contracts.yml` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/manifest.yml` | MERGE | Retain upstream map plus exact template-support Python import helper mappings. |
| `.template-sync/scripts/bootstrap_first_adoption.py` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/scripts/first_adoption_quality_reports.py` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/scripts/first_adoption_state.py` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/scripts/generate_sync_candidates.py` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/scripts/initialize_adoption_journal.py` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/scripts/instruction_profile_migration.py` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/scripts/materialize_downstream_adoption.py` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/scripts/report_excluded_module_references.py` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/scripts/run_first_adoption_checks.py` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/scripts/template_sync_materialization_helpers.py` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/scripts/validate_downstream_adoption.py` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/scripts/validate_instruction_contracts.py` | TAKE | Retain exact pinned upstream bytes. |
| `.template-sync/scripts/validate_marker.py` | TAKE | Retain exact pinned upstream bytes. |
| `.tflint.hcl` | SKIP | Unselected module; do not create this upstream path. |
| `.vscode/settings.json` | SKIP | Retain exact pre-existing bytes, including ROMSorter-BF-Part3 window.title. |
| `.yamllint.yml` | TAKE | Retain exact pinned upstream bytes. |
| `AGENTS.md` | MERGE | Prune complete unselected blocks and Azure host section; keep GitHub/Codex protocols and retained anchors; bump metadata. |
| `CLAUDE.md` | SKIP | Unselected module; do not create this upstream path. |
| `CODE_OF_CONDUCT.md` | MERGE | Direct reports to the maintainer using contact information on their GitHub profile. |
| `CONTRIBUTING.md` | MERGE | Prune unselected stacks, render identity, route support to Discussions and document host/test/gate policy. |
| `COPILOT_CHAT_PROMPTS.md` | MERGE | Reviewed deterministic selected-module pruning, identity substitution or companion rendering. |
| `GEMINI.md` | SKIP | Unselected module; do not create this upstream path. |
| `GETTING_STARTED_EXISTING_REPO.md` | SKIP | Unselected module; do not create this upstream path. |
| `GETTING_STARTED_NEW_REPO.md` | SKIP | Unselected module; do not create this upstream path. |
| `LICENSE` | MERGE | Rename the existing LICENSE.txt; retain its exact bytes, MIT text and Frank Lesniak author. It was already LF. |
| `OPTIONAL_CONFIGURATIONS.md` | SKIP | Unselected module; do not create this upstream path. |
| `README.md` | MERGE | Preserve project identity and examples; fix filename case and add narrow development, support, security and compatibility guidance. |
| `SECURITY.md` | MERGE | PVR-only, 48h/7d response and latest-version support; content assumes enabled with immediate manual activation in TODO. |
| `TEMPLATE_MAINTENANCE.md` | SKIP | Unselected module; do not create this upstream path. |
| `TEMPLATE_UPDATE_PROCEDURE.md` | TAKE | Retain exact pinned upstream bytes. |
| `docs/ISSUE_EVALUATION_PROMPT.md` | TAKE | Retain exact pinned upstream bytes. |
| `docs/PR_REVIEW_PROMPTS.md` | MERGE | Reviewed deterministic selected-module pruning, identity substitution or companion rendering. |
| `docs/azure-devops-support.md` | SKIP | Unselected module; do not create this upstream path. |
| `docs/copilot-setup.md` | MERGE | Reviewed deterministic selected-module pruning, identity substitution or companion rendering. |
| `docs/instruction-enforcement.md` | TAKE | Retain exact pinned upstream bytes. |
| `docs/terraform/TERRAFORM_COPILOT_INSTRUCTIONS_GUIDE.md` | SKIP | Unselected module; do not create this upstream path. |
| `docs/terraform/TERRAFORM_LINTING_GUIDE.md` | SKIP | Unselected module; do not create this upstream path. |
| `docs/terraform/TERRAFORM_TESTING_GUIDE.md` | SKIP | Unselected module; do not create this upstream path. |
| `docs/upstream-style-guides.md` | MERGE | Reviewed deterministic selected-module pruning, identity substitution or companion rendering. |
| `docs/workflow-security.md` | TAKE | Retain exact pinned upstream bytes. |
| `package-lock.json` | MERGE | Synchronize root package identity and Node minimum; preserve all dependency entries. |
| `package.json` | MERGE | Render ROMSorter development identity and align Node development minimum with CI at 24. |
| `pyproject.toml` | SKIP | Unselected module; do not create this upstream path. |
| `pyrightconfig.json` | SKIP | Unselected module; do not create this upstream path. |
| `requirements-pre-commit.txt` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/README.md` | SKIP | Unselected module; do not create this upstream path. |
| `schemas/example-config.schema.json` | SKIP | Unselected module; do not create this upstream path. |
| `schemas/examples/example-config/invalid/extra-property.json` | SKIP | Unselected module; do not create this upstream path. |
| `schemas/examples/example-config/invalid/missing-required.json` | SKIP | Unselected module; do not create this upstream path. |
| `schemas/examples/example-config/invalid/wrong-type.json` | SKIP | Unselected module; do not create this upstream path. |
| `schemas/examples/example-config/valid/full.json` | SKIP | Unselected module; do not create this upstream path. |
| `schemas/examples/example-config/valid/minimal.json` | SKIP | Unselected module; do not create this upstream path. |
| `schemas/examples/first-adoption-quality-suppressions/invalid/missing-path-reference-section.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/first-adoption-quality-suppressions/invalid/missing-suppression-reason.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/first-adoption-quality-suppressions/invalid/missing-suppression-selector.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/first-adoption-quality-suppressions/invalid/unknown-path-reference-rule.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/first-adoption-quality-suppressions/valid/full.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/first-adoption-quality-suppressions/valid/minimal.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-contracts/invalid/minimal.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-contracts/invalid/trailing-carriage-return-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-contracts/invalid/trailing-newline-contract-key.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-contracts/invalid/trailing-newline-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-contracts/valid/minimal.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-profile/invalid/missing-modules.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-profile/invalid/trailing-carriage-return-exception-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-profile/invalid/trailing-carriage-return-local-override-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-profile/invalid/trailing-carriage-return-local-ownership-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-profile/invalid/trailing-newline-exception-content-sha256.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-profile/invalid/trailing-newline-exception-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-profile/invalid/trailing-newline-local-override-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-profile/invalid/trailing-newline-local-ownership-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-profile/invalid/trailing-newline-protected-guide-waiver-contract-key.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/instruction-profile/valid/standalone.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-placeholders/invalid/missing-token-scope.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-placeholders/invalid/trailing-carriage-return-path-pattern.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-placeholders/invalid/trailing-newline-path-pattern.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-placeholders/invalid/trailing-newline-source-attribute.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-placeholders/invalid/trailing-newline-source-name.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-placeholders/invalid/unknown-finding-kind.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-placeholders/invalid/unsafe-path.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-placeholders/valid/full.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-placeholders/valid/minimal.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/duplicate-table-condition.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/empty-scoped-policy.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/empty-section-modules.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/missing-anchor-list.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/noncanonical-table-cell.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/path-traversal.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/protected-guide-reference-missing-target.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/trailing-carriage-return-path.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/trailing-heading-space.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/trailing-newline-contract-key.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/trailing-newline-module-name.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/trailing-newline-path.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/trailing-successor-newline.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/invalid/unknown-block-kind.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/valid/canonical-heading.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/valid/interleaved-blocks.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/valid/multi-agent.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/valid/protected-guide-obligations.yaml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-instruction-contracts/valid/scoped-policy.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/azure-devops-credential-url.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/azure-fields-without-azure-host.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/azure-org-url-extra-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/azure-project-url-missing-project.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/azure-repository-url-missing-git.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/azure-url-non-azure-host.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/branch-last-reviewed-sha.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/deferred-candidate-directory-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/duplicate-module.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/empty-instruction-waiver-reason.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/invalid-protected-adoption-mode.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/local-ownership-extra-property.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/local-ownership-glob-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/malformed-protected-decision.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/malformed-protected-guide-waiver-contract-key.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/missing-collaboration-policy-follow-up.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/missing-instruction-waiver-authorization.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/missing-placeholder-waiver-scope.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/missing-protected-decision-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/missing-protected-edit-authorization.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/missing-protected-guide-waiver-target.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/missing-remove-local-reason.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/missing-tailored-authorization.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/multiline-issue-label.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/protected-decision-directory-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/protected-decision-path-traversal.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/short-deferred-source-commit.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/short-last-reviewed-sha.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/tailored-non-content-decision.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/trailing-carriage-return-local-override-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/trailing-carriage-return-local-ownership-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/trailing-carriage-return-protected-decision-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/trailing-newline-azure-organization-url.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/trailing-newline-azure-project-url.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/trailing-newline-azure-repository-url.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/trailing-newline-last-reviewed-sha.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/trailing-newline-local-override-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/trailing-newline-local-ownership-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/trailing-newline-protected-decision-path.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/trailing-newline-protected-guide-waiver-contract-key.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/unknown-azure-module.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/unknown-default-decision.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/unknown-module.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/invalid/unknown-template-sync-field.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/valid/azure-devops-only-host-modules.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/valid/full-reconciliation-in-progress.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/valid/full.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/valid/github-only-host-modules.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/valid/instruction-contract-waiver.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/valid/minimal.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/valid/mixed-host-modules.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/valid/placeholder-waiver.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/valid/protected-decisions.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/valid/protected-guide-contract-waiver.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/template-sync-marker/valid/source-ref-first-adoption.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/workflow-security-contract/invalid/absolute-example.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/workflow-security-contract/invalid/contract.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/workflow-security-contract/invalid/empty-component-example.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/workflow-security-contract/invalid/interior-dot-example.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/workflow-security-contract/invalid/leading-dot-example.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/workflow-security-contract/invalid/trailing-newline-example.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/workflow-security-contract/invalid/trailing-newline-run-sha256.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/workflow-security-contract/invalid/trailing-newline-workflow-path.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/workflow-security-contract/valid/contract.yml` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/examples/workflow-security-contract/valid/example-paths.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/first-adoption-quality-suppressions.schema.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/instruction-contracts.schema.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/instruction-profile.schema.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/template-placeholders.schema.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/template-sync-instruction-contracts.schema.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/template-sync-manifest.schema.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/template-sync-marker.schema.json` | TAKE | Retain exact pinned upstream bytes. |
| `schemas/workflow-security-contract.schema.json` | TAKE | Retain exact pinned upstream bytes. |
| `src/copilot_repo_template/__init__.py` | SKIP | Unselected module; do not create this upstream path. |
| `src/copilot_repo_template/example.py` | SKIP | Unselected module; do not create this upstream path. |
| `src/tools/Resolve-PSScriptAnalyzerCandidate.ps1` | TAKE | Retain exact pinned upstream bytes. |
| `src/tools/Resolve-PSScriptAnalyzerGate.ps1` | TAKE | Retain exact pinned upstream bytes. |
| `templates/adoption/_TEMPLATE-ADOPTION-DIFFICULTIES.md` | TAKE | Retain exact pinned upstream bytes. |
| `templates/json/README.md` | MERGE | Point schema-worked-example references to pinned upstream because schema module is excluded. |
| `templates/json/examples/starter/invalid/missing-required.json` | TAKE | Retain exact pinned upstream bytes. |
| `templates/json/examples/starter/valid/minimal.json` | TAKE | Retain exact pinned upstream bytes. |
| `templates/json/starter.schema.json` | TAKE | Retain exact pinned upstream bytes. |
| `templates/markdown/.markdownlint-cli2.jsonc` | TAKE | Retain exact pinned upstream bytes. |
| `templates/powershell/Example.Tests.ps1` | TAKE | Retain exact pinned upstream bytes. |
| `templates/python/README.md` | SKIP | Unselected module; do not create this upstream path. |
| `templates/python/pyproject.toml` | SKIP | Unselected module; do not create this upstream path. |
| `templates/python/tests/__init__.py` | SKIP | Unselected module; do not create this upstream path. |
| `templates/python/tests/test_placeholder.py` | SKIP | Unselected module; do not create this upstream path. |
| `templates/python/tests/test_schema_examples.py` | SKIP | Unselected module; do not create this upstream path. |
| `templates/terraform/Example.tftest.hcl` | SKIP | Unselected module; do not create this upstream path. |
| `templates/yaml/README.md` | TAKE | Retain exact pinned upstream bytes. |
| `templates/yaml/starter-config.yaml` | TAKE | Retain exact pinned upstream bytes. |
| `templates/yaml/yamllint.lenient.yml` | TAKE | Retain exact pinned upstream bytes. |
| `templates/yaml/yamllint.strict.yml` | TAKE | Retain exact pinned upstream bytes. |
| `tests/PowerShell/Placeholder.Tests.ps1` | SKIP | Omit example assertions; use real offline ProductSyntax.Tests.ps1 instead. |
| `tests/PowerShell/Resolve-PSScriptAnalyzerCandidate.Tests.ps1` | TAKE | Retain exact pinned upstream bytes. |
| `tests/PowerShell/Resolve-PSScriptAnalyzerGate.Tests.ps1` | TAKE | Retain exact pinned upstream bytes. |
| `tests/__init__.py` | MERGE | Required test import package; describe retained tooling without Python example-product claims. |
| `tests/_pytest_compat.py` | TAKE | Required by retained support tests; narrow downstream manifest correction. |
| `tests/fixtures/dependabot/auto-assignment.yml` | SKIP | Unselected module; do not create this upstream path. |
| `tests/test_bootstrap_first_adoption.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_check_prohibited_placeholders.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_contract_wiring.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_copilot_setup.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_dependabot_schema.py` | SKIP | Unselected module; do not create this upstream path. |
| `tests/test_example.py` | SKIP | Unselected module; do not create this upstream path. |
| `tests/test_first_adoption_quality_reports.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_first_adoption_state.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_generate_sync_candidates.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_gitattributes.py` | SKIP | Unselected module; do not create this upstream path. |
| `tests/test_initialize_adoption_journal.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_instruction_input_limits.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_instruction_profile.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_instruction_profile_output.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_instruction_profile_references.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_instruction_yaml_duplicates.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_materialize_downstream_adoption.py` | MERGE | Mark eight cases requiring the full upstream README or excluded Azure assets upstream_template_only; retain all applicable materialization and security coverage. |
| `tests/test_precommit_runner.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_replace_template_placeholders.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_report_excluded_module_references.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_run_first_adoption_checks.py` | MERGE | Allow the retained downstream pytest.ini fixture configuration while preserving strict marker and selection assertions; retain upstream pyproject fallback. |
| `tests/test_schema_examples.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_template_manifest.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_template_sync_materialization_helpers.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_terraform_hooks.py` | SKIP | Unselected module; do not create this upstream path. |
| `tests/test_validate_downstream_adoption.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_validate_instruction_contracts.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_validate_marker.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_workflow_security_contract.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/test_workflow_security_lifecycle.py` | TAKE | Retain exact pinned upstream bytes. |
| `tests/toolchain-eol/check-toolchain-eol.test.js` | TAKE | Retain exact pinned upstream bytes. |
| `tests/toolchain-eol/fixtures/node-schedule.json` | TAKE | Retain exact pinned upstream bytes. |

## Downstream-Owned Paths

All 25 root PowerShell scripts keep their public paths and executable content. The four root CSVs keep exact bytes.
In addition to the separately reviewed EOL changes, two scripts each lose spaces from one blank line outside strings.
The existing arcade documentation keeps its meaning; any mechanical Markdown fixes are separately described in the summary.
The additional syntax test only parses public scripts and never dot-sources or invokes them.
`pytest.ini`, `_TODO-repo-init.md`, `_ADOPTION-SUMMARY.md`, and `docs/adoption/` are downstream review/test records.

## Workflow Review

PowerShell CI keeps push/pull-request triggers, contents-read permissions, credential-free checkout, full action pins,
test-root scoping and fail-closed analyzer errors. The reviewed override changes the gate mode, expands the matrix with
Windows PowerShell 5.1, verifies host release lines and writes discovery outputs explicitly as UTF-8. Only the PowerShell and Copilot setup workflow entries were updated after semantic review. PowerShell CI and Copilot setup also pin Pester 5.9.1 to enforce the retained 5.x guide;
their reviewed install/import fingerprints were updated. No action identity or release pin was changed.
All retained action-release annotations were verified.

The supported PowerShell 7 release line on 2026-11-11 is 7.6. Microsoft's lifecycle table ends 7.4 and 7.5 support on
2026-11-10. See [Microsoft's lifecycle table](https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-support-lifecycle).
The owner must keep hosted runner patches current and add newly supported release lines to the matrix.
