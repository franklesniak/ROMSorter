<!-- markdownlint-disable MD013 -->

# Workflow Security Contract

- **Status:** Active
- **Owner:** Repository Maintainers
- **Last Updated:** 2026-09-23
- **Scope:** Documents the retained workflow contract, validation commands, security limits, and module lifecycle.

The `github-actions` module owns the workflow contract, this guide, and the standalone Workflow Security job. The validator, its parsing helper, and its schema are shared with `template-sync-support` as trusted materializer dependencies. The same helper supports instruction enforcement, but Actions does not require that module or its core/profile files. Actions validation remains usable without the Markdown, Python project, baseline, or template-sync-support modules. Install the validator dependencies declared in the standalone workflow before running the direct command:

```text
python .github/scripts/validate_workflow_security.py
python .github/scripts/validate_workflow_security.py --verify-releases
```

The first command checks local inputs without network access. The second also resolves each annotated upstream release through GitHub's public API and fails on a mismatched SHA or unavailable verification. It sends no credentials. API rate limits can require retrying the explicit verification later; a failed lookup is never reported as verified. Routine local and CI checks are offline. A maintainer must run release verification for new or changed pins before accepting them.

For branch protection or rulesets, select the **Workflow Security** job check after it has run successfully. The optional instruction validator exposes the separate **Instruction Contracts** check. Existing rulesets that required `validate` need an owner to select the intended named checks; workflow updates do not change repository settings. The reviewed contract protects these job names while preserving their internal `validate` job IDs.

## Ownership and selection

The [upstream manifest](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/.template-sync/manifest.yml) is the only module/path authority. The source contract lists reviewed workflow paths and execution controls, without duplicating module selections or action versions. Before staging, materialization derives retained template-owned workflows from the selected source inventory and manifest relations. It rejects a missing contract, a missing retained entry, or an empty retained workflow inventory. Materialization validates the source contract, then applies existing module markers to selected workflows and records the resulting controls in a self-contained downstream contract. Omitted language workflows and their optional steps are not required downstream. Runtime validation requires every listed file; missing owned files fail.

Materialization also scans retained manifest-owned Markdown (`.md` and `.mdc`) after module-marker pruning with the installed validator's documentation parser. This includes retained Cursor rules when they contain governed action references. The detected files containing governed action references must exactly match the contract's retained `examples` list. A missing entry or a declared file without such references fails before staging. For example, removing a retained guide from `examples` while leaving its copyable action references causes a missing-entry error. Documents omitted by module selection and unmapped adopter documents are outside discovery. Successful rendering preserves the declared example order, and downstream validation uses the resulting self-contained list without requiring the manifest.

The contract is protected policy data. First adoption and replacement require a path-scoped `protected_file_decisions` entry for `.github/workflow-security-contract.yml`, with explicit owner authorization for `TAKE` or the specific `MERGE`. Materialization preserves changed files pending recorded decisions. This candidate-side check does not independently establish human authorization or protect against a coordinated edit to the validator and its tests.

Adopter-created workflows are ignored by default and remain untouched during synchronization. To opt into universal checking of every workflow, run:

```text
python .github/scripts/validate_workflow_security.py --strict
```

To keep that profile in CI, explicitly authorize adding `--strict` to the standalone workflow's validation step and updating its reviewed contract fingerprint. Record the corresponding workflow override and protected contract decision during synchronization. Strict checking applies universal security rules to additional workflows; it does not invent required commands for adopter jobs. A renamed template-owned workflow requires a reviewed contract and manifest update; it cannot silently leave the owned inventory.

## Enforced properties and limits

External actions and reusable workflows require lowercase full commit SHAs and exact same-line `# vMAJOR.MINOR.PATCH` release comments. Commented workflow examples and listed copyable documentation examples follow that rule. Documentation checks decode YAML keys and values in complete YAML/YML/untyped fences and individual fragments, including quoted and flow forms. YAML/YML language labels may have trailing metadata. Other labeled fences contain literal samples and do not contribute action references; matching delimiters and quote/list boundaries determine their extent. Invalid fence openers and examples outside those boundaries remain checked. This is a bounded example scanner, not a complete Markdown renderer. Each reference needs its own physical annotation line; split scalars and multiple references on one line are rejected. Tagged documentation nodes remain inert and are never constructed as objects. The annotation is descriptive; the `uses:` line remains the action identity. There is no separate version ledger. Offline checks establish syntax, while explicit online verification establishes the release-to-commit correspondence.

In executable workflow files, reference checks follow the parsed job and step `uses` fields. Action-like text inside a scalar, such as a `uses: owner/action@v1` line inside `run: |`, remains data. An action input named `with.uses` also remains data. Hash-prefixed scalar content is not a YAML comment; genuine comments outside scalars still receive example validation, including immediately after a scalar ends. Command fingerprints continue to cover reviewed shell content.

Top-level and job permissions must explicitly be empty or `contents: read`. Checkout must set literal `persist-credentials: false`. The baseline rejects `pull_request_target` and `workflow_run`; it does not ship privileged exceptions. Local actions and Docker references are currently rejected because the template has no reviewed local-action or container surface. Supporting either requires an explicit policy extension and tests rather than an uninspected exemption.

Required workflow and job permissions, job and step controls, order, conditions, environments, and shell-body SHA-256 fingerprints must match the reviewed contract. Workflow and job concurrency, explicit cache access modes, job outputs and snapshots, and background/wait/cancel step controls are covered. Parallel step groups are rejected until recursive command and action validation is supported. Job display names and fallback job IDs are covered because they identify required checks. Top-level workflow display names remain free to change. Fingerprints normalize CRLF to LF and preserve all other decoded scalar whitespace, including leading indentation and trailing newlines. A removed permission, removed check, advisory conversion, changed shell, skipped condition, or altered failure aggregator fails validation. Fingerprints detect reviewed-byte drift; they do not prove arbitrary shell semantics. `actionlint` remains the syntax check.

Existing exceptions are narrowly recorded: Markdown checks capture individual failures and aggregate their outcomes; the auto-fix preview restores pre-commit's native failure after collecting its artifact; Python type checks intentionally remain advisory in downstream repositories. Their exact conditions and shell bodies are covered by the contract. Changing these exceptions requires review and passing negative tests.

Governed fences inside nested list and quote containers receive the same complete YAML checks. The scanner removes container prefixes while preserving physical line positions, so splitting a `uses` key from its value cannot bypass validation. Indented governed examples following list prose remain checked. An action-containing block that ends without a matching closing fence fails, including when its containing list or quote ends.

The executable-workflow parser bounds each input to 1 MiB, limits nesting, rejects duplicate keys, anchors, aliases, custom tags, multiple YAML documents, traversal, and symlink inputs. It uses YAML 1.2 boolean behavior so `on` is a string key. Documentation composition uses the same input and nesting bounds with cycle-safe node traversal. Separate prose comments from commented YAML examples with a blank comment line so each example can parse as one fragment. The hook installs its exact Python parser and schema-validator dependencies independently of language scaffolding.

## Updates and removal

Materialization executes the validator and placeholder helper from the running tool installation. It validates selected contract data against that installation's trusted schema. The selected workflow validator and schema are inert staged data during the current run. Keep the installed tool bundle and its dependencies reviewed together; an incompatible source contract requires an explicit tool upgrade. Removing Actions while retaining template-sync support keeps the shared validator and schema for future materialization. Removing both modules omits the validator and schema. The parsing helper remains while any of Actions, template-sync support, or instruction enforcement is retained.

Review changes to workflow paths, scheduling, cache access, job outputs and snapshots, asynchronous steps, contract/schema, manifest mappings, and wiring together. For intentional command changes, compute the normalized command's SHA-256 with the validator's `describe_workflow` function after semantic review and update the corresponding contract entry. Never regenerate the contract merely to make a failed check pass. Run the positive/negative, wiring, and lifecycle suites, including the guard-removal test. Action-only updates do not require contract fingerprint changes.

Retain Dependabot's `github-actions` updater when both `github-platform` and `github-actions` are selected. Actions-only adopters that omit platform automation must review upstream releases manually. SHA pins support Dependabot version updates, but GitHub's Actions vulnerability alerts depend on semantic-version references; review upstream advisories as part of maintenance. See [GitHub's dependency guidance](https://docs.github.com/en/actions/reference/security/secure-use#understanding-dependencies-in-your-workflows).

When removing the module, materialize the remaining modules, review the explicit cleanup, and remove old standalone Actions assets only with the required authorization. Materialization deliberately does not delete excluded files. Replace retained shared files with their pruned versions so hooks, updater entries, documentation, and validation references disappear together. The downstream adoption validator must report no excluded-module leftovers after cleanup. Keep unrelated adopter workflows under their recorded local ownership.
