<!-- markdownlint-disable MD013 -->

# YAML Template Files

This directory contains starter YAML linting configurations and a small starter config example for downstream consumers that want to adopt the template's YAML toolchain.

## Purpose

These template files demonstrate how to configure `yamllint` and how to author a minimal human-readable YAML file that aligns with the YAML authoring standards defined in [`.github/instructions/yaml.instructions.md`](../../.github/instructions/yaml.instructions.md). They are copy-and-adapt material, not enforced repository contracts: while these files remain under `templates/`, **none of them is loaded as the active `.yamllint.yml` and their settings are not applied as the linting policy for the rest of the repository**. The active YAML linting configuration for this repository is `.yamllint.yml` at the repository root, which mirrors `templates/yaml/yamllint.lenient.yml`. The starter YAML files themselves are still parsed for syntax by the repository's `check-yaml` pre-commit hook and linted by `yamllint` against the active root configuration, like every other YAML file in the tree; the carve-out is about *which configuration is active*, not about whether these files are validated as YAML at all.

## Files Included

- **`yamllint.lenient.yml`**: Mirrors the active repository-root `.yamllint.yml` rule settings. The `rules:` block is byte-identical to the active root configuration; only the leading comment block differs (it documents the file's role as starter content rather than the active config). Use this when you want the template's current defaults (`truthy.check-keys: false`, `comments-indentation: disable`, `line-length.level: warning`).
- **`yamllint.strict.yml`**: Stricter alternative with `truthy.check-keys: true` and `line-length` at default error severity. Adopting it requires quoting `"on":` in every GitHub Actions workflow file. Documented by the upstream template's [yamllint ADRs](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/.github/TEMPLATE_DESIGN_DECISIONS.md).
- **`starter-config.yaml`**: Minimal human-authored YAML config example that demonstrates 2-space indentation, lowercase `true` / `false`, quoted version pins (for example, `"3.13"`, `"1.0"`), quoted YAML 1.1 truthy-looking strings (`"no"`, `"on"`) when they are intended as strings, and a short block scalar.
- **`README.md`**: This file.

## When to Choose Lenient vs. Strict

Choose **`yamllint.lenient.yml`** when:

- You are starting a new repository and want zero-friction GitHub Actions authoring (idiomatic unquoted `on:` keys).
- Your YAML files include long URLs, descriptive prose inside issue forms, or multi-line shell commands inside workflow `run:` blocks where forcing lines to wrap would change the literal output downstream users see.
- You want long-line warnings to surface in CI annotations without blocking PRs on cosmetic wraps.

Choose **`yamllint.strict.yml`** when:

- You want maximum YAML 1.2 hygiene and are willing to quote `"on":` in every workflow file under `.github/workflows/*.yml` at adoption time.
- You want every long line to fail the build rather than surface as a warning, and you are willing to reformat any offending YAML files (issue templates, workflow `run:` blocks, etc.) before adoption.
- Your project does not contain non-breakable user-facing YAML content where the existing line-length softening matters.

The trade-offs and rationale for each setting are documented in the two ADRs cited above. The strict starter file itself passes the repository's active root `.yamllint.yml`; it does not need to be self-applied as the active yamllint configuration during pre-commit.

## Copy/Adapt the Starter Config

To use one of the yamllint configurations in a downstream repository, copy it to the repository root:

```bash
# Use the template defaults verbatim:
cp templates/yaml/yamllint.lenient.yml .yamllint.yml

# Or adopt the stricter alternative (quote "on": in workflow files first):
cp templates/yaml/yamllint.strict.yml .yamllint.yml
```

To use `starter-config.yaml` as a starting point for a real project config file, copy it under your project's chosen configuration directory and rename it to match your file family (for example, `config/service.yaml`). Update keys, values, and the comment block to reflect the real shape of the file. Keep the conservative subset rules (lowercase booleans, quoted version pins, quoted YAML 1.1 truthy-looking strings, block scalars for multi-line strings).

## Pre-commit Snippet for `yamllint`

If your downstream repository does not already have a `yamllint` pre-commit hook, add one to `.pre-commit-config.yaml` after copying one of the starter configurations to `.yamllint.yml`:

```yaml
- repo: https://github.com/adrienverge/yamllint
  rev: v1.38.0
  hooks:
    - id: yamllint
      args: [-c, .yamllint.yml]
```

> **Version pinning.** Look up the latest tagged release of [`yamllint`](https://github.com/adrienverge/yamllint) and pin a current version rather than copying the example `rev:` value above. Update the pin via your normal dependency-update process.

## Optional Pre-commit Snippet for YAML Schema Validation with `check-jsonschema`

This snippet is **optional** and is intended for downstream projects that have a concrete schema-backed YAML file family to validate (for example, an application configuration file with a published JSON Schema). Do not adopt it speculatively. After authoring or adopting the schema, scope the hook to the file family it covers:

```yaml
- repo: https://github.com/python-jsonschema/check-jsonschema
  rev: 0.33.3
  hooks:
    - id: check-jsonschema
      name: Validate <your-name> YAML config
      files: ^config/.*\.ya?ml$
      args:
        - --schemafile
        - schemas/<your-name>.schema.json
```

> **Version pinning.** As with the `yamllint` hook above, pin a current upstream release of `check-jsonschema` rather than copying the example `rev:` value.

## Ecosystem-Specific Validators Stay Opt-in

Ecosystem-specific YAML validators (for example, `kubeconform` for Kubernetes manifests, `spectral` for OpenAPI/AsyncAPI, `cfn-lint` for AWS CloudFormation, `ansible-lint` for Ansible, `helm lint` for Helm charts) **SHOULD** be adopted only when the downstream repository actually uses those file types. Adding a validator for a stack the project does not use creates noise without enforcing anything useful. See the upstream template's [ecosystem-specific validator guidance](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/OPTIONAL_CONFIGURATIONS.md#ecosystem-specific-yaml-validators-opt-in) for the catalog and adoption guidance.

## Starter Content, Not an Active Contract

The files under `templates/yaml/` are **starter content**. They are not loaded by `yamllint` as the active linting configuration, and they are not consumed as application or runtime configurations or as `check-jsonschema` schema contracts. The repository's `check-yaml` and `yamllint` pre-commit hooks do still parse and lint the starter YAML files for syntax and style like every other YAML file in the tree; the carve-out is about whether the starter files drive a tool's runtime behavior, not about whether they are validated as YAML at all. The active YAML linting configuration for this repository is `.yamllint.yml` at the repository root.

## Additional Resources

- [`.github/instructions/yaml.instructions.md`](../../.github/instructions/yaml.instructions.md) — YAML authoring standards
- [`.yamllint.yml`](../../.yamllint.yml) — the active YAML linting configuration for this repository
- [Upstream yamllint ADRs](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/.github/TEMPLATE_DESIGN_DECISIONS.md) — rationale for the truthy.check-keys default and line-length warning level
- [`yamllint` documentation](https://yamllint.readthedocs.io/)
- [`check-jsonschema` documentation](https://check-jsonschema.readthedocs.io/)
- [Upstream JSON/YAML starter content adoption guidance](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/OPTIONAL_CONFIGURATIONS.md#jsonyaml-starter-content-opt-in) — adoption guidance for these template starters
