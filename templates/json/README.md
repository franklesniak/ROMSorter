<!-- markdownlint-disable MD013 -->

# JSON Template Files

This directory contains starter JSON Schema content and example fixtures for downstream consumers that want to add schema-backed JSON contracts to their repositories.

## Purpose

These template files demonstrate how to author a JSON Schema Draft 2020-12 contract and example fixtures that align with the JSON authoring standards defined in [`.github/instructions/json.instructions.md`](../../.github/instructions/json.instructions.md) and the schema conventions documented in [`schemas/README.md`](https://github.com/franklesniak/copilot-repo-template/blob/2f6aa17caa063eaa0b39afb3afb1d27757b4eec3/schemas/README.md). They are copy-and-adapt material, not enforced repository contracts. While these files remain under `templates/`, the repository's **schema validation** hooks (`check-jsonschema` for valid example data and `check-metaschema` for schema self-validation) and the schema-example pytest contract in [`tests/test_schema_examples.py`](../../tests/test_schema_examples.py) are intentionally not extended to cover them; the active hook scopes (`^schemas/example-config\.schema\.json$` and `^schemas/examples/example-config/valid/.*\.json$`) and the test discovery root (`schemas/`) are anchored to the active worked example, not to `templates/**`. When baseline pre-commit validation is retained, the starter JSON files themselves are still validated for **syntax** by the repository's `check-json` pre-commit hook (anchored to `\.json$`) and, when both baseline and GitHub Actions are retained, re-run in [`.github/workflows/data-ci.yml`](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/.github/workflows/data-ci.yml) like every other `.json` file in the tree; the carve-out is about *schema contract validation*, not about whether these files are parsed as JSON at all.

The active worked example that exercises the same shape under enforcement is `schemas/example-config.schema.json` together with its fixtures under `schemas/examples/example-config/`. The template starter under this directory is intentionally smaller and is meant to be lifted into a downstream `schemas/` directory rather than studied in place.

## Files Included

- **`starter.schema.json`**: Minimal JSON Schema Draft 2020-12 starter with `$schema`, `$id`, `title`, `description`, `type: "object"`, `additionalProperties: false`, one required property (`name`) plus a required `enabled` boolean, and a small set of optional properties. Values are obviously fake (for example, `https://example.invalid/...`, `"example-service"`).
- **`examples/starter/valid/minimal.json`**: Smallest valid instance for `starter.schema.json` (only the required properties).
- **`examples/starter/invalid/missing-required.json`**: Smallest invalid instance for `starter.schema.json`, omitting the required `enabled` property. It is syntactically valid JSON but is rejected by the schema.
- **`README.md`**: This file.

## How to Use the Starter Schema and Examples

1. Copy `starter.schema.json` into your downstream repository as `schemas/<your-name>.schema.json` and rename it to match your file family (for example, `schemas/feature-flags.schema.json`).
2. Update `$id`, `title`, and `description` to reflect your project. Keep `$schema` set to JSON Schema Draft 2020-12 unless a specific consumer requires another draft.
3. Replace the starter's `properties`, `required`, and `additionalProperties` with the real shape of your file family. Keep `additionalProperties: false` for project-owned closed contracts.
4. Copy `examples/starter/valid/minimal.json` and `examples/starter/invalid/missing-required.json` into `schemas/examples/<your-name>/valid/` and `schemas/examples/<your-name>/invalid/` respectively. Add additional fixtures (for example, a `full.json` that exercises every optional property) as the schema grows.
5. Verify the fixtures behave as expected by running `check-jsonschema` directly against them (see the snippet below) before wiring anything into pre-commit.

## Copy/Adapt Into a Downstream `schemas/` Directory

The recommended layout in the downstream repository mirrors the layout used by this template's worked example:

```text
schemas/
├── <your-name>.schema.json
└── examples/
    └── <your-name>/
        ├── valid/
        │   └── minimal.json
        └── invalid/
            └── missing-required.json
```

Once the schema and fixtures are in place, follow the upstream template's [JSON/YAML starter content adoption guidance](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/OPTIONAL_CONFIGURATIONS.md#jsonyaml-starter-content-opt-in) to wire a scoped `check-jsonschema` pre-commit hook for the new schema-backed file family. Do not broaden the active `Validate example-config valid examples` hook to cover `templates/**`; the starter content under this directory is intentionally outside the active hook's scope.

## Pre-commit Snippet for `check-jsonschema`

After copying the starter schema into `schemas/<your-name>.schema.json` and the fixtures into `schemas/examples/<your-name>/{valid,invalid}/`, add a scoped hook to `.pre-commit-config.yaml` that validates only the **valid** fixtures (invalid fixtures MUST NOT be wired into a normal `check-jsonschema` hook because their expected non-zero exit would be reported as a hook failure):

```yaml
- repo: https://github.com/python-jsonschema/check-jsonschema
  rev: 0.33.3
  hooks:
    - id: check-jsonschema
      name: Validate <your-name> valid examples
      files: ^schemas/examples/<your-name>/valid/.*\.json$
      args:
        - --schemafile
        - schemas/<your-name>.schema.json
```

> **Version pinning.** Look up the latest tagged release of [`check-jsonschema`](https://github.com/python-jsonschema/check-jsonschema) and pin a current version rather than copying the example `rev:` value above. Update the pin via your normal dependency-update process.

Invalid fixtures should be exercised by a test that asserts validation fails. The active worked example in this template uses [`tests/test_schema_examples.py`](../../tests/test_schema_examples.py) for this purpose, and an upstream Python starter version of that test lives at [`templates/python/tests/test_schema_examples.py`](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/templates/python/tests/test_schema_examples.py).

You can verify any single fixture from the command line at any time without pre-commit:

```bash
check-jsonschema \
  --schemafile schemas/<your-name>.schema.json \
  schemas/examples/<your-name>/valid/minimal.json
```

A valid fixture must exit with code `0`; an invalid fixture must exit with a non-zero code.

## Starter Content, Not an Active Contract

The files under `templates/json/` are **starter content**. They are not validated by the repository's active `check-jsonschema` or `check-metaschema` hooks, and they are not discovered by `tests/test_schema_examples.py`. Both the active hook scopes (`^schemas/example-config\.schema\.json$` and `^schemas/examples/example-config/valid/.*\.json$`) and the test discovery root (`schemas/`) are intentionally narrow so that template starter content can evolve independently of the repository's enforced contracts.

If you want the starter content under this directory to be validated locally before copying it into `schemas/`, run `check-jsonschema` directly against it from the command line as shown above. Do not broaden the active hooks or the test discovery root to include `templates/**`.

## Additional Resources

- [JSON Schema specification](https://json-schema.org/specification) — canonical entry point that lists Draft 2020-12 and other published drafts
- [JSON Schema documentation](https://json-schema.org/)
- [`check-jsonschema` documentation](https://check-jsonschema.readthedocs.io/)
- [`.github/instructions/json.instructions.md`](../../.github/instructions/json.instructions.md) — JSON authoring standards
- [`schemas/README.md`](https://github.com/franklesniak/copilot-repo-template/blob/2f6aa17caa063eaa0b39afb3afb1d27757b4eec3/schemas/README.md) — schema conventions, the active worked example, and the canonical downstream-removal checklist
- [Upstream JSON/YAML starter content adoption guidance](https://github.com/franklesniak/copilot-repo-template/blob/HEAD/OPTIONAL_CONFIGURATIONS.md#jsonyaml-starter-content-opt-in) — adoption guidance for these template starters
