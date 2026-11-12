---
applyTo: "**/*.json,**/*.jsonc"
description: "JSON authoring standards: strict-by-default, schema-backed, deterministic, and secure."
---

<!-- markdownlint-disable MD013 -->

# JSON Writing Style

**Version:** 1.4.20260929.0

## Metadata

- **Status:** Active
- **Owner:** Repository Maintainers
- **Last Updated:** 2026-09-29
- **Scope:** Defines authoring standards for JSON and JSONC files in this repository, including configuration, schemas, fixtures, generated metadata, and machine-readable contracts. Covers dialect policy, formatting, key ordering, naming, data modeling, schema usage, comments, security, and generated output.
- **Related:** [Repository Copilot Instructions](../copilot-instructions.md)
<!-- template-sync: begin baseline-reference-only -->
- **Related baseline guidance:** [`.gitattributes` Rules](./gitattributes.instructions.md)
<!-- template-sync: end baseline-reference-only -->
<!-- template-sync: begin yaml-reference-only -->
- **Related YAML guidance:** [YAML Writing Style](./yaml.instructions.md)
<!-- template-sync: end yaml-reference-only -->

## Purpose and Scope

This guide establishes how JSON and JSONC files are authored in this repository. JSON is treated as a **machine-readable contract**: it MUST be unambiguous, deterministic, and safe to consume by tools and downstream automation.

The guide applies to every `.json` and `.jsonc` file in the repository, including configuration, schemas, fixtures, examples, policy documents, generated metadata, and any other JSON-shaped artifact, regardless of programming language or surrounding stack.

> **Note:** This document uses [RFC 2119](https://www.rfc-editor.org/rfc/rfc2119) keywords (**MUST**, **MUST NOT**, **SHOULD**, **SHOULD NOT**, **MAY**) to indicate requirement levels.

<!-- template-sync: begin baseline-reference-only -->
Line-ending pinning, BOM behavior, end-of-file newline, and trailing whitespace policy for JSON files are governed by [`.gitattributes` Rules](./gitattributes.instructions.md); this guide does not duplicate or override those rules.
<!-- template-sync: end baseline-reference-only -->

## Quick Reference Checklist

- **[All]** `.json` files **MUST** be strict JSON; `.jsonc` **MAY** be used only when the consuming tool explicitly supports JSONC.
- **[All]** **MUST** use 2-space indentation; **MUST NOT** use tabs.
- **[All]** **MUST NOT** include trailing commas in strict JSON.
- **[All]** Keys and string values **MUST** be double-quoted.
- **[All]** **MUST NOT** blindly sort keys; **MUST** preserve intentional grouping and tool-managed ordering (for example, `package.json`, JSON/JSONC lockfiles, and other generated files).
- **[All]** Strict JSON **MUST NOT** contain comments.
- **[All]** Production or load-bearing JSON files **MUST** have schema validation; durable fixtures, examples, policy documents, and config contracts **SHOULD** have schema validation.
- **[All]** **MUST NOT** commit secrets; example values **MUST** be obviously fake.
- **[All]** Untrusted JSON **MUST** be validated before use and **MUST NOT** be evaluated as executable code.
- **[All]** Generated JSON **MUST** be reproducible and **SHOULD** identify its source or generation command.

## Dialect Policy

This repository recognizes two JSON dialects: strict JSON and JSONC. Other dialects are out of scope by default.

- Files with the `.json` extension **MUST** be strict JSON as defined by [RFC 8259](https://www.rfc-editor.org/rfc/rfc8259). Strict JSON **MUST NOT** contain comments, trailing commas, unquoted keys, single-quoted strings, or any other non-RFC 8259 syntax.
- Files with the `.jsonc` extension **MAY** be used **only** when the consuming tool explicitly documents support for JSONC (for example, the TypeScript compiler reading `tsconfig.json`, and some VS Code settings files). When in doubt, prefer `.json`.
- The common `check-json` pre-commit hook validates `.json` files only. JSONC is **not** validated by `check-json`. Repositories that need stronger enforcement for `.jsonc` files **SHOULD** add JSONC-aware tooling (for example, a JSONC-aware parser, linter, or schema validator) rather than retrofitting `check-json`.
- JSON5 is **not** included in this repository's defaults and **MUST NOT** be introduced without an explicit, documented project decision. The `applyTo` glob for this guide intentionally omits `.json5`.

## Formatting

- Indentation **MUST** be 2 spaces per level. Tabs **MUST NOT** be used.
- Keys and string values **MUST** be double-quoted. Single quotes and unquoted keys **MUST NOT** be used (this is required by strict JSON and is also the recommended style for JSONC).
- Trailing commas **MUST NOT** appear in strict JSON. In JSONC files, trailing commas **MAY** be used only when the consuming tool documents support for them; otherwise they **SHOULD** be avoided to ease later conversion to strict JSON.
- Objects with more than one entry **SHOULD** use one key-value pair per line to keep diffs reviewable.
- Files **MUST** end with a single newline.
<!-- template-sync: begin baseline-reference-only -->
- When baseline is retained, its `end-of-file-fixer` pre-commit hook enforces the single-newline requirement. Line-ending and trailing-whitespace policy is governed by [`.gitattributes` Rules](./gitattributes.instructions.md).
<!-- template-sync: end baseline-reference-only -->

**Example (strict JSON):**

```json
{
  "name": "example",
  "version": "1.0.0",
  "isEnabled": true
}
```

## Key Ordering

JSON does not assign semantic meaning to key order, but tools, humans, and diffs do. Treat key order as part of the file's contract.

- Most formatters, including Prettier, do not sort JSON keys by default. Do not assume any formatter will reorder keys for you.
- Keys **MUST NOT** be blindly sorted alphabetically across an entire file. Intentional grouping (for example, `name` before `version` in `package.json`, or `$schema` first in a schema-bearing file) **MUST** be preserved.
- Tool-managed and generated files **MUST NOT** be reordered by hand. This includes, but is not limited to:
  - `package.json` (npm/Yarn/pnpm reorder fields on install per their own rules)
  - `package-lock.json` and other JSON/JSONC lockfiles or generated manifests
  - Compiler/build tool JSON/JSONC manifests where the tool documents an ordering convention
- Generated files **SHOULD** preserve their generator's ordering exactly. If you need a different order, change the generator, not the output.
- Within a hand-authored object, related keys **SHOULD** be grouped (for example, all identity fields, then all behavior fields, then all metadata fields). New keys **SHOULD** be added next to related keys, not appended at the end "because diffs are smaller."

## Naming Conventions

- Object keys in project-owned JSON **SHOULD** use `camelCase` by default. Downstream repositories **MAY** adopt a different convention (for example, `snake_case` to match a Python-heavy stack) provided the choice is documented and applied consistently within that repository.
- JSON file names **SHOULD** use lowercase `kebab-case` (for example, `release-notes.json`, `feature-flags.json`). Ecosystem-mandated names (for example, `package.json`, `tsconfig.json`) **MUST** be used as-is.
- Boolean keys **SHOULD** use a clear affirmative prefix such as `is`, `has`, `can`, `should`, or `enable` (for example, `isEnabled`, `hasOwner`, `canRetry`, `shouldFail`, `enableTelemetry`). Negated booleans (for example, `isNotReady`) **SHOULD** be avoided to prevent double-negation in consumers.

## Data Modeling

- Prefer **explicit objects** over positional arrays for records. Use `{ "first": "...", "last": "..." }` rather than `["...", "..."]` so that field meaning is self-describing.
- Identifiers (IDs, account numbers, version numbers used as identifiers, ZIP codes) **MUST** be encoded as strings, not numbers. JSON numbers are IEEE 754 doubles and silently lose precision beyond 2^53; identifier-shaped numbers can also acquire leading-zero or formatting issues when round-tripped.
- Timestamps **MUST** be RFC 3339-compatible strings (for example, `"2026-05-01T12:34:56Z"`). Epoch integers **MAY** be used only when explicitly required by the consuming system and **SHOULD** be documented.
- Money and other exact-quantity values **SHOULD NOT** be encoded as JSON floats. Prefer either an integer in the smallest unit (for example, cents) or a decimal-as-string with a documented format and currency.
- Arrays **SHOULD** be homogeneous. Heterogeneous arrays **SHOULD** be replaced with arrays of tagged objects (for example, `[{ "kind": "...", ... }, ...]`).
- `null` **SHOULD** be used only when "absent or unknown" is a meaningful, documented state. Otherwise, omit the key.

## Schema Policy

JSON schemas make JSON contracts checkable. The amount of schema rigor SHOULD scale with the file's importance.

- **Production or load-bearing JSON files MUST have schema validation.** This includes any JSON consumed by build, deploy, runtime configuration, or release automation where a malformed value would cause incorrect behavior.
- **Durable fixtures, examples, policy documents, and config contracts SHOULD have schema validation.** A "durable" artifact is one that other code or other repositories depend on remaining shaped a certain way.
- **Simple tool-owned configs MAY rely on the owning tool's validator** (for example, `tsconfig.json` is validated by the TypeScript compiler; `.eslintrc.json` is validated by ESLint). A separate schema **SHOULD NOT** be added when it would only duplicate the tool's own validation.

Schema location and shape:

- Schemas **SHOULD** live under a root-level `schemas/` directory, not under `.github/schemas/`. Keeping schemas at the repository root makes them discoverable to downstream consumers and to non-Git tooling.
- Schema files **SHOULD** use the `.schema.json` suffix (for example, `schemas/feature-flags.schema.json`).
- [JSON Schema Draft 2020-12](https://json-schema.org/draft/2020-12/schema) **SHOULD** be the default `$schema`, unless a specific consumer requires an earlier draft (for example, an OpenAPI version pinned to Draft-07). The chosen draft **SHOULD** be stated in the schema's `$schema` field and called out in the schema's accompanying documentation when it is not Draft 2020-12.
- Project-owned **closed** schemas **SHOULD** set `"additionalProperties": false` so that unknown keys are caught early.
- Ecosystem-mirroring schemas (schemas that describe an external format the project does not own, for example, a third-party config) **MAY** leave additional properties open and **SHOULD** document why in the schema's `description` or in a sibling `README.md`.

Reference and discriminant constraints:

When a property's value must correspond to an identifier, key, reference value, tag or discriminant value (such as a `kind` field; see [Data Modeling](#data-modeling) for tagged objects), module name, path-like key, or similar project-owned contract value whose shape or allowed values are constrained elsewhere in the same schema, a sibling schema, or another schema in the same schema-backed file family, that property **SHOULD** reuse the constraint rather than being left as a permissive non-empty string. Apply these rules:

- **Reuse before restating.** Prefer a shared `$defs` subschema referenced with `$ref` when the shared constraint lives in the same schema. Use a cross-file `$ref` only when the repository's validator setup reliably resolves that reference.
- **Closed value sets.** Prefer `enum` for a closed set, or `const` for a single fixed value, especially a branch-specific tag or discriminant value such as `kind`, when the value is stable and maintainable in schema.
- **Mirror with a drift guard.** When a constraint must be mirrored instead of reused directly, mirror the exact project-owned constraint, including case sensitivity, separator style, and path-safety rules, and guard against drift with an enforced check when practical, such as a schema-example test, semantic validator, generated closed set with a reproducibility check, or unit test. If no executable drift guard is practical, document the source of truth and why manual synchronization is acceptable.
- **Shape is not existence.** A `pattern` constrains lexical shape only; it does not prove that the referenced identifier exists. Whole-value patterns **SHOULD** be anchored with `^` and `$` unless substring matching is intentional and documented. When a validator may read the pattern in Python's dialect (python-jsonschema, or check-jsonschema with `--regex-variant python`), end it with `$(?![\s\S])`, because Python's `$` also matches before a final line break. A class that negates `\s` **SHOULD** also name `\ufeff`, which ECMA-262's `\s` holds and Python's does not. Confirming referential existence requires a closed `enum` or `const`, a semantic validator, or a test.
- **`format` is not a substitute.** Do not rely on `format` alone for project-specific identifier or reference constraints. JSON Schema Draft 2020-12 treats `format` as annotation by default unless assertion behavior is enabled by the validator or dialect; this repository's `check-jsonschema` usage checks supported formats by default and does not pass `--disable-formats`, while many JSON Schema validators do not check formats by default. When the contract requires a narrower constraint than `format` provides, pair `format` with an anchored project-owned `pattern`, a closed `enum` or `const`, or semantic validation.

Validation tooling:

- **Syntax validation.** Strict `.json` files **MUST** pass the repository's configured JSON syntax validator. When `check-json` is used, it validates strict JSON only; JSONC files need a JSONC-aware parser, linter, or schema validator if the repository wants automated JSONC enforcement.
- **Schema validation.** Schema-backed JSON files **MUST** pass every schema validator configured for their file family in the repository's pre-commit hooks or CI. Where no hook exists yet for a schema-backed file family, authors **SHOULD** run the applicable validator locally before committing and SHOULD add a file-family-scoped validator when the contract becomes durable.
- **Schema self-validation.** Project-owned schemas **SHOULD** be self-validated against their declared metaschema when the repository's validator supports that check.
- **Example fixture tests.** Repositories that keep valid and invalid schema examples **SHOULD** test both sides of the contract: valid examples pass, and invalid examples fail. Invalid examples **MUST NOT** be wired directly into a normal passing schema-validation hook unless that hook is explicitly designed to expect failure.
- **Active hook list.** When pre-commit or CI validation is retained, its configuration is the authoritative inventory of active validators. Repository-specific schema inventories, worked examples, removal checklists, and future validation candidates SHOULD live in the repository's schema documentation, if present.

## Comments and Documentation

- Strict JSON **MUST NOT** contain comments of any kind, including `//` line comments, `/* ... */` block comments, dummy `"_comment"` keys used as a comment workaround, or trailing-string hacks. JSONC **MAY** contain comments only when the consuming tool documents support for them.
- Documentation about a JSON file's shape and meaning **SHOULD** live in the schema's `description` and `title` fields and in a sibling `README.md`, not in the JSON file itself.
- Inline rationale for individual fields **SHOULD** be expressed in the schema's `description` for that property, so it is discoverable by anyone who reads the schema or uses a schema-aware editor.

## Security

- Secrets (API keys, tokens, connection strings, passwords, signing keys) **MUST NOT** be committed in any JSON file, including examples, fixtures, and tests.
- Example values **MUST** be obviously fake (for example, `"REPLACE_ME"`, `"example-token-not-real"`, `"example-api-key-not-real"`). Fake values **SHOULD NOT** resemble real credentials closely enough to trigger secret scanners or to mislead a reader into thinking they are real.
- Untrusted JSON (input from network, user, or other untrusted source) **MUST** be parsed using a safe parser and **MUST** be validated against a schema or explicit type checks before its values are used.
- JSON input **MUST NOT** be evaluated as executable code. In particular, `eval`, `Function`, `exec`, `Invoke-Expression`, and equivalent constructs in any language **MUST NOT** be used to "parse" JSON.
- JSON consumers **SHOULD** apply size and depth limits to defend against pathological inputs (deeply nested objects, very large arrays, very long strings).

## Generated JSON

- Generation **MUST** be reproducible: re-running the generator on the same inputs **MUST** produce byte-identical output. This is what makes diffs meaningful and what allows generated JSON to be safely committed.
- Generated files **MUST** use stable formatting (consistent indentation, consistent quoting, consistent line endings). The producer **SHOULD** write files with LF line endings explicitly; when baseline is retained, this also follows the repository's `.gitattributes` policy.
- Where ordering is non-semantic (for example, the order of keys in an object), the generator **MUST** apply a stable, documented ordering (for example, sorted by key, or grouped in a documented order). Where ordering **is** semantic (for example, an array of pipeline steps), the generator **MUST** preserve input order.
- Generated files **SHOULD** identify their source or generation command, either via a sibling `README.md` that names the generator, a top-level `"$generatedBy"` or similar property when the schema permits it, or a comment in the generator's own README. The identification **MUST NOT** violate the strict-JSON-no-comments rule: if the file is `.json`, use a property allowed by the schema or a sibling document, not an inline comment.

## Definition of Done for JSON Changes

A JSON change is considered done when **all** of the following hold:

- The file's dialect (strict JSON vs. JSONC) is correct for its extension and consumer.
- Formatting matches this guide (2-space indentation, no tabs, no trailing commas in strict JSON, double-quoted keys and strings).
- Key order is intentional: tool-managed and generated files have not been reordered by hand, and hand-authored objects preserve documented grouping.
- Naming follows the project's chosen convention (default `camelCase` for project-owned JSON; lowercase `kebab-case` filenames; affirmative boolean prefixes).
- Data modeling rules are followed: explicit objects over positional arrays, identifiers as strings, RFC 3339 timestamps, no floats for money or exact quantities.
- Schema validation is in place at the level required by the [Schema Policy](#schema-policy) tier.
- Strict JSON contains no comments; documentation lives in schemas and sibling docs.
- No secrets are committed; example values are obviously fake.
- Generated JSON is reproducible, stably formatted, stably ordered when ordering is non-semantic, and identifies its source or generation command.
- The repository's configured JSON or JSONC validators pass; schema and metaschema validators pass for any schema-backed file family wired into pre-commit or CI; retained schema-example tests pass after any schema or schema-example change. When pre-commit or Markdown checks are retained, the applicable checks pass for associated documentation changes.
