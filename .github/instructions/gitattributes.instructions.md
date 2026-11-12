---
applyTo: "**/.gitattributes"
description: "Rules for .gitattributes entries, including line-ending pinning, ordinary binary declassification, and opt-in Git LFS-managed binaries."
---

<!-- markdownlint-disable MD013 -->

# `.gitattributes` Rules

**Version:** 1.4.20260617.0

## Metadata

- **Status:** Active
- **Owner:** Repository Maintainers
- **Last Updated:** 2026-06-17
- **Scope:** Applies to any `.gitattributes` file in repositories that adopt these instructions, independent of programming language. Governs how committed text artifacts, linter-enforced LF file families, template-managed text formats, ordinary binary declassification, and opt-in Git LFS-managed binaries are protected against platform-dependent checkout rewriting or incorrect Git attribute handling.
- **Related:** [Repository Copilot Instructions](../copilot-instructions.md)

## Purpose and Scope

This file defines the normative rule for entries in `.gitattributes` that protect byte-exact text artifacts, linter-enforced LF file families, and template-managed text formats from platform-dependent line-ending rewriting. It also distinguishes ordinary binary declassification from opt-in Git LFS-managed binary attributes. It applies to every `.gitattributes` file in any repository that adopts these instructions, regardless of the programming languages used in the repository.

> **Note:** This document uses [RFC 2119](https://www.rfc-editor.org/rfc/rfc2119) keywords (**MUST**, **MUST NOT**, **SHOULD**, **SHOULD NOT**, **MAY**) to indicate requirement levels.

## Quick Reference Checklist

- **[All]** Any committed text file whose identity is its exact byte sequence **MUST** be pinned to LF in `.gitattributes` using a path pattern as specific as practical. See [Rule: Pin Byte-Exact Text Artifacts to LF](#rule-pin-byte-exact-text-artifacts-to-lf).
- **[All]** Any committed text file family whose repository-enforced linter or validator requires LF line endings **MUST** be pinned to LF in `.gitattributes`. See [Rule: Pin Linter-Enforced LF File Families to LF](#rule-pin-linter-enforced-lf-file-families-to-lf).
- **[Templates]** Common template-managed text formats that are edited or pruned across platforms **SHOULD** be pinned to LF when host-dependent CRLF conversion would create broad, non-semantic churn. See [Rule: Pin Template-Managed Text Formats to LF](#rule-pin-template-managed-text-formats-to-lf).
- **[All]** Ordinary binary formats that should remain in Git but must not receive text conversion **SHOULD** use Git's built-in `binary` macro, while formats intentionally stored through Git LFS **MUST** use explicit LFS attributes. See [Rule: Separate Ordinary Binary Overrides From Git LFS Tracking](#rule-separate-ordinary-binary-overrides-from-git-lfs-tracking).

## Rule: Pin Byte-Exact Text Artifacts to LF

Any committed text file whose identity is its exact byte sequence **MUST** be pinned to LF line endings in `.gitattributes` using a pattern as specific as practical to the affected artifact class. Examples of byte-exact artifacts include, but are not limited to:

- Golden or snapshot test baselines
- Expected-output fixtures compared by exact equality
- Files used as hash inputs (for example, inputs to checksum or digest computation)
- Signed payloads or signature verification inputs

The pattern **MUST** be expressed using standard `.gitattributes` syntax with both the `text` and `eol=lf` attributes. The pattern **SHOULD** be as narrow as practical (for example, scoping to a specific directory and extension) so that rules remain intentional and easy to audit.

**Example:**

```gitattributes
tests/**/golden/*.json text eol=lf
```

### Blanket Rules Are Not a Substitute

A blanket rule such as `* text=auto` **MUST NOT** be treated as a substitute for per-path `eol=lf` pinning. `text=auto` lets Git auto-detect text files and normalize to LF in the repository, but it does **not** force LF on Windows checkouts when `core.autocrlf=true` is in effect: Git still applies the working-tree line-ending conversion configured on the host and can rewrite LF to CRLF on checkout. Only explicit `eol=lf` on an artifact path guarantees that the bytes in the working tree match the bytes in the repository on every platform.

## Rule: Pin Linter-Enforced LF File Families to LF

Any committed text file family whose repository-enforced linter or validator requires LF line endings **MUST** be pinned to LF line endings in `.gitattributes` using a pattern as specific as practical to the affected file family. The rule applies even when the file's semantic consumer parses line endings equivalently, because the linter or validator has made the working-tree newline style part of the repository contract.

**Example:** If a retained YAML style validator rejects CRLF YAML files through a `new-lines: type: unix` rule, the affected repository should pin YAML files to LF:

```gitattributes
*.yml  text eol=lf
*.yaml text eol=lf
```

This rule does not imply that every parsed data format must be pinned by extension. File families whose configured tools parse or validate independently of line-ending style **SHOULD NOT** be LF-pinned solely for symmetry with another format.

## Rule: Pin Template-Managed Text Formats to LF

Template repositories that ship common text file families for downstream adoption **SHOULD** pin those file families to LF when platform-dependent checkout rewriting would create broad, non-semantic CRLF churn during adoption, stack pruning, or repeated cross-platform edits. This rationale is **CRLF-churn prevention for template-managed text formats**. It is distinct from byte-exact artifact protection and linter-enforced LF requirements: the consumer may parse CRLF correctly, but the template's review and adoption workflow benefits from stable LF working-tree bytes.

The pattern **SHOULD** use a low-risk extension or path family that the template actually owns. Do not add a broad extension pin merely because the format is text; document why the file family is template-managed and keep binary override behavior clear.

**Example:** This template pins Markdown, PowerShell, JSON, JSONC, TOML, JavaScript (`*.js`), JavaScript module (`*.mjs`), Python, shell script (`*.sh`), HCL, Terraform, dot-ignore, CODEOWNERS, `.gitattributes`, and root-license file families because they are shipped as template-managed documentation, scripts, configuration, control files, or examples and are commonly edited during downstream adoption. Shell scripts have an additional correctness rationale: a CRLF shebang can make Unix-like script execution look for an interpreter path containing a carriage return.

```gitattributes
*.md                          text eol=lf
*.mdc                         text eol=lf
*.ps1                         text eol=lf
*.psd1                        text eol=lf
*.psm1                        text eol=lf
*.json                        text eol=lf
*.jsonc                       text eol=lf
*.toml                        text eol=lf
*.js                          text eol=lf
*.mjs                         text eol=lf
*.py                          text eol=lf
*.sh                          text eol=lf
*.hcl                         text eol=lf
*.tf                          text eol=lf
*.tfvars                      text eol=lf
*.tftpl                       text eol=lf
*.tfbackend                   text eol=lf
.gitattributes                text eol=lf
.*ignore                      text eol=lf
CODEOWNERS                    text eol=lf
/LICENSE                      text eol=lf
```

## Rule: Separate Ordinary Binary Overrides From Git LFS Tracking

Ordinary binary overrides and Git LFS-managed attributes serve different contracts and **MUST NOT** be collapsed into one another.

- Use Git's built-in `binary` macro for binary formats that should stay as normal Git blobs while avoiding text conversion, text diffs, and text merges. The `binary` macro expands to `-diff -merge -text`; it does not configure Git LFS and should leave `filter` unspecified.
- Use explicit Git LFS attributes only for path families that the repository intentionally stores through Git LFS:

```gitattributes
*.psd filter=lfs diff=lfs merge=lfs -text
```

Repositories **SHOULD** adopt the optional `git-lfs` template module only when they intentionally store large opaque authoring or project files that benefit from LFS object storage, and when maintainers are prepared to support the required Git LFS workflow for contributors. Good candidates are project files such as Photoshop, Illustrator, InDesign, design-tool, 3D scene, CAD, or similarly opaque authoring formats. The initial template-managed LFS block is deliberately narrow and covers only these opaque project formats:

- `*.psd`
- `*.psb`
- `*.ai`
- `*.indd`
- `*.sketch`
- `*.fig`
- `*.blend`
- `*.fbx`
- `*.dwg`

Repositories **SHOULD NOT** move common media, archive, image, font, or PDF patterns into Git LFS solely because they are binary; this template already declassifies those formats as ordinary Git binaries. Text-ish asset formats such as `*.obj`, `*.gltf`, `*.dxf`, and `*.svg` **SHOULD NOT** be added to LFS rules unless maintainers deliberately decide those exact project formats should be LFS-managed and add tests for the effective attributes.

Project-specific LFS additions **SHOULD** use the narrowest practical pattern, one extension or path family per line. `.gitattributes` does not use shell brace expansion for path patterns, so authors **MUST NOT** write grouped forms such as `*.{psd,psb}`. Prefer explicit lines instead:

```gitattributes
assets/source/*.psd filter=lfs diff=lfs merge=lfs -text
assets/source/*.psb filter=lfs diff=lfs merge=lfs -text
```

Adding LFS attributes affects future Git clean/smudge behavior for matching paths; it does **not** migrate existing committed blobs, rewrite repository history, or convert already-tracked files into LFS objects. Repositories that want existing content stored as LFS objects still need Git LFS installed and must re-add or migrate those files through an explicit, reviewed Git LFS workflow.

## Defaults Shipped by This Template

This template ships a repo-root `.gitattributes` file with LF-pinning defaults for common byte-exact fixture locations, linter-enforced LF file families, and CRLF-churn prevention for template-managed text formats:

- `tests/**/golden/**`
- `tests/**/goldens/**`
- `tests/**/snapshots/**`
- `tests/**/__snapshots__/**`
- `tests/**/fixtures/**`
- `testdata/**`
- `*.yml`
- `*.yaml`
- `*.md`
- `*.mdc`
- `*.ps1`
- `*.psd1`
- `*.psm1`
- `*.json`
- `*.jsonc`
- `*.toml`
- `*.js`
- `*.mjs`
- `*.py`
- `*.sh`
- `*.hcl`
- `*.tf`
- `*.tfvars`
- `*.tftpl`
- `*.tfbackend`
- `.gitattributes`
- `.*ignore`
- `CODEOWNERS`
- `/LICENSE`

The fixture paths are assumed to contain **text** fixtures. To keep the defaults safe when binary assets are committed under the same directories (for example, `.png` screenshots under `__snapshots__/`), the shipped `.gitattributes` also declassifies a curated list of common binary extensions (images, documents and archives, compiled artifacts, audio and video, fonts) using the `binary` macro so that Git does not apply line-ending conversion to them.

The shipped root `.gitattributes` also contains an opt-in `git-lfs` inline block for the narrow opaque authoring/project formats listed in [Rule: Separate Ordinary Binary Overrides From Git LFS Tracking](#rule-separate-ordinary-binary-overrides-from-git-lfs-tracking). Downstream repositories that do not adopt the `git-lfs` module remove that delimited block during template materialization and retain only the ordinary binary overrides. Downstream repositories that do adopt `git-lfs` receive the explicit `filter=lfs diff=lfs merge=lfs -text` rules, but must still install and operate Git LFS intentionally; the template does not make Git LFS required for every adopter.

Downstream adopters **MUST** extend these entries whenever they introduce a new byte-exact fixture location that is not already covered (for example, a project-specific `expected/` directory, a `golden_files/` tree, or signed payloads under a custom path). New entries **SHOULD** follow the "as narrow as practical" guidance above. Existing template entries **SHOULD NOT** be removed unless the maintainer has confirmed that no byte-exact comparison exists in the repository that depends on those paths, no retained linter or validator requires LF for the file family, and the extension is no longer a template-managed text format whose cross-platform CRLF churn matters to adoption review.

### Excluding Binary Files Under Fixture Paths

If a project stores binary assets (for example, `.png` screenshots or `.zip` archives) inside a directory that is pinned to `text eol=lf`, those binaries **MUST** be declassified explicitly so Git does not rewrite them as if they were text. The standard idiom is a later-evaluated pattern that sets Git's built-in `binary` macro. Later patterns override earlier ones per attribute, so the binary override wins over the directory-wide pin:

```gitattributes
tests/**/snapshots/**    text eol=lf
*.png                    binary
*.zip                    binary
```

The shipped `.gitattributes` already declassifies a curated list of common binary extensions. Adopters **SHOULD** extend that ordinary-binary list if they commit binary fixtures in formats not covered and do not intend those files to be Git LFS-managed (for example, custom binary serialization formats, proprietary container formats, or simulator traces stored as opaque blobs). A narrower path-scoped form **MAY** be used (for example, `tests/**/fixtures/*.bin binary`) when a binary extension is project-specific and scoping it globally would be too broad.

## Interaction With Language-Specific Producer/Consumer Rules

The Git-layer rule defined here is necessary but not sufficient for byte-exact stability. Code that produces or compares these artifacts may also require language-specific normalization or read rules, for example:

- Tools that **generate** fixtures **SHOULD** write LF line endings explicitly, rather than relying on the host default.
- Tools that **compare** fixtures **SHOULD** read bytes in a mode that does not perform its own newline translation (for example, binary mode) when the comparison is byte-exact.
- Hashing and signing tools **SHOULD** operate on raw bytes and **MUST NOT** depend on on-disk text normalization.

These language-specific concerns are out of scope for this instructions file; they are addressed in the relevant language instructions that a repository retains. The Git-layer rule and the producer/consumer rules are complementary: each alone is insufficient, and both are needed for stable byte-exact artifacts across platforms.

## Rationale

Git's end-of-line handling is configurable per host. On Windows, the common default `core.autocrlf=true` rewrites LF to CRLF on checkout and CRLF to LF on commit. For a source file this is usually harmless; for a fixture whose identity is its exact byte sequence (for example, a file hashed into a test oracle), this silent rewriting breaks byte-exact comparisons — hash equality, signature verification, and snapshot diffs will fail on Windows even though the committed bytes are correct.

Per-path `eol=lf` pinning in `.gitattributes` is the durable Git-layer fix because it overrides `core.autocrlf` and any other host-level configuration for the specified paths. Producer-side normalization alone is insufficient: even if a generator writes LF bytes and a comparator reads in binary mode, a Windows checkout with `core.autocrlf=true` will still present CRLF bytes on disk, and any tool that reads the on-disk file (including hashing pipelines that are not explicitly reading in binary mode) will observe the rewritten bytes. Pinning the path to `eol=lf` is what guarantees that the bytes written to the working tree match the bytes stored in the repository, independent of host configuration.

The same Git-layer guarantee is required when a repository-enforced linter or validator makes LF line endings part of the contract. YAML style validation is a common example: when a retained validator rejects CRLF YAML through a `new-lines` rule, leaving YAML subject to host checkout conversion makes standard validation fail on Windows even though the parsed YAML data is unchanged. Pinning `*.yml` and `*.yaml` to `eol=lf` aligns the working tree with the configured validation contract.

Template-managed text formats have a third, weaker but still durable rationale: CRLF-churn prevention. Markdown, Cursor MDC, PowerShell, JSON, JSONC, TOML, JavaScript (`*.js`), JavaScript module (`*.mjs`), Python, HCL, Terraform, dot-ignore, CODEOWNERS, `.gitattributes`, and root `LICENSE` files are frequently touched during downstream adoption and stack pruning. Allowing host-specific checkout conversion for those file families creates large non-semantic diffs and can obscure the intended template change. LF pinning keeps adoption review focused on content while leaving binary safety to the explicit ordinary-binary and Git LFS-managed overrides. Shell scripts also need LF for direct execution on Unix-like hosts because a carriage return in the shebang line can become part of the interpreter path.
