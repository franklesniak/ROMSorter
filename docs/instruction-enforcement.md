<!-- markdownlint-disable MD013 -->
# Static Instruction Enforcement

## Metadata

- **Status:** Active
- **Owner:** Repository Maintainers
- **Last Updated:** 2026-09-23
- **Scope:** Optional static instruction checks, applicability, local declarations, and migration.
- **Related:** [Repository instructions](../.github/copilot-instructions.md)

Select `agent-instructions` and `instruction-enforcement` to retain the checker, protected local profile and reviewed catalog. Select each desired agent module explicitly: `agent-copilot`, `agent-codex`, `agent-claude`, `agent-cursor`, `agent-gemini`, or `agent-hermes`. The language modules independently control their instruction contracts. Selecting only `agent-instructions` retains shared guidance; omitting enforcement is an explicit **policy-only** choice without the standalone validation runtime.

| Selected modules | Execution route |
| --- | --- |
| Instructions, enforcement, baseline | `pre-commit run validate-instruction-profile --all-files` |
| Instructions, enforcement, GitHub Actions | Dedicated Instruction Contracts workflow |
| Instructions, enforcement, Azure Pipelines | Dedicated instruction-contracts pipeline, registered through Azure DevOps Services |
| Instructions, enforcement, both hosts | Both dedicated routes |
| Instructions, enforcement, no baseline or host | Invalid; select a runner or choose policy-only |
| Enforcement without instructions | Invalid |

The optional Python **project** module is never required. Baseline reuses pre-commit's Python environment; each dedicated host route installs Python, PyYAML and jsonschema. GitHub is the default host, and Azure is optional additive support. Azure YAML parsing locally does not establish service-backed pipeline or branch-policy validation.

For GitHub branch protection or rulesets, select the **Instruction Contracts** job check after it has run successfully. Its job name distinguishes it from **Workflow Security**. Existing rulesets that required the previous `validate` check need an owner to select the intended named checks; updating workflow files does not update repository settings.

## Explicit modes

Run `python .github/scripts/validate_instruction_profile.py` after installing its declared PyYAML and jsonschema dependencies. Missing, malformed or contradictory applicability data fails with a nonzero exit. Filesystem errors while resolving the default or explicitly supplied repository root also fail with a concise diagnostic that omits the exception's filename fields.

The shared loaders reject repeated explicit YAML mapping keys and JSON object names before schema validation, including nested mappings and objects. For example, two `modules` keys fail instead of silently retaining the second selection. Supported safe YAML aliases and merge overrides remain valid: an explicit key may override a value inherited from a merge. The same parsers check YAML and JSON argument files supplied to the materializer before adoption.

Instruction profile and catalog schemas permit fragment references within the supplied schema, such as `#/$defs/module`, and reject nonfragment `$ref` and `$dynamicRef` declarations even in unused schema branches. Schema resolution performs no file or network retrieval. Reference-looking keys inside annotation data remain data. An unresolved local fragment fails with a schema-resolution diagnostic rather than an uncaught traceback.

Protected-guide section and reference obligations share one `(path, key)` identity space. A catalog cannot reuse that pair across obligation types, even when their target modules differ or neither obligation currently fails. The same key on different source paths remains valid. This prevents one marker waiver from silently covering two different obligations; standalone exceptions retain their distinct section and reference anchors.

- **Marker mode:** `.github/instruction-profile.yml` contains `version: 1`, `mode: marker` and an explicit `context`: `upstream-template` in the source template, or `downstream` after adoption. The downstream context requires its marker; absence never falls back to upstream applicability. The marker-aware adapter remains authoritative for retained modules, protected decisions, local ownership and waivers. The upstream template validates every catalog obligation; a downstream marker uses its explicit selection.
- **Standalone mode:** the protected profile contains `version: 1`, `mode: standalone`, explicit `modules` and `exceptions` lists. It reads the protected `.github/instruction-contracts.yml` catalog and its standalone schema. Its validation engine and bounded helpers live under `.github/scripts/`; the standalone check runs after sync-support files have been physically removed.

Do not activate both modes. Standalone performs only an existence check for a conflicting marker; it never reads or requires that marker, and removing all support files does not prevent validation. The marker adapter rejects a standalone profile, and a standalone profile cannot select sync support. When changing modes, use reviewed materialization and resolve protected-file decisions before applying the resulting candidate.

Relative-link checks recognize multiline labels and titles, including reference definitions whose destination starts on the next line. An escaped exclamation mark before a link does not turn it into an image: `\![Guide](target.md)` checks the link, while `![Image](target.md)` remains an image. They preserve the opening line number and do not combine fragments across fenced examples or blank lines. This is bounded target extraction for static checks, not a complete Markdown renderer.

Relative links in indented code blocks are literal examples. Code indentation is measured within quote and list containers; indentation that continues an open paragraph remains live. Tabs use structural column stops without changing the reported destination. For embedded Markdown in YAML, structural YAML indentation keeps its existing live-link behavior.

An empty list marker, such as a lone `-`, retains its continuation margin without causing an index error. A following blank line ends that empty item while preserving established outer containers. Subsequent indentation is measured against those remaining containers. A single-dash setext underline after paragraph text forms a heading without adding a list margin; links in the heading remain checked, while following indented code remains literal.

Raw HTML blocks, including comments, contain literal examples rather than Markdown links. Inline HTML tokens also keep their contents literal, such as a link-shaped string inside a tag attribute; ordinary text between inline tags remains live. HTML inside a link label does not supply label brackets. Fences inside raw HTML do not hide links after that HTML block ends. Destinations and titles retain their own syntax and original spelling. Embedded YAML keeps its existing fence context rather than treating the entire file as a Markdown document.

Matched code spans inside an inline link label do not supply label brackets. For example, ``[a `]`](target.md)`` still checks `target.md`; the bracket inside code does not end the label. Link-shaped text wholly inside a code span is literal and produces no target. Unmatched backticks remain literal. Destinations, titles and reference-definition labels retain their own syntax, so a backtick in those components does not hide a later link.

When an ordinary link label contains another live link, the inner link wins: `[outer [Guide](inner.md)](outer.md)` checks `inner.md`, while the outer destination is literal text. Image descriptions remain non-anchor text, including links inside their descriptions; an image inside a real link does not hide that enclosing link. Recognized full, collapsed and shortcut reference links also affect enclosing link activity. Genuine reference-definition destinations remain part of the target inventory even when unused, without adding duplicate targets at each reference use. Definition-shaped text cannot interrupt an open paragraph. For example, `[Guide]` immediately followed by `[Guide]: target.md` on the next line, without a blank line, supplies no definition target. Real inline links in that paragraph remain checked.

Before comparing a destination with a catalog path, the checker decodes Markdown punctuation escapes and valid character references in one pass. For example, `target\.md` and `target&#46;md` both compare as `target.md`. Replacements are not decoded again, and ordinary URL percent decoding follows afterward. Reports and exception anchors keep the exact original destination spelling; equivalent rendered paths do not share an exception automatically.

## Local exceptions and migration

Each standalone exception names one exact `path`, one exact reported `anchor`, the file's `content_sha256`, a `reason`, and an `authorization_basis` declaration. The content digest is SHA-256 over strict UTF-8 text with CRLF/CR normalized to LF. `file:absent` uses the literal digest `absent`. Paths cannot name directories or traversal. An unrelated content edit invalidates the declaration; one exception cannot excuse another file or anchor. Active Claude imports and tracked local memory remain failures and cannot be excepted.

Every applied exception is reported. A schema-valid declaration is auditable local data, **not independent proof of owner authorization**. Owner approval remains an external repository process. Candidate-owned static checks cannot establish human permission, arbitrary natural-language compliance, future agent behavior, or resistance to an attacker editing the checker and its catalog together.

Initial materialization renders the explicit selected profile. Removing sync support translates relevant anchor/removal declarations to exact local exceptions and scopes stale-section waivers to observed content. It preserves source decisions as migration evidence without treating their historical authority as a new grant. Review the protected profile candidate and any reported failure before removing the old support paths. Reintroducing sync with local declarations requires explicitly translating them back to reviewed marker decisions first; the tool refuses an implicit loss of declarations.

Before changing downstream files, materialization checks that the selected enforcement profile, runtime and mode-specific schemas and catalog are regular files. `SKIP` therefore requires existing local files. A skipped profile must match the requested mode. Standalone mode requires the same module set; marker mode requires `downstream` context. These checks reject missing inputs and conflicting applicability before target writes. Review the preserved profile or requested selection before retrying. Valid local bytes and declarations remain owner-controlled; these checks do not validate every preserved runtime file's contents.

Standalone migration uses the catalog selected for installation, including a locally preserved catalog. When retaining a local catalog or explicitly skipping a standalone profile, materialization checks catalog obligations against the selected guide content and effective profile exceptions before changing downstream files. A skipped profile supplies its own exceptions; it cannot borrow declarations from a discarded profile candidate, even when taking the catalog. An effective profile in another mode requires review before migration proceeds. Compatible local catalogs and skipped profiles remain unchanged. Incompatible combinations fail before installation, using the tool's reviewed schemas. Catalog semantic errors produce controlled failure diagnostics rather than uncaught tracebacks.

An explicit applicable `REMOVE-LOCAL` decision records a removal; materialization does not delete the local file. Standalone migration requires that removal to be completed before treating the selected content as absent. If the target file remains, migration fails before changing downstream files. When the target is already absent, a matching original `file:absent` declaration keeps its digest and rationale; other content-bound declarations cannot be renewed to excuse the removal. No stale-section or reference exception is generated for the absent file.

Newly migrated waivers follow the file selection: protected files use their explicit protected decisions, and other files use the most specific matching local override. `TAKE` uses the staged candidate when present, and `SKIP` uses the preserved local content. Section or reference waivers produce exceptions only for failures in that selected content. An applicable direct instruction waiver must also match a current failure for either selection. If its anchor no longer fails, migration rejects the conflicting waiver before writing the profile. Review or remove that waiver explicitly before retrying.

Existing standalone exceptions retain their original content hashes; changing the selected content does not silently renew them. Migration rejects an applicable retained exception before writing the profile when its original hash or failure anchor does not match the selected content. For example, `TAKE` of a changed guide requires review of its old exception even when the same anchor still fails. Review or remove the incompatible declaration explicitly, then rerun materialization. Declarations for excluded module scopes still retire into the recorded source decisions.

Repeated materialization preserves identical content and surfaces changed protected local profiles through ordinary protected-file reconciliation. Later module removal uses the same reviewed cleanup: remove the owned runtime, schemas, hooks, host routes and references together. A legacy selection that predates per-agent modules must explicitly select the agents to retain before cleanup; no agent deletion should be inferred from the new taxonomy.

The standalone catalog is seeded from the reviewed template obligations. Template maintainers update that seed with the marker catalog and test their equality. Downstream owners control their retained protected profile/catalog; weakening that catalog is a governance change requiring explicit authorization.
