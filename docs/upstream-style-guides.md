<!-- markdownlint-disable MD013 -->

# Upstream Style Guide Origins

The retained language instruction files are reviewed, vendored snapshots of generated upstream consumer guides. The template does not run the upstream generators. Reading and applying the retained guidance requires no live upstream fetch or source-repository write access. The immutable links below record origin; they grant no authority to change any repository.

Template maintainers refresh these snapshots. Downstream owners normally review later changes through their existing selective template-update process, preserving local ownership, overrides, and explicit protected-file decisions. A direct upstream refresh is a separate owner-chosen action. This origin record does not replace `.template-sync/marker.yml` ownership or decision records when sync support is retained. Each source blob identifies the imported upstream content, not any later locally customized destination bytes; record deliberate local differences alongside the affected entry.

<!-- template-sync: begin powershell-reference-only -->

## PowerShell

- **Source repository:** [franklesniak/PSStyleGuide](https://github.com/franklesniak/PSStyleGuide).
- **Source commit:** `986a78cfad02abe9698ee258735d8451abeb9249`.
- **Imported consumer path:** [`powershell.instructions.md`](https://github.com/franklesniak/PSStyleGuide/blob/986a78cfad02abe9698ee258735d8451abeb9249/powershell.instructions.md).
- **Imported consumer Git blob:** `534762988c0634d34c01059c9acb310e14c24371`.
- **Generated status:** Upstream generates the consumer from normative `STYLE_GUIDE.md`; see its [source-ownership instructions](https://github.com/franklesniak/PSStyleGuide/blob/986a78cfad02abe9698ee258735d8451abeb9249/AGENTS.md).
- **Destination:** [`.github/instructions/powershell.instructions.md`](../.github/instructions/powershell.instructions.md).
- **Deliberate local content differences:** None at import. The destination is a protected vendored snapshot, not a locally generated file.

<!-- template-sync: end powershell-reference-only -->
