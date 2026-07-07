---
name: onenote-migrate-obsidian
description: Migrate a desktop OneNote notebook into an Obsidian-friendly Markdown vault, in bulk and fully automated. Preserves page hierarchy (nested/indented pages become folder notes with backlinks), extracts embedded images and file attachments, converts tables/lists/headings, and rewrites internal OneNote links to Obsidian wikilinks. Use this skill whenever the user wants to export, migrate, convert, or back up OneNote to Markdown, Obsidian, or plain files — even if they don't name the script — including phrases like "get my notes out of OneNote", "move to Obsidian", "OneNote to Markdown", or "export my notebook". This runs locally against the OneNote desktop app via COM; it does NOT work with the OneNote Store/UWP app or personal-account cloud-only notebooks that lack a synced desktop copy.
---

# OneNote → Obsidian migration

This skill drives the **desktop** OneNote application through its COM object model to
extract pages as clean Markdown, producing a vault Obsidian can open directly. It was
built and proven on a real 935-page work/school notebook (0 errors, ~3 min).

## Critical constraints (read first — these cause the most failures)

1. **Windows PowerShell 5.1 ONLY.** COM activation of `OneNote.Application` fails under
   PowerShell 7 with `E_FAIL`. Always invoke the script via the 5.1 host explicitly:
   ```powershell
   & "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass `
     -File .\scripts\Export-OneNoteToObsidian.ps1 -OutputRoot C:\path\to\Vault
   ```
   Do not run it with the default `powershell`/`pwsh` on modern systems if that resolves to 7.x.
2. **Desktop OneNote 2016/2021 must be installed and running**, with the target notebooks
   open and synced. The Microsoft Store / UWP "OneNote for Windows 10" app has **no COM
   interface** and will not work. If the user only has that, tell them plainly.
3. **We use the XML route, not Publish-to-DOCX.** `GetPageContent` (XML) is reliable;
   `Publish` to DOCX/HTML/MHTML fails on many machines (`0x80042031`). Off-the-shelf tools
   that rely on the DOCX path (e.g. alxnbl/onenote-md-exporter) can fail for this reason —
   that is why this skill converts the XML ourselves.

## How to run it

Always test on a single section before a full run, so the user can review the output
shape in Obsidian first.

```powershell
# 1. Smoke test — one section, capped
& "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass `
  -File .\scripts\Export-OneNoteToObsidian.ps1 -OutputRoot C:\test\Vault -SectionFilter "Work" -MaxPages 5

# 2. Full section
& "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" ... -SectionFilter "Work"

# 3. Whole notebook (all sections)
& "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" ... -OutputRoot C:\test\Vault
```

Parameters (`scripts\Export-OneNoteToObsidian.ps1`):

| Param            | Meaning                                                            |
|------------------|-------------------------------------------------------------------|
| `-OutputRoot`    | **Required.** Destination vault folder (created if missing).      |
| `-SectionFilter` | Only export sections matching this (wildcards, e.g. `Work`, `*`).  |
| `-NotebookFilter`| Only export notebooks matching this.                              |
| `-MaxPages`      | Stop after N pages — for smoke testing.                           |

The script **deletes and recreates nothing outside `-OutputRoot`**, but it does write
freely inside it. Point it at a scratch folder (e.g. `C:\test\Vault`) for testing, and
only at the user's real Obsidian vault once they've reviewed the output.

## What the output looks like

- **Hierarchy preserved.** OneNote nests pages via the `pageLevel` attribute (indentation
  in the page tab). A page that has deeper-level pages after it becomes a *folder note*:
  a folder named after the page, containing the page's own note plus its children. This is
  the standard Obsidian "folder note" convention and keeps the graph tidy.
- **Backlinks.** Every child page carries a visible `**Parent:** [[...]]` link. Top-level pages link up to a per-section **hub note**
  (e.g. `Work/Work.md`, `type: section`) that lists all its pages. This lets the user
  navigate purely by links and eventually drop folders if they want.
- **Metadata at the end.** OneNote provenance (`title`, `created`, `updated`,
  `onenote-id`, `parent`, `source`) is preserved in a fenced
  `## OneNote conversion metadata` block at the end of each note, so the note opens on
  its actual content rather than a YAML block.
- **Images & attachments** go to a single vault-root `_resources/` folder and are embedded
  with `![[...]]`. Image format is sniffed from magic bytes (OneNote leaves the `format`
  attribute empty).
- **Internal links** (`onenote:` / `onenote://`) are rewritten to `[[wikilinks]]` using a
  GUID→path map built in pass 1. Unresolved ones (page OneNote can't cross-match) fall back
  to plain text rather than breaking.
- **Recycle bin is skipped.**

## How it works (for maintenance)

Two-pass design, because OneNote's `GetHyperlinkToObject` returns a `page-id` GUID that
**differs** from the hierarchy page ID — so links can only be resolved after every page's
final path is known.

- **Pass 1** walks `GetHierarchy(..., 4, ...)` (4 = `hsPages`), computes each page's final
  vault path (applying nesting/folder-note logic), pre-registers section hubs, and builds
  the GUID→path link map.
- **Pass 2** pulls each page via `GetPageContent(id, ..., 1)` (1 = `piBinaryData`, embeds
  image base64), converts the `one:*` XML tree to Markdown, saves resources, rewrites links,
  writes files, then writes the section hub notes.

Key gotchas already handled in the script (don't reintroduce these bugs):
- Use `$doc.SelectNodes("//one:Node", $ns)` — dotted access like `$hier.Notebooks.Notebook`
  returns an adapted object missing the 2-arg `SelectNodes` overload.
- `New-Ns` must `return ,$ns` — `XmlNamespaceManager` is enumerable, so a bare `return`
  unrolls it into an `Object[]`.
- Sanitise filenames: strip `[\/:*?"<>|#^\[\]]`, de-collide with `(2)`, `(3)` suffixes.

## Reusability notes

The script is self-contained and portable — hand someone `scripts\Export-OneNoteToObsidian.ps1`
plus this SKILL.md and they can migrate their own notebook. The only environment
requirements are: Windows, desktop OneNote, and Windows PowerShell 5.1 (ships with Windows).
