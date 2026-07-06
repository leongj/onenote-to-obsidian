# OneNote → Obsidian

Two GitHub Copilot **skills** for getting your notes out of Microsoft OneNote — either as a
one-time bulk migration to an [Obsidian](https://obsidian.md)-friendly Markdown vault, or as
live, ad-hoc queries against the running OneNote app.

Both work by driving the **desktop OneNote application** through its COM automation API
(the same `OneNote.Application` object model that tools like Excel expose). Nothing is sent
anywhere — it all runs locally against your own machine.

## What's in here

| Skill | What it does |
|-------|--------------|
| [`onenote-migrate-obsidian`](skills/onenote-migrate-obsidian) | Bulk-export a whole notebook (or a single section) to Markdown: preserves page hierarchy, extracts images and attachments, converts tables/lists/headings, and rewrites internal links to Obsidian wikilinks. |
| [`onenote-query`](skills/onenote-query) | Read your notes on demand without exporting: list the notebook tree, search pages by title or body text, and print any page as Markdown. |

## What it can do

- **Full-fidelity Markdown** from OneNote page XML — headings, bullet/numbered lists, tables,
  and text formatting.
- **Preserves your hierarchy.** OneNote lets you indent pages under other pages (subpages).
  The migration turns each parent into an Obsidian [*folder note*](https://help.obsidian.md/folder-notes):
  a folder named after the page, containing the page's own note plus its children.
- **Backlinks & navigation.** Every child page links back to its parent, and every top-level
  page links up to a per-section *hub note*. You can navigate the whole vault by links (and
  it looks great in Obsidian's graph view).
- **Images & attachments** are extracted to a single `_resources/` folder and embedded with
  `![[...]]`.
- **Internal OneNote links** (`onenote:` / `onenote://`) are rewritten to `[[wikilinks]]`
  where they can be resolved.
- Skips the OneNote recycle bin.

Proven on a real 935-page work notebook: full run in ~3 minutes, 0 errors.

## What it can't do (yet)

- **No handwriting/ink recognition, drawings, or embedded Office objects.** Images are
  extracted as-is; ink strokes are not converted to text.
- **Internal-link resolution is best-effort.** OneNote's hyperlink API returns a page GUID
  that sometimes doesn't match the page in the hierarchy; unresolved links fall back to plain
  text rather than breaking. In practice these are rare, but they exist.
- **No incremental / sync mode.** Each run is a full re-export. It's a migration tool, not a
  two-way sync.
- **Nesting depth** was tested to two levels (parent → child). Deeper nesting should work but
  is less battle-tested.
- The `onenote-query` page reader shows images as `*(image)*` placeholders — use the
  migration skill if you need the actual image files.

## Assumptions & requirements

This was originally built for one person's setup, so it makes some assumptions. It should
work for anyone who matches them:

- **Windows.** COM automation is Windows-only.
- **Desktop OneNote 2016 or 2021**, installed and running, with the target notebook open and
  synced. **The Microsoft Store / "OneNote for Windows 10" (UWP) app will NOT work** — it has
  no COM interface. If that's all you have, install desktop OneNote first.
- **Windows PowerShell 5.1** (the `powershell.exe` that ships with Windows). This is critical:
  COM activation of `OneNote.Application` **fails under PowerShell 7** (`pwsh`) with `E_FAIL`.
  The scripts must be run with the 5.1 host — the examples below show how.
- Works with both personal and work/school accounts, as long as the notebook is available in
  the *desktop* app (Obsidian's own built-in importer only handles personal accounts, which is
  part of why this exists).

## Installing the skills

These are [GitHub Copilot](https://github.com/features/copilot) skills. To install, copy the
skill folder(s) into your Copilot skills directory (`~/.copilot/skills/` on most setups):

```powershell
git clone https://github.com/leongj/onenote-to-obsidian.git
Copy-Item .\onenote-to-obsidian\skills\onenote-migrate-obsidian $env:USERPROFILE\.copilot\skills\ -Recurse
Copy-Item .\onenote-to-obsidian\skills\onenote-query           $env:USERPROFILE\.copilot\skills\ -Recurse
```

Then just ask Copilot something like *"migrate my OneNote to Obsidian"* or *"what did I write
about X in OneNote?"* and the relevant skill triggers.

## Running the scripts directly (without Copilot)

The scripts are self-contained and can be run on their own. Note the explicit path to the
Windows PowerShell 5.1 host:

```powershell
# Migrate a single section first so you can review the output shape:
& "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass `
  -File .\skills\onenote-migrate-obsidian\scripts\Export-OneNoteToObsidian.ps1 `
  -OutputRoot C:\test\Vault -SectionFilter "Work" -MaxPages 5

# Then the whole notebook:
& "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass `
  -File .\skills\onenote-migrate-obsidian\scripts\Export-OneNoteToObsidian.ps1 `
  -OutputRoot C:\test\Vault

# Query live:
& "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass `
  -File .\skills\onenote-query\scripts\Search-OneNote.ps1 -Query "project kickoff"
```

**Tip:** always export to a scratch folder (e.g. `C:\test\Vault`) and review it in Obsidian
before pointing it at your real vault.

## How it works (the short version)

The migration is a two-pass design, because OneNote's `GetHyperlinkToObject` returns a page
GUID that differs from the hierarchy page ID — so links can only be resolved once every
page's final path is known:

1. **Pass 1** walks the hierarchy, computes each page's final vault path (applying the
   folder-note nesting logic), and builds a GUID → path link map.
2. **Pass 2** pulls each page's XML, converts it to Markdown, saves images/attachments,
   rewrites links, and writes the files.

Full technical notes (COM method signatures, the PowerShell/COM gotchas, XML structure) live
in each skill's `SKILL.md`.

## License

[MIT](LICENSE) © 2026 Jason Leong. Provided as-is; not affiliated with or endorsed by Microsoft.
