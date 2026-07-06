---
name: onenote-query
description: Query a live desktop OneNote notebook directly from the terminal — list the notebook/section/page tree, search pages by title or full-text content, and read any page's content as Markdown — WITHOUT exporting or migrating anything. Use this skill whenever the user wants to look something up in their OneNote, find a note, "what did I write about X", pull the content of a specific page, or browse their notebook structure, and they haven't asked for a full migration/export. This reads the OneNote desktop app via COM in real time, so results always reflect the current notebook. It does NOT work with the OneNote Store/UWP app. For a one-off bulk conversion to Markdown/Obsidian, use the onenote-migrate-obsidian skill instead.
---

# OneNote live query

Read your OneNote notebook on demand — no migration required. Useful when you want the
answer *now* and don't want to maintain a separate copy of your notes.

## Critical constraints (same as the migration skill)

- **Windows PowerShell 5.1 ONLY.** COM activation fails under PowerShell 7 (`E_FAIL`).
  Always invoke via the 5.1 host explicitly:
  ```powershell
  & "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass `
    -File .\scripts\<script>.ps1 <args>
  ```
- **Desktop OneNote 2016/2021 must be running** with the notebook synced. The Microsoft
  Store / UWP "OneNote for Windows 10" app has no COM interface.

## The three scripts (in `scripts/`)

### 1. List the structure — `List-OneNoteHierarchy.ps1`
Prints notebooks → sections → (optionally) pages. Start here to see what exists and grab
page IDs. Nesting is shown by indentation (`pageLevel`).
```powershell
... -File .\scripts\List-OneNoteHierarchy.ps1                       # notebooks + sections
... -File .\scripts\List-OneNoteHierarchy.ps1 -SectionFilter Work -IncludePages
```

### 2. Find a page — `Search-OneNote.ps1`
Two modes:
- **Title scan (default)** — walks the hierarchy and matches page *titles* (case-insensitive
  substring). Reliable and index-independent; best when the user names a page.
- **`-Content`** — uses OneNote's native full-text index to match page *bodies*. Fast, but
  only covers indexed pages and uses word/AND semantics.
```powershell
... -File .\scripts\Search-OneNote.ps1 -Query "project kickoff"          # by title
... -File .\scripts\Search-OneNote.ps1 -Query "budget forecast" -Content   # by body text
```
Each hit prints the page title, its section breadcrumb, and its page ID.

### 3. Read a page — `Get-OneNotePage.ps1`
Fetches a page and converts it to readable Markdown (headings, lists, tables; images are
shown as `*(image)*` placeholders — use the migration skill if you need the actual image
files). Accept a page ID (exact) or a title (first substring match).
```powershell
... -File .\scripts\Get-OneNotePage.ps1 -Title "project kickoff"
... -File .\scripts\Get-OneNotePage.ps1 -PageId "{GUID}{1}{...}"
```

## Typical workflow

To answer "what did I note about the budget planning session?":
1. `Search-OneNote.ps1 -Query "budget" -Content` (or by title if you know the page name).
2. Take the page ID from the result.
3. `Get-OneNotePage.ps1 -PageId "<id>"` and read/summarise the returned Markdown.

Prefer **title search** when the user names a specific note, and **`-Content`** when they
describe a topic. If a title search returns nothing, fall back to a broader single-word
title query or a content search — OneNote's phrase matching is narrow.

## When to use the migration skill instead

This skill is for ad-hoc lookups against the live app. If the user wants a durable,
browsable, linked copy of their notes (Obsidian vault, Markdown files, images on disk),
that's a bulk export — use **onenote-migrate-obsidian**.

## Notes for maintenance

All three scripts share the same COM pattern (`New-Object -ComObject OneNote.Application`)
and namespace (`http://schemas.microsoft.com/office/onenote/2013/onenote`, prefix `one`).
Key gotchas already handled: `New-Ns` returns `,$ns` (the namespace manager is enumerable);
`FindPages`' schema arg is an enum 0–3 (pass `2` = xs2013, not `4`); use
`SelectNodes("//one:...", $ns)` rather than dotted property access.
