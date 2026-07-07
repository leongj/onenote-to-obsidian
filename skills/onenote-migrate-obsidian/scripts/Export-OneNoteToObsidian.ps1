<#
.SYNOPSIS
    Export OneNote notebooks to an Obsidian-friendly Markdown vault via COM automation.

.DESCRIPTION
    Drives the *desktop* OneNote application (2016/2021) through its COM object model
    (OneNote.Application) to extract pages as Markdown. Uses GetPageContent XML (no Word /
    no Publish dependency) so it works even where Publish-to-DOCX fails.

    Two-pass design:
      Pass 1 - walk the hierarchy, compute final file paths, and build an internal-link map
               (OneNote hyperlink page-id GUID -> target note) so cross-page links resolve.
      Pass 2 - convert each page's XML to Markdown, extract images/attachments to a central
               _resources/ folder, and rewrite onenote:// links to [[wikilinks]].

    MUST be run with Windows PowerShell 5.1 (NOT PowerShell 7 - COM activation fails there).

.PARAMETER OutputRoot
    Destination vault folder. Created if missing.

.PARAMETER SectionFilter
    Optional. Only export sections whose name matches this (wildcards allowed, e.g. 'Work').

.PARAMETER NotebookFilter
    Optional. Only export notebooks whose name matches this.

.PARAMETER MaxPages
    Optional. Stop after N pages (for smoke testing).

.EXAMPLE
    powershell.exe -File Export-OneNoteToObsidian.ps1 -OutputRoot C:\test\ObsidianVault -SectionFilter Work
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $OutputRoot,
    [string] $SectionFilter = "*",
    [string] $NotebookFilter = "*",
    [int]    $MaxPages = 0
)

$ErrorActionPreference = "Stop"
$oneNS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

# ------------------------------------------------------------------ helpers
function New-Ns([xml]$doc) {
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("one", $oneNS)
    return ,$ns
}

function Get-SafeName([string]$name, [int]$maxLen = 120) {
    if ([string]::IsNullOrWhiteSpace($name)) { $name = "Untitled" }
    $clean = [regex]::Replace($name, '[\\/:*?"<>|#\^\[\]]', '_')
    $clean = $clean -replace '\s+', ' '
    $clean = $clean.Trim().TrimEnd('.')
    if ($clean.Length -gt $maxLen) { $clean = $clean.Substring(0, $maxLen).Trim() }
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = "Untitled" }
    return $clean
}

function Get-UniqueName([string]$base, [System.Collections.Generic.HashSet[string]]$used) {
    $candidate = $base; $i = 2
    while ($used.Contains($candidate.ToLowerInvariant())) {
        $candidate = "$base ($i)"; $i++
    }
    [void]$used.Add($candidate.ToLowerInvariant())
    return $candidate
}

function Get-ImageExt([byte[]]$b) {
    if ($b.Length -ge 4 -and $b[0] -eq 0x89 -and $b[1] -eq 0x50 -and $b[2] -eq 0x4E -and $b[3] -eq 0x47) { return "png" }
    if ($b.Length -ge 3 -and $b[0] -eq 0xFF -and $b[1] -eq 0xD8 -and $b[2] -eq 0xFF) { return "jpg" }
    if ($b.Length -ge 3 -and $b[0] -eq 0x47 -and $b[1] -eq 0x49 -and $b[2] -eq 0x46) { return "gif" }
    if ($b.Length -ge 2 -and $b[0] -eq 0x42 -and $b[1] -eq 0x4D) { return "bmp" }
    return "bin"
}

# ------------------------------------------------------------------ inline text -> markdown
# Uses script-scoped $script:LinkMap to rewrite onenote: links to [[wikilinks]].
function Convert-Inline([string]$t) {
    if ($null -eq $t) { return "" }

    # onenote: internal links -> [[Target]]  (match by page-id GUID)
    $t = [regex]::Replace($t, '(?is)<a[^>]*href="(onenote:[^"]*)"[^>]*>(.*?)</a>', {
        param($m)
        $href = $m.Groups[1].Value
        $disp = $m.Groups[2].Value
        $guidMatch = [regex]::Match($href, 'page-id=\{([0-9A-Fa-f\-]+)\}')
        if ($guidMatch.Success) {
            $guid = $guidMatch.Groups[1].Value.ToUpperInvariant()
            if ($script:LinkMap.ContainsKey($guid)) {
                $script:LinksResolved++
                $target = $script:LinkMap[$guid]
                # strip inline tags from display text
                $dispClean = ([regex]::Replace($disp, '(?is)<[^>]+>', '')).Trim()
                if ($dispClean -and $dispClean -ne $target) { return "[[$target|$dispClean]]" }
                return "[[$target]]"
            }
        }
        $script:LinksUnresolved++
        return ([regex]::Replace($disp, '(?is)<[^>]+>', '')).Trim()
    })

    # regular hyperlinks
    $t = [regex]::Replace($t, '(?is)<a[^>]*href="([^"]*)"[^>]*>(.*?)</a>', {
        param($m)
        $u = $m.Groups[1].Value
        $d = ([regex]::Replace($m.Groups[2].Value, '(?is)<[^>]+>', '')).Trim()
        if (-not $d) { $d = $u }
        "[$d]($u)"
    })

    # bold / italic
    $t = [regex]::Replace($t, "(?is)<span[^>]*font-weight:\s*bold[^>]*>(.*?)</span>", '**$1**')
    $t = [regex]::Replace($t, "(?is)<span[^>]*font-style:\s*italic[^>]*>(.*?)</span>", '_$1_')
    # line breaks
    $t = [regex]::Replace($t, "(?is)<br[^>]*>", "`n")
    # strip remaining tags
    $t = [regex]::Replace($t, "(?is)<[^>]+>", "")
    # decode entities
    $t = $t -replace "&lt;", "<" -replace "&gt;", ">" -replace "&quot;", '"' `
            -replace "&#39;", "'" -replace "&apos;", "'" -replace "&nbsp;", " " -replace "&amp;", "&"
    return $t.Trim()
}

# ------------------------------------------------------------------ walk one page's outline tree
function Walk-OEChildren($oeChildren, [int]$depth, [hashtable]$styleMap, [ref]$sb, $page, $ns) {
    foreach ($oe in $oeChildren.SelectNodes("one:OE", $ns)) {
        $indent  = "  " * $depth
        $table   = $oe.SelectSingleNode("one:Table", $ns)
        $image   = $oe.SelectSingleNode("one:Image", $ns)
        $insFile = $oe.SelectSingleNode("one:InsertedFile", $ns)
        $tNode   = $oe.SelectSingleNode("one:T", $ns)
        $listNode= $oe.SelectSingleNode("one:List", $ns)

        if ($image) {
            $fn = Save-EmbeddedImage $image $page $ns
            if ($fn) { $sb.Value.AppendLine("$indent![[$fn]]") | Out-Null; $sb.Value.AppendLine("") | Out-Null }
        }

        if ($insFile) {
            $fn = Save-Attachment $insFile
            if ($fn) { $sb.Value.AppendLine("$indent[[$fn]]") | Out-Null; $sb.Value.AppendLine("") | Out-Null }
        }

        if ($table) {
            $rows = $table.SelectNodes("one:Row", $ns); $r = 0
            foreach ($row in $rows) {
                $cells = @()
                foreach ($cell in $row.SelectNodes("one:Cell", $ns)) {
                    foreach ($cimg in $cell.SelectNodes(".//one:Image", $ns)) { [void](Save-EmbeddedImage $cimg $page $ns) }
                    $cellText = ($cell.SelectNodes(".//one:T", $ns) | ForEach-Object { (Convert-Inline $_.InnerText) -replace '\r?\n', ' ' -replace '\|', '\|' }) -join " "
                    $cells += $cellText
                }
                $sb.Value.AppendLine("| " + ($cells -join " | ") + " |") | Out-Null
                if ($r -eq 0) { $sb.Value.AppendLine("| " + (($cells | ForEach-Object { "---" }) -join " | ") + " |") | Out-Null }
                $r++
            }
            $sb.Value.AppendLine("") | Out-Null
            continue
        }

        if ($tNode) {
            $text = Convert-Inline $tNode.InnerText
            if ($text -ne "") {
                # heading via quick style (h1..h6)?
                $qsi = $oe.quickStyleIndex
                $styleName = if ($qsi -ne $null -and $styleMap.ContainsKey([string]$qsi)) { $styleMap[[string]$qsi] } else { $null }
                if ($styleName -match '^h([1-6])$') {
                    $level = [int]$Matches[1]
                    $sb.Value.AppendLine(("#" * $level) + " " + $text) | Out-Null
                    $sb.Value.AppendLine("") | Out-Null
                }
                elseif ($listNode) {
                    $marker = if ($listNode.SelectSingleNode("one:Number", $ns)) { "1." } else { "-" }
                    $sb.Value.AppendLine("$indent$marker $text") | Out-Null
                }
                else {
                    $sb.Value.AppendLine("$indent$text") | Out-Null
                    $sb.Value.AppendLine("") | Out-Null
                }
            }
        }

        $childWrap = $oe.SelectSingleNode("one:OEChildren", $ns)
        if ($childWrap) { Walk-OEChildren $childWrap ($depth + 1) $styleMap $sb $page $ns }
    }
}

# ------------------------------------------------------------------ save an embedded image (base64)
function Save-EmbeddedImage($img, $page, $ns) {
    $dataNode = $img.SelectSingleNode("one:Data", $ns)
    if (-not $dataNode -or [string]::IsNullOrEmpty($dataNode.InnerText)) { return $null }
    try { $bytes = [Convert]::FromBase64String($dataNode.InnerText) } catch { return $null }
    $ext  = Get-ImageExt $bytes
    $name = "$($page.Slug)-img-$($script:ImgCounter).$ext"
    $script:ImgCounter++
    [System.IO.File]::WriteAllBytes((Join-Path $script:ResDir $name), $bytes)
    $script:ImagesSaved++
    return $name
}

# ------------------------------------------------------------------ save an inserted file attachment
function Save-Attachment($insFile) {
    $src = $insFile.pathCache
    if ([string]::IsNullOrEmpty($src)) { $src = $insFile.pathSource }
    if ([string]::IsNullOrEmpty($src) -or -not (Test-Path $src)) {
        $script:AttachMissing++
        return $null
    }
    $preferred = $insFile.preferredName
    if ([string]::IsNullOrEmpty($preferred)) { $preferred = Split-Path $src -Leaf }
    $safe = Get-UniqueName (Get-SafeName $preferred) $script:ResUsed
    $dest = Join-Path $script:ResDir $safe
    Copy-Item -LiteralPath $src -Destination $dest -Force
    $script:AttachSaved++
    return $safe
}

# ================================================================== MAIN
$swAll = [System.Diagnostics.Stopwatch]::StartNew()
Write-Host "Connecting to OneNote (COM)..." -ForegroundColor Cyan
$on = New-Object -ComObject OneNote.Application

$hierXml = ""
$on.GetHierarchy("", 4, [ref]$hierXml)   # 4 = hsPages
[xml]$hier = $hierXml
$ns = New-Ns $hier

# Collect target pages (respecting notebook/section filters)
$targets = New-Object System.Collections.ArrayList
foreach ($nb in $hier.SelectNodes("//one:Notebook", $ns)) {
    if ($nb.name -notlike $NotebookFilter) { continue }
    foreach ($sec in $nb.SelectNodes(".//one:Section", $ns)) {
        if ($sec.name -notlike $SectionFilter) { continue }
        if ($sec.isInRecycleBin -eq "true") { continue }
        # skip sections inside the recycle bin section group
        $inBin = $false
        $anc = $sec
        while ($anc -ne $null -and $anc.LocalName -ne "Notebook") {
            if ($anc.isRecycleBin -eq "true" -or $anc.isInRecycleBin -eq "true") { $inBin = $true; break }
            $anc = $anc.ParentNode
        }
        if ($inBin) { continue }
        $sectionPath = @()
        # build section-group path within notebook
        $node = $sec
        while ($node -ne $null -and $node.LocalName -ne "Notebook") {
            if ($node.LocalName -eq "Section" -or $node.LocalName -eq "SectionGroup") { $sectionPath = ,($node.name) + $sectionPath }
            $node = $node.ParentNode
        }
        $relFolder = ($sectionPath | ForEach-Object { Get-SafeName $_ }) -join [IO.Path]::DirectorySeparatorChar
        foreach ($pg in $sec.SelectNodes("one:Page", $ns)) {
            [void]$targets.Add([pscustomobject]@{
                Id            = $pg.ID
                Title         = $pg.name
                Notebook      = $nb.name
                PageLevel     = $pg.pageLevel
                SectionName   = $sec.name
                SectionFolder = Join-Path (Get-SafeName $nb.name) $relFolder
            })
        }
    }
}
if ($MaxPages -gt 0 -and $targets.Count -gt $MaxPages) {
    $targets = $targets[0..($MaxPages-1)]
}
Write-Host ("Target pages: {0}" -f $targets.Count) -ForegroundColor Green

# Output dirs
$null = New-Item -ItemType Directory -Force -Path $OutputRoot
$script:ResDir = Join-Path $OutputRoot "_resources"
$null = New-Item -ItemType Directory -Force -Path $script:ResDir
$script:ResUsed = New-Object 'System.Collections.Generic.HashSet[string]'

# ---------- PASS 1: compute nested paths + parent/child tree + link map ----------
Write-Host "Pass 1/2 - mapping pages, nesting and links..." -ForegroundColor Cyan
$script:LinkMap = @{}                       # hyperlink page-id GUID -> note vault path
$usedPerFolder = @{}                        # folder -> hashset of used names (files + subfolders)
$pagePlan = New-Object System.Collections.ArrayList
$stack = @{}          # level -> folder that children of that page live in
$lastAtLevel = @{}    # level -> pagePlan index of most recent page at that level

# Pre-register a "section note" for every distinct section (its own folder note + hub).
$sectionInfo = @{}    # sectionFolder -> @{ Title; BaseName; VaultPath; Children(ArrayList) }
foreach ($t in $targets) {
    if ($sectionInfo.ContainsKey($t.SectionFolder)) { continue }
    if (-not $usedPerFolder.ContainsKey($t.SectionFolder)) {
        $usedPerFolder[$t.SectionFolder] = New-Object 'System.Collections.Generic.HashSet[string]'
    }
    $secBase = Get-UniqueName (Get-SafeName $t.SectionName) $usedPerFolder[$t.SectionFolder]
    $secVault = (($t.SectionFolder + [IO.Path]::DirectorySeparatorChar + $secBase) -replace '\\', '/')
    $sectionInfo[$t.SectionFolder] = [pscustomobject]@{
        Title = $t.SectionName; BaseName = $secBase; VaultPath = $secVault
        Children = (New-Object System.Collections.ArrayList)
    }
}

$p1 = 0
$n = $targets.Count
for ($idx = 0; $idx -lt $n; $idx++) {
    $t = $targets[$idx]
    $p1++
    $level = 1
    if ($t.PageLevel) { $level = [int]$t.PageLevel }

    # A page is a "parent" (folder note) if the NEXT page is nested deeper.
    $hasChildren = $false
    if ($idx + 1 -lt $n) {
        $nextLevel = 1
        if ($targets[$idx + 1].PageLevel) { $nextLevel = [int]$targets[$idx + 1].PageLevel }
        if ($nextLevel -gt $level) { $hasChildren = $true }
    }

    # Parent directory = the folder owned by the ancestor one level up, else the section root.
    $parentDir = $t.SectionFolder
    if ($level -gt 1 -and $stack.ContainsKey($level - 1)) { $parentDir = $stack[$level - 1] }
    if (-not $usedPerFolder.ContainsKey($parentDir)) {
        $usedPerFolder[$parentDir] = New-Object 'System.Collections.Generic.HashSet[string]'
    }

    $safeTitle = Get-SafeName $t.Title
    if ($hasChildren) {
        # Folder-note: create a folder, put the note of the same name inside it.
        $folderName = Get-UniqueName $safeTitle $usedPerFolder[$parentDir]
        $dirRel   = Join-Path $parentDir $folderName
        $baseName = $folderName
        $stack[$level] = $dirRel
        if (-not $usedPerFolder.ContainsKey($dirRel)) {
            $usedPerFolder[$dirRel] = New-Object 'System.Collections.Generic.HashSet[string]'
        }
        [void]$usedPerFolder[$dirRel].Add($baseName.ToLowerInvariant())   # reserve folder-note name
    }
    else {
        $baseName = Get-UniqueName $safeTitle $usedPerFolder[$parentDir]
        $dirRel   = $parentDir
    }

    # Drop any stale deeper levels now that we're at $level.
    foreach ($k in @($stack.Keys)) { if ($k -gt $level) { $stack.Remove($k) } }

    # Parent linkage for the Subpages index.
    $parentIdx = -1
    if ($level -gt 1 -and $lastAtLevel.ContainsKey($level - 1)) { $parentIdx = $lastAtLevel[$level - 1] }
    $lastAtLevel[$level] = $pagePlan.Count
    foreach ($k in @($lastAtLevel.Keys)) { if ($k -gt $level) { $lastAtLevel.Remove($k) } }

    $vaultPath = (($dirRel + [IO.Path]::DirectorySeparatorChar + $baseName) -replace '\\', '/')
    $slug = (Get-SafeName $t.Title 60) -replace '\s', '-'
    if ([string]::IsNullOrWhiteSpace($slug)) { $slug = "page" }

    # map internal-link GUID -> vault path
    try {
        $link = ""
        $on.GetHyperlinkToObject($t.Id, "", [ref]$link)
        $gm = [regex]::Match($link, 'page-id=\{([0-9A-Fa-f\-]+)\}')
        if ($gm.Success) { $script:LinkMap[$gm.Groups[1].Value.ToUpperInvariant()] = $vaultPath }
    } catch { }

    [void]$pagePlan.Add([pscustomobject]@{
        Id = $t.Id; Title = $t.Title; DirRel = $dirRel; BaseName = $baseName
        Level = $level; HasChildren = $hasChildren; VaultPath = $vaultPath
        ParentIdx = $parentIdx; Slug = "$slug-$p1"; SectionFolder = $t.SectionFolder
    })
    # Top-level pages belong to the section hub.
    if ($level -eq 1) { [void]$sectionInfo[$t.SectionFolder].Children.Add($pagePlan[$pagePlan.Count - 1]) }
    if ($p1 % 25 -eq 0) { Write-Host "  mapped $p1/$n" }
}

# Build parent -> children map for the Subpages index.
$childrenMap = @{}
for ($i = 0; $i -lt $pagePlan.Count; $i++) {
    $pi = $pagePlan[$i].ParentIdx
    if ($pi -ge 0) {
        if (-not $childrenMap.ContainsKey($pi)) { $childrenMap[$pi] = New-Object System.Collections.ArrayList }
        [void]$childrenMap[$pi].Add($pagePlan[$i])
    }
}
$parentCount = @($pagePlan | Where-Object { $_.HasChildren }).Count
Write-Host ("Pass 1 done: {0} pages, {1} parents (folder notes), {2} link targets." -f $pagePlan.Count, $parentCount, $script:LinkMap.Count) -ForegroundColor Green

# ---------- PASS 2: convert ----------
Write-Host "Pass 2/2 - converting pages..." -ForegroundColor Cyan
$script:ImagesSaved = 0; $script:AttachSaved = 0; $script:AttachMissing = 0
$script:LinksResolved = 0; $script:LinksUnresolved = 0
$errors = New-Object System.Collections.ArrayList
$done = 0

for ($pi = 0; $pi -lt $pagePlan.Count; $pi++) {
    $plan = $pagePlan[$pi]
    $done++
    try {
        $pc = ""
        $on.GetPageContent($plan.Id, [ref]$pc, 1)   # 1 = piBinaryData (embed images)
        [xml]$pdoc = $pc
        $pns = New-Ns $pdoc
        $pageNode = $pdoc.SelectSingleNode("//one:Page", $pns)

        # quick-style index -> name
        $styleMap = @{}
        foreach ($qs in $pdoc.SelectNodes("//one:QuickStyleDef", $pns)) { $styleMap[[string]$qs.index] = $qs.name }

        $script:ImgCounter = 0
        $pageMeta = [pscustomobject]@{ Slug = $plan.Slug }

        # Resolve parent (for backlink): nested page -> its parent page; top-level -> section hub.
        $parentPlan = $null
        if ($plan.ParentIdx -ge 0) { $parentPlan = $pagePlan[$plan.ParentIdx] }
        elseif ($plan.Level -eq 1 -and $sectionInfo.ContainsKey($plan.SectionFolder)) { $parentPlan = $sectionInfo[$plan.SectionFolder] }

        $sb = New-Object System.Text.StringBuilder
        $sb.AppendLine("# $($plan.Title)") | Out-Null
        $sb.AppendLine("") | Out-Null
        if ($parentPlan) {
            $sb.AppendLine("**Parent:** [[$($parentPlan.VaultPath)|$($parentPlan.Title)]]") | Out-Null
            $sb.AppendLine("") | Out-Null
        }

        $sbRef = [ref]$sb
        foreach ($outline in $pdoc.SelectNodes("//one:Outline", $pns)) {
            $oec = $outline.SelectSingleNode("one:OEChildren", $pns)
            if ($oec) { Walk-OEChildren $oec 0 $styleMap $sbRef $pageMeta $pns }
        }

        # Subpages index (folder notes only) - links to nested child pages.
        if ($plan.HasChildren -and $childrenMap.ContainsKey($pi)) {
            $sb.AppendLine("") | Out-Null
            $sb.AppendLine("## Subpages") | Out-Null
            $sb.AppendLine("") | Out-Null
            foreach ($child in $childrenMap[$pi]) {
                $sb.AppendLine("- [[$($child.VaultPath)|$($child.Title)]]") | Out-Null
            }
            $sb.AppendLine("") | Out-Null
        }

        $sb.AppendLine("") | Out-Null
        $sb.AppendLine("---") | Out-Null
        $sb.AppendLine("") | Out-Null
        $sb.AppendLine("## OneNote conversion metadata") | Out-Null
        $sb.AppendLine("") | Out-Null
        $sb.AppendLine('```yaml') | Out-Null
        $sb.AppendLine("title: " + ($plan.Title -replace '"', '\"')) | Out-Null
        $sb.AppendLine("created: $($pageNode.dateTime)") | Out-Null
        $sb.AppendLine("updated: $($pageNode.lastModifiedTime)") | Out-Null
        $sb.AppendLine("onenote-id: $($plan.Id)") | Out-Null
        if ($parentPlan) { $sb.AppendLine('parent: "[[' + $parentPlan.VaultPath + '|' + $parentPlan.Title + ']]"') | Out-Null }
        $sb.AppendLine("source: OneNote") | Out-Null
        $sb.AppendLine('```') | Out-Null

        $destDir = Join-Path $OutputRoot $plan.DirRel
        $null = New-Item -ItemType Directory -Force -Path $destDir
        $destFile = Join-Path $destDir ($plan.BaseName + ".md")
        [System.IO.File]::WriteAllText($destFile, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
    }
    catch {
        [void]$errors.Add([pscustomobject]@{ Title = $plan.Title; Error = $_.Exception.Message })
        Write-Host ("  ! error on '{0}': {1}" -f $plan.Title, $_.Exception.Message) -ForegroundColor Yellow
    }
    if ($done % 25 -eq 0) { Write-Host "  converted $done/$($pagePlan.Count)" }
}

# ---------- write section hub notes ----------
foreach ($secFolder in $sectionInfo.Keys) {
    $sec = $sectionInfo[$secFolder]
    $sb = New-Object System.Text.StringBuilder
    $sb.AppendLine("# $($sec.Title)") | Out-Null
    $sb.AppendLine("") | Out-Null
    $sb.AppendLine("## Pages") | Out-Null
    $sb.AppendLine("") | Out-Null
    foreach ($child in $sec.Children) {
        $sb.AppendLine("- [[$($child.VaultPath)|$($child.Title)]]") | Out-Null
    }
    $sb.AppendLine("") | Out-Null
    $sb.AppendLine("---") | Out-Null
    $sb.AppendLine("") | Out-Null
    $sb.AppendLine("## OneNote conversion metadata") | Out-Null
    $sb.AppendLine("") | Out-Null
    $sb.AppendLine('```yaml') | Out-Null
    $sb.AppendLine("title: " + ($sec.Title -replace '"', '\"')) | Out-Null
    $sb.AppendLine("source: OneNote") | Out-Null
    $sb.AppendLine("type: section") | Out-Null
    $sb.AppendLine('```') | Out-Null

    $destDir = Join-Path $OutputRoot $secFolder
    $null = New-Item -ItemType Directory -Force -Path $destDir
    $destFile = Join-Path $destDir ($sec.BaseName + ".md")
    [System.IO.File]::WriteAllText($destFile, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
}

$swAll.Stop()

# ---------- summary ----------
Write-Host ""
Write-Host "==================== SUMMARY ====================" -ForegroundColor Cyan
Write-Host ("Pages converted   : {0} / {1}" -f ($pagePlan.Count - $errors.Count), $pagePlan.Count)
Write-Host ("Folder notes      : {0}" -f $parentCount)
Write-Host ("Section hubs      : {0}" -f $sectionInfo.Count)
Write-Host ("Images extracted  : {0}" -f $script:ImagesSaved)
Write-Host ("Attachments saved : {0}  (missing cache: {1})" -f $script:AttachSaved, $script:AttachMissing)
Write-Host ("Internal links    : {0} resolved, {1} unresolved" -f $script:LinksResolved, $script:LinksUnresolved)
Write-Host ("Errors            : {0}" -f $errors.Count)
Write-Host ("Elapsed           : {0:n1} s" -f $swAll.Elapsed.TotalSeconds)
Write-Host ("Output            : {0}" -f (Resolve-Path $OutputRoot))
if ($errors.Count -gt 0) {
    $errLog = Join-Path $OutputRoot "_export-errors.txt"
    $errors | ForEach-Object { "$($_.Title)`t$($_.Error)" } | Set-Content $errLog -Encoding UTF8
    Write-Host ("Error log         : {0}" -f $errLog) -ForegroundColor Yellow
}
Write-Host "=================================================" -ForegroundColor Cyan
