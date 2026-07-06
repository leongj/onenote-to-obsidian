<#
.SYNOPSIS
    Read a single OneNote page as Markdown/plain text (live, via COM).
.DESCRIPTION
    Fetches a page's XML with GetPageContent and converts the text/heading/list/table
    structure to readable Markdown. Intended for quickly surfacing note content to an
    agent or user WITHOUT a full migration. Images are noted as placeholders (use the
    onenote-migrate-obsidian skill if you need the actual image files). MUST run under
    Windows PowerShell 5.1.
.PARAMETER PageId
    The OneNote page ID (get one from Search-OneNote.ps1 or List-OneNoteHierarchy.ps1).
.PARAMETER Title
    Alternative to -PageId: search for a page by title and read the first match.
.EXAMPLE
    powershell.exe -File Get-OneNotePage.ps1 -Title "project kickoff"
    powershell.exe -File Get-OneNotePage.ps1 -PageId "{GUID}{1}{...}"
#>
[CmdletBinding(DefaultParameterSetName = "ById")]
param(
    [Parameter(Mandatory = $true, ParameterSetName = "ById")]  [string] $PageId,
    [Parameter(Mandatory = $true, ParameterSetName = "ByTitle")] [string] $Title
)

$ErrorActionPreference = "Stop"
$oneNS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

function New-Ns([xml]$doc) {
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("one", $oneNS)
    return ,$ns
}

# Strip inline HTML tags OneNote embeds inside one:T CDATA, keep the text.
function Convert-Inline([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return "" }
    $s = $s -replace '<br[^>]*>', "`n"
    $s = $s -replace '<[^>]+>', ''
    $s = $s -replace '&nbsp;', ' '
    $s = $s -replace '&lt;', '<' -replace '&gt;', '>' -replace '&amp;', '&' -replace '&quot;', '"' -replace '&#39;', "'"
    return $s.Trim()
}

$app = New-Object -ComObject OneNote.Application

if ($PSCmdlet.ParameterSetName -eq "ByTitle") {
    $hout = ""
    $app.GetHierarchy("", 4, [ref]$hout)  # 4 = hsPages
    $hres = [xml]$hout
    $hns = New-Ns $hres
    $first = $hres.SelectNodes("//one:Page", $hns) | Where-Object {
        $_.name -and $_.name.ToLower().Contains($Title.ToLower()) -and $_.isInRecycleBin -ne "true"
    } | Select-Object -First 1
    if (-not $first) { Write-Host "No page titled like '$Title'." -ForegroundColor Yellow; return }
    $PageId = $first.ID
    Write-Host ("# Reading: {0}" -f $first.name) -ForegroundColor Cyan
}

$out = ""
$app.GetPageContent($PageId, [ref]$out, 0)  # 0 = no binary; we only want text here
$page = [xml]$out
$ns = New-Ns $page

# Quick-style index -> markdown heading prefix
$styleMap = @{}
foreach ($qs in $page.SelectNodes("//one:QuickStyleDef", $ns)) {
    $name = $qs.name
    $prefix = switch -Regex ($name) {
        '^h1$' { "# " }  '^h2$' { "## " }  '^h3$' { "### " }
        '^h4$' { "#### " } '^h5$' { "##### " } '^h6$' { "###### " }
        default { "" }
    }
    $styleMap[$qs.index] = $prefix
}

$sb = New-Object System.Text.StringBuilder
$title = $page.SelectSingleNode("//one:Title//one:T", $ns)
if ($title) {
    $tt = Convert-Inline $title.'#cdata-section'
    if ($tt) { [void]$sb.AppendLine("# " + $tt); [void]$sb.AppendLine() }
}

function Walk($node) {
    foreach ($child in $node.ChildNodes) {
        switch ($child.LocalName) {
            "OE" {
                $t = $child.SelectSingleNode("one:T", $ns)
                if ($t) {
                    $text = Convert-Inline $t.'#cdata-section'
                    $prefix = ""
                    if ($child.quickStyleIndex -ne $null -and $styleMap.ContainsKey($child.quickStyleIndex)) {
                        $prefix = $styleMap[$child.quickStyleIndex]
                    }
                    $bullet = ""
                    if ($child.SelectSingleNode("one:List/one:Bullet", $ns)) { $bullet = "- " }
                    elseif ($child.SelectSingleNode("one:List/one:Number", $ns)) { $bullet = "1. " }
                    if ($text) { [void]$sb.AppendLine($prefix + $bullet + $text) }
                }
                if ($child.SelectSingleNode("one:Image", $ns)) { [void]$sb.AppendLine("*(image)*") }
                if ($child.SelectSingleNode("one:Table", $ns)) {
                    foreach ($row in $child.SelectNodes("one:Table/one:Row", $ns)) {
                        $cells = @()
                        foreach ($cell in $row.SelectNodes("one:Cell", $ns)) {
                            $ct = $cell.SelectNodes(".//one:T", $ns) | ForEach-Object { Convert-Inline $_.'#cdata-section' }
                            $cells += (($ct -join " ") -replace '\|', '\|')
                        }
                        [void]$sb.AppendLine("| " + ($cells -join " | ") + " |")
                    }
                }
                Walk $child
            }
            "OEChildren" { Walk $child }
            default { if ($child.HasChildNodes) { Walk $child } }
        }
    }
}

foreach ($outline in $page.SelectNodes("//one:Outline/one:OEChildren", $ns)) { Walk $outline }
Write-Output $sb.ToString()
