<#
.SYNOPSIS
    List the OneNote notebook / section / page tree (live, via COM).
.DESCRIPTION
    Connects to the desktop OneNote application and prints the hierarchy so an agent or
    user can see what notebooks, sections and pages exist and grab page IDs for other
    scripts. MUST run under Windows PowerShell 5.1 (COM fails under PowerShell 7).
.PARAMETER NotebookFilter
    Optional wildcard to limit to matching notebooks (default '*').
.PARAMETER SectionFilter
    Optional wildcard to limit to matching sections (default '*').
.PARAMETER IncludePages
    Include page titles under each section (indented by pageLevel).
.EXAMPLE
    powershell.exe -File List-OneNoteHierarchy.ps1 -SectionFilter Work -IncludePages
#>
[CmdletBinding()]
param(
    [string] $NotebookFilter = "*",
    [string] $SectionFilter  = "*",
    [switch] $IncludePages
)

$ErrorActionPreference = "Stop"
$oneNS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

function New-Ns([xml]$doc) {
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("one", $oneNS)
    return ,$ns
}

$app = New-Object -ComObject OneNote.Application
# GetHierarchy(bstrStartNodeID, hsScope, [ref]xmlOut); hsScope 4 = hsPages (full tree)
$out = ""
$app.GetHierarchy("", 4, [ref]$out)
$hier = [xml]$out
$ns = New-Ns $hier

foreach ($nb in $hier.SelectNodes("//one:Notebook", $ns)) {
    if ($nb.name -notlike $NotebookFilter) { continue }
    Write-Host ("[NOTEBOOK] {0}" -f $nb.name) -ForegroundColor Cyan
    foreach ($sec in $nb.SelectNodes(".//one:Section", $ns)) {
        if ($sec.name -notlike $SectionFilter) { continue }
        if ($sec.isInRecycleBin -eq "true" -or $sec.isRecycleBin -eq "true") { continue }
        $pageCount = $sec.SelectNodes("one:Page", $ns).Count
        Write-Host ("  [SECTION] {0}  ({1} pages)" -f $sec.name, $pageCount) -ForegroundColor Yellow
        if ($IncludePages) {
            foreach ($pg in $sec.SelectNodes("one:Page", $ns)) {
                $lvl = [int]($pg.pageLevel)
                if ($lvl -lt 1) { $lvl = 1 }
                $indent = "    " + ("  " * ($lvl - 1))
                Write-Host ("{0}- {1}" -f $indent, $pg.name)
                Write-Host ("{0}  id: {1}" -f $indent, $pg.ID) -ForegroundColor DarkGray
            }
        }
    }
}
