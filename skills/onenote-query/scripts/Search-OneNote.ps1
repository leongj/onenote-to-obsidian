<#
.SYNOPSIS
    Search OneNote pages by title or full-text content (live, via COM).
.DESCRIPTION
    Two search modes:
      * Default (title scan): walks the notebook hierarchy and matches page TITLES against
        the query (case-insensitive substring). Deterministic and reliable — always finds a
        page you can name, including recently-created/unindexed ones.
      * -Content: uses OneNote's native FindPages full-text index to find pages whose BODY
        matches. Fast, but only covers pages OneNote has indexed and uses AND/word semantics.
    Prints each hit with its section breadcrumb and page ID. MUST run under Windows PowerShell 5.1.
.PARAMETER Query
    The text to look for.
.PARAMETER Content
    Search page body text (FindPages index) instead of titles.
.PARAMETER Max
    Max results to print (default 50).
.EXAMPLE
    powershell.exe -File Search-OneNote.ps1 -Query "project kickoff"
    powershell.exe -File Search-OneNote.ps1 -Query "budget forecast" -Content
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $Query,
    [switch] $Content,
    [int] $Max = 50
)

$ErrorActionPreference = "Stop"
$oneNS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

function New-Ns([xml]$doc) {
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("one", $oneNS)
    return ,$ns
}

function Get-Breadcrumb($pageNode) {
    $crumbs = @()
    $node = $pageNode.ParentNode
    while ($node -ne $null -and $node.NodeType -eq "Element") {
        if ($node.Attributes -and $node.Attributes["name"]) { $crumbs = ,($node.name) + $crumbs }
        $node = $node.ParentNode
    }
    return ($crumbs -join " / ")
}

function Show-Hits($pages, $label) {
    if (-not $pages -or $pages.Count -eq 0) {
        Write-Host "No pages matched '$Query' ($label)." -ForegroundColor Yellow
        return
    }
    Write-Host ("{0} match(es) for '{1}' ({2}):" -f $pages.Count, $Query, $label) -ForegroundColor Cyan
    $shown = 0
    foreach ($pg in $pages) {
        if ($shown -ge $Max) { Write-Host ("  ... {0} more" -f ($pages.Count - $Max)); break }
        Write-Host ("  {0}" -f $pg.name) -ForegroundColor White
        Write-Host ("     in: {0}" -f (Get-Breadcrumb $pg)) -ForegroundColor DarkGray
        Write-Host ("     id: {0}" -f $pg.ID) -ForegroundColor DarkGray
        $shown++
    }
    Write-Host ""
    Write-Host "Tip: read one with  Get-OneNotePage.ps1 -PageId '<id>'" -ForegroundColor DarkGray
}

$app = New-Object -ComObject OneNote.Application

if ($Content) {
    # Full-text search via OneNote's index (indexed pages only, AND word semantics).
    $out = ""
    $app.FindPages("", $Query, [ref]$out, $false, $false, 2)  # 2 = xs2013
    $res = [xml]$out
    $ns = New-Ns $res
    Show-Hits ($res.SelectNodes("//one:Page", $ns)) "content"
}
else {
    # Title scan across the full hierarchy — reliable, index-independent.
    $out = ""
    $app.GetHierarchy("", 4, [ref]$out)  # 4 = hsPages
    $hier = [xml]$out
    $ns = New-Ns $hier
    $all = $hier.SelectNodes("//one:Page", $ns)
    $matches = @($all | Where-Object {
        $_.name -and $_.name.ToLower().Contains($Query.ToLower()) `
        -and $_.isInRecycleBin -ne "true"
    })
    Show-Hits $matches "title"
}
