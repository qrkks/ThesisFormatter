$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$macro = Get-Content -Raw -Path (Join-Path $repoRoot "format_macro.bas")

function Get-SubBody {
    param([string]$Name)

    $match = [regex]::Match(
        $macro,
        "(?:Public |Private )?Sub\s+$Name\b.*?\r?\nEnd Sub",
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    if (-not $match.Success) { throw "Could not find Sub $Name." }
    $match.Value
}

function Get-FunctionBody {
    param([string]$Name)

    $match = [regex]::Match(
        $macro,
        "(?:Public |Private )?Function\s+$Name\b.*?\r?\nEnd Function",
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    if (-not $match.Success) { throw "Could not find Function $Name." }
    $match.Value
}

$pipeline = Get-SubBody "RunSDUTCMFormatting"
$abstractFormatter = Get-SubBody "MergeAndFormatAbstract"
$abstractFinder = Get-FunctionBody "FindAbstractLabelParagraphStart"

if ($pipeline -match "For\s+i\s*=|FormatTitleParagraph|FormatLevel[123]Paragraph|FormatBodyParagraph") {
    throw "The aggressive default pipeline must not directly format every paragraph."
}
if ($pipeline -match "ProcessTables") {
    throw "Table formatting should be opt-in in the aggressive default pipeline."
}
if ($pipeline -notmatch "ConfigureSDUTCMStyles") {
    throw "The aggressive default pipeline must configure styles."
}
if ($abstractFormatter -match "ActiveDocument\.Paragraphs\s*\(") {
    throw "Abstract formatting must not scan paragraphs by index."
}
if ($abstractFormatter -notmatch "FindAbstractLabelParagraphStart") {
    throw "Abstract formatting should target known labels."
}
if ($abstractFinder -notmatch "\.Find") {
    throw "Abstract label lookup should use Range.Find."
}
if ($abstractFinder -notmatch "nextStart\s*=\s*searchRange\.End\s*\+\s*1") {
    throw "Abstract label lookup must advance past rejected matches."
}

Write-Host "Aggressive performance regression checks passed."
