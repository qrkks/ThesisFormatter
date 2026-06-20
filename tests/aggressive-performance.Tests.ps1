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
$pageFormatter = Get-SubBody "SetPageAndBodyFormat"
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
if ($pageFormatter -match "For\s+Each\s+para") {
    throw "Default line spacing must not be applied paragraph by paragraph."
}
if ($pageFormatter -notmatch "ActiveDocument\.Content\.ParagraphFormat\.LineSpacingRule") {
    throw "Default line spacing should be applied to the document range once."
}
if ($abstractFormatter -match "ActiveDocument\.Paragraphs\s*\(") {
    throw "Abstract formatting must not scan paragraphs by index."
}
if ($abstractFormatter -notmatch "FindAbstractLabelParagraphStart") {
    throw "Abstract formatting should target known labels."
}
if ($abstractFormatter -notmatch "ActiveDocument\.Content\.Text") {
    throw "Abstract formatting should read the document text once."
}
if ($abstractFinder -notmatch "InStr") {
    throw "Abstract label lookup should search the in-memory document text."
}
if ($abstractFinder -match "\.Find") {
    throw "Abstract label lookup must not use Word Range.Find."
}

Write-Host "Aggressive performance regression checks passed."
