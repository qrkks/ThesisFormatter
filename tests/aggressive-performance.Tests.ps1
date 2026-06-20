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
$styleConfiguration = Get-SubBody "ConfigureSDUTCMStyles"
$titleStyleConfiguration = Get-SubBody "ConfigureTitleStyleIfExists"
$headingStyleConfiguration = Get-SubBody "ConfigureHeadingStyleIfExists"
$headingIndentCleanup = Get-SubBody "ClearHeadingParagraphIndents"

foreach ($styleName in @(
    'ZhBodyTextStyleName\(\)',
    'ZhBodyStyleName\(\)',
    '"Normal"',
    '"First Paragraph"'
)) {
    if ($styleConfiguration -notmatch "ConfigureBodyStyleIfExists\s+$styleName,\s*24") {
        throw "Body style $styleName should use a 24 pt first-line indent."
    }
}

foreach ($styleBody in @($titleStyleConfiguration, $headingStyleConfiguration)) {
    foreach ($property in @("FirstLineIndent", "LeftIndent", "RightIndent")) {
        if ($styleBody -notmatch "\.$property\s*=\s*0") {
            throw "Title and heading styles should explicitly set $property to zero."
        }
    }
}

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
if ($pipeline -notmatch "ConfigureSDUTCMStyles[\s\S]*ClearHeadingParagraphIndents") {
    throw "The default pipeline should clear direct heading indents after configuring styles."
}
if ($headingIndentCleanup -notmatch "For\s+Each\s+para\s+In\s+ActiveDocument\.Paragraphs") {
    throw "Heading indent cleanup should enumerate document paragraphs once."
}
if ($headingIndentCleanup -notmatch "para\.Style\s*=\s*ZhTitleStyleName\(\)") {
    throw "Heading indent cleanup should include the document title style."
}
if ($headingIndentCleanup -notmatch "wdOutlineLevel1" -or
    $headingIndentCleanup -notmatch "wdOutlineLevel3") {
    throw "Heading indent cleanup should include outline levels 1 through 3."
}
foreach ($property in @("FirstLineIndent", "LeftIndent", "RightIndent")) {
    if ($headingIndentCleanup -notmatch "\.$property\s*=\s*0") {
        throw "Heading indent cleanup should set $property to zero."
    }
}
if ($headingIndentCleanup -match "\.Font\.|\.Alignment\s*=|\.LineSpacing") {
    throw "Heading indent cleanup should not alter non-indent formatting."
}

Write-Host "Aggressive performance regression checks passed."
