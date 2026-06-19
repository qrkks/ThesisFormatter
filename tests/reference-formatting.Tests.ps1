$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$macroPath = Join-Path $repoRoot "format_macro.bas"
$macro = Get-Content -Raw -Path $macroPath

function Get-SubBody {
    param(
        [string]$Name
    )

    $match = [regex]::Match(
        $macro,
        "Sub\s+$Name\b.*?\r?\nEnd Sub",
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    if (-not $match.Success) {
        throw "Could not find Sub $Name in format_macro.bas."
    }

    return $match.Value
}

function Assert-Contains {
    param(
        [string]$Text,
        [string]$Pattern,
        [string]$Message
    )

    if ($Text -notmatch $Pattern) {
        throw $Message
    }
}

function Assert-NotContains {
    param(
        [string]$Text,
        [string]$Pattern,
        [string]$Message
    )

    if ($Text -match $Pattern) {
        throw $Message
    }
}

$fontHelper = Get-SubBody "ApplyRangeFontPreservingItalic"
$bodyFormatter = Get-SubBody "FormatBodyParagraph"
$referenceFormatter = Get-SubBody "FormatReferenceEntries"
$processReferences = Get-SubBody "ProcessReferences"
$processReferencesWithSort = Get-SubBody "ProcessReferencesWithSort"
$optimizedFormatter = Get-SubBody "FormatReferenceSection"
$rangeFormatter = Get-SubBody "ApplyReferenceEntriesFormat"

foreach ($process in @($processReferences, $processReferencesWithSort)) {
    Assert-Contains `
        -Text $process `
        -Pattern "FormatReferenceSection" `
        -Message "Complete reference processing should use the bounded reference-section formatter."

    Assert-NotContains `
        -Text $process `
        -Pattern "\b(?:NormalizeReferenceHeadingParagraphs|FormatReferences|FormatReferenceEntries)\b" `
        -Message "Complete reference processing should not chain legacy full-document passes."
}

Assert-Contains `
    -Text $optimizedFormatter `
    -Pattern "For\s+Each\s+para\s+In\s+searchRange\.Paragraphs" `
    -Message "Reference boundary detection should enumerate only the range after the heading."

Assert-NotContains `
    -Text $optimizedFormatter `
    -Pattern "Paragraphs\s*\(" `
    -Message "Optimized reference formatting should not use indexed paragraph access."

Assert-Contains `
    -Text $rangeFormatter `
    -Pattern "With\s+entriesRange\.ParagraphFormat" `
    -Message "Reference paragraph formatting should be applied to one range."

Assert-Contains `
    -Text $rangeFormatter `
    -Pattern "With\s+entriesRange\.Font" `
    -Message "Reference font formatting should be applied to one range."

Assert-NotContains `
    -Text $rangeFormatter `
    -Pattern "\.Italic\s*=|\.Color\s*=" `
    -Message "Reference range formatting should preserve italic and color by leaving them untouched."

Assert-NotContains `
    -Text $fontHelper `
    -Pattern "\.Characters\b|For\s+Each\s+charRange" `
    -Message "Font formatting must operate on the whole range, not character by character."

Assert-Contains `
    -Text $fontHelper `
    -Pattern "With\s+sourceRange\.Font" `
    -Message "Font formatting should update the supplied range in one operation."

foreach ($property in @("NameFarEast", "Name", "Size", "Bold")) {
    Assert-Contains `
        -Text $fontHelper `
        -Pattern "\.$property\s*=" `
        -Message "Range font formatting should set $property."
}

Assert-NotContains `
    -Text $fontHelper `
    -Pattern "\.Italic\s*=" `
    -Message "Italic formatting should be preserved by leaving it untouched."

Assert-NotContains `
    -Text $fontHelper `
    -Pattern "\.Color\s*=" `
    -Message "Font color should be preserved by leaving it untouched."

Assert-Contains `
    -Text $bodyFormatter `
    -Pattern "ApplyRangeFontPreservingItalic\s+para\.Range" `
    -Message "Body paragraph formatting should preserve italic text because references are formatted as body text before reference processing."

Assert-Contains `
    -Text $referenceFormatter `
    -Pattern "ApplyRangeFontPreservingItalic\s+para\.Range" `
    -Message "Reference entry formatting should preserve existing italic spans."

Assert-NotContains `
    -Text $referenceFormatter `
    -Pattern "With\s+para\.Range\.Font" `
    -Message "Reference entry formatting should not apply font formatting to the whole paragraph range at once."

Assert-Contains `
    -Text $referenceFormatter `
    -Pattern 'Author Bio"\)\s+Then[ \t]*\r?\n[ \t]*EnsurePageBreakBeforeParagraph\s+para[ \t]*\r?\n[ \t]*foundReferences\s*=\s*False' `
    -Message "Recognized content after the reference list should start on a new page even when it is not styled as a heading."

Write-Host "Reference formatting regression checks passed."
