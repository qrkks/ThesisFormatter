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

Assert-Contains `
    -Text $fontHelper `
    -Pattern "wasItalic\s*=\s*charRange\.Font\.Italic" `
    -Message "The font helper should snapshot each character's italic state before applying base font formatting."

Assert-Contains `
    -Text $fontHelper `
    -Pattern "wasColor\s*=\s*charRange\.Font\.Color" `
    -Message "The font helper should snapshot each character's color before applying base font formatting."

Assert-Contains `
    -Text $fontHelper `
    -Pattern "\.Italic\s*=\s*wasItalic" `
    -Message "The font helper should restore each character's italic state after applying base font formatting."

Assert-Contains `
    -Text $fontHelper `
    -Pattern "\.Color\s*=\s*wasColor" `
    -Message "The font helper should restore each character's color after applying base font formatting."

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
