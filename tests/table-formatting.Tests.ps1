$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$macroPath = Join-Path $repoRoot "format_macro.bas"
$macro = Get-Content -Raw -Path $macroPath
$tableFormatterMatch = [regex]::Match(
    $macro,
    "Private Sub ApplyThreeLineTableStyle\(ByVal tbl As Table\).*?\r?\nEnd Sub",
    [System.Text.RegularExpressions.RegexOptions]::Singleline
)

if (-not $tableFormatterMatch.Success) {
    throw "Could not find ApplyThreeLineTableStyle in format_macro.bas."
}

$tableFormatter = $tableFormatterMatch.Value

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

Assert-Contains `
    -Text $tableFormatter `
    -Pattern "\.AllowAutoFit\s*=\s*True" `
    -Message "Table formatting should allow AutoFit so hidden/narrow columns do not squeeze content."

Assert-Contains `
    -Text $tableFormatter `
    -Pattern "\.AutoFitBehavior\s+wdAutoFitContent" `
    -Message "Table formatting should fit columns to content before centering the table."

Assert-Contains `
    -Text $tableFormatter `
    -Pattern "\.FirstLineIndent\s*=\s*0" `
    -Message "Table cell paragraphs should not inherit first-line indentation."

Assert-Contains `
    -Text $tableFormatter `
    -Pattern "\.LeftIndent\s*=\s*0" `
    -Message "Table cell paragraphs should not inherit left indentation."

Assert-Contains `
    -Text $tableFormatter `
    -Pattern "\.RightIndent\s*=\s*0" `
    -Message "Table cell paragraphs should not inherit right indentation."

Write-Host "Table formatting regression checks passed."
