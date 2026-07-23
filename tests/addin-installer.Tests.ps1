$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$installerPath = Join-Path $repoRoot "install.ps1"

if (-not (Test-Path -LiteralPath $installerPath)) {
    throw "Could not find install.ps1."
}

$installer = Get-Content -Raw -LiteralPath $installerPath

function Assert-Contains {
    param(
        [string]$Pattern,
        [string]$Message
    )

    if ($installer -notmatch $Pattern) {
        throw $Message
    }
}

Assert-Contains `
    -Pattern "DefaultFilePath\(8\)" `
    -Message "The installer should ask Word for its global STARTUP directory."

Assert-Contains `
    -Pattern 'ThesisFormatter\.dotm' `
    -Message "The installer should create a dedicated global Word add-in."

Assert-Contains `
    -Pattern 'wdFormatXMLTemplateMacroEnabled\s*=\s*15' `
    -Message "The installer should save a macro-enabled Word template."

Assert-Contains `
    -Pattern 'VBComponents\.Add\(1\)' `
    -Message "The installer should create a standard VBA module."

Assert-Contains `
    -Pattern 'CodeModule\.AddFromString' `
    -Message "The installer should inject format_macro.bas into the add-in."

Assert-Contains `
    -Pattern 'AddIns\.Add\(' `
    -Message "The installer should load the add-in immediately."

Assert-Contains `
    -Pattern '\[switch\]\$Uninstall' `
    -Message "The installer should support command-line uninstall."

Assert-Contains `
    -Pattern 'FormatThesisToSDUTCM[\s\S]*ApplySingleSpacingToTables' `
    -Message "Normal.dotm migration should identify only a known ThesisFormatter module."

if ($installer -match 'VBComponents\s*\|\s*ForEach-Object\s*\{\s*\$[^\r\n]+\.Remove') {
    throw "The installer must not remove every VBA component from Normal.dotm."
}

Write-Host "Global Word add-in installer regression checks passed."
