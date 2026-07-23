[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

$wdStartupPath = 8
$wdFormatXMLTemplateMacroEnabled = 15
$addinFileName = "ThesisFormatter.dotm"
$moduleName = "ThesisFormatter"
$repoRoot = $PSScriptRoot
$macroPath = Join-Path $repoRoot "format_macro.bas"
$backupRoot = Join-Path $env:LOCALAPPDATA "ThesisFormatter\backups"
$word = $null
$createdWord = $false
$templateDocument = $null
$newComponent = $null
$installedAddIn = $null
$tempAddInPath = Join-Path $env:TEMP ("ThesisFormatter-" + [guid]::NewGuid() + ".dotm")

function Release-ComObject {
    param([object]$Value)

    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Get-RunningOrNewWord {
    try {
        $application = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
        return @($application, $false)
    }
    catch {
        $application = New-Object -ComObject Word.Application
        $application.Visible = $false
        return @($application, $true)
    }
}

function Get-AddInByPath {
    param(
        [object]$Application,
        [string]$FullPath
    )

    foreach ($addIn in $Application.AddIns) {
        $isMatch = $false
        try {
            $isMatch = [string]::Equals(
                [IO.Path]::GetFullPath($addIn.Path + "\" + $addIn.Name),
                $FullPath,
                [StringComparison]::OrdinalIgnoreCase
            )
            if ($isMatch) {
                return $addIn
            }
        }
        finally {
            if ($null -ne $addIn -and -not $isMatch) {
                Release-ComObject $addIn
            }
        }
    }

    return $null
}

function Backup-File {
    param(
        [string]$Path,
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    [void](New-Item -ItemType Directory -Path $backupRoot -Force)
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $extension = [IO.Path]::GetExtension($Path)
    $backupPath = Join-Path $backupRoot ("${Label}-${timestamp}${extension}")
    Copy-Item -LiteralPath $Path -Destination $backupPath
    return $backupPath
}

function Migrate-ThesisFormatterFromNormalTemplate {
    param([object]$Application)

    $normalTemplate = $Application.NormalTemplate
    $project = $normalTemplate.VBProject
    $componentsToRemove = @()

    foreach ($component in $project.VBComponents) {
        try {
            if ($component.Type -ne 1) {
                continue
            }

            $lineCount = $component.CodeModule.CountOfLines
            if ($lineCount -eq 0) {
                continue
            }

            $code = $component.CodeModule.Lines(1, $lineCount)
            if ($code.Contains("FormatThesisToSDUTCM") -and
                $code.Contains("ApplySingleSpacingToTables")) {
                $componentsToRemove += $component
                continue
            }
        }
        finally {
            if ($componentsToRemove -notcontains $component) {
                Release-ComObject $component
            }
        }
    }

    foreach ($component in $componentsToRemove) {
        [void](New-Item -ItemType Directory -Path $backupRoot -Force)
        $safeName = $component.Name -replace '[^\p{L}\p{Nd}_.-]', '_'
        $backupPath = Join-Path $backupRoot ("Normal-${safeName}-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".bas")
        $component.Export($backupPath)
        $project.VBComponents.Remove($component)
        Write-Host "Migrated old Normal.dotm module; backup: $backupPath"
        Release-ComObject $component
    }

    if ($componentsToRemove.Count -gt 0) {
        $normalTemplate.Save()
    }

    Release-ComObject $project
    Release-ComObject $normalTemplate
}

if (-not $Uninstall -and -not (Test-Path -LiteralPath $macroPath)) {
    throw "Could not find format_macro.bas next to install.ps1."
}

try {
    $wordResult = Get-RunningOrNewWord
    $word = $wordResult[0]
    $createdWord = [bool]$wordResult[1]
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 1

    $startupPath = [IO.Path]::GetFullPath($word.Options.DefaultFilePath(8))
    $destinationPath = [IO.Path]::GetFullPath((Join-Path $startupPath $addinFileName))
    if (-not [string]::Equals(
            [IO.Path]::GetDirectoryName($destinationPath),
            $startupPath.TrimEnd('\'),
            [StringComparison]::OrdinalIgnoreCase
        )) {
        throw "Refusing to install outside Word's STARTUP directory."
    }

    $installedAddIn = Get-AddInByPath -Application $word -FullPath $destinationPath

    if ($Uninstall) {
        if (-not $PSCmdlet.ShouldProcess($destinationPath, "Uninstall the global ThesisFormatter Word add-in")) {
            return
        }

        if ($null -ne $installedAddIn) {
            $installedAddIn.Installed = $false
        }
        if (Test-Path -LiteralPath $destinationPath) {
            $backupPath = Backup-File -Path $destinationPath -Label "ThesisFormatter-before-uninstall"
            Remove-Item -LiteralPath $destinationPath -Force
            Write-Host "Backup: $backupPath"
        }
        Write-Host "ThesisFormatter global Word add-in uninstalled."
        return
    }

    if (-not $PSCmdlet.ShouldProcess($destinationPath, "Install or update the global ThesisFormatter Word add-in")) {
        return
    }

    if ($null -ne $installedAddIn) {
        $installedAddIn.Installed = $false
    }

    $existingBackup = Backup-File -Path $destinationPath -Label "ThesisFormatter-before-update"

    try {
        $templateDocument = $word.Documents.Add()
        $newComponent = $templateDocument.VBProject.VBComponents.Add(1)
        $newComponent.Name = $moduleName
        $newComponent.CodeModule.AddFromString(
            (Get-Content -Raw -Encoding UTF8 -LiteralPath $macroPath)
        )
    }
    catch {
        throw "Word blocked programmatic VBA access. In Word, enable Trust Center > Macro Settings > Trust access to the VBA project object model, then rerun install.ps1. Details: $($_.Exception.Message)"
    }

    $templateDocument.SaveAs2($tempAddInPath, $wdFormatXMLTemplateMacroEnabled)
    $templateDocument.Close(0)
    Release-ComObject $templateDocument
    $templateDocument = $null

    [void](New-Item -ItemType Directory -Path $startupPath -Force)
    Copy-Item -LiteralPath $tempAddInPath -Destination $destinationPath -Force
    if ($null -eq $installedAddIn) {
        $installedAddIn = $word.AddIns.Add($destinationPath, $true)
    }
    $installedAddIn.Installed = $true

    Migrate-ThesisFormatterFromNormalTemplate -Application $word

    if (-not $installedAddIn.Installed -or -not (Test-Path -LiteralPath $destinationPath)) {
        throw "Word did not load the installed ThesisFormatter add-in."
    }

    Write-Host "ThesisFormatter global Word add-in installed."
    Write-Host "Path: $destinationPath"
    if ($existingBackup) {
        Write-Host "Previous add-in backup: $existingBackup"
    }
    Write-Host "Use Alt+F8 in any document and run FormatThesisToSDUTCM."
}
finally {
    if ($templateDocument) {
        $templateDocument.Close(0)
    }
    if ($createdWord -and $word) {
        $word.Quit()
    }
    Release-ComObject $newComponent
    Release-ComObject $templateDocument
    Release-ComObject $installedAddIn
    Release-ComObject $word
    if (Test-Path -LiteralPath $tempAddInPath) {
        Remove-Item -LiteralPath $tempAddInPath -Force
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
