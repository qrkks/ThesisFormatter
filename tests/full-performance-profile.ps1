param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,
    [double]$MaxSeconds = 20,
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$macroPath = Join-Path $repoRoot "format_macro.bas"
$resolvedDocument = (Resolve-Path -LiteralPath $DocumentPath).Path
$tempPath = Join-Path $env:TEMP ("full-profile-" + [guid]::NewGuid() + ".docx")

Copy-Item -LiteralPath $resolvedDocument -Destination $tempPath

$word = $null
$document = $null
$component = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.ScreenUpdating = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 1

    $document = $word.Documents.Open($tempPath, $false, $false, $false)
    $component = $document.VBProject.VBComponents.Add(1)
    $component.Name = "FullProfiler"
    $component.CodeModule.AddFromString(
        (Get-Content -Raw -Encoding UTF8 -LiteralPath $macroPath)
    )
    $component.CodeModule.AddFromString(@'
Public Function ProfileFullFormatting() As Double
    Dim started As Single

    Application.ScreenUpdating = False
    started = Timer
    RunSDUTCMFormatting
    ProfileFullFormatting = Timer - started
End Function
'@)

    $document.Activate()
    $seconds = [double]$word.Run("ProfileFullFormatting")
    Write-Host ("Full formatting: {0:N3} s" -f $seconds)

    if ($seconds -gt $MaxSeconds) {
        throw "Full formatting exceeded ${MaxSeconds}s (actual: $seconds s)."
    }

    if ($OutputPath) {
        $resolvedOutput = [IO.Path]::GetFullPath($OutputPath)
        if (Test-Path -LiteralPath $resolvedOutput) {
            Remove-Item -LiteralPath $resolvedOutput -Force
        }
        $document.SaveAs2($resolvedOutput, 16)
        Write-Host "Output: $resolvedOutput"
    }
}
finally {
    if ($document) { $document.Close(0) }
    if ($word) { $word.Quit() }
    if ($component) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($component) }
    if ($document) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
    if ($word) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
