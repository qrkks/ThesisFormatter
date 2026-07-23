$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$macroPath = Join-Path $repoRoot "format_macro.bas"
$tempRoot = Join-Path $env:TEMP ("thesis-table-preservation-" + [guid]::NewGuid())
$sourcePath = Join-Path $tempRoot "source.docx"
$outputPath = Join-Path $tempRoot "output.docx"
$word = $null
$document = $null
$component = $null

[void](New-Item -ItemType Directory -Path $tempRoot)

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.ScreenUpdating = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 1

    $document = $word.Documents.Add()
    $tocLabel = [string][char]0x76EE + [char]0x5F55
    $document.Content.Text = "Thesis Title`r${tocLabel}`rChapter One`rBody paragraph`r"
    $document.Paragraphs.Item(1).Style = -63
    $document.Paragraphs.Item(3).Style = -2

    $tableRange = $document.Range($document.Content.End - 1, $document.Content.End - 1)
    $table = $document.Tables.Add($tableRange, 2, 2)
    $table.Cell(1, 1).Range.Text = "A"
    $table.Cell(1, 2).Range.Text = "B"
    $table.Cell(2, 1).Range.Text = "C"
    $table.Cell(2, 2).Range.Text = "D"
    $table.Range.Font.Name = "Arial"
    $table.Range.Font.Size = 9
    $table.Range.Font.Color = 255
    $table.Range.ParagraphFormat.Alignment = 0
    $table.Range.ParagraphFormat.FirstLineIndent = 7
    $table.AllowAutoFit = $false
    $table.Rows.Alignment = 2
    $table.Borders.Item(-1).LineStyle = 7
    $document.SaveAs2($sourcePath, 16)
    $document.Close(0)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    $document = $null

    $sourceHashBefore = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
    Copy-Item -LiteralPath $sourcePath -Destination $outputPath
    $document = $word.Documents.Open($outputPath, $false, $false, $false)
    $component = $document.VBProject.VBComponents.Add(1)
    $component.Name = "TablePreservationProbe"
    $component.CodeModule.AddFromString((Get-Content -Raw -Encoding UTF8 -LiteralPath $macroPath))
    $component.CodeModule.AddFromString(@'
Public Function ProbeFormatting() As Double
    Dim started As Single

    Application.ScreenUpdating = False
    started = Timer
    RunSDUTCMFormatting
    ProbeFormatting = Timer - started
End Function
'@)

    $document.Activate()
    $seconds = [double]$word.Run("ProbeFormatting")
    $table = $document.Tables.Item($document.Tables.Count)

    if ($table.Range.ParagraphFormat.LineSpacingRule -ne 0) {
        throw "Expected table paragraphs to use wdLineSpaceSingle."
    }
    if ($table.Range.Font.Name -ne "Arial" -or
        $table.Range.Font.Size -ne 9 -or
        $table.Range.Font.Color -ne 255) {
        throw "Default formatting changed the table font."
    }
    if ($table.Range.ParagraphFormat.Alignment -ne 0 -or
        $table.Range.ParagraphFormat.FirstLineIndent -ne 7) {
        throw "Default formatting changed table paragraph formatting other than line spacing."
    }
    if ($table.AllowAutoFit -ne 0 -or
        $table.Rows.Alignment -ne 2 -or
        $table.Borders.Item(-1).LineStyle -ne 7) {
        throw "Default formatting changed table layout or borders."
    }

    $document.Close(0)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    $document = $null

    $sourceHashAfter = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
    if ($sourceHashBefore -ne $sourceHashAfter) {
        throw "The source document was modified."
    }

    Write-Host ("Word table preservation check passed in {0:N3} s." -f $seconds)
}
finally {
    if ($document) { $document.Close(0) }
    if ($word) { $word.Quit() }
    if ($component) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($component) }
    if ($document) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
    if ($word) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    if (Test-Path -LiteralPath $sourcePath) { Remove-Item -LiteralPath $sourcePath -Force }
    if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
    if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Force }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
