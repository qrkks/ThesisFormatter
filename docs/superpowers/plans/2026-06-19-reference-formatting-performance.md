# Reference Formatting Performance Optimization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Format the reference section through one heading lookup, one bounded section scan, and one range-level formatting operation.

**Architecture:** Add a private optimized reference-section pipeline used by both complete reference entry points. Keep the existing standalone macros for compatibility, while tests ensure the main formatting path no longer chains three full-document passes.

**Tech Stack:** Word VBA, PowerShell regression scripts, Word COM automation, Git

---

### Task 1: Define the optimized reference-section contract

**Files:**
- Modify: `tests/reference-formatting.Tests.ps1`
- Create: `tests/reference-performance-profile.ps1`

- [ ] **Step 1: Add static assertions for the optimized main path**

After extracting `$referenceFormatter`, add:

```powershell
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
        -Pattern "NormalizeReferenceHeadingParagraphs|FormatReferences|FormatReferenceEntries" `
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
```

Retain all existing assertions for the legacy helper and page-break behavior.

- [ ] **Step 2: Add a reusable real-document performance profiler**

Create `tests/reference-performance-profile.ps1`:

```powershell
param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,
    [double]$MaxSeconds = 15
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$macroPath = Join-Path $repoRoot "format_macro.bas"
$resolvedDocument = (Resolve-Path -LiteralPath $DocumentPath).Path
$tempPath = Join-Path $env:TEMP ("reference-profile-" + [guid]::NewGuid() + ".docx")

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
    $component.Name = "ReferenceProfiler"
    $component.CodeModule.AddFromString(
        (Get-Content -Raw -Encoding UTF8 -LiteralPath $macroPath)
    )
    $component.CodeModule.AddFromString(@'
Public Function ProfileReferenceFormatting() As Double
    Dim started As Single

    Application.ScreenUpdating = False
    started = Timer
    ProcessReferencesWithSort
    ProfileReferenceFormatting = Timer - started
End Function
'@)

    $document.Activate()
    $seconds = [double]$word.Run("ProfileReferenceFormatting")
    Write-Host ("Reference formatting: {0:N3} s" -f $seconds)

    if ($seconds -gt $MaxSeconds) {
        throw "Reference formatting exceeded ${MaxSeconds}s (actual: $seconds s)."
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
```

- [ ] **Step 3: Run the static regression test and verify RED**

Run:

```powershell
& ./tests/reference-formatting.Tests.ps1
```

Expected: FAIL because `FormatReferenceSection` does not exist.

- [ ] **Step 4: Run the performance profiler and verify the old path exceeds the budget**

Run:

```powershell
& ./tests/reference-performance-profile.ps1 `
  -DocumentPath 'C:\Coding\学术\master-thesis\research\literature\review_book\_output\COST-量表结构与经济毒性文献综述工作稿.docx'
```

Expected: FAIL after reporting a reference-formatting duration substantially above 15 seconds.

- [ ] **Step 5: Commit the failing tests**

```powershell
git add tests/reference-formatting.Tests.ps1 tests/reference-performance-profile.ps1
git commit -m "test: define bounded reference formatting"
```

### Task 2: Implement bounded reference-section formatting

**Files:**
- Modify: `format_macro.bas:798-906`
- Modify: `format_macro.bas:1381-1466`
- Test: `tests/reference-formatting.Tests.ps1`

- [ ] **Step 1: Add the optimized reference-section helpers**

Add these private procedures near the existing reference helpers:

```vb
Private Sub FormatReferenceSection()
    Dim headingPara As Paragraph
    Dim para As Paragraph
    Dim searchRange As Range
    Dim entriesRange As Range
    Dim insertRange As Range
    Dim txt As String
    Dim headingLabel As String
    Dim remainder As String
    Dim entriesStart As Long
    Dim entriesEnd As Long

    Set headingPara = FindReferenceHeadingParagraph()
    If headingPara Is Nothing Then Exit Sub

    txt = Trim(Replace(headingPara.Range.Text, vbCr, ""))
    remainder = GetReferenceHeadingRemainder(txt)
    If Len(remainder) > 0 Then
        headingLabel = GetReferenceHeadingLabel(txt)
        headingPara.Range.Text = headingLabel & vbCr
        Set insertRange = ActiveDocument.Range(headingPara.Range.End, headingPara.Range.End)
        insertRange.InsertAfter remainder & vbCr
    End If

    FormatReferenceHeadingParagraph headingPara

    entriesStart = headingPara.Range.End
    entriesEnd = ActiveDocument.Content.End
    Set searchRange = ActiveDocument.Range(entriesStart, entriesEnd)

    For Each para In searchRange.Paragraphs
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        If IsReferenceSectionEndParagraph(para, txt) Then
            entriesEnd = para.Range.Start
            EnsurePageBreakBeforeParagraph para
            Exit For
        End If
    Next para

    If entriesEnd <= entriesStart Then Exit Sub

    Set entriesRange = ActiveDocument.Range(entriesStart, entriesEnd)
    ApplyReferenceEntriesFormat entriesRange
End Sub

Private Function FindReferenceHeadingParagraph() As Paragraph
    Dim para As Paragraph
    Dim txt As String

    For Each para In ActiveDocument.Paragraphs
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        If IsReferenceHeadingText(txt) Then
            Set FindReferenceHeadingParagraph = para
            Exit Function
        End If
    Next para
End Function

Private Function IsReferenceSectionEndParagraph(ByVal para As Paragraph, ByVal txt As String) As Boolean
    If txt = "附录" Or txt = "Appendix" Or _
       Left(txt, 3) = "图 " Or Left(txt, 3) = "表 " Or _
       Left(txt, 4) = "Figure" Or Left(txt, 4) = "Table" Or _
       Left(txt, 5) = "致谢" Or Left(txt, 5) = "Acknowledgments" Or _
       Left(txt, 6) = "作者简介" Or Left(txt, 6) = "Author Bio" Then
        IsReferenceSectionEndParagraph = True
        Exit Function
    End If

    IsReferenceSectionEndParagraph = IsHeadingLevelParagraph(para, 1) Or _
                                     IsHeadingLevelParagraph(para, 2) Or _
                                     IsHeadingLevelParagraph(para, 3)
End Function

Private Sub FormatReferenceHeadingParagraph(ByVal para As Paragraph)
    On Error Resume Next
    para.Style = ActiveDocument.Styles("标题 1")
    If Err.Number <> 0 Then
        Err.Clear
        para.Style = ActiveDocument.Styles("Heading 1")
    End If
    On Error GoTo 0

    With para.Range.Font
        .NameFarEast = "宋体"
        .Name = "宋体"
        .Size = 18
        .Bold = True
        .Color = wdColorBlack
    End With

    With para.Range.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .FirstLineIndent = 0
        .LeftIndent = 0
        .RightIndent = 0
    End With

    EnsurePageBreakBeforeParagraph para
End Sub

Private Sub ApplyReferenceEntriesFormat(ByVal entriesRange As Range)
    If entriesRange Is Nothing Then Exit Sub

    With entriesRange.ParagraphFormat
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = -36
        .LeftIndent = 36
        .LineSpacingRule = wdLineSpace1pt5
    End With

    With entriesRange.Font
        .NameFarEast = "宋体"
        .Name = "Times New Roman"
        .Size = 12
        .Bold = False
    End With
End Sub
```

- [ ] **Step 2: Route both complete reference entry points through the optimized formatter**

Replace both complete-process bodies with:

```vb
Sub ProcessReferences()
    FormatReferenceSection
End Sub

Sub ProcessReferencesWithSort()
    FormatReferenceSection
End Sub
```

Do not modify `FormatReferences`, `FormatReferenceEntries`, `AutoNumberReferences`, or `SortReferences`; they remain compatibility entry points.

- [ ] **Step 3: Run the focused static test and verify GREEN**

Run:

```powershell
& ./tests/reference-formatting.Tests.ps1
```

Expected: `Reference formatting regression checks passed.`

- [ ] **Step 4: Run all static regression tests**

Run:

```powershell
& ./tests/reference-formatting.Tests.ps1
& ./tests/table-formatting.Tests.ps1
git diff --check
```

Expected: both test scripts pass and `git diff --check` exits 0.

- [ ] **Step 5: Commit the implementation**

```powershell
git add format_macro.bas
git commit -m "perf: format references as one range"
```

### Task 3: Verify performance on the real document

**Files:**
- Verify: `format_macro.bas`
- Verify: `tests/reference-performance-profile.ps1`

- [ ] **Step 1: Run the real-document performance profiler and verify GREEN**

Run:

```powershell
& ./tests/reference-performance-profile.ps1 `
  -DocumentPath 'C:\Coding\学术\master-thesis\research\literature\review_book\_output\COST-量表结构与经济毒性文献综述工作稿.docx'
```

Expected: prints a duration no greater than 15 seconds and exits 0.

- [ ] **Step 2: Run final verification from a clean command**

Run:

```powershell
& ./tests/reference-formatting.Tests.ps1
& ./tests/table-formatting.Tests.ps1
git diff --check HEAD~2..HEAD
git status --short
```

Expected: both static tests pass, diff check exits 0, and the worktree is clean.

- [ ] **Step 3: Commit the profiler if it was not included in the test commit**

```powershell
git add tests/reference-performance-profile.ps1
git commit -m "test: add reference performance profiler"
```

Skip this commit only when `git status --short` confirms the profiler was already committed in Task 1.
