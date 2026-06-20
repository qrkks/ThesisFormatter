# Aggressive Performance Experiment Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reduce the default full-document formatting path below 20 seconds by relying on Word styles, targeting abstract labels with `Range.Find`, and making table formatting opt-in.

**Architecture:** The default pipeline becomes style-driven and removes the all-paragraph direct-formatting pass. Abstract handling retains its current formatting body but feeds it at most four paragraph positions discovered by `Range.Find`; `ProcessTables` remains public but leaves the default pipeline.

**Tech Stack:** Word VBA, PowerShell regression scripts, Word COM automation, Git worktree

---

### Task 1: Create the isolated experiment workspace

**Files:**
- Verify: `.gitignore`
- Worktree: `.worktrees/aggressive-performance`

- [ ] **Step 1: Create the experiment branch and worktree**

Run from `C:\Coding\ThesisFormatter`:

```powershell
git check-ignore .worktrees/test
git worktree add .worktrees/aggressive-performance -b codex/aggressive-performance
```

Expected: `.worktrees/test` is ignored and the new worktree is created from the latest `main`.

- [ ] **Step 2: Verify the baseline in the worktree**

Run from `C:\Coding\ThesisFormatter\.worktrees\aggressive-performance`:

```powershell
& ./tests/reference-formatting.Tests.ps1
& ./tests/table-formatting.Tests.ps1
git status --short
```

Expected: both tests pass and status is clean.

### Task 2: Define the aggressive fast-path contract

**Files:**
- Create: `tests/aggressive-performance.Tests.ps1`
- Create: `tests/full-performance-profile.ps1`

- [ ] **Step 1: Add static fast-path assertions**

Create `tests/aggressive-performance.Tests.ps1`:

```powershell
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
if ($abstractFormatter -match "ActiveDocument\.Paragraphs\s*\(") {
    throw "Abstract formatting must not scan paragraphs by index."
}
if ($abstractFormatter -notmatch "FindAbstractLabelParagraphStart") {
    throw "Abstract formatting should target known labels."
}
if ($abstractFinder -notmatch "\.Find") {
    throw "Abstract label lookup should use Range.Find."
}

Write-Host "Aggressive performance regression checks passed."
```

- [ ] **Step 2: Add the full-pipeline profiler**

Create `tests/full-performance-profile.ps1`:

```powershell
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
```

The source document is copied and never opened for writing.

- [ ] **Step 3: Verify RED**

Run:

```powershell
& ./tests/aggressive-performance.Tests.ps1
```

Expected: FAIL because the current pipeline contains the direct paragraph loop and `FindAbstractLabelParagraphStart` does not exist.

- [ ] **Step 4: Verify the existing pipeline exceeds the performance budget**

Run:

```powershell
& ./tests/full-performance-profile.ps1 `
  -DocumentPath 'C:\Coding\学术\master-thesis\research\literature\review_book\_output\COST-量表结构与经济毒性文献综述工作稿.docx'
```

Expected: FAIL after reporting approximately 113 seconds, above the 20-second budget.

- [ ] **Step 5: Commit the failing tests**

```powershell
git add tests/aggressive-performance.Tests.ps1 tests/full-performance-profile.ps1
git commit -m "test: define aggressive formatting fast path"
```

### Task 3: Remove default direct paragraph and table formatting

**Files:**
- Modify: `format_macro.bas:1159-1210`
- Test: `tests/aggressive-performance.Tests.ps1`

- [ ] **Step 1: Replace the default pipeline**

Replace `RunSDUTCMFormatting` with:

```vb
Private Sub RunSDUTCMFormatting()
    SetPageAndBodyFormat
    ConfigureSDUTCMStyles
    MergeAndFormatAbstract
    ProcessTableOfContents
    ProcessReferencesWithSort
    ProcessImages
    ApplyMixedPageNumbersByTOC
End Sub
```

Do not delete any standalone paragraph-formatting or table-formatting macro.

- [ ] **Step 2: Run the static test**

```powershell
& ./tests/aggressive-performance.Tests.ps1
```

Expected: still FAIL only because `FindAbstractLabelParagraphStart` is missing or unused.

### Task 4: Replace the abstract full scan with targeted lookups

**Files:**
- Modify: `format_macro.bas:125-277`
- Test: `tests/aggressive-performance.Tests.ps1`

- [ ] **Step 1: Change the setup and loop in `MergeAndFormatAbstract`**

Retain the existing formatting body inside the loop, but replace its declarations and `For i = ActiveDocument.Paragraphs.Count To 1 Step -1` setup with:

```vb
    Dim i As Integer
    Dim para As Paragraph
    Dim nextPara As Paragraph
    Dim txt As String
    Dim contentTxt As String
    Dim rngEnd As Range
    Dim rng As Range
    Dim targetStarts(1 To 4) As Long
    Dim targetCount As Integer
    Dim sortI As Integer
    Dim sortJ As Integer
    Dim swapStart As Long
    Dim targetStart As Long

    targetStart = FindAbstractLabelParagraphStart("摘要")
    If targetStart > 0 Then
        targetCount = targetCount + 1
        targetStarts(targetCount) = targetStart
    End If

    targetStart = FindAbstractLabelParagraphStart("关键词")
    If targetStart > 0 Then
        targetCount = targetCount + 1
        targetStarts(targetCount) = targetStart
    End If

    targetStart = FindAbstractLabelParagraphStart("Abstract")
    If targetStart > 0 Then
        targetCount = targetCount + 1
        targetStarts(targetCount) = targetStart
    End If

    targetStart = FindAbstractLabelParagraphStart("Keywords")
    If targetStart > 0 Then
        targetCount = targetCount + 1
        targetStarts(targetCount) = targetStart
    End If

    For sortI = 1 To targetCount - 1
        For sortJ = sortI + 1 To targetCount
            If targetStarts(sortI) < targetStarts(sortJ) Then
                swapStart = targetStarts(sortI)
                targetStarts(sortI) = targetStarts(sortJ)
                targetStarts(sortJ) = swapStart
            End If
        Next sortJ
    Next sortI

    For i = 1 To targetCount
        Set para = ActiveDocument.Range(targetStarts(i), targetStarts(i)).Paragraphs(1)
```

Keep the existing `Next i` at the end. Descending positions preserve range validity when a later paragraph is deleted during merging.

- [ ] **Step 2: Add the targeted finder**

Add immediately after `MergeAndFormatAbstract`:

```vb
Private Function FindAbstractLabelParagraphStart(ByVal label As String) As Long
    Dim searchRange As Range
    Dim para As Paragraph
    Dim txt As String
    Dim nextStart As Long

    Set searchRange = ActiveDocument.Content.Duplicate
    With searchRange.Find
        .ClearFormatting
        .Text = label
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
    End With

    Do While searchRange.Find.Execute
        Set para = searchRange.Paragraphs(1)
        txt = Trim(Replace(para.Range.Text, vbCr, ""))

        If txt = label Or Left(txt, Len(label) + 1) = label & ":" Or _
           Left(txt, Len(label) + 1) = label & "：" Then
            FindAbstractLabelParagraphStart = para.Range.Start
            Exit Function
        End If

        nextStart = searchRange.End
        searchRange.SetRange Start:=nextStart, End:=ActiveDocument.Content.End
    Loop
End Function
```

- [ ] **Step 3: Verify GREEN**

Run:

```powershell
& ./tests/aggressive-performance.Tests.ps1
& ./tests/reference-formatting.Tests.ps1
& ./tests/table-formatting.Tests.ps1
git diff --check
```

Expected: all three scripts pass and diff check exits 0.

- [ ] **Step 4: Commit the implementation**

```powershell
git add format_macro.bas
git commit -m "perf: add style-driven formatting fast path"
```

### Task 5: Verify performance and produce a comparison document

**Files:**
- Verify: `format_macro.bas`
- Verify: `tests/full-performance-profile.ps1`
- Output: `%TEMP%\COST-aggressive-performance-output.docx`

- [ ] **Step 1: Run the full profiler and save the output copy**

```powershell
$output = Join-Path $env:TEMP 'COST-aggressive-performance-output.docx'
& ./tests/full-performance-profile.ps1 `
  -DocumentPath 'C:\Coding\学术\master-thesis\research\literature\review_book\_output\COST-量表结构与经济毒性文献综述工作稿.docx' `
  -OutputPath $output
```

Expected: total duration is no greater than 20 seconds and the output document exists.

- [ ] **Step 2: Run final verification**

```powershell
& ./tests/aggressive-performance.Tests.ps1
& ./tests/reference-formatting.Tests.ps1
& ./tests/table-formatting.Tests.ps1
git diff --check HEAD~2..HEAD
git status --short
```

Expected: all tests pass, diff check exits 0, and the worktree is clean.
