# Clear Direct Heading Indents Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Force the document title and outline levels 1–3 to have zero first-line, left, and right indentation even when source paragraphs contain direct formatting.

**Architecture:** Add one targeted paragraph pass immediately after style configuration. The pass reads each paragraph's title style or outline level and writes only three indentation properties for matching headings.

**Tech Stack:** Word VBA, PowerShell regression scripts, Word COM automation, Git

---

### Task 1: Define direct heading-indent cleanup

**Files:**
- Modify: `tests/aggressive-performance.Tests.ps1`

- [ ] **Step 1: Add cleanup extraction and assertions**

After existing body extraction, add:

```powershell
$headingIndentCleanup = Get-SubBody "ClearHeadingParagraphIndents"
```

Add these assertions before the final success message:

```powershell
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
```

- [ ] **Step 2: Run the focused test and verify RED**

```powershell
& ./tests/aggressive-performance.Tests.ps1
```

Expected: FAIL because `ClearHeadingParagraphIndents` does not exist.

- [ ] **Step 3: Commit the failing test**

```powershell
git add tests/aggressive-performance.Tests.ps1
git commit -m "test: require direct heading indent cleanup"
```

### Task 2: Clear direct heading indentation

**Files:**
- Modify: `format_macro.bas:1220-1232`
- Modify: `format_macro.bas:1278-1293`
- Test: `tests/aggressive-performance.Tests.ps1`

- [ ] **Step 1: Add the cleanup call after style configuration**

The start of `RunSDUTCMFormatting` must be:

```vb
Private Sub RunSDUTCMFormatting()
    SetPageAndBodyFormat
    ConfigureSDUTCMStyles
    ClearHeadingParagraphIndents
```

- [ ] **Step 2: Add the targeted cleanup procedure**

Add after `ConfigureSDUTCMStyles`:

```vb
Private Sub ClearHeadingParagraphIndents()
    Dim para As Paragraph
    Dim outlineLevel As WdOutlineLevel

    For Each para In ActiveDocument.Paragraphs
        outlineLevel = para.OutlineLevel
        If para.Style = ZhTitleStyleName() Or _
           (outlineLevel >= wdOutlineLevel1 And outlineLevel <= wdOutlineLevel3) Then
            With para.Range.ParagraphFormat
                .FirstLineIndent = 0
                .LeftIndent = 0
                .RightIndent = 0
            End With
        End If
    Next para
End Sub
```

- [ ] **Step 3: Run all static tests and verify GREEN**

```powershell
& ./tests/aggressive-performance.Tests.ps1
& ./tests/reference-formatting.Tests.ps1
& ./tests/table-formatting.Tests.ps1
git diff --check
```

Expected: all three tests pass and diff check exits 0.

- [ ] **Step 4: Commit the implementation**

```powershell
git add format_macro.bas
git commit -m "fix: clear direct heading indents"
```

### Task 3: Verify real-document output

**Files:**
- Verify: `format_macro.bas`
- Output: `%TEMP%\COST-aggressive-performance-output.docx`

- [ ] **Step 1: Regenerate the comparison document**

```powershell
$output = Join-Path $env:TEMP 'COST-aggressive-performance-output.docx'
& ./tests/full-performance-profile.ps1 `
  -DocumentPath 'C:\Coding\学术\master-thesis\research\literature\review_book\_output\COST-量表结构与经济毒性文献综述工作稿.docx' `
  -MaxSeconds 120 `
  -OutputPath $output
```

Expected: exits 0 and writes the output document.

- [ ] **Step 2: Inspect every outline level 1–3 paragraph**

Open the generated output read-only through Word automation. For each paragraph whose `OutlineLevel` is 1, 2, or 3, assert:

```powershell
$paragraph.Range.ParagraphFormat.FirstLineIndent -eq 0
$paragraph.Range.ParagraphFormat.LeftIndent -eq 0
$paragraph.Range.ParagraphFormat.RightIndent -eq 0
```

Expected: no heading violates any zero-indent assertion.

- [ ] **Step 3: Confirm the worktree is clean**

```powershell
git status --short
```

Expected: no output.
