# Font Formatting Performance Fix Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace character-by-character Word font formatting with one range-level operation while preserving italic and color by leaving those properties untouched.

**Architecture:** Keep the existing shared `ApplyRangeFontPreservingItalic` entry point so its body and reference callers remain stable. Change only its implementation and the structural regression test that defines which font properties it may touch.

**Tech Stack:** Word VBA, PowerShell regression scripts, Git

---

### Task 1: Define the range-level formatting contract

**Files:**
- Modify: `tests/reference-formatting.Tests.ps1:49-88`
- Test: `tests/reference-formatting.Tests.ps1`

- [ ] **Step 1: Replace the character-preservation assertions with the desired contract**

Keep the existing helper/body/reference extraction and replace the helper assertions with:

```powershell
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
```

Retain the assertions that `FormatBodyParagraph` and `FormatReferenceEntries` call `ApplyRangeFontPreservingItalic`, plus the existing page-break assertion.

- [ ] **Step 2: Run the focused test and verify RED**

Run:

```powershell
& ./tests/reference-formatting.Tests.ps1
```

Expected: FAIL with `Font formatting must operate on the whole range, not character by character.` because the current helper traverses `contentRange.Characters`.

- [ ] **Step 3: Commit the failing regression test**

```powershell
git add tests/reference-formatting.Tests.ps1
git commit -m "test: guard range-level font formatting"
```

### Task 2: Replace character-level formatting with one range operation

**Files:**
- Modify: `format_macro.bas:1539-1565`
- Test: `tests/reference-formatting.Tests.ps1`
- Test: `tests/table-formatting.Tests.ps1`

- [ ] **Step 1: Implement the minimal range-level helper**

Replace the body of `ApplyRangeFontPreservingItalic` with:

```vb
Private Sub ApplyRangeFontPreservingItalic(ByVal sourceRange As Range, ByVal eastAsianFont As String, ByVal latinFont As String, ByVal fontSize As Single, ByVal isBold As Boolean)
    If sourceRange Is Nothing Then Exit Sub

    With sourceRange.Font
        .NameFarEast = eastAsianFont
        .Name = latinFont
        .Size = fontSize
        .Bold = isBold
    End With
End Sub
```

Do not change either caller. Because neither `Italic` nor `Color` is assigned, Word retains those properties.

- [ ] **Step 2: Run all regression tests and verify GREEN**

Run:

```powershell
& ./tests/reference-formatting.Tests.ps1
& ./tests/table-formatting.Tests.ps1
```

Expected: both scripts print their `regression checks passed` message and exit 0.

- [ ] **Step 3: Verify the resulting diff**

Run:

```powershell
git diff --check
rg -n "contentRange\.Characters|For Each charRange|\.Italic\s*=|\.Color\s*=" format_macro.bas
```

Expected: `git diff --check` exits 0. The search may find color assignments elsewhere in the macro, but none inside `ApplyRangeFontPreservingItalic`; inspect the helper diff to confirm its only assignments are `NameFarEast`, `Name`, `Size`, and `Bold`.

- [ ] **Step 4: Commit the implementation**

```powershell
git add format_macro.bas
git commit -m "fix: batch body font formatting"
```

### Task 3: Final verification

**Files:**
- Verify: `format_macro.bas`
- Verify: `tests/reference-formatting.Tests.ps1`
- Verify: `tests/table-formatting.Tests.ps1`

- [ ] **Step 1: Run the complete available test suite from a clean command**

```powershell
& ./tests/reference-formatting.Tests.ps1
& ./tests/table-formatting.Tests.ps1
git diff --check HEAD~2..HEAD
git status --short
```

Expected: both test scripts pass, diff check exits 0, and `git status --short` is empty except for the implementation plan if it has not yet been committed.

- [ ] **Step 2: Review scope against the design**

Confirm that the final two implementation commits change only `tests/reference-formatting.Tests.ps1` and `format_macro.bas`; paragraph layout, tables, page breaks, and style configuration remain untouched.
