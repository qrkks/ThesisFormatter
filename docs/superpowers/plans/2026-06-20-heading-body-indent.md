# Heading and Body Indentation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Keep all title and heading styles explicitly unindented while making every supported body style use a 24 pt first-line indent.

**Architecture:** Change only the body-style configuration arguments in `ConfigureSDUTCMStyles`. Add structural regression assertions so future performance work cannot silently restore no-indent `Normal` or `First Paragraph` behavior.

**Tech Stack:** Word VBA, PowerShell regression scripts, Git

---

### Task 1: Define indentation rules in tests

**Files:**
- Modify: `tests/aggressive-performance.Tests.ps1`

- [ ] **Step 1: Add style-body extraction and assertions**

After the current body extraction, add:

```powershell
$styleConfiguration = Get-SubBody "ConfigureSDUTCMStyles"
$titleStyleConfiguration = Get-SubBody "ConfigureTitleStyleIfExists"
$headingStyleConfiguration = Get-SubBody "ConfigureHeadingStyleIfExists"

foreach ($styleName in @(
    'ZhBodyTextStyleName\(\)',
    'ZhBodyStyleName\(\)',
    '"Normal"',
    '"First Paragraph"'
)) {
    if ($styleConfiguration -notmatch "ConfigureBodyStyleIfExists\s+$styleName,\s*24") {
        throw "Body style $styleName should use a 24 pt first-line indent."
    }
}

foreach ($styleBody in @($titleStyleConfiguration, $headingStyleConfiguration)) {
    foreach ($property in @("FirstLineIndent", "LeftIndent", "RightIndent")) {
        if ($styleBody -notmatch "\.$property\s*=\s*0") {
            throw "Title and heading styles should explicitly set $property to zero."
        }
    }
}
```

- [ ] **Step 2: Run the focused test and verify RED**

```powershell
& ./tests/aggressive-performance.Tests.ps1
```

Expected: FAIL because `Normal` and `First Paragraph` currently pass `0`.

- [ ] **Step 3: Commit the failing test**

```powershell
git add tests/aggressive-performance.Tests.ps1
git commit -m "test: define body indentation rules"
```

### Task 2: Configure all body styles with first-line indentation

**Files:**
- Modify: `format_macro.bas:1278-1293`
- Test: `tests/aggressive-performance.Tests.ps1`

- [ ] **Step 1: Change the two body-style arguments**

Use these four calls in `ConfigureSDUTCMStyles`:

```vb
    ConfigureBodyStyleIfExists ZhBodyTextStyleName(), 24
    ConfigureBodyStyleIfExists ZhBodyStyleName(), 24
    ConfigureBodyStyleIfExists "Normal", 24
    ConfigureBodyStyleIfExists "First Paragraph", 24
```

Leave title and heading indentation assignments unchanged.

- [ ] **Step 2: Run all static tests and verify GREEN**

```powershell
& ./tests/aggressive-performance.Tests.ps1
& ./tests/reference-formatting.Tests.ps1
& ./tests/table-formatting.Tests.ps1
git diff --check
```

Expected: all three scripts pass and diff check exits 0.

- [ ] **Step 3: Commit the implementation**

```powershell
git add format_macro.bas
git commit -m "fix: indent all supported body styles"
```

### Task 3: Verify the experiment branch

**Files:**
- Verify: `format_macro.bas`
- Verify: `tests/aggressive-performance.Tests.ps1`

- [ ] **Step 1: Run performance verification with the existing experimental budget**

```powershell
& ./tests/full-performance-profile.ps1 `
  -DocumentPath 'C:\Coding\学术\master-thesis\research\literature\review_book\_output\COST-量表结构与经济毒性文献综述工作稿.docx' `
  -MaxSeconds 120
```

Expected: exits 0 and remains near the current experimental result of roughly 42 seconds.

- [ ] **Step 2: Confirm the worktree is clean**

```powershell
git status --short
```

Expected: no output.
