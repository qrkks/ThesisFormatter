# Project AI Guidance

## Scope

These instructions apply to the entire repository. Preserve existing user changes and keep modifications focused on the requested behavior.

## Performance Principles

Performance is a requirement for every new feature and behavior change, not a later cleanup step.

- Before implementation, identify whether the change adds a full-document scan, repeated scan, per-paragraph formatting pass, table layout recalculation, TOC refresh, pagination, or other expensive Word operation.
- Do not iterate through `Range.Characters` to preserve character formatting. Properties that must remain unchanged should not be assigned.
- Avoid indexed paragraph loops such as `ActiveDocument.Paragraphs(i)`, especially repeated full-document loops. Prefer one `For Each` pass, a bounded `Range`, or one in-memory `ActiveDocument.Content.Text` lookup.
- Prefer Word style configuration and range-level batch assignment over repeated paragraph or character property writes.
- Apply only properties required by the feature. Do not set a value and then restore it.
- Keep expensive optional behavior, such as table AutoFit or full table restyling, outside the default path unless the feature explicitly requires it.
- When several features need the same document traversal, consolidate them where doing so does not make behavior harder to verify.

## Required Testing

Use test-driven development for code changes: add a failing regression test, verify the expected failure, implement the smallest fix, and rerun all tests.

Run the static regression suite:

```powershell
& ./tests/aggressive-performance.Tests.ps1
& ./tests/reference-formatting.Tests.ps1
& ./tests/table-formatting.Tests.ps1
```

Static tests are not sufficient for changes to the main formatting path. Such changes must also be tested through Word automation on a representative real document.

## Real-Document Performance Testing

- Never modify or overwrite the source thesis document. Open it read-only or copy it to a temporary file before running the macro.
- Use `tests/full-performance-profile.ps1` for full-flow timing and `tests/reference-performance-profile.ps1` for reference-only timing.
- Prefer supplying the benchmark path through a local variable or environment variable instead of committing a machine-specific document path.

Example:

```powershell
$benchmark = $env:THESIS_FORMATTER_BENCHMARK_DOC
& ./tests/full-performance-profile.ps1 -DocumentPath $benchmark -MaxSeconds 120
```

For a new feature that affects the main path:

1. Record the baseline duration before the implementation.
2. Record the duration after the implementation using the same document and environment.
3. If the total time regresses materially, collect per-stage timings and find the cause before merging.
4. Do not raise a performance threshold merely to make a regression pass without documenting and approving the tradeoff.

## Output Verification

Performance improvements must preserve formatting behavior. Inspect the generated temporary document, not only the source code.

Verify the areas affected by the change, including where relevant:

- title and heading styles, alignment, and effective indentation;
- body first-line indentation and line spacing;
- abstract and keyword text merging;
- TOC fields and page numbers;
- reference heading, reference range, italics, and hanging indentation;
- table count and table formatting;
- paragraph, word, character, field, and section counts.

When direct formatting may override a style, verify the effective property on actual output paragraphs through Word automation.

## Completion Criteria

A main-path change is complete only when:

- the new regression test was observed failing before the implementation;
- all static regression tests pass;
- `git diff --check` passes;
- the real-document performance test completes successfully;
- before-and-after timing is reported;
- generated output has been checked for the affected formatting behavior;
- the source document remains unchanged.
