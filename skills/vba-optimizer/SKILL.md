---
name: vba-optimizer
description: Reviews, profiles, and refactors Excel VBA for speed, correctness, and maintainability. Use when a user asks to optimize slow macros, reduce worksheet interaction, replace Select/Activate or Copy/Paste, batch range operations, use arrays or dictionaries, control calculation/events/screen updating safely, or evaluate VBA performance claims.
argument-hint: "[VBA file, pasted code, workbook task, or performance problem]"
---

# VBA Optimizer

Optimize Excel VBA by removing unnecessary work before applying micro-optimizations. Preserve behavior, workbook semantics, and maintainability. Never claim a speedup without measurement or a clearly labeled estimate.

## Core principle

The dominant cost in many Excel macros is not VBA syntax. It is repeated communication with the Excel object model, recalculation, events, rendering, clipboard operations, file I/O, and inefficient algorithms.

Prefer this order:

1. Measure and identify the bottleneck.
2. Preserve required behavior and outputs.
3. Remove unnecessary work.
4. Reduce Excel/VBA boundary crossings.
5. Replace cell-by-cell work with bulk operations or in-memory algorithms.
6. Disable expensive Excel features only when the workload triggers them.
7. Restore all application state on every exit path.
8. Re-measure and report the result.

## Loop-elimination decision order

When optimizing a loop, apply this order before tuning its body:

1. **Eliminate the loop**: remove iteration entirely, for example with bulk range assignment.
2. **Delegate the loop**: use Excel's built-in operations where their semantics fit, such as `SUM`, `Find`, `Sort`, `AutoFilter`, or `RemoveDuplicates`.
3. **Move the loop into memory**: replace worksheet iteration with arrays or dictionaries.
4. **Eliminate nested loops**: replace repeated searches with a `Dictionary` or another index to reduce algorithmic complexity.
5. **Reduce unavoidable loops**: limit the range, exit early, hoist loop-invariant work, and cache repeated calculations.
6. **Optimize the remaining loop body**: only then tune work that must remain inside the loop.

Array processing does not eliminate all iteration. It eliminates worksheet loops and moves the remaining iteration into memory, where it is typically far cheaper.

For a repeated lookup, replace an `O(n × m)` nested scan with a `Dictionary`-based index and lookup structure that is approximately `O(n + m)`, while preserving duplicate and ordering semantics.

## Workflow

### 1. Establish correctness and scope

Before editing, determine:

- What the macro must produce.
- Which values, formulas, formats, comments, validation, names, shapes, filters, and workbook events must be preserved.
- Whether the code operates on `ThisWorkbook`, `ActiveWorkbook`, external workbooks, or add-ins.
- Data size, formula density, volatile formulas, external links, file locations, and expected Excel version.
- Whether the user wants a minimal patch, a full rewrite, or a review only.

If the code is incomplete, optimize the visible portion and state assumptions. Do not invent workbook structure.

### 2. Measure before changing

Instrument meaningful phases rather than timing only the whole procedure. Measure repeated runs and prefer the median. Keep test data and output checks identical.

For simple runs, `Timer` is acceptable, but handle midnight rollover. For higher resolution on Windows, use `QueryPerformanceCounter` only when necessary and declare it correctly for 32-bit and 64-bit Office.

Record at least:

- Total elapsed time.
- Read time.
- in-memory processing time.
- Write time.
- Recalculation time.
- File or external-application I/O time.

Do not benchmark with dead code whose result is unused; an optimizer or different execution path may remove the work.

### 3. Find the highest-impact problems

Inspect in this priority order:

#### A. Worksheet interaction inside loops

Flag repeated calls such as:

```vba
For i = 2 To lastRow
    ws.Cells(i, 1).Value2 = ws.Cells(i, 2).Value2 * 2
Next i
```

Prefer one bulk read, in-memory processing, and one bulk write:

```vba
Dim data As Variant
Dim output() As Variant
Dim i As Long

 data = ws.Range("B2:B" & lastRow).Value2
ReDim output(1 To UBound(data, 1), 1 To 1)

For i = 1 To UBound(data, 1)
    If IsNumeric(data(i, 1)) Then
        output(i, 1) = data(i, 1) * 2
    End If
Next i

ws.Range("A2").Resize(UBound(output, 1), 1).Value2 = output
```

Range reads return a one-based, two-dimensional array, even for one column.

#### B. Clipboard and recorded UI operations

Remove unnecessary `Select`, `Activate`, `Selection`, `ActiveCell`, `Application.GoTo`, and sheet switching.

For values only:

```vba
 destination.Value2 = source.Value2
```

For formulas only:

```vba
 destination.Formula2 = source.Formula2
```

Use `Copy Destination:=...` only when the full cell content or formatting is truly required. Use `PasteSpecial` only for semantics that direct assignment cannot preserve.

#### C. Repeated algorithms

Replace nested scans with appropriate structures:

- `Scripting.Dictionary` for repeated exact-key lookups, uniqueness, grouping, and counts.
- `Range.Find` for locating cells in a range.
- Excel worksheet functions for one-shot aggregate or lookup operations when their semantics match.
- Sorting plus linear processing when order enables a simpler algorithm.
- Early exit when only the first match is required.

Do not call a worksheet function thousands of times in a loop without comparing an in-memory alternative.

Calling-convention traps when you replace a nested scan:

- `Range.Find` inherits `LookIn`, `LookAt`, `SearchOrder`, `MatchCase`, and `MatchByte` from the last use, including the user's last manual Ctrl+F. Pass all of them explicitly every time. When looping with `FindNext`, store the first hit's `.Address` and stop when it comes back around.
- `Application.Match` / `Application.VLookup` return a `Variant` error on no match (test with `IsError`); `WorksheetFunction.Match` / `.VLookup` raise a run-time error instead. Omitting the third `Match` argument (or passing 1) does an approximate match on unsorted data and returns wrong hits — pass 0 for exact.
- `Scripting.Dictionary`: early binding (`New Scripting.Dictionary`, needs the "Microsoft Scripting Runtime" reference) gives IntelliSense and slightly faster calls; late binding (`CreateObject("Scripting.Dictionary")`) needs no reference and runs on a machine that lacks it. It does not exist in Mac Excel — fall back to `Collection` (with an error-trapped key probe) or parallel key/value arrays.

#### D. Excessively broad ranges

Avoid whole-column, whole-row, and stale `UsedRange` processing unless required. Determine the real data bounds from a reliable key column or table. Account for blanks and multiple data blocks.

`UsedRange` does not shrink after rows are deleted until the workbook is saved. `ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row` finds the last row of one column but stops at its first trailing blank. For the true last used row anywhere on the sheet:

```vba
Dim f As Range
Dim lastRow As Long
Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
    LookIn:=xlFormulas, LookAt:=xlPart, _
    SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
If Not f Is Nothing Then lastRow = f.Row
```

With internal blank rows or several data blocks, walk a key column explicitly rather than trusting a single `End` jump.

#### E. Repeated object resolution

Cache stable objects:

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Data")
```

Use `With` when it improves clarity and avoids repeated property chains. Treat this as a secondary optimization, not a substitute for batching.

#### F. Unnecessary work

Remove:

- Recorder-generated property assignments unrelated to the requested change.
- Repeated last-row or last-column searches when a running index is sufficient.
- Repeated formatting when an entire range can be formatted once.
- Unused values, calls, `UsedRange` evaluations, and redundant conversions.
- Per-iteration `Debug.Print`, status updates, and message boxes.

For large logs, collect records in an array or buffer and emit summaries or a batch at the end. Avoid repeated growth of one huge string with `result = result & ...`.

#### G. Array rewrite correctness traps

The bulk read / in-memory process / bulk write in A is the highest-impact change and the easiest to get subtly wrong. Before shipping one, check:

- A single-cell range returns a scalar from `.Value2`, not a 2D array. `data = rng.Value2` then `UBound(data, 1)` raises error 13. Branch on `rng.Count` first, and special-case one cell, an empty range, and a single row or column.
- Cells holding an error (`#N/A`, `#VALUE!`, …) become `Variant/vbError` elements. `IsNumeric(v)` is `False`, but `v * 2` or `v > 0` raises error 13. Guard with `If Not IsError(v) Then`.
- `.Value2` returns dates and currency as `Double`. `.Value` returns `Date` and `Currency` (4-decimal rounding). Round-tripping through the wrong property changes the stored type; pick the property that matches how the values are compared or formatted downstream.
- Writing a 1D array to a range fills one row horizontally. To write a column, build a 2D `(1 To n, 1 To 1)` array directly. `Application.Transpose` has element-count, type, and 255-character string limits; do not rely on it here.
- If the destination range and the array differ in size, Excel silently truncates the array or fills the surplus cells with `#N/A`. Size the write exactly: `ws.Range("A2").Resize(UBound(a, 1), UBound(a, 2)).Value2 = a`.
- `ReDim Preserve` can only change the last dimension. For a growing result, over-allocate and trim, or make rows the last dimension and transpose once at the end.
- `.Value2` on a filtered or partly hidden range still returns every row. For visible rows only, iterate `rng.SpecialCells(xlCellTypeVisible).Areas`.
- In a merged range only the top-left cell holds the value; an array read returns `Empty` for the rest, and an array write errors or sets only the top-left.
- Reading a very large range into a Variant array can raise `Out of memory` (error 7); 32-bit Excel is capped near 1.3 GB in practice. Read in row chunks.

### 4. Control Excel state only when relevant

Potentially expensive features include:

- `Application.ScreenUpdating`
- `Application.Calculation`
- `Application.EnableEvents`
- `Application.DisplayAlerts`
- `Application.StatusBar`
- worksheet `DisplayPageBreaks`

Use them conditionally:

- Disable screen updating when code causes visible changes, selection, activation, or frequent redraws.
- Use manual calculation when many dependent formulas would otherwise recalculate repeatedly.
- Disable events when edits trigger event procedures or recursion.
- Disable alerts only when the expected dialogs are understood and intentionally handled.
- Disable page-break display only for page-layout-heavy work.

Always save and restore the original state, not assumed defaults. Restore state after errors too.

A leaked `Application.EnableEvents = False` is the worst case: if the macro exits without restoring it, the user's own `Worksheet_Change` and `Workbook_*` handlers stay dead for the rest of the Excel session, with no error and no visible cause. `ScreenUpdating` and `Calculation` leak the same way. This is why the restore must run on every exit path, including the error path.

```vba
Public Sub RunOptimized()
    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts

    On Error GoTo Fail

    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    ' Main work here.

CleanExit:
    With Application
        .Calculation = oldCalc
        .EnableEvents = oldEvents
        .ScreenUpdating = oldScreen
        .DisplayAlerts = oldAlerts
    End With

    If errNumber <> 0 Then
        Err.Raise errNumber, , errDescription
    End If
    Exit Sub

Fail:
    errNumber = Err.Number
    errDescription = Err.Description
    Resume CleanExit
End Sub
```

If current formula results are needed while calculation is manual, call `.Calculate` on the specific range, worksheet, or workbook at the right point; reading `.Value2` alone returns the stale result. After a `Copy`, set `Application.CutCopyMode = False` so Excel does not prompt about a large clipboard on close.

### 5. Choose properties by semantics

Prefer `.Value2` for underlying values when date/currency coercion and displayed text are not required.

Use:

- `.Text` only when the formatted display string is required.
- `.Value` when VBA Date or Currency conversion is intentionally required.
- `.Formula2` for modern formula semantics when supported and required.
- `.Formula` when compatibility with legacy implicit-intersection behavior is required.

Do not remove explicit `.Value` merely because it can be omitted. Choose explicitness and semantics over tiny or unverified differences.

### 6. Treat micro-optimizations as low priority

Do not prioritize these before algorithm and boundary improvements:

- `For Next` versus `For Each` versus `Do Loop`.
- `Long` versus `Variant` solely for speed in object-model-heavy code.
- `Left` versus `Left$` solely for speed.
- `Sheets(1)` versus `Sheets("Data")` solely for speed.
- `Range` versus `Cells` solely for speed.
- Explicit `.Value` versus the default property.
- `With` versus an object variable solely for speed.

Still use explicit types and `Option Explicit` for correctness, intent, maintainability, API contracts, and compatibility. Remember that in `Dim a, b, c As Long`, only `c` is a `Long`.

Choose `Cells(row, column)` for variable indexes and `Range` for fixed addresses, named ranges, and multi-cell ranges. Use whichever expresses the requirement clearly.

### 7. Be cautious with aggressive or external optimization

Do not recommend registry changes, antivirus exclusions, DLL compilation, or external executables as routine optimization.

A native/DLL compiler may help CPU-bound, strongly typed, in-memory algorithms, but it may provide little benefit for Excel object-model calls. Evaluate deployment, signing, security policy, architecture compatibility, maintainability, and reproducible real-workbook benchmarks.

Never generalize extreme benchmark claims from loops whose outputs are unused.

### 8. Validate after refactoring

Check:

- Same values and error values.
- Same formulas and formula semantics.
- Same row ordering and duplicate behavior.
- Same formatting where required.
- Same handling of blanks, text numbers, dates, currency, booleans, and errors.
- Same filter, hidden-row, merged-cell, protection, and event behavior.
- Correct workbook and worksheet qualification.
- Application state restored after success, error, and user cancellation.

For array rewrites, test empty ranges, one-cell ranges, one-row ranges, a range containing an error value, and the maximum expected size.

## When the worksheet itself is slow

Some workbooks are slow to open, edit, calculate, and save regardless of any macro. If the macro spends its time in recalculation, look here before tuning VBA:

- Volatile functions recalculate on every change: `NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `OFFSET`, `INDIRECT`, `CELL`, `INFO`.
- Array formulas, `SUMPRODUCT`, and `VLOOKUP` / `MATCH` over entire columns (`A:A`) instead of the real data range.
- Conditional formatting rule count growing through copy/paste, especially rules with volatile or whole-sheet formulas. Check `Cells.FormatConditions.Count`.
- Data validation lists that reference large or volatile ranges.
- External workbook links, including stale ones under Edit Links.
- Defined-name bloat and thousands of distinct cell styles ("Too many different cell formats", error 1004).
- Oversized or duplicated pivot caches inflating file size and open time.

These are workbook-repair items, not code changes. Report them separately from the macro optimization.

## Practical references

- `references/practical.md` — a drop-in range-comparison routine for proving a refactor preserved behavior, plus three worked before/after case studies (recorder cleanup, nested-loop join, filtered-row extraction).

## Review output format

When reviewing or rewriting code, return:

1. **Diagnosis**: likely bottlenecks ranked by expected impact.
2. **Correctness risks**: behavior that must not change.
3. **Optimized code**: complete, runnable code where enough context exists.
4. **Why it is faster**: tie each change to fewer calls, less recalculation, less I/O, or lower algorithmic complexity.
5. **Measurement plan**: how to compare before and after.
6. **Trade-offs and assumptions**: formulas, formats, events, compatibility, and memory.

Avoid unsupported promises such as “100x faster.” If no benchmark was run, say “expected to reduce…” rather than claiming a measured gain.

Label each issue by severity so reviews read consistently:

- **Critical**: wrong output, data loss, or corruption — an array write that truncates rows, a leaked `EnableEvents = False`, unqualified `ActiveSheet` writing to the wrong sheet.
- **High**: a real bottleneck that will not scale — worksheet access inside a large loop, `O(n × m)` nested matching, whole-column processing.
- **Medium**: technical debt with limited blast radius — repeated last-row searches, recorder cruft, un-batched formatting.
- **Low**: naming, micro-optimizations, style.
- **Info**: noted only; may be an acceptable choice.

Commonly missed in VBA reviews:

- Application state (`EnableEvents`, `Calculation`, `ScreenUpdating`) not restored on the error path.
- Worksheet access still inside the loop after a partial rewrite.
- `data = rng.Value2` breaking on a one-cell range (error 13).
- `Range.Find` relying on inherited `LookIn` / `LookAt` state.
- `WorksheetFunction.*` raising instead of returning an error, with no handler.
- Volatile worksheet functions, not the macro, driving the recalculation cost.
- `UsedRange` still reporting deleted rows.

## Quick audit checklist

- [ ] Is workbook/sheet access fully qualified?
- [ ] Are `Select`, `Activate`, and `Selection` unnecessary?
- [ ] Are values or formulas copied through the clipboard unnecessarily?
- [ ] Is worksheet access happening inside a large loop?
- [ ] Can the range be read and written once with a 2D Variant array?
- [ ] Can repeated lookup be replaced by `Dictionary`, `Find`, or one worksheet function call?
- [ ] Does the loop stop when its objective is met?
- [ ] Is the processed range larger than the actual data?
- [ ] Are last-row/last-column calculations repeated unnecessarily?
- [ ] Are formatting and logging operations batched?
- [ ] Are calculation, events, and rendering actual bottlenecks?
- [ ] Are original application settings restored on all exits?
- [ ] Are `.Value2`, `.Value`, `.Text`, `.Formula`, and `.Formula2` chosen by semantics?
- [ ] Has correctness been tested before and after?
- [ ] Has performance been measured with representative data?

