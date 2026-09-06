# VBA Optimizer — practical reference

Companion code for `../SKILL.md`. Two things:

1. A range-comparison routine for proving a refactor did not change behavior.
2. Three worked before/after case studies.

No measurement or benchmark tooling here — see `SKILL.md` §2 for that. These are
about correctness and structure.

---

## 1. Prove a refactor preserved behavior

Run the old macro against one range and the new macro against an identical copy,
then call `CompareRanges` on the two result ranges. It prints the first mismatch
(address, both values, both types) to the Immediate window and returns the number
of differing cells. `0` means cell-for-cell identical; `-1` means the ranges are
not the same shape.

Paste into a standard module. No references required.

```vba
Option Explicit

Public Function CompareRanges(ByVal a As Range, ByVal b As Range) As Long
    Dim va As Variant, vb As Variant
    Dim r As Long, c As Long, rowCount As Long, colCount As Long
    Dim mismatches As Long

    If a.Rows.Count <> b.Rows.Count Or a.Columns.Count <> b.Columns.Count Then
        Debug.Print "Shape differs: " & a.Address(False, False) & " is " & _
            a.Rows.Count & "x" & a.Columns.Count & ", " & _
            b.Address(False, False) & " is " & b.Rows.Count & "x" & b.Columns.Count
        CompareRanges = -1
        Exit Function
    End If

    va = ReadAs2D(a)
    vb = ReadAs2D(b)
    rowCount = UBound(va, 1)
    colCount = UBound(va, 2)

    For r = 1 To rowCount
        For c = 1 To colCount
            If Not CellsMatch(va(r, c), vb(r, c)) Then
                mismatches = mismatches + 1
                If mismatches = 1 Then
                    Debug.Print "First mismatch at " & _
                        a.Cells(r, c).Address(False, False) & ": " & _
                        Describe(va(r, c)) & "  vs  " & Describe(vb(r, c))
                End If
            End If
        Next c
    Next r

    If mismatches = 0 Then
        Debug.Print "Ranges identical: " & rowCount & "x" & colCount
    Else
        Debug.Print "Mismatched cells: " & mismatches & " of " & (rowCount * colCount)
    End If
    CompareRanges = mismatches
End Function

' Always return a 1-based 2D array, even for a single cell or a single row/column.
Private Function ReadAs2D(ByVal rng As Range) As Variant
    Dim out() As Variant
    If rng.Count = 1 Then
        ReDim out(1 To 1, 1 To 1)
        out(1, 1) = rng.Value2
        ReadAs2D = out
    Else
        ReadAs2D = rng.Value2
    End If
End Function

Private Function CellsMatch(ByVal x As Variant, ByVal y As Variant) As Boolean
    Dim mag As Double

    If IsError(x) Or IsError(y) Then
        ' ponytail: treat any two errors as equal; flag only error-vs-value.
        CellsMatch = IsError(x) And IsError(y)
    ElseIf IsEmpty(x) Or IsEmpty(y) Then
        CellsMatch = IsEmpty(x) And IsEmpty(y)
    ElseIf VarType(x) = vbDouble And VarType(y) = vbDouble Then
        mag = 1
        If Abs(x) > mag Then mag = Abs(x)
        If Abs(y) > mag Then mag = Abs(y)
        CellsMatch = (Abs(x - y) <= 0.0000000001 * mag)
    Else
        CellsMatch = (VarType(x) = VarType(y))
        If CellsMatch Then CellsMatch = (x = y)
    End If
End Function

Private Function Describe(ByVal v As Variant) As String
    If IsError(v) Then
        Describe = "<error>"
    ElseIf IsEmpty(v) Then
        Describe = "<empty>"
    Else
        Describe = TypeName(v) & " '" & CStr(v) & "'"
    End If
End Function
```

Notes:

- Pass a single contiguous range on each side. A multi-area (`Union`) range does
  not read cleanly with `.Value2`.
- The `vbDouble` branch uses a relative tolerance so the last-bit differences
  between a cell-by-cell write and an array round-trip do not register as a
  regression. Tighten or drop it if you need bit-exact equality.
- Errors are compared only as "both are errors". If a refactor could turn `#N/A`
  into `#VALUE!` and that matters, compare the error text separately.

---

## 2. Case studies

Each shows the original, the rewrite, the boundary crossings removed, and the
behavior that stays the same. Before and after share a procedure name by design —
paste them into separate modules if you want both.

### 2.1 Recorder macro: sheet activation + clipboard

**Before** — two sheet activations, three selections, one clipboard round-trip.

```vba
Sub CopyTotals()
    Sheets("Source").Select
    Range("A2:A1000").Select
    Selection.Copy
    Sheets("Report").Select
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub
```

**After** — one qualified array read, one write.

```vba
Sub CopyTotals()
    Dim src As Range, dst As Range
    Set src = ThisWorkbook.Worksheets("Source").Range("A2:A1000")
    Set dst = ThisWorkbook.Worksheets("Report").Range("B2")
    dst.Resize(src.Rows.Count, src.Columns.Count).Value2 = src.Value2
End Sub
```

- **Removed**: 2 `.Select` of sheets, 3 `.Select` of ranges, `Copy` +
  `PasteSpecial` clipboard round-trip, dependence on the active sheet.
- **Preserved**: values only (matches `xlPasteValues`), same source and target
  cells. Formats and formulas are still not copied — same as the original paste.

### 2.2 Nested loop → dictionary join

**Before** — `O(orders × customers)` cell reads.

```vba
' For each order row, look up the customer name and write it into column C.
Sub FillCustomerNames()
    Dim wsO As Worksheet, wsC As Worksheet
    Dim o As Long, c As Long, lastO As Long, lastC As Long
    Set wsO = ThisWorkbook.Worksheets("Orders")
    Set wsC = ThisWorkbook.Worksheets("Customers")
    lastO = wsO.Cells(wsO.Rows.Count, 1).End(xlUp).Row
    lastC = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row

    For o = 2 To lastO
        For c = 2 To lastC
            If wsC.Cells(c, 1).Value2 = wsO.Cells(o, 2).Value2 Then
                wsO.Cells(o, 3).Value2 = wsC.Cells(c, 2).Value2
                Exit For
            End If
        Next c
    Next o
End Sub
```

**After** — `O(orders + customers)`; one index build, two bulk reads, one write.

```vba
Sub FillCustomerNames()
    Dim wsO As Worksheet, wsC As Worksheet
    Dim i As Long, lastO As Long, lastC As Long
    Dim cust As Variant, ord As Variant, out() As Variant
    Dim map As Object

    Set wsO = ThisWorkbook.Worksheets("Orders")
    Set wsC = ThisWorkbook.Worksheets("Customers")
    lastO = wsO.Cells(wsO.Rows.Count, 1).End(xlUp).Row
    lastC = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    If lastO < 2 Then Exit Sub

    Set map = CreateObject("Scripting.Dictionary")
    If lastC >= 2 Then
        cust = wsC.Range("A2:B" & lastC).Value2
        For i = 1 To UBound(cust, 1)
            If Not map.Exists(cust(i, 1)) Then map(cust(i, 1)) = cust(i, 2)
        Next i
    End If

    ' Read 2 columns so a single data row still yields a 2D array
    ' (see SKILL.md "Array rewrite correctness traps").
    ord = wsO.Range("B2:C" & lastO).Value2
    ReDim out(1 To UBound(ord, 1), 1 To 1)
    For i = 1 To UBound(ord, 1)
        If map.Exists(ord(i, 1)) Then out(i, 1) = map(ord(i, 1))
    Next i

    wsO.Range("C2").Resize(UBound(out, 1), 1).Value2 = out
End Sub
```

- **Removed**: up to `lastO × lastC` cell reads and `lastO` scattered writes →
  two bulk reads, one bulk write, plus an in-memory pass. Complexity
  `O(n × m) → O(n + m)`.
- **Preserved**: first match wins (`If Not map.Exists` keeps the first customer
  row for a duplicate key; `Exit For` did the same); a blank result when no
  customer matches.
- **Watch**: `.Value2` key comparison matches the original `=` comparison only
  when the key columns hold the same type on both sheets (e.g. both numbers or
  both text). Normalize the key if one side is text-numbers.

### 2.3 Filtered rows copied one at a time

**Before** — one `.Copy` per visible row.

```vba
Sub ExtractVisible()
    Dim ws As Worksheet, wsOut As Worksheet
    Dim last As Long, r As Long, outRow As Long
    Set ws = ThisWorkbook.Worksheets("Data")
    Set wsOut = ThisWorkbook.Worksheets("Extract")
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    outRow = 1
    For r = 2 To last
        If Not ws.Rows(r).Hidden Then
            ws.Range("A" & r & ":E" & r).Copy wsOut.Range("A" & outRow)
            outRow = outRow + 1
        End If
    Next r
End Sub
```

**After** — one `.Copy` per contiguous visible block (usually a handful).

```vba
Sub ExtractVisible()
    Dim ws As Worksheet, wsOut As Worksheet
    Dim last As Long, outRow As Long
    Dim src As Range, vis As Range, area As Range
    Set ws = ThisWorkbook.Worksheets("Data")
    Set wsOut = ThisWorkbook.Worksheets("Extract")
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Exit Sub

    Set src = ws.Range("A2:E" & last)
    On Error Resume Next
    Set vis = src.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If vis Is Nothing Then Exit Sub

    outRow = 1
    For Each area In vis.Areas
        area.Copy wsOut.Range("A" & outRow)
        outRow = outRow + area.Rows.Count
    Next area
End Sub
```

- **Removed**: one clipboard round-trip per visible row → one per visible block.
  A filter that leaves 5 contiguous stretches turns thousands of `.Copy` calls
  into 5.
- **Preserved**: same rows, same columns, full cell content and formatting
  (still `.Copy`, not value assignment), same output order.
- **Watch**: `SpecialCells(xlCellTypeVisible)` raises when nothing is visible —
  trapped above. For values only, read `src.Value2` and filter in memory instead
  of `.Copy`.
