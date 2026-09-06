# VBA オプティマイザー — 実践リファレンス

`../SKILL.md` の付属コード。内容は 2 つ。

1. リファクタで挙動が変わっていないことを証明する範囲比較ルーチン。
2. Before/After のケーススタディ 3 本。

計測・ベンチマークのツールはここにはありません（`SKILL.md` §2 を参照）。ここは
正確性と構造の話です。

---

## 1. リファクタで挙動が変わっていないことを証明する

旧マクロをある範囲に対して実行し、新マクロを同一のコピーに対して実行してから、
2 つの結果範囲に対して `CompareRanges` を呼びます。最初の不一致（アドレス、両方の
値、両方の型）をイミディエイトウィンドウに出力し、異なるセルの数を返します。
`0` はセル単位で完全一致、`-1` は形が違うことを表します。

標準モジュールに貼り付けます。参照設定は不要です。

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

' 単一セルや 1 行/1 列でも、常に 1 始まりの 2 次元配列を返す。
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
        ' ponytail: 2 つのエラーは等しいとみなし、エラー対値だけを検出する。
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

補足。

- 両側とも連続した単一範囲を渡します。複数エリア（`Union`）の範囲は `.Value2` で
  きれいに読めません。
- `vbDouble` の分岐は相対許容差を使い、セル単位の書込と配列往復の最下位ビットの
  差を退行として拾わないようにしています。ビット完全一致が必要なら、狭めるか外し
  ます。
- エラーは「両方エラーかどうか」だけを比較します。`#N/A` が `#VALUE!` に変わり得て
  それが重要なら、エラーの文言を別途比較します。

---

## 2. ケーススタディ

それぞれ、元のコード、書換え、消えた境界越え、変わらない挙動を示します。Before と
After は意図的に手続き名が同じです。両方を使うなら別々のモジュールに貼り付けます。

### 2.1 記録マクロ: シートのアクティブ化 + クリップボード

**Before** — シートのアクティブ化 2 回、選択 3 回、クリップボード往復 1 回。

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

**After** — 修飾済みの配列読込 1 回、書込 1 回。

```vba
Sub CopyTotals()
    Dim src As Range, dst As Range
    Set src = ThisWorkbook.Worksheets("Source").Range("A2:A1000")
    Set dst = ThisWorkbook.Worksheets("Report").Range("B2")
    dst.Resize(src.Rows.Count, src.Columns.Count).Value2 = src.Value2
End Sub
```

- **消えたもの**: シートの `.Select` 2 回、範囲の `.Select` 3 回、`Copy` +
  `PasteSpecial` のクリップボード往復、アクティブシートへの依存。
- **保たれるもの**: 値のみ（`xlPasteValues` と一致）、同じ元セルと先セル。書式と
  数式は元の貼り付けと同じくコピーされません。

### 2.2 二重ループ → 辞書による結合

**Before** — `O(注文数 × 顧客数)` のセル読込。

```vba
' 各注文行について顧客名を引き、C 列に書き込む。
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

**After** — `O(注文数 + 顧客数)`。インデックス構築 1 回、一括読込 2 回、書込 1 回。

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

    ' 2 列読むことで、データ行が 1 行でも 2 次元配列になる
    ' （SKILL.md「配列書換えの正確性のワナ」を参照）。
    ord = wsO.Range("B2:C" & lastO).Value2
    ReDim out(1 To UBound(ord, 1), 1 To 1)
    For i = 1 To UBound(ord, 1)
        If map.Exists(ord(i, 1)) Then out(i, 1) = map(ord(i, 1))
    Next i

    wsO.Range("C2").Resize(UBound(out, 1), 1).Value2 = out
End Sub
```

- **消えたもの**: 最大 `lastO × lastC` のセル読込と `lastO` 回の散発的な書込 →
  一括読込 2 回、一括書込 1 回、メモリ内の 1 パス。計算量 `O(n × m) → O(n + m)`。
- **保たれるもの**: 最初の一致を採用（`If Not map.Exists` が重複キーの最初の顧客行
  を残す。`Exit For` と同じ）。一致しないときは空。
- **注意**: `.Value2` のキー比較が元の `=` 比較と一致するのは、両シートのキー列が
  同じ型（どちらも数値、またはどちらも文字列）のときだけです。片方が文字列の数値
  ならキーを正規化します。

### 2.3 フィルター行を 1 行ずつコピー

**Before** — 可視行 1 行ごとに `.Copy` 1 回。

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

**After** — 連続した可視ブロック 1 つごとに `.Copy` 1 回（通常は数個）。

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

- **消えたもの**: 可視行ごとのクリップボード往復 → 可視ブロックごと。連続した
  5 区画を残すフィルターなら、数千回の `.Copy` が 5 回になります。
- **保たれるもの**: 同じ行、同じ列、セル内容と書式の全体（値代入ではなく `.Copy`）、
  同じ出力順。
- **注意**: `SpecialCells(xlCellTypeVisible)` は可視セルが無いと送出します（上で
  トラップ済み）。値だけなら `src.Value2` を読み、`.Copy` ではなくメモリ内で
  絞り込みます。
```
