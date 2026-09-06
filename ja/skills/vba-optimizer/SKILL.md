---
name: vba-optimizer
description: Excel VBA を速度、正確性、保守性の観点からレビュー、計測、リファクタリングするスキルです。遅いマクロの最適化、ワークシート操作の削減、Select/Activate や Copy/Paste の置換、範囲操作の一括化、配列や辞書の使用、計算・イベント・画面更新の安全な制御、VBA の性能主張の評価を依頼されたときに使用してください。
argument-hint: "[VBA ファイル、貼り付けたコード、ブックの作業内容、または性能上の問題]"
---

# VBA オプティマイザー

細かな最適化の前に、不要な処理を取り除いて Excel VBA を最適化します。挙動、ブックの意味論、保守性を守ってください。計測、または明確に「推定」と示した根拠なしに高速化を主張してはいけません。

## 基本原則

多くの Excel マクロで支配的なコストは VBA 構文ではありません。Excel オブジェクトモデルとの反復通信、再計算、イベント、描画、クリップボード操作、ファイル I/O、非効率なアルゴリズムです。

次の順序を優先します。

1. 計測してボトルネックを特定する。
2. 必要な挙動と出力を守る。
3. 不要な処理を削除する。
4. Excel/VBA 境界の往復を減らす。
5. セル単位の処理を一括操作またはメモリ内アルゴリズムに置き換える。
6. 高コストな Excel 機能は、負荷が発生する場合だけ無効化する。
7. すべての終了経路でアプリケーション状態を復元する。
8. 再計測して結果を報告する。

## ループ排除の判断順序

ループを最適化するときは、ループ本体を調整する前に次の順序で判断します。

1. **ループをなくす**: 範囲の一括代入などで、反復そのものを削除する。
2. **ループを Excel に委譲する**: 意味論が合う場合は、`SUM`、`Find`、`Sort`、`AutoFilter`、`RemoveDuplicates` など Excel の組み込み処理を使う。
3. **ループをメモリ内へ移す**: ワークシート上の反復を、配列または Dictionary に置き換える。
4. **二重ループをなくす**: 繰り返す検索を `Dictionary` などのインデックスに置き換え、計算量を減らす。
5. **避けられないループを減らす**: 処理範囲の限定、早期終了、ループ不変処理の外出し、重複計算のキャッシュを行う。
6. **残るループ本体を最適化する**: 最後に、どうしてもループ内に残る処理だけを調整する。

配列処理は、すべての反復をなくす手法ではありません。遅いワークシート上のループを排除し、通常ははるかに低コストなメモリ内へ残る反復を移す手法です。

繰り返し検索では、`O(n × m)` の二重走査を `Dictionary` によるインデックスと検索構造へ置き換え、重複時の扱いと順序の意味論を保ちながら、おおむね `O(n + m)` にします。

## 進め方

### 1. 正確性と対象範囲を確立する

編集前に、マクロが作るべき結果、保持すべき値・数式・書式・コメント・検証・名前・図形・フィルター・ブックイベントを確認します。`ThisWorkbook`、`ActiveWorkbook`、外部ブック、アドインのどれを扱うか、データ量、数式密度、揮発性数式、外部リンク、ファイル場所、想定 Excel バージョンも確認します。

最小変更、全面書換え、レビューのみのどれを望むかを確認します。コードが不完全なら見えている部分だけを最適化し、前提を明示します。ブック構造を捏造してはいけません。

### 2. 変更前に計測する

手続き全体だけではなく意味のある工程を計測します。反復実行して中央値を使い、テストデータと出力検証は同一にします。単純な実行では `Timer` を使えますが日付またぎを考慮します。Windows で高分解能が必要な場合だけ、32/64 ビット Office に正しく対応した `QueryPerformanceCounter` を使用します。

少なくとも合計経過時間、読込時間、メモリ内処理時間、書込時間、再計算時間、ファイルまたは外部アプリケーション I/O 時間を記録します。結果が未使用のデッドコードは、最適化や実行経路の違いで除去され得るためベンチマークに使いません。

### 3. 影響が大きい問題を先に探す

#### A. ループ内のワークシート操作

次のような反復呼出しを指摘します。

```vba
For i = 2 To lastRow
    ws.Cells(i, 1).Value2 = ws.Cells(i, 2).Value2 * 2
Next i
```

一度の一括読込、メモリ内処理、一度の一括書込を優先します。

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

1 列だけでも、範囲の読込結果は 1 始まりの 2 次元配列です。

#### B. クリップボードと記録マクロの UI 操作

不要な `Select`、`Activate`、`Selection`、`ActiveCell`、`Application.GoTo`、シート切替を取り除きます。値だけなら `destination.Value2 = source.Value2`、数式だけなら `destination.Formula2 = source.Formula2` を使います。セル内容や書式全体が必要な場合だけ `Copy Destination:=...` を使い、直接代入では保持できない意味論だけに `PasteSpecial` を使います。

#### C. 繰り返しアルゴリズム

入れ子の走査を、目的に合う構造へ置換します。

- 繰り返す完全一致検索、重複排除、グループ化、件数集計には `Scripting.Dictionary`
- 範囲内のセル探索には `Range.Find`
- 意味論が合う一度きりの集計・検索にはワークシート関数
- 順序が単純な処理を可能にするならソート後の線形処理
- 最初の一致だけが必要なら早期終了

メモリ内の代替案と比較せず、ワークシート関数をループ内で数千回呼んではいけません。

入れ子の走査を置き換えるときの呼び出し規約のワナ。

- `Range.Find` の `LookIn`、`LookAt`、`SearchOrder`、`MatchCase`、`MatchByte` は前回値（ユーザーの直前の手動 Ctrl+F を含む）が残ります。毎回すべて明示します。`FindNext` でループするときは最初の一致の `.Address` を保持し、一巡して戻ったら止めます。
- `Application.Match` / `Application.VLookup` は不一致で `Variant` のエラーを返します（`IsError` で判定）。`WorksheetFunction.Match` / `.VLookup` は代わりに実行時エラーを送出します。`Match` の第 3 引数を省略（または 1 を指定）すると未ソートデータに対する近似一致になり誤った結果を返すため、完全一致には 0 を渡します。
- `Scripting.Dictionary`: 事前バインド（`New Scripting.Dictionary`、参照設定「Microsoft Scripting Runtime」が必要）は入力補完が効き呼び出しがわずかに速く、遅延バインド（`CreateObject("Scripting.Dictionary")`）は参照不要で未設定のマシンでも動きます。Mac 版 Excel には存在しないため、`Collection`（キー探索はエラートラップ）またはキー／値の並行配列で代替します。

#### D. 過度に広い範囲

必要でない限り、列全体・行全体・古い `UsedRange` を処理しません。信頼できるキー列またはテーブルから実データ範囲を決め、空白と複数データブロックを考慮します。

`UsedRange` は行を削除してもブックを保存するまで縮みません。`ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row` は 1 列の最終行を返しますが、その列の最初の末尾空白で止まります。シート上のどこであっても本当の最終使用行を得るには次のようにします。

```vba
Dim f As Range
Dim lastRow As Long
Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
    LookIn:=xlFormulas, LookAt:=xlPart, _
    SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
If Not f Is Nothing Then lastRow = f.Row
```

途中に空白行があるときや複数ブロックがあるときは、1 回の `End` ジャンプに頼らずキー列を明示的に走査します。

#### E. オブジェクト解決の繰り返し

安定したオブジェクトはキャッシュします。

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Data")
```

可読性が上がり、プロパティ連鎖を減らせるときは `With` を使います。ただし一括処理の代替となる主最適化ではありません。

#### F. 不要な処理

要求に関係ない記録マクロ由来のプロパティ代入、繰り返される最終行・最終列検索、範囲全体で一度できる書式設定、未使用の値・呼出し・`UsedRange` 評価・重複変換、反復ごとの `Debug.Print`・状態更新・メッセージボックスを削除します。大量ログは配列やバッファに集め、最後に要約または一括出力します。`result = result & ...` で巨大な文字列を繰り返し伸長しません。

#### G. 配列書換えの正確性のワナ

A の一括読込・メモリ内処理・一括書込は最も効果が大きく、かつ最も気づきにくく壊しやすい変更です。出す前に次を確認します。

- 単一セルの範囲は `.Value2` からスカラーを返し、2 次元配列ではありません。`data = rng.Value2` のあと `UBound(data, 1)` はエラー 13。先に `rng.Count` で分岐し、1 セル・空範囲・1 行/1 列を特別扱いします。
- エラー（`#N/A`、`#VALUE!` など）を持つセルは `Variant/vbError` 要素になります。`IsNumeric(v)` は `False` ですが `v * 2` や `v > 0` はエラー 13。`If Not IsError(v) Then` で防御します。
- `.Value2` は日付と通貨を `Double` で返します。`.Value` は `Date` と `Currency`（小数 4 桁で丸め）を返します。誤ったプロパティで往復すると格納型が変わります。値を後段でどう比較・書式設定するかに合わせて選びます。
- 1 次元配列を範囲へ書くと横 1 行に展開されます。列に書くには `(1 To n, 1 To 1)` の 2 次元配列を直接構築します。`Application.Transpose` は要素数・型・255 文字の制限があり、ここでは頼りません。
- 書込先の範囲と配列のサイズが違うと、Excel は配列を静かに切り詰めるか余りのセルを `#N/A` で埋めます。書込サイズを厳密に合わせます。`ws.Range("A2").Resize(UBound(a, 1), UBound(a, 2)).Value2 = a`。
- `ReDim Preserve` は最終次元しか変えられません。伸ばす出力は過剰確保して切り詰めるか、行を最終次元にして最後に一度転置します。
- フィルターや一部非表示の範囲でも `.Value2` は全行を返します。可視行だけなら `rng.SpecialCells(xlCellTypeVisible).Areas` を反復します。
- 結合範囲では左上セルだけが値を持ちます。配列読込は残りを `Empty` で返し、配列書込はエラーになるか左上だけを設定します。
- 非常に大きな範囲を Variant 配列へ読み込むと `Out of memory`（エラー 7）。32 ビット Excel は実質 1.3 GB 付近が上限です。行チャンクで読みます。

### 4. 必要な場合だけ Excel の状態を制御する

高コストになり得るものは、`Application.ScreenUpdating`、`Application.Calculation`、`Application.EnableEvents`、`Application.DisplayAlerts`、`Application.StatusBar`、ワークシートの `DisplayPageBreaks` です。

表示変更や頻繁な再描画があるときだけ画面更新を止めます。多数の依存数式が繰り返し再計算されるときだけ手動計算にします。編集がイベント処理や再帰を起こすときだけイベントを止めます。想定ダイアログを理解して意図的に処理するときだけ警告を止めます。改ページ表示はページレイアウト中心の作業だけで止めます。

想定した既定値ではなく、元の状態を保存して、エラー時にも復元します。

最悪なのは `Application.EnableEvents = False` の復元漏れです。復元せずに終了すると、ユーザー自身の `Worksheet_Change` や `Workbook_*` ハンドラーが、エラーも目に見える原因もないまま、その Excel セッション中ずっと動かなくなります。`ScreenUpdating` と `Calculation` も同様に漏れます。だからこそ復元はエラー経路を含むすべての終了経路で実行しなければなりません。

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

    ' ここに主処理を記述する。

CleanExit:
    With Application
        .Calculation = oldCalc
        .EnableEvents = oldEvents
        .ScreenUpdating = oldScreen
        .DisplayAlerts = oldAlerts
    End With
    If errNumber <> 0 Then Err.Raise errNumber, , errDescription
    Exit Sub

Fail:
    errNumber = Err.Number
    errDescription = Err.Description
    Resume CleanExit
End Sub
```

手動計算中に現在の数式結果が必要なら、適切な時点で必要な範囲、ワークシート、またはブックだけを `.Calculate` します。`.Value2` を読むだけでは古い結果が返ります。`Copy` のあとは `Application.CutCopyMode = False` にして、終了時に大容量クリップボードの確認が出ないようにします。

### 5. 意味論に応じてプロパティを選ぶ

日付・通貨の強制変換や表示文字列が不要な基礎値には `.Value2` を優先します。

- 書式済みの表示文字列が必要なときだけ `.Text`
- VBA の Date または Currency 変換を意図して必要とするときだけ `.Value`
- 対応環境で現代的な数式意味論が必要なときは `.Formula2`
- 従来の暗黙の交差動作との互換性が必要なときは `.Formula`

省略できるからという理由だけで明示的な `.Value` を削除しません。ごく小さい未検証の差より、明示性と意味論を優先します。

### 6. 細かな最適化は優先度を下げる

アルゴリズムと Excel/VBA 境界の改善より、`For Next`・`For Each`・`Do Loop` の比較、オブジェクトモデル中心コードでの `Long` と `Variant`、`Left` と `Left$`、`Sheets(1)` と `Sheets("Data")`、`Range` と `Cells`、明示的 `.Value` と既定プロパティ、`With` とオブジェクト変数の速度差を優先してはいけません。

ただし正確性、意図、保守性、API 契約、互換性のために明示的な型と `Option Explicit` を使います。`Dim a, b, c As Long` では `Long` は `c` だけです。可変インデックスには `Cells(row, column)`、固定アドレス・名前付き範囲・複数セル範囲には `Range` を使い、要件を最も明瞭に表せる方を選びます。

### 7. 攻撃的・外部的な最適化に注意する

レジストリ変更、ウイルス対策除外、DLL コンパイル、外部実行ファイルを通常の最適化として推奨しません。ネイティブ/DLL コンパイラは CPU 集約的で強い型付けのメモリ内アルゴリズムには有効な場合がありますが、Excel オブジェクトモデル呼出しにはほとんど効果がないことがあります。デプロイ、署名、セキュリティポリシー、アーキテクチャ互換性、保守性、再現可能な実ブックでの測定を評価します。未使用の出力を持つループから極端なベンチマーク結果を一般化してはいけません。

### 8. リファクタリング後に検証する

値・エラー値、数式と数式意味論、行順と重複の扱い、必要な書式、空白・文字列の数値・日付・通貨・真偽値・エラーの扱い、フィルター・非表示行・結合セル・保護・イベント動作、ブックとワークシートの修飾、成功・エラー・ユーザーキャンセル後の状態復元を確認します。配列への書換えでは、空範囲、1 セル、1 行、エラー値を含む範囲、想定最大サイズをテストします。

## ワークシート自体が遅い場合

マクロと関係なく、開く・編集・計算・保存が遅いブックがあります。マクロの時間が再計算に費やされているなら、VBA を調整する前にここを見ます。

- 揮発性関数は変更のたびに再計算します。`NOW`、`TODAY`、`RAND`、`RANDBETWEEN`、`OFFSET`、`INDIRECT`、`CELL`、`INFO`。
- 実データ範囲ではなく列全体（`A:A`）に対する配列数式、`SUMPRODUCT`、`VLOOKUP` / `MATCH`。
- コピー＆ペーストで増殖する条件付き書式のルール数。特に揮発性またはシート全体の数式を持つルール。`Cells.FormatConditions.Count` を確認します。
- 大きな範囲や揮発性の範囲を参照するデータ入力規則リスト。
- 外部ブックリンク。「リンクの編集」にある古いものを含む。
- 定義名の肥大、数千の異なるセルスタイル（「セルの書式が多すぎます」、エラー 1004）。
- ファイルサイズと開く時間を膨らませる過大または重複したピボットキャッシュ。

これらはコード変更ではなくブック修復の項目です。マクロの最適化とは分けて報告します。

## 実践リファレンス

- `references/practical.md` — リファクタで挙動が変わっていないことを証明する貼り付け用の範囲比較ルーチンと、Before/After のケーススタディ 3 本（記録マクロの整理、二重ループの結合、フィルター行の抽出）。

## レビューの出力形式

コードのレビューまたは書換えでは、次を返します。

1. **診断**: 想定される影響度順のボトルネック
2. **正確性のリスク**: 変えてはいけない挙動
3. **最適化済みコード**: 十分な文脈があれば、完全で実行可能なコード
4. **高速化の理由**: 呼出し、再計算、I/O、アルゴリズム計算量の削減と変更を対応付ける
5. **計測計画**: 変更前後の比較方法
6. **トレードオフと前提**: 数式、書式、イベント、互換性、メモリ

根拠のない「100 倍高速」などの約束を避けます。ベンチマークを実行していないなら、計測済みの改善を主張せず「削減が見込まれる」と表現します。

レビューの表記を揃えるため、各指摘に深刻度を付けます。

- **Critical**: 出力の誤り、データ損失、破損。行を切り詰める配列書込、`EnableEvents = False` の復元漏れ、修飾漏れの `ActiveSheet` で別シートに書き込む。
- **High**: スケールしない実ボトルネック。大きなループ内のワークシートアクセス、`O(n × m)` の入れ子照合、列全体の処理。
- **Medium**: 影響範囲が限定的な技術的負債。最終行検索の繰り返し、記録マクロの残骸、未一括化の書式設定。
- **Low**: 命名、細かな最適化、スタイル。
- **Info**: 記録のみ。許容できる選択のこともある。

VBA レビューで見落としやすい点。

- エラー経路でアプリケーション状態（`EnableEvents`、`Calculation`、`ScreenUpdating`）が復元されない。
- 部分的な書換えのあと、ループ内にワークシートアクセスが残っている。
- `data = rng.Value2` が 1 セル範囲で壊れる（エラー 13）。
- `Range.Find` が継承した `LookIn` / `LookAt` の状態に依存している。
- `WorksheetFunction.*` がエラーを返さず送出し、ハンドラーがない。
- マクロではなく揮発性ワークシート関数が再計算コストを生んでいる。
- `UsedRange` が削除済みの行をまだ報告している。

## 簡易監査チェックリスト

- [ ] ブック／シートへのアクセスは完全修飾されているか
- [ ] `Select`、`Activate`、`Selection` は不要ではないか
- [ ] 値や数式を不要にクリップボード経由でコピーしていないか
- [ ] 大きなループ内でワークシートにアクセスしていないか
- [ ] 2 次元 Variant 配列で 1 回の読込・書込にできるか
- [ ] 繰り返す検索を `Dictionary`、`Find`、1 回のワークシート関数に置換できるか
- [ ] 目的達成時にループを停止するか
- [ ] 処理範囲が実データより大きすぎないか
- [ ] 最終行・最終列の計算を不要に繰り返していないか
- [ ] 書式設定とログ出力を一括化しているか
- [ ] 計算、イベント、描画が実際のボトルネックか
- [ ] すべての終了経路で元のアプリケーション設定を復元するか
- [ ] `.Value2`、`.Value`、`.Text`、`.Formula`、`.Formula2` を意味論で選んでいるか
- [ ] 変更前後の正確性をテストしたか
- [ ] 代表的なデータで性能を測定したか
