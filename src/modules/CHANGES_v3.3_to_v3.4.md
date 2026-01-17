# ver3.3 → ver3.4 変更点詳細

## 概要
左右端部のピーク検出を **1ピーク → 最大2ピーク** に拡張

---

## 1. Config シート

### ver3.3
```
A列                    | B列
-----------------------|-------
Parameter              | Value
L_mm                   | 15
CenterFrac             | 0.1
Hist_BinCount          | 20
```

### ver3.4 ★変更
```
A列                    | B列
-----------------------|-------
Parameter              | Value
L_mm                   | 15
CenterFrac             | 0.1
Hist_BinCount          | 20
MinPeakSeparation_mm   | 0        ← ★新規追加
```

**追加コード（EnsureSheets 内）:**
```vba
' ver3.4 で追加
If .Range("A5").Value = "" Then .Range("A5").Value = "MinPeakSeparation_mm"
If .Range("B5").Value = "" Then .Range("B5").Value = 0
```

---

## 2. RunEdgePeakAnalysis メイン処理

### ver3.3 - パラメータ読み込み
```vba
Dim Lmm As Double, centerFrac As Double, binCount As Long
Lmm = CDbl(GetConfigValue("Config", "B2", 15#))
centerFrac = CDbl(GetConfigValue("Config", "B3", 0.1))
binCount = CLng(GetConfigValue("Config", "B4", 20))
```

### ver3.4 - パラメータ読み込み ★変更
```vba
Dim Lmm As Double, centerFrac As Double, binCount As Long, minPeakSep As Double
Lmm = CDbl(GetConfigValue("Config", "B2", 15#))
centerFrac = CDbl(GetConfigValue("Config", "B3", 0.1))
binCount = CLng(GetConfigValue("Config", "B4", 20))
minPeakSep = CDbl(GetConfigValue("Config", "B5", 0#))  ' ★新規

' ★新規: 負の値を0に丸める
If minPeakSep < 0 Then minPeakSep = 0#
```

### ver3.3 - ピーク検出
```vba
Dim xL As Double, yL As Double, xR As Double, yR As Double

If Not MaxYInRange(xArr, yArr, leftMin, leftMax, xL, yL) Then
    Err.Raise vbObjectError + 2005, , "左端部範囲に点がありません。"
End If

If Not MaxYInRange(xArr, yArr, rightMin, rightMax, xR, yR) Then
    Err.Raise vbObjectError + 2006, , "右端部範囲に点がありません。"
End If

Dim hL As Double, hR As Double
hL = (yL - baseline) / baseline
hR = (yR - baseline) / baseline
```

### ver3.4 - ピーク検出 ★変更
```vba
' ★新規: Peak1/Peak2 用の変数
Dim xL1 As Double, yL1 As Double, xL2 As Variant, yL2 As Variant
Dim xR1 As Double, yR1 As Double, xR2 As Variant, yR2 As Variant

' ★新規: Top2YInRange 呼び出し
If Not Top2YInRange(xArr, yArr, leftMin, leftMax, minPeakSep, xL1, yL1, xL2, yL2) Then
    Err.Raise vbObjectError + 2005, , "左端部範囲に点がありません。"
End If

If Not Top2YInRange(xArr, yArr, rightMin, rightMax, minPeakSep, xR1, yR1, xR2, yR2) Then
    Err.Raise vbObjectError + 2006, , "右端部範囲に点がありません。"
End If

' Peak1 の無次元化高さ
Dim hL As Double, hR As Double
hL = (yL1 - baseline) / baseline
hR = (yR1 - baseline) / baseline

' ★新規: Peak2 の無次元化高さ（存在する場合のみ）
Dim hL2 As Variant, hR2 As Variant
If Not IsEmpty(xL2) Then
    hL2 = (CDbl(yL2) - baseline) / baseline
Else
    hL2 = Empty
End If

If Not IsEmpty(xR2) Then
    hR2 = (CDbl(yR2) - baseline) / baseline
Else
    hR2 = Empty
End If

' ★新規: PeakStatus の判定
Dim peakStatus As String
If Not IsEmpty(xL2) And Not IsEmpty(xR2) Then
    peakStatus = "OK_2PEAK"
ElseIf Not IsEmpty(xL2) Or Not IsEmpty(xR2) Then
    peakStatus = "WARN_1PEAK"
Else
    peakStatus = "WARN_1PEAK"
End If
```

### ver3.3 - 結果書き込み
```vba
AppendResultEx runTime, filePath, Lmm, centerFrac, baseline, _
               xL, yL, hL, xR, yR, hR, "OK", ""

chartIndex = chartIndex + 1
AddProfileChartToChartsSheet chartIndex, Dir(filePath), xArr, yArr, baseline, _
                             xL, yL, xR, yR, Lmm, centerFrac
```

### ver3.4 - 結果書き込み ★変更
```vba
' ★新規: AppendResultEx2（拡張版）
AppendResultEx2 runTime, filePath, Lmm, centerFrac, baseline, _
                xL1, yL1, hL, xR1, yR1, hR, _
                xL2, yL2, hL2, xR2, yR2, hR2, _
                "OK", "", peakStatus

chartIndex = chartIndex + 1
' ★新規: AddProfileChartToChartsSheet2（拡張版）
AddProfileChartToChartsSheet2 chartIndex, Dir(filePath), xArr, yArr, baseline, _
                               xL1, yL1, xR1, yR1, xL2, yL2, xR2, yR2, _
                               Lmm, centerFrac
```

---

## 3. Result シート列定義

### ver3.3
```
A: Datetime
B: File
C: L_mm
D: CenterFrac
E: Baseline_um
F: x_L_mm
G: yPeak_L_um
H: h_L_(y-baseline)/baseline
I: x_R_mm
J: yPeak_R_um
K: h_R_(y-baseline)/baseline
L: Status
M: Error
```
**合計: 13列（A～M）**

### ver3.4 ★変更
```
A: Datetime
B: File
C: L_mm
D: CenterFrac
E: Baseline_um
F: x_L_mm
G: yPeak_L_um
H: h_L_(y-baseline)/baseline
I: x_R_mm
J: yPeak_R_um
K: h_R_(y-baseline)/baseline
L: Status
M: Error
N: x_L2_mm                         ← ★新規
O: yPeak_L2_um                     ← ★新規
P: h_L2_(y-baseline)/baseline      ← ★新規
Q: x_R2_mm                         ← ★新規
R: yPeak_R2_um                     ← ★新規
S: h_R2_(y-baseline)/baseline      ← ★新規
T: PeakStatus                      ← ★新規
```
**合計: 20列（A～T）**

---

## 4. 新規関数: Top2YInRange

### 関数シグネチャ
```vba
Private Function Top2YInRange(ByRef xArr() As Double, ByRef yArr() As Double, _
                               ByVal xMin As Double, ByVal xMax As Double, _
                               ByVal minSep As Double, _
                               ByRef x1 As Double, ByRef y1 As Double, _
                               ByRef x2 As Variant, ByRef y2 As Variant) As Boolean
```

### アルゴリズム
```vba
' Step 1: Peak1（最大値）を検出
For i = LBound(xArr) To UBound(xArr)
    If xArr(i) >= xMin And xArr(i) <= xMax Then
        If (Not found) Or (yArr(i) > bestY) Then
            bestY = yArr(i)
            bestX = xArr(i)
            found = True
        End If
    End If
Next i

If Not found Then
    Top2YInRange = False  ' 点がない
    Exit Function
End If

x1 = bestX
y1 = bestY

' Step 2: Peak2（Peak1から minSep 以上離れた最大値）を検出
For i = LBound(xArr) To UBound(xArr)
    If xArr(i) >= xMin And xArr(i) <= xMax Then
        If Abs(xArr(i) - x1) >= minSep Then  ' ★距離条件
            If (Not found2) Or (yArr(i) > best2Y) Then
                best2Y = yArr(i)
                best2X = xArr(i)
                found2 = True
            End If
        End If
    End If
Next i

If found2 Then
    x2 = best2X
    y2 = best2Y
Else
    x2 = Empty  ' ★Peak2なし
    y2 = Empty
End If

Top2YInRange = True
```

---

## 5. AppendResultEx → AppendResultEx2

### ver3.3 - AppendResultEx
```vba
Private Sub AppendResultEx( _
    ByVal runTime As Date, _
    ByVal filePath As String, _
    ByVal Lmm As Double, ByVal centerFrac As Double, _
    ByVal baseline As Variant, _
    ByVal xL As Variant, ByVal yL As Variant, ByVal hL As Variant, _
    ByVal xR As Variant, ByVal yR As Variant, ByVal hR As Variant, _
    ByVal status As String, ByVal errMsg As String)

    ' ... A～M 列のみ書き込み
End Sub
```

### ver3.4 - AppendResultEx2 ★新規
```vba
Private Sub AppendResultEx2( _
    ByVal runTime As Date, _
    ByVal filePath As String, _
    ByVal Lmm As Double, ByVal centerFrac As Double, _
    ByVal baseline As Variant, _
    ByVal xL As Variant, ByVal yL As Variant, ByVal hL As Variant, _
    ByVal xR As Variant, ByVal yR As Variant, ByVal hR As Variant, _
    ByVal xL2 As Variant, ByVal yL2 As Variant, ByVal hL2 As Variant, _  ' ★追加
    ByVal xR2 As Variant, ByVal yR2 As Variant, ByVal hR2 As Variant, _  ' ★追加
    ByVal status As String, ByVal errMsg As String, ByVal peakStatus As String)  ' ★追加

    ' ... A～M 列書き込み（従来通り）

    If status = "OK" Then
        ' ... E～K 列書き込み

        ' ★新規: Peak2 データ（N～S 列）
        If Not IsEmpty(xL2) Then
            ws.Cells(r, 14).Value = xL2  ' N
            ws.Cells(r, 15).Value = yL2  ' O
            ws.Cells(r, 16).Value = hL2  ' P
        End If

        If Not IsEmpty(xR2) Then
            ws.Cells(r, 17).Value = xR2  ' Q
            ws.Cells(r, 18).Value = yR2  ' R
            ws.Cells(r, 19).Value = hR2  ' S
        End If

        ws.Cells(r, 20).Value = peakStatus  ' T
    Else
        ' ERROR時: E～K と N～S をクリア
        ws.Range(ws.Cells(r, 5), ws.Cells(r, 11)).ClearContents
        ws.Range(ws.Cells(r, 14), ws.Cells(r, 20)).ClearContents  ' ★追加
    End If
End Sub
```

---

## 6. Charts - マーカー追加

### ver3.3 - AddProfileChartToChartsSheet
```vba
Private Sub AddProfileChartToChartsSheet( _
    ByVal chartIndex As Long, ByVal fileName As String, _
    ByRef xArr() As Double, ByRef yArr() As Double, _
    ByVal baseline As Double, _
    ByVal xL As Double, ByVal yL As Double, _
    ByVal xR As Double, ByVal yR As Double, _
    ByVal Lmm As Double, ByVal centerFrac As Double)

    ' ... Profile, Baseline 追加

    ' Left Peak marker
    Dim sl As Series
    Set sl = ch.SeriesCollection.NewSeries
    sl.name = "LeftPeak"
    sl.XValues = Array(xL)
    sl.Values = Array(yL)

    ' Right Peak marker
    Dim sr As Series
    Set sr = ch.SeriesCollection.NewSeries
    sr.name = "RightPeak"
    sr.XValues = Array(xR)
    sr.Values = Array(yR)
End Sub
```

### ver3.4 - AddProfileChartToChartsSheet2 ★変更
```vba
Private Sub AddProfileChartToChartsSheet2( _
    ByVal chartIndex As Long, ByVal fileName As String, _
    ByRef xArr() As Double, ByRef yArr() As Double, _
    ByVal baseline As Double, _
    ByVal xL1 As Double, ByVal yL1 As Double, _           ' ★変数名変更
    ByVal xR1 As Double, ByVal yR1 As Double, _           ' ★変数名変更
    ByVal xL2 As Variant, ByVal yL2 As Variant, _         ' ★追加
    ByVal xR2 As Variant, ByVal yR2 As Variant, _         ' ★追加
    ByVal Lmm As Double, ByVal centerFrac As Double)

    ' ... Profile, Baseline 追加（従来通り）

    ' Left Peak1 marker
    Dim sl As Series
    Set sl = ch.SeriesCollection.NewSeries
    sl.name = "LeftPeak"
    sl.XValues = Array(xL1)  ' ★変数名変更
    sl.Values = Array(yL1)   ' ★変数名変更

    ' Right Peak1 marker
    Dim sr As Series
    Set sr = ch.SeriesCollection.NewSeries
    sr.name = "RightPeak"
    sr.XValues = Array(xR1)  ' ★変数名変更
    sr.Values = Array(yR1)   ' ★変数名変更

    ' ★新規: Left Peak2 marker（存在する場合のみ）
    If Not IsEmpty(xL2) Then
        Dim sl2 As Series
        Set sl2 = ch.SeriesCollection.NewSeries
        sl2.name = "LeftPeak2"
        sl2.XValues = Array(CDbl(xL2))
        sl2.Values = Array(CDbl(yL2))
        sl2.ChartType = xlXYScatter
        sl2.MarkerStyle = xlMarkerStyleCircle
        sl2.MarkerSize = 5
    End If

    ' ★新規: Right Peak2 marker（存在する場合のみ）
    If Not IsEmpty(xR2) Then
        Dim sr2 As Series
        Set sr2 = ch.SeriesCollection.NewSeries
        sr2.name = "RightPeak2"
        sr2.XValues = Array(CDbl(xR2))
        sr2.Values = Array(CDbl(yR2))
        sr2.ChartType = xlXYScatter
        sr2.MarkerStyle = xlMarkerStyleCircle
        sr2.MarkerSize = 5
    End If
End Sub
```

---

## 7. 変更なし（互換性維持）

以下の機能は **ver3.3 と完全に同じ** です：

### CSV 読み込み
- `ReadCsvXY`
- `ReadAllTextRobust`
- `TryParseXY`
- `TryParseDoubleLocale`

### 数学関数
- `MeanYInRange`
- `MaxYInRange` ★残存（後方互換性のため）
- `QuickSortXY`

### Histogram
- `BuildHLHRHistogramsLatest` ★Peak1（h_L, h_R）のみ使用
- `PrepareHistSheet`
- `BuildHistogramChart`
- `CalcStatsHLHR`
- `ToVariant1D_FromDoubleArrayHLHR`

### その他
- `EnsureSheetExists`
- `GetConfigValue`
- `PickCsvFiles`

---

## 8. エラーハンドリングの違い

### ver3.3
```vba
' Peak が取れない場合 → ERROR
If Not MaxYInRange(..., xL, yL) Then
    Err.Raise vbObjectError + 2005, , "左端部範囲に点がありません。"
End If
```

### ver3.4
```vba
' Peak1 が取れない場合 → ERROR（同じ）
If Not Top2YInRange(..., xL1, yL1, xL2, yL2) Then
    Err.Raise vbObjectError + 2005, , "左端部範囲に点がありません。"
End If

' Peak2 が取れない場合 → ★ERRORにしない
' → xL2/yL2 が Empty になるだけ
' → PeakStatus = "WARN_1PEAK" として処理継続
```

---

## 9. 使用例

### MinPeakSeparation_mm = 0 の場合
```
範囲内のデータ:
x:  1.0,  2.0,  3.0,  4.0,  5.0
y: 10.0, 15.0,  8.0, 12.0,  9.0

結果:
Peak1: x=2.0, y=15.0（最大値）
Peak2: x=4.0, y=12.0（2番目に大きい値、Peak1と異なる位置）
```

### MinPeakSeparation_mm = 3.0 の場合
```
範囲内のデータ:
x:  1.0,  2.0,  3.0,  4.0,  5.0
y: 10.0, 15.0,  8.0, 12.0,  9.0

結果:
Peak1: x=2.0, y=15.0（最大値）
Peak2: x=5.0, y=9.0（|5.0-2.0|=3.0 ≥ 3.0 を満たす点の中で最大）
       ※ x=4.0 は |4.0-2.0|=2.0 < 3.0 なので除外
```

### MinPeakSeparation_mm = 10.0（大きすぎる場合）
```
範囲内のデータ:
x:  1.0,  2.0,  3.0,  4.0,  5.0
y: 10.0, 15.0,  8.0, 12.0,  9.0

結果:
Peak1: x=2.0, y=15.0（最大値）
Peak2: Empty（条件を満たす点がない）
PeakStatus: "WARN_1PEAK"
```

---

## まとめ

### 主な変更箇所（コード行数）
| 項目 | 変更内容 | 行数 |
|------|----------|------|
| Config 初期化 | MinPeakSeparation_mm 追加 | +2 |
| Result ヘッダ | N～T 列追加 | +7 |
| RunEdgePeakAnalysis | パラメータ読み込み、ピーク検出、結果書き込み | +40 |
| Top2YInRange | 新規関数 | +60 |
| AppendResultEx2 | 新規関数（拡張版） | +30 |
| AddProfileChartToChartsSheet2 | 新規関数（拡張版） | +30 |
| **合計** | | **約 170 行** |

### 後方互換性
- **Config A2～A4**: 変更なし
- **Result A～M 列**: 変更なし
- **Hist シート**: Peak1 データを使用（変更なし）
- **CSV 読み込み**: 変更なし

### 移行方法
1. ver3.3 のデータは ver3.4 でそのまま読める
2. ver3.4 で作成したデータの A～M 列は ver3.3 互換
3. N～T 列は ver3.3 では単に無視される

### テスト推奨項目
- [ ] MinPeakSeparation_mm = 0 で実行
- [ ] MinPeakSeparation_mm = 2.0 で実行
- [ ] 1ピークのみのデータで実行（WARN_1PEAK 確認）
- [ ] 2ピークのデータで実行（OK_2PEAK 確認）
- [ ] Charts で Peak2 マーカー表示確認
- [ ] Hist が Peak1 ベースで正常動作確認
