# 端部ピーク解析 ver3.4 - 変更概要

## バージョン情報
- **バージョン**: 3.4
- **ベース**: ver3.3
- **主な変更**: 左右端部のピーク検出を1ピーク→最大2ピークに拡張

## 主な変更点

### 1. Config シート - 新パラメータ追加

**新規追加パラメータ:**
- `Config!A5`: "MinPeakSeparation_mm"
- `Config!B5`: 数値入力（Double）、デフォルト値 0.0

**説明:**
- 2つのピークを区別するための最小距離（mm単位）
- 0.0 の場合、ピーク間の距離制約なし
- 負の値が入力された場合は自動的に 0 に丸められます

**使用例:**
```
MinPeakSeparation_mm = 2.0
→ 第1ピークから2.0mm以上離れた位置にある最大値を第2ピークとして検出
```

### 2. ピーク検出ロジック - Top2YInRange 関数

**新関数:**
```vba
Function Top2YInRange(xArr(), yArr(), xMin As Double, xMax As Double, minSep As Double, _
                      ByRef x1 As Double, ByRef y1 As Double, _
                      ByRef x2 As Variant, ByRef y2 As Variant) As Boolean
```

**仕様:**
1. **Peak1（第1ピーク）の検出:**
   - 指定範囲内で y 値が最大の点を検出
   - 従来の MaxYInRange と同じアルゴリズム
   - 同じy値の場合は最初に見つかった点を選択

2. **Peak2（第2ピーク）の検出:**
   - Peak1 確定後、同じ範囲内で以下の条件を満たす点を検索:
     - `abs(x - x1) >= minSep` （Peak1からの距離が minSep 以上）
     - その中で y 値が最大の点
   - 条件を満たす点がない場合: `x2 = Empty`, `y2 = Empty`

3. **戻り値:**
   - `True`: 範囲内に少なくとも1点あり、Peak1が検出できた
   - `False`: 範囲内に点がない（エラーケース）

**注意:**
- 既存の MaxYInRange 関数は後方互換性のため残してあります

### 3. Result シート - 列の拡張

**既存列（A～M）:** そのまま維持

**新規追加列（N～T）:**
| 列 | ヘッダ名 | 内容 |
|----|----------|------|
| N | x_L2_mm | 左側第2ピークのx座標（mm） |
| O | yPeak_L2_um | 左側第2ピークのy値（μm） |
| P | h_L2_(y-baseline)/baseline | 左側第2ピークの無次元化高さ |
| Q | x_R2_mm | 右側第2ピークのx座標（mm） |
| R | yPeak_R2_um | 右側第2ピークのy値（μm） |
| S | h_R2_(y-baseline)/baseline | 右側第2ピークの無次元化高さ |
| T | PeakStatus | ピーク検出ステータス |

**PeakStatus の値:**
- `OK_2PEAK`: 左右両側で2ピーク検出成功
- `WARN_1PEAK`: 左右どちらか（または両方）で1ピークのみ
- （空白）: ERROR ケースの場合

**データ格納ルール:**
- Peak2 が存在しない場合: N～S 列は空欄
- Peak2 が存在する場合: h_L2/h_R2 を baseline で無次元化して格納
- Status="ERROR" の場合: E～K列および N～S列をすべて空欄化

### 4. Charts シート - マーカー追加

**既存マーカー:**
- Profile（プロファイル線）
- Baseline（ベースライン）
- LeftPeak（左側第1ピーク）
- RightPeak（右側第1ピーク）

**新規追加マーカー:**
- **LeftPeak2**: 左側第2ピーク（存在する場合のみ）
- **RightPeak2**: 右側第2ピーク（存在する場合のみ）

**マーカースタイル:**
- 形状: Circle（○）
- サイズ: 5
- 色: Excel デフォルト（自動割り当て）

### 5. 内部関数の変更

#### AppendResultEx2（新関数）
- AppendResultEx を拡張した新バージョン
- Peak2 データ（xL2, yL2, hL2, xR2, yR2, hR2）を追加パラメータとして受け取る
- PeakStatus を T列に出力

#### AddProfileChartToChartsSheet2（新関数）
- AddProfileChartToChartsSheet を拡張した新バージョン
- Peak2 座標（xL2, yL2, xR2, yR2）を追加パラメータとして受け取る
- Peak2 が Empty でない場合のみマーカーを追加

## 互換性と既存機能の維持

### 変更なし（ver3.3 の仕様を維持）
1. **CSV 読み込み:**
   - BOM対応、複数区切り文字対応、ロケール対応
   - ReadCsvXY, TryParseXY, TryParseDoubleLocale

2. **Baseline 計算:**
   - 中央部分の平均値計算
   - MeanYInRange

3. **Hist シート:**
   - hL/hR のヒストグラム生成
   - 統計値（Mean/Std/Min/Max/p95/p99）
   - **重要**: Hist は Peak1（h_L, h_R）のデータを使用

4. **エラーハンドリング:**
   - Peak1 が取れない場合: ファイル単位で ERROR
   - Peak2 が取れない場合: ERROR にせず PeakStatus で警告

5. **その他:**
   - QuickSortXY（ソート処理）
   - PickCsvFiles（ファイル選択）
   - 50ファイル一括処理

## 使用方法

### 1. セットアップ
1. Excel ファイルを開く
2. VBA エディタ（Alt+F11）を起動
3. 既存の ver3.3 モジュールを削除（または無効化）
4. `コードver3.4_2peaks.bas` をインポート

### 2. Config シート設定
```
Parameter              | Value
-----------------------|-------
L_mm                   | 15      （端部解析範囲）
CenterFrac             | 0.1     （中央部分の割合）
Hist_BinCount          | 20      （ヒストグラムのビン数）
MinPeakSeparation_mm   | 2.0     （ピーク間最小距離、新規）
```

### 3. 実行
1. マクロ `RunEdgePeakAnalysis` を実行
2. CSV ファイルを選択（最大50件）
3. 処理完了後、以下のシートを確認:
   - **Result**: 解析結果（Peak1/Peak2 データ）
   - **Charts**: プロファイルグラフ（Peak1/Peak2 マーカー付き）
   - **Hist**: hL/hR ヒストグラム（Peak1 ベース）

## テストシナリオ

### シナリオ1: 1ピークのみのデータ
**期待結果:**
- Peak1: 正常検出
- Peak2: N～S列は空欄
- PeakStatus: "WARN_1PEAK"
- 処理: 正常完了（エラーにならない）

### シナリオ2: 2ピークが離れているデータ
**条件:** MinPeakSeparation_mm = 2.0
**期待結果:**
- Peak1: 最大ピーク検出
- Peak2: Peak1 から 2.0mm 以上離れた次の最大ピーク検出
- PeakStatus: "OK_2PEAK"
- Charts: 4つのマーカー表示（LeftPeak, LeftPeak2, RightPeak, RightPeak2）

### シナリオ3: MinPeakSeparation を大きくする
**条件:** MinPeakSeparation_mm = 10.0
**期待結果:**
- Peak2 が抑止される（条件を満たす点がない）
- N～S列は空欄
- PeakStatus: "WARN_1PEAK"

### シナリオ4: MinPeakSeparation = 0
**期待結果:**
- Peak1: 最大ピーク
- Peak2: Peak1 と同じ位置以外の最大ピーク（実質的に2番目に大きいピーク）

## トラブルシューティング

### Q1: Peak2 が検出されない
**確認項目:**
1. MinPeakSeparation_mm の値が大きすぎないか？
2. データに本当に2つのピークがあるか？
3. Result シート T列の PeakStatus を確認

### Q2: 既存の ver3.3 との互換性
**対応:**
- Result シートの A～M 列は完全互換
- Hist シートは Peak1（h_L, h_R）を使用するため互換
- Charts は Peak2 マーカーが追加されるだけ

### Q3: Config!B5 が空欄の場合
**動作:**
- 自動的に 0.0 がデフォルト値として使用されます
- エラーにはなりません

## ファイル構成

```
src/modules/
├── コードver3.3.pdf          （旧バージョン、参照用）
├── コードver3.4_2peaks.bas   （新バージョン、★使用推奨）
└── README_ver3.4.md          （このファイル）
```

## 変更履歴

### ver3.4 (2026-01-17)
- 左右端部のピーク検出を1ピーク→最大2ピークに拡張
- Config に MinPeakSeparation_mm パラメータ追加
- Result シートに N～T 列追加（Peak2 データ、PeakStatus）
- Charts に LeftPeak2/RightPeak2 マーカー追加
- Top2YInRange 関数追加
- AppendResultEx2, AddProfileChartToChartsSheet2 関数追加

### ver3.3
- ベースバージョン（1ピーク検出）
- CSV読み込み、baseline、Hist、Charts の基本機能

## ライセンス・著作権

このコードは既存の ver3.3 をベースに拡張したものです。

## サポート

問題や質問がある場合は、プロジェクトの Issue トラッカーに報告してください。
