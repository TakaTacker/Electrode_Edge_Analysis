# 要件定義書（更新版）：端部形状解析ツール
## Excel VBA ver3.3 → Codex 移行

---

## 1. 目的・背景

幅方向プロファイル（`x=位置[mm]`, `y=厚み[µm]`）をCSVから読み込み、左右端部のピーク厚みを中央部の基準厚みに対して無次元化した指標 **hL/hR** として算出し、複数ファイルをバッチで解析・可視化・集計する。

---

## 2. スコープ

### 2.1 対象（In Scope）

- **CSVファイル処理**
  - CSV（最大50ファイル）をユーザーが選択し、一括解析を実行
  - 各CSVについて以下を計算し、結果を蓄積（OK/ERROR混在対応）

- **計算項目**
  - `Baseline`：中央部平均厚み
  - `左端部ピーク`：最大yの位置/値
  - `右端部ピーク`：最大yの位置/値
  - `無次元端部突出度`：hL/hR

- **出力**
  - **Charts出力**：全プロファイルをグリッド状に描画（最大50）
  - **Hist出力**：最新OK runの hL/hR 分布ヒストグラムを作成
  - **統計表示**：Mean、Std、Min、Max、p95、p99を表示
  - **エラー処理**：対象ファイルのみERRORとして記録し、処理は継続

### 2.2 対象外（Out of Scope）

- セキュリティ・ブラウザ互換・監査要件
- UI カスタマイズ（通知音など）
- PDF自動出力
- 解析ルールの可変化（MVP は固定ルール）

---

## 3. 入力仕様

### 3.1 入力データ形式（CSV）

| 項目 | 仕様 | 備考 |
|------|------|------|
| ファイル単位 | 1ファイル = 1本の幅方向プロファイル | - |
| **A列** | `x（mm）` | 位置情報 |
| **B列** | `y（µm）` | 厚み情報 |
| ヘッダ | 1行目は必ずヘッダ。データは2行目以降 | スキップ必須 |
| 区切り文字 | `,` / `;` / `TAB` に対応 | 自動判定推奨 |
| 文字コード | UTF-8（BOM有無両対応） | - |
| 小数点 | OS ロケール（`.` / `,`）差を許容 | ロケール対応必須 |
| 空行・欠損行 | 無視。x/yが数値としてパースできる行のみ採用 | 堅牢性重視 |

### 3.2 設定値（Config）

| パラメータ | 型 | デフォルト | 制約 |
|-----------|-----|-----------|------|
| `L_mm` | Double | 15 | L_mm > 0 |
| `CenterFrac` | Double | 0.1 | 0 < CenterFrac < 0.5 |
| `Hist_BinCount` | Long | 20 | 5～200に丸める |

**パラメータ説明：**
- **L_mm**：左右端部として扱う範囲幅（mm）
- **CenterFrac**：中央部の幅割合（全幅に対して±割合）
- **Hist_BinCount**：ヒストグラムのビン数

---

## 4. 出力仕様

### 4.1 Result（集計テーブル）

1行 = 1ファイルの解析結果（OK/ERROR を含む）  
ファイル識別は**ファイル名のみ**（フルパスは保持しない）

> **注記**：同名ファイルが混在する運用は想定しない（必要なら拡張要件）

| 列 | 見出し | 型 | 内容 |
|----|---------|----|------|
| A | Datetime | Date | 実行時刻（runTime） |
| B | File | String | ファイル名のみ |
| C | L_mm | Double | Config値 |
| D | CenterFrac | Double | Config値 |
| E | Baseline_um | Double | 中央部平均y（µm） |
| F | x_L_mm | Double | 左端部最大値のx |
| G | yPeak_L_um | Double | 左端部最大y |
| H | h_L_(y-baseline)/baseline | Double | hL = (yL - baseline) / baseline |
| I | x_R_mm | Double | 右端部最大値のx |
| J | yPeak_R_um | Double | 右端部最大y |
| K | h_R_(y-baseline)/baseline | Double | hR = (yR - baseline) / baseline |
| L | Status | String | "OK" または "ERROR" |
| M | Error | String | エラー理由（ERROR時のみ） |

> **ERROR時の処理**：E～Kは空欄とする

### 4.2 Charts（全プロファイル可視化）

- **レイアウト**：最大50チャートをグリッド配置（例：5列×10行）

**1チャートの表示要素**
- プロファイル（x vs y）折れ線
- Baseline水平線
- 左右ピーク点（マーカー）

**タイトル要素**
- ファイル名
- L_mm値
- CenterFrac値（中央部±）

### 4.3 Hist（分布＋統計）

**対象データ**  
Result内の最新OK runTimeに一致する行（Status="OK"）

**ヒストグラムテーブル**

| 列 | 内容 |
|-----|------|
| BinUpper | 各binの上限値 |
| Count_hL | hLのカウント |
| Count_hR | hRのカウント |

**付帯情報ブロック**
- RunTime / N / Min / Max / BinCount / BinWidth

**グラフ要素**
- hLヒストグラム（棒グラフ）
- hRヒストグラム（棒グラフ）

**統計量（hL/hR別）**

| 統計量 | 内容 |
|--------|------|
| Mean | 平均値 |
| Std | 標本標準偏差 |
| Min | 最小値 |
| Max | 最大値 |
| p95 | 95パーセンタイル |
| p99 | 99パーセンタイル |

**bin設計**
- bin境界は Min～Max を等幅で BinCount 分割
- BinUpper は各binの上限値
- Min/Max は hL と hR を**両方含む全体から決定**（左右を同一スケールで比較するため）

> **重要**：Excel FREQUENCYで発生し得る**overflow bin（上限超過）**は、現仕様では表示・集計対象から除外（=先頭 BinCount のみ出力）

---

## 5. 処理仕様（アルゴリズム）

### 5.1 全体フロー（バッチ実行）

1. Config/Result/Charts/Hist の存在確認と初期化
2. Charts/Hist は再生成のためクリア（既存グラフ削除）
3. Config値を取得（空ならデフォルト投入）
4. ファイル選択（最大50）
5. `runTime = Now` を取得
6. ファイルごとに解析（例外はファイル単位で捕捉して継続）
7. 最後にHistを生成（最新OK runに対して）
8. ERROR一覧をまとめて通知（あれば）

### 5.2 ファイル解析（1CSV あたり）

#### 前処理

- CSVを読み込み、数値行のみ抽出（1行目ヘッダは必ずスキップ）
- x昇順にソート（yを追随）

#### 範囲定義

```
xMin = min(x)
xMax = max(x)
width = xMax - xMin
xMid = (xMin + xMax) / 2

left range:   [xMin, xMin + L_mm]
right range:  [xMax - L_mm, xMax]
center range: [xMid - CenterFrac*width, xMid + CenterFrac*width]
```

#### Baseline（中央部平均）

```
baseline = mean(y in center range)
```

> **Baseline閾値**：`abs(baseline) < 1e-7` の場合は除算不可として ERROR

#### ピーク検出

- **左ピーク**：left range内の最大y（同値なら先に見つかった点で良い）
- **右ピーク**：right range内の最大y

#### 指標算出

```
hL = (yL - baseline) / baseline
hR = (yR - baseline) / baseline
```

#### 出力

- **OK**：Result に1行追加 ＋ Charts にチャート追加
- **ERROR**：Result に Status=ERROR ＋ Errorメッセージ記録（Charts追加なし）

---

## 6. エラーハンドリング要件（ファイル単位継続）

> **基本方針**：**全体停止はしない**。対象ファイルのみERRORとして記録し次へ進む

### エラー条件

以下の場合、該当ファイルをERRORとして記録：

- 有効な数値行が0
- データ点数が少なすぎる（目安：4点未満）
- x幅が0以下
- center範囲に点がない（baseline算出不可）
- baseline が閾値以下（`abs < 1e-7`）
- 左端部範囲に点がない
- 右端部範囲に点がない

### エラー情報の保存

Result の該当行に理由を保存し、最後に一覧表示する

---

## 7. 非機能要件

| 要件 | 内容 |
|------|------|
| ファイル処理数 | 最大50ファイルを処理可能 |
| 進捗表示 | StatusBar等で表示 |
| UI応答性 | DoEvents相当で適宜イベント処理 |
| 再実行性 | Charts/Hist はクリーンに再生成される |
| デフォルト値 | Config が空ならデフォルトを自動投入 |

---

## 8. Codex移行のためのモジュール分割案（推奨）

### 8.1 io_csv モジュール

**責務**：ロバストなCSV読み込み

```
robust read（delimiter/BOM/locale/空行/欠損行）
read_csv_xy(path) -> (x[], y[])
```

### 8.2 analysis モジュール

**責務**：解析処理の実装

```
sort_xy(x, y)
mean_y_in_range(y, x_min, x_max)
max_y_in_range(y, x_min, x_max)
compute_metrics(x, y, L_mm, CenterFrac) 
  -> (baseline, xL, yL, hL, xR, yR, hR)
```

### 8.3 batch_runner モジュール

**責務**：バッチ実行制御

```
pick_files(max=50)
run_batch(files, config) -> (results, errors)
latest_ok_run(results) -> runTime
```

### 8.4 reporting モジュール

**責務**：結果出力と可視化

```
write_result_table(results)
plot_profiles_grid(profiles)
build_histogram(hL, hR, binCount, min/max unified, overflow excluded)
stats(h) -> (mean, std, min, max, p95, p99)
```

---

## 9. 受け入れ基準（Acceptance Criteria）

以下のすべてを満たす必要がある：

- ✅ ver3.0互換のCSVが読み込める（BOM/区切り/小数点/空行/欠損耐性）
- ✅ 9ファイル以上のバッチ実行でも落ちず、ERRORは記録して継続する
- ✅ Result に OK/ERROR が正しく蓄積される
- ✅ Charts にプロファイル（baseline線＋左右ピーク点）が出る
- ✅ Hist に最新OK run の hL/hR ヒストグラムが出る
- ✅ Hist に hL/hR の統計値（Mean/Std/Min/Max/p95/p99）が出る
- ✅ Hist のbin は hL+hR全体の min/max で統一され、overflow binは表示しない

---

## 備考

この更新版を Codex に渡す場合、以下の形式に整形すると実装タスク（チケット）として直接活用可能：

- API仕様
- 関数シグネチャ
- テストケース表
- エッジケース一覧

移行開発をそのまま開始できる形での提供を推奨。
