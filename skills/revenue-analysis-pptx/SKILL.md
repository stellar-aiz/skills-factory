---
name: revenue-analysis-pptx
description: >
  売上分析ー売上高・EBITDAの推移（Revenue Analysis）のPowerPointスライドを生成するスキル。
  BDD（ビジネスデュー・ディリジェンス）において、対象会社の過去3年間＋今期着地見込みの
  売上高・EBITDA・EBITDA率をネイティブ棒グラフ＋Shape描画の折れ線＋CAGR注釈で
  1枚のスライドにまとめる。上部にEBITDA率の折れ線、下部に売上高の棒グラフ＋CAGRの構成。
  Excelファイルから財務データを読み取り、自動でスライドを生成できる。
  左側Y軸は売上高を表示し、対象会社の売上高規模に応じて適切なレンジを自動決定する。
  company-history-pptx / customer-profile-pptx と同じテンプレート・フォントを使用。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「売上分析」「売上高・EBITDAの推移」「売上推移」「Revenue Analysis」という言葉が出た場合
  - 「売上高とEBITDAをスライドにして」「売上の推移をパワポで」「財務推移のスライド」という要望
  - 「売上高・営業利益の推移」「売上高推移チャート」「業績推移スライド」という要望
  - BDD（ビジネスDD）の文脈で対象会社の売上高・EBITDA推移のスライド化を求められた場合
  - ユーザーがExcelファイルや財務データを提供して、売上分析スライドの作成を求めた場合
  - 「対象会社の財務推移をスライドにまとめて」「EBITDAマージンの推移」という要望
  - 「CAGR付きの売上チャート」「売上高の成長率をスライドに」という要望
supported_brands: [stellar_aiz, roleup]

---

# 売上分析 PowerPoint ジェネレーター

BDD（ビジネスデュー・ディリジェンス）において、対象会社の過去3年間＋今期着地見込みの
売上高・EBITDA・EBITDA率を1枚のPowerPointスライドに整理するスキル。

---

## スライド構成

| セクション | 位置 | 内容 |
|---|---|---|
| **メインメッセージ** | 最上部 | 最大70文字。「〜すべき」で締める |
| **チャートタイトル** | メインメッセージ直下 | 「売上分析ー売上高・EBITDAの推移」 |
| **単位表記・凡例** | チャート上部 | 左: 単位、右: 凡例（■売上高 ●━EBITDA率） |
| **EBITDA率（上部）** | チャート最上部 | 折れ線＋マーカー＋データラベル（Shape描画） |
| **CAGR注釈（中部）** | 棒の上端付近 | 矢印＋楕円で売上高CAGRを表示 |
| **売上高（下部）** | チャート下部 | ネイティブ棒グラフ（左Y軸に目盛表示） |
| **出典** | 左下 | 情報ソースの記載 |

### レイアウトの3層構造

```
┌─────────────────────────────────────┐
│  EBITDA率 折れ線 + ラベル（上部）    │  ← Shape描画
│                                     │
│  CAGR矢印 + 楕円（中部）           │  ← Shape描画
│  ┌───┐  ┌───┐  ┌───┐  ┌───┐       │
│  │   │  │   │  │   │  │   │       │  ← 棒グラフ（ネイティブ）
│  │   │  │   │  │   │  │   │       │
│  └───┘  └───┘  └───┘  └───┘       │
└─────────────────────────────────────┘
```

### 年度表記ルール

- **通常年度**: `yy/mm期` 形式（例: "22/3期", "23/3期"）
- **最終年度（今期着地見込み）**: `yy/mm期(見込み)` 形式（例: "25/3期(見込み)"）
- DDの対象期間は通常、過去3年＋今期見込みの4年分

### Y軸レンジの自動決定ルール

売上高の最大値に基づき、適切な軸レンジを自動計算する：
- 最大値の1.3倍を軸上限候補とする
- 「切りの良い数値」に丸める（10, 50, 100, 500, 1000, 5000, 10000単位等）
- 目盛間隔は軸上限を5等分した値

---

## 入力パターンと処理フロー

### パターンA：Excelファイルが入力された場合

1. **Step 1: Excelファイルを読み取る**（openpyxlまたはpandasで読み込み）
2. **Step 2: 売上高・EBITDAデータを抽出・整理**し、Markdownでユーザーに提示
3. **Step 3: ユーザーの承認後**、JSONを作成してPowerPointを生成

### パターンB：データが整理済みで入力された場合

1. **Step 1: 内容を確認**し、年度・売上高・EBITDAが揃っているか確認
2. **Step 2: 確認後、JSONを作成してPowerPointを生成**

### 年度表記の変換ルール

Excelから読み取った年度を `yy/mm期` 形式に変換する：
- "2022年3月期" → "22/3期"
- "FY2022" → "22/3期"（決算月がわかれば付与）
- 最終年度（着地見込み）は末尾に "(見込み)" を付与

---

## JSONデータ仕様

`{{WORK_DIR}}/revenue_analysis_data.json` に以下の形式で保存する：

```json
{
  "source": "出典：対象会社提供資料",
  "main_message": "対象会社は過去3年間で売上高CAGRx.x%と安定成長しており、収益基盤を評価すべき",
  "chart_title": "売上分析ー売上高・EBITDAの推移",
  "unit_label": "（単位：百万円、%）",
  "bar_label": "売上高",
  "line_label": "EBITDA率",
  "data": [
    {"year": "22/3期", "revenue": 5000, "ebitda": 800},
    {"year": "23/3期", "revenue": 5200, "ebitda": 850},
    {"year": "24/3期", "revenue": 5500, "ebitda": 900},
    {"year": "25/3期(見込み)", "revenue": 5800, "ebitda": 950}
  ]
}
```

### JSONフィールド仕様

| フィールド | 型 | 必須 | 説明 |
|---|---|---|---|
| `source` | string | 任意 | 出典テキスト（左下に表示） |
| `main_message` | string | 必須 | メインメッセージ。最大70文字、「〜すべき」で締める |
| `chart_title` | string | 任意 | チャートタイトル。デフォルト「売上分析ー売上高・EBITDAの推移」 |
| `unit_label` | string | 任意 | 単位表記 |
| `bar_label` | string | 任意 | 棒グラフ系列名。デフォルト「売上高」 |
| `line_label` | string | 任意 | 折れ線系列名。デフォルト「EBITDA率」 |
| `data` | array | 必須 | 年度別データの配列（3〜4要素を想定） |
| `data[].year` | string | 必須 | 年度ラベル（例: "22/3期", "25/3期(見込み)"） |
| `data[].revenue` | number | 必須 | 売上高 |
| `data[].ebitda` | number | 必須 | EBITDA |

EBITDA率は `ebitda / revenue * 100` で自動計算される。

---

## スクリプト実行コマンド

```bash
pip install python-pptx openpyxl -q --break-system-packages

# 推奨: brand 指定で起動（template は brand_resolver で自動解決）
python <SKILL_DIR>/scripts/fill_revenue_analysis.py \
  --brand stellar_aiz \
  --data {{WORK_DIR}}/revenue_analysis_data.json \
  --output {{OUTPUT_DIR}}/RevenueAnalysis_output.pptx

# Roleup 出力
python <SKILL_DIR>/scripts/fill_revenue_analysis.py \
  --brand roleup \
  --data {{WORK_DIR}}/revenue_analysis_data.json \
  --output {{OUTPUT_DIR}}/RevenueAnalysis_output.pptx
```

`--template` を明示指定すると brand 解決を上書きできる（任意）。
`<SKILL_DIR>` は実際のスキルインストールパスに置き換えること。

---

## デザイン仕様

### 色

| 要素 | カラーコード |
|---|---|
| テキスト | #333333 |
| 棒グラフ（売上高） | #4E79A7 |
| 折れ線・マーカー（EBITDA率） | #003366 |
| CAGR矢印・楕円 | #333333 |
| 出典テキスト | #666666 |

### レイアウト

| 要素 | 値 |
|---|---|
| チャート開始X | 0.80in |
| チャート幅 | 11.73in |
| チャート開始Y | 1.70in |
| チャート高さ | 5.20in |
| EBITDA率ゾーン | 軸上限の90-98%領域 |
| CAGR矢印 | 棒上端の0.35in上 |

---

## アセット

| ファイル名 | 用途 |
|---|---|
| `assets/stellar_aiz/revenue-analysis-template.pptx` | Stella 16:9 用テンプレート |
| `assets/stellar_aiz/layout.json` | Stella 用レイアウト座標 |
| `assets/roleup/revenue-analysis-template.pptx` | Roleup A4 横用テンプレート（cp roleup template から派生） |
| `assets/roleup/layout.json` | Roleup 用レイアウト座標 |

## スクリプト

| ファイル名 | 用途 |
|---|---|
| `scripts/fill_revenue_analysis.py` | JSONデータからPPTXを生成するスクリプト |

## 参考

| ファイル名 | 内容 |
|---|---|
| `references/sample_data.json` | サンプルJSONデータ |
