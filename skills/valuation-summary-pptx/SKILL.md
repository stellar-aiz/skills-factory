---
name: valuation-summary-pptx
description: >
  バリュエーション・財務サマリー（Valuation Summary）のPowerPointスライドを生成するスキル。
  M&A DDの提案書・最終報告で使用する3種類のチャートをJSON側のchart_typeで切り替え可能:
  (1) フットボールフィールド（手法別バリュエーションレンジ比較）
  (2) 株式価値ブリッジ（EV→Equity Valueのウォーターフォール）
  (3) 財務サマリーテーブル（主要KPI一覧）
  全てPowerPointネイティブオブジェクトで作成。Accent2色を自動適用。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「バリュエーション」「Valuation」「株式価値」「事業価値」「EV」という言葉が出た場合
  - 「フットボールフィールド」「バリュエーションレンジ」「DCF」「マルチプル比較」という言葉が出た場合
  - 「株式価値ブリッジ」「エクイティブリッジ」「ウォーターフォール」「EV→Equity」という要望
  - 「財務サマリー」「Financial Summary」「DD財務分析」「主要財務指標」という要望
  - M&A DD提案書やDD最終報告書のバリュエーション・財務セクションのスライド作成を求められた場合
  - 「DD提案書のバリュエーションスライド」「DD最終報告のバリュエーション」という要望
supported_brands: [stellar_aiz, roleup]

---

# バリュエーション・財務サマリー PowerPoint ジェネレーター

M&A DDで使用するバリュエーション・財務分析のスライドを生成するスキル。
3種類のチャートタイプをJSON側の `chart_type` で切り替え可能。

---

## チャートタイプ一覧

| chart_type | 名称 | 用途 |
|---|---|---|
| `football_field` | フットボールフィールド | 手法別バリュエーションレンジ比較。DD提案書・最終報告の定番 |
| `equity_bridge` | 株式価値ブリッジ | EV→Equity Valueのウォーターフォール。DD最終報告で調整項目を説明 |
| `financial_summary` | 財務サマリーテーブル | 主要財務指標の時系列一覧。実績/計画の区別、CAGR表示対応 |

---

## 共通仕様

- **Main Message**: 最大70文字。必ず「〜すべき」で終える
- **Chart Title**: 10〜20文字程度
- **Source**: roleup では必須 (`source` フィールド、Source 3 placeholder に書き込み)
- **アクセント色**: stella=テンプレ accent2 自動抽出、roleup=brand `accent_revenue_bar` (#7C4C2C)
- **フォント**: stella=Meiryo / roleup=Yu Gothic UI (theme.font_ea)

---

## 入力パターンと処理フロー

### パターンA：議事録・文字起こしが入力された場合

**いきなりPowerPointを作成しない**

1. **Step 1: バリュエーション/財務データのドラフトを作成**
2. **Step 2: Markdownでユーザーに提示**し、chart_typeを含めて確認・修正を求める
3. **Step 3: ユーザーの承認後**、PowerPointを生成する

### パターンB：データが明確に入力された場合

1. **Step 1: 内容を確認**し、chart_typeに応じた必須フィールドが揃っているか確認
2. **Step 2: 確認後、PowerPointを生成する**

---

## JSON仕様

### 1. フットボールフィールド（chart_type: "football_field"）

```json
{
  "main_message": "〜すべき",
  "chart_title": "バリュエーション・レンジ分析",
  "chart_type": "football_field",
  "football_field": {
    "unit": "億円",
    "methods": [
      {"name": "DCF法（WACC 8-10%）", "low": 85, "high": 120, "mid": 102},
      {"name": "類似企業比較法（EV/EBITDA）", "low": 88, "high": 115},
      {"name": "先例取引比較法", "low": 95, "high": 130},
      {"name": "修正純資産法", "low": 60, "high": 75}
    ],
    "recommended_range": {"low": 90, "high": 110},
    "axis_min": 50,
    "axis_max": 140
  }
}
```

- `methods[].mid`（任意）: ミッドポイント。白い縦線で表示
- `recommended_range`（任意）: 推定レンジ。薄い背景帯で表示
- `axis_min/axis_max`（任意）: 軸の範囲。未指定時は自動計算

### 2. 株式価値ブリッジ（chart_type: "equity_bridge"）

```json
{
  "main_message": "〜すべき",
  "chart_title": "株式価値ブリッジ",
  "chart_type": "equity_bridge",
  "equity_bridge": {
    "unit": "億円",
    "items": [
      {"name": "事業価値\n(EV)", "value": 150, "type": "start"},
      {"name": "純有利子\n負債", "value": -42, "type": "adjust"},
      {"name": "非事業用\n資産", "value": 8, "type": "adjust"},
      {"name": "DD調整\n項目", "value": -8, "type": "adjust"},
      {"name": "株式価値", "value": 100, "type": "total"}
    ]
  }
}
```

- `type: "start"`: 開始バー（Accent2色）
- `type: "adjust"`: 調整項目（正=緑、負=赤）
- `type: "total"`: 合計バー（Accent2色）

### 3. 財務サマリーテーブル（chart_type: "financial_summary"）

```json
{
  "main_message": "〜すべき",
  "chart_title": "財務サマリー",
  "chart_type": "financial_summary",
  "financial_summary": {
    "periods": ["FY2022\n(実績)", "FY2023\n(実績)", "FY2024\n(計画)"],
    "metrics": [
      {"name": "【損益計算書】", "is_section": true, "values": ["", "", ""]},
      {"name": "  売上高", "values": [8500, 9520, 10660], "cagr": "12.0%"},
      {"name": "  営業利益", "values": [850, 1050, 1280], "is_total": true, "cagr": "22.7%"},
      {"name": "  EBITDA", "values": [1200, 1450, 1730], "is_total": true}
    ]
  }
}
```

- `is_section: true`: セクションヘッダー行（Accent2薄色背景）
- `is_total: true`: 合計行（太字、上に太線）
- `cagr`（任意）: CAGR列を自動追加

---

## Step 3: PowerPointの生成

### スクリプト実行コマンド

```bash
pip install python-pptx -q --break-system-packages

python <SKILL_DIR>/scripts/fill_valuation.py \
  --brand stellar_aiz \
  --data {{WORK_DIR}}/valuation_data.json \
  --output {{OUTPUT_DIR}}/Valuation_output.pptx
```

`--brand` を `roleup` にすると Roleup Standard Format (vF 20250928、A4 横、Yu Gothic UI、本文 10pt 統一、accent_revenue_bar #7C4C2C) で出力する。
`--template` は通常省略 (brand_resolver が `assets/<brand>/valuation-template.pptx` を解決)。

---

## アセット

| ファイル名 | 用途 |
|---|---|
| `assets/stellar_aiz/valuation-template.pptx` | stella V1 テンプレート (16:9、Title 1 + Text Placeholder 2、accent2 自動抽出) |
| `assets/stellar_aiz/layout.json` | stella レイアウト座標 (chart area 等) |
| `assets/roleup/valuation-template.pptx` | Roleup テンプレート (cp roleup 派生、A4 横、Yu Gothic UI、Source 3 placeholder) |
| `assets/roleup/layout.json` | Roleup レイアウト座標 |

## スクリプト

| ファイル名 | 用途 |
|---|---|
| `scripts/fill_valuation.py` | JSONデータからchart_typeに応じた3種類のチャートを動的生成してPPTXを出力する |
