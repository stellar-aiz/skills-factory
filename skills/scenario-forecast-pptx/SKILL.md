---
name: scenario-forecast-pptx
description: >
  BDD向けシナリオ別見立てチャート（売上高/EBITDA）のPowerPointスライドを生成するスキル。
  Base/Upside/Downsideの3シナリオ折れ線チャートを左右2つ並べ、
  各シリーズ凡例を各チャート右側に縦配置、
  実績・見込・計画の期間種別凡例をチャート下部に配置し、出典を含む1枚スライドを作成する。
  customer-profile-pptxと同じテンプレート・フォントを使用。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「シナリオ別見立て」「ケース別見立て」「Base/Upside/Downside」「3ケース」という言葉が出た場合
  - 「売上とEBITDAのチャート」「売上高とEBITDAの推移」「P/L見立て」という要望
  - 「今後の見立て」「財務見立て」「Financial Projection」「事業計画チャート」という言葉が出た場合
  - BDD（ビジネスDD）の文脈で対象会社の将来見立て・シナリオ分析のスライド化を求められた場合
  - ユーザーが財務モデル・事業計画のデータを貼り付けて、シナリオ別チャートのスライド化を求めた場合
  - 「折れ線チャートで3シナリオを比較」「Base/Upside/Downsideのグラフ」という要望
  - 「実績・見込・計画」を含む時系列チャートのスライドを作りたいという要望
supported_brands: [stellar_aiz, roleup]

---

# BDD シナリオ別見立てチャート PowerPoint ジェネレーター

BDD（ビジネスデュー・ディリジェンス）において、対象会社の将来見立てを
Base/Upside/Downsideの3シナリオで可視化するスライドを生成するスキル。

---

## スライド構成

| セクション | 位置 | 内容 |
|---|---|---|
| **メインメッセージ** | 最上部 | 最大70文字。「〜すべき」で締める |
| **チャートタイトル** | メインメッセージ直下 | 例:「今後の見立て ー ケース別 ー 全体売上/EBITDA」 |
| **左チャート** | 左半分 | タイトル＋横線＋単位＋折れ線チャート＋右側凡例＋下部期間種別凡例 |
| **右チャート** | 右半分 | 同上 |
| **出典** | 左下 | 情報ソースの記載 |

### チャート仕様 (3 シナリオ色 — brand 別)

| シナリオ | stellar_aiz | roleup | マーカー |
|---|---|---|---|
| Base | #4E79A7 (青) | #7C4C2C (chart_palette[0]) | ◆ダイヤ |
| Upside | #ED7D31 (オレンジ) | #897141 (chart_palette[1]) | ●丸 |
| Downside | #A5A5A5 (グレー) | #CDCECE (highlight_other) | ●丸 |
- 3シリーズは全て独立した折れ線グラフとして作成（実績期間は同じ値にすることで重なって見える）
- データラベル: 全データポイントに表示（各シナリオの色で太字）
- チャートタイトル下: 横線オブジェクト（チャート幅）
- 凡例: 各チャートの右側に縦並びで配置

### 期間種別の定義（チャート下部の凡例）

- **実績（actual）**: 過去の確定期
- **見込（forecast）**: 現在進行中の期
- **計画（plan）**: 来期以降の計画期

---

## 入力パターンと処理フロー

### パターンA：財務モデル・事業計画データが入力された場合

**いきなりPowerPointを作成しない**

1. **Step 1: データを抽出・整理する**
   - 売上高・EBITDA等の Base / Upside / Downside 値を期間別に抽出
   - 期間種別を判定: 過去期=actual, 当期=forecast, 来期以降=plan
2. **Step 2: Markdownテーブルでユーザーに提示**し、確認・修正を求める
3. **Step 3: ユーザーの承認後**、JSONを作成しPowerPointを生成する

### パターンB：データが整理済みで入力された場合

1. **Step 1: 内容を確認**
2. **Step 2: 確認後、PowerPointを生成する**

---

## JSONデータ仕様

`{{WORK_DIR}}/scenario_forecast_data.json` に以下の形式で保存する：

```json
{
  "main_message": "ケース別に全体売上・EBITDAの見立てを整理し、投資判断の前提とすべき",
  "chart_title": "今後の見立て ー ケース別 ー 全体売上/EBITDA",
  "source": "出典：対象会社提供資料、マネジメントインタビュー",
  "series_labels": {"base": "Base", "upside": "Upside", "downside": "Downside"},
  "period_type_labels": {"actual": "実績", "forecast": "見込", "plan": "計画"},
  "left_chart": {
    "title": "売上高",
    "unit": "（単位：百万円）",
    "y_max": 1400,
    "y_step": 200,
    "periods": [
      {"label": "21/6期", "type": "actual",   "base": 374, "upside": 374,  "downside": 374},
      {"label": "22/6期", "type": "actual",   "base": 423, "upside": 423,  "downside": 423},
      {"label": "23/6期", "type": "forecast", "base": 664, "upside": 664,  "downside": 664},
      {"label": "24/6期", "type": "plan",     "base": 554, "upside": 554,  "downside": 543},
      {"label": "25/6期", "type": "plan",     "base": 589, "upside": 854,  "downside": 553}
    ]
  },
  "right_chart": {
    "title": "調整後EBITDA",
    "unit": "（単位：百万円）",
    "y_max": 300,
    "y_step": 50,
    "periods": [
      {"label": "21/6期", "type": "actual", "base": 90, "upside": 90, "downside": 90}
    ]
  }
}
```

### 重要ルール

- **全3シリーズに全期間の値を設定すること**。実績期間では Base / Upside / Downside を同じ値にする（重なって1本の線に見える）
- `null` は「その期間にデータなし」を意味し、折れ線が途切れる。通常は使わない
- `type` フィールドは下部凡例の範囲に対応:
  - `"actual"` = 実績（過去の確定期）
  - `"forecast"` = 見込（現在進行中の期）
  - `"plan"` = 計画（来期以降）

### JSONフィールド仕様

| フィールド | 型 | 必須 | 説明 |
|---|---|---|---|
| `main_message` | string | 必須 | メインメッセージ。最大70文字、「〜すべき」で締める |
| `chart_title` | string | 任意 | チャートタイトル |
| `source` | string | 任意 | 出典テキスト（左下に表示） |
| `series_labels` | object | 任意 | 凡例ラベル。デフォルト: Base/Upside/Downside |
| `period_type_labels` | object | 任意 | 期間種別ラベル。デフォルト: 実績/見込/計画 |
| `left_chart.title` | string | 必須 | 左チャートタイトル（例: "売上高"） |
| `left_chart.unit` | string | 任意 | 単位表記（例: "（単位：百万円）"） |
| `left_chart.y_max` | number | 任意 | Y軸最大値 |
| `left_chart.y_step` | number | 任意 | Y軸目盛り間隔 |
| `left_chart.periods` | array | 必須 | 期間別データの配列 |
| `left_chart.periods[].label` | string | 必須 | 期間ラベル（例: "21/6期"） |
| `left_chart.periods[].type` | string | 必須 | 期間種別: "actual" / "forecast" / "plan" |
| `left_chart.periods[].base` | number | 必須 | Base値 |
| `left_chart.periods[].upside` | number | 必須 | Upside値（実績期はBaseと同値） |
| `left_chart.periods[].downside` | number | 必須 | Downside値（実績期はBaseと同値） |
| `right_chart` | object | 必須 | 右チャート（left_chartと同じ構造） |

---

## スクリプト実行コマンド

```bash
pip install python-pptx -q --break-system-packages

python <SKILL_DIR>/scripts/fill_scenario_forecast.py \
  --brand stellar_aiz \
  --data {{WORK_DIR}}/scenario_forecast_data.json \
  --output {{OUTPUT_DIR}}/ScenarioForecast_output.pptx
```

`--brand` を `roleup` にすると Roleup Standard Format (vF 20250928、A4 横、Yu Gothic UI、本文 10pt 統一、3 シナリオは brand chart_palette[0..2] + highlight_other) で出力する。
`--template` は通常省略 (brand_resolver が `assets/<brand>/scenario-forecast-template.pptx` を解決)。

※ `<SKILL_DIR>` は実際のスキルインストールパスに置き換えること。

---

## デザイン仕様

### レイアウト定数

`assets/<brand>/layout.json` に外出し。

| 要素 | stellar_aiz | roleup |
|---|---|---|
| 左チャート開始X | 0.41 in | 0.41 in |
| 左チャート幅 | 5.20 in | 4.55 in |
| 左凡例X | 5.65 in | 5.00 in |
| 右チャート開始X | 6.80 in | 6.07 in |
| 右チャート幅 | 5.20 in | 4.55 in |
| 右凡例X | 12.04 in | 10.66 in |
| チャートタイトルY | 1.40 in | 1.45 in |
| タイトル下横線Y | 1.75 in | n/a (テンプレ object 8 が下線役) |
| チャート領域Y | 1.95 in | 2.00 in |
| チャート高さ | 4.50 in | 5.00 in |
| 期間種別凡例Y | 6.60 in | 7.10 in |

### フォント / 色

| 要素 | stellar_aiz | roleup |
|---|---|---|
| 日本語 (ea / latin) | Meiryo UI / Arial | Yu Gothic UI / Yu Gothic UI |
| チャートタイトル | 14pt Bold center | 12pt subtitle 色 (#897141) left |
| 単位ラベル | 9pt #666666 | 10pt #3E3A39 (source 色) |
| シリーズ凡例 | 9pt | 10pt (font_size_body_pt) |
| 期間種別凡例 | 10pt | 10pt |
| データラベル / 軸 | 10pt | 10pt |
| 出典 | 10pt textbox (#666666) | 6pt Source 3 placeholder (#3E3A39) |

---

## アセット

| ファイル名 | 用途 |
|---|---|
| `assets/stellar_aiz/scenario-forecast-template.pptx` | stella V1 テンプレート (16:9、Title 1 + Text Placeholder 2 + Table 1) |
| `assets/stellar_aiz/layout.json` | stella レイアウト座標 |
| `assets/roleup/scenario-forecast-template.pptx` | Roleup テンプレート (cp roleup 派生、A4 横、Yu Gothic UI、Source 3 placeholder) |
| `assets/roleup/layout.json` | Roleup レイアウト座標 |

## スクリプト

| ファイル名 | 用途 |
|---|---|
| `scripts/fill_scenario_forecast.py` | JSONデータからPPTXを生成するスクリプト |

## 参考

| ファイル名 | 内容 |
|---|---|
| `references/sample_data.json` | サンプルJSONデータ |
