---
name: customer-profile-pptx
description: >
  主要顧客プロファイル（Customer Profile）のPowerPointスライドを生成するスキル。
  BDD（ビジネスデュー・ディリジェンス）において、対象会社の主要顧客の企業概要と業績を
  左右2カラムで1枚のスライドにまとめる。
  左側: 企業の概要（ブレットポイント形式のテーブル）
  右側: 業績チャート（PowerPointネイティブ複合チャート: 棒=売上高, 折れ線=営業利益率）＋CAGR注釈
  company-history-pptxと同じテンプレート・フォントを使用。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「顧客プロファイル」「Customer Profile」「主要顧客」「顧客情報」という言葉が出た場合
  - 「顧客の概要をスライドにして」「取引先の情報をパワポで」「顧客企業の業績をスライドに」という要望
  - BDD（ビジネスDD）の文脈で対象会社の主要顧客のプロファイルスライド化を求められた場合
  - ユーザーがIM（インフォメーション・メモランダム）や顧客HPの情報を貼り付けて、顧客プロファイルのスライド化を求めた場合
  - 「売上高と営業利益率のチャート」「業績チャート付きの顧客スライド」という要望
---

# 主要顧客プロファイル PowerPoint ジェネレーター

BDD（ビジネスデュー・ディリジェンス）において、対象会社の主要顧客の企業概要と業績を
左右2カラムで1枚のPowerPointスライドに整理するスキル。

---

## スライド構成

| セクション | 位置 | 内容 |
|---|---|---|
| **メインメッセージ** | 最上部 | 最大70文字。「〜すべき」で締める |
| **チャートタイトル** | メインメッセージ直下 | 「主要顧客プロファイル：○○社」 |
| **企業の概要** | 左側 | ブレットポイント形式のkey-valueテーブル（枠線なし） |
| **業績** | 右側 | ネイティブ複合チャート（棒=売上高 + 折れ線=営業利益率）＋CAGR注釈 |
| **出典** | 左下 | 情報ソースの記載 |

### 左側: 企業の概要

- ブレットポイント形式（「• ラベル　値」）
- 枠線なし、14ptフォント
- 典型的な項目: 商号、事業内容、本社所在地、拠点、設立年、資本金、売上高、代表者、上場、従業員数、仕入先、販売先、大株主

### 右側: 業績

- **棒グラフ（青）**: 売上高
- **折れ線（紺色）**: 営業利益率（白文字データラベル、マーカーサイズ9）
- **CAGR注釈**: 矢印＋楕円で自動計算値を表示
- **凡例**: 単位表記（左）と系列凡例（右）を分離配置
- 目盛線なし、年ラベル縦書き

### 年度範囲のルール

- **End year**: 実行年の前年（2026年実行なら2025年）
- **Start year**: そこから最大10年遡る（データがある限り）
- **CAGR**: data配列の最初と最後の売上高から自動計算

---

## 入力パターンと処理フロー

### パターンA：IM・顧客HP・TSR等の情報が入力された場合

**いきなりPowerPointを作成しない**

1. **Step 1: 顧客情報を抽出・整理する**
2. **Step 2: Markdownでユーザーに提示**し、確認・修正を求める
3. **Step 3: ユーザーの承認後**、PowerPointを生成する

### パターンB：データが整理済みで入力された場合

1. **Step 1: 内容を確認**し、項目が揃っているか確認
2. **Step 2: 確認後、PowerPointを生成する**

---

## JSONデータ仕様

`{{WORK_DIR}}/customer_profile_data.json` に以下の形式で保存する：

```json
{
  "source": "出典：対象会社HP、東京商工リサーチ、日経バリューサーチ",
  "main_message": "○○社はBDD対象会社の最大顧客であり、安定成長する取引基盤を評価すべき",
  "chart_title": "主要顧客プロファイル：株式会社○○",
  "company_overview": {
    "section_title": "企業の概要",
    "items": [
      {"label": "商号", "value": "株式会社 ○○"},
      {"label": "事業内容", "value": "○○業"},
      {"label": "本社所在地", "value": "○○県○○市"},
      {"label": "設立年", "value": "19XX年XX月"},
      {"label": "資本金", "value": "X,XXX万円"},
      {"label": "売上高", "value": "XX億円（XX/X期）"},
      {"label": "代表者", "value": "代表取締役 ○○ ○○"},
      {"label": "上場", "value": "未上場"},
      {"label": "従業員数", "value": "XXX名"},
      {"label": "大株主", "value": "○○(XX.X%)、○○(XX.X%)"}
    ]
  },
  "performance": {
    "section_title": "業績",
    "unit_label": "（単位：十億円、%）",
    "bar_label": "売上高",
    "line_label": "営業利益率",
    "data": [
      {"year": "2015", "revenue": 10.0, "op_margin": 3.5},
      {"year": "2016", "revenue": 11.0, "op_margin": 4.0},
      {"year": "2017", "revenue": 12.0, "op_margin": 3.8}
    ]
  }
}
```

### JSONフィールド仕様

| フィールド | 型 | 必須 | 説明 |
|---|---|---|---|
| `source` | string | 任意 | 出典テキスト（左下に表示） |
| `main_message` | string | 必須 | メインメッセージ。最大70文字、「〜すべき」で締める |
| `chart_title` | string | 任意 | チャートタイトル。デフォルト「主要顧客プロファイル」 |
| `company_overview.section_title` | string | 任意 | 左側セクションタイトル。デフォルト「企業の概要」 |
| `company_overview.items` | array | 必須 | ラベル＋値の配列 |
| `company_overview.items[].label` | string | 必須 | 項目ラベル（例: "商号"） |
| `company_overview.items[].value` | string | 必須 | 項目の値 |
| `performance.section_title` | string | 任意 | 右側セクションタイトル。デフォルト「業績」 |
| `performance.unit_label` | string | 任意 | 単位表記 |
| `performance.bar_label` | string | 任意 | 棒グラフの系列名。デフォルト「売上高」 |
| `performance.line_label` | string | 任意 | 折れ線の系列名。デフォルト「営業利益率」 |
| `performance.data` | array | 必須 | 年度別データの配列 |
| `performance.data[].year` | string | 必須 | 年度（例: "2023"） |
| `performance.data[].revenue` | number | 必須 | 売上高 |
| `performance.data[].op_margin` | number | 必須 | 営業利益率（%） |

---

## スクリプト実行コマンド

```bash
pip install python-pptx -q --break-system-packages

# ブランド未指定（デフォルト = stellar_aiz）
python <SKILL_DIR>/scripts/fill_customer_profile.py \
  --data {{WORK_DIR}}/customer_profile_data.json \
  --output {{OUTPUT_DIR}}/CustomerProfile_output.pptx

# Rollup 社向け
python <SKILL_DIR>/scripts/fill_customer_profile.py \
  --data {{WORK_DIR}}/customer_profile_data.json \
  --brand rollup \
  --output {{OUTPUT_DIR}}/CustomerProfile_output.pptx

# テンプレートを明示指定（オプション）
python <SKILL_DIR>/scripts/fill_customer_profile.py \
  --data {{WORK_DIR}}/customer_profile_data.json \
  --template <SKILL_DIR>/assets/stellar_aiz/customer-profile-template.pptx \
  --output {{OUTPUT_DIR}}/CustomerProfile_output.pptx
```

※ `<SKILL_DIR>` は実際のスキルインストールパスに置き換えること。
`--brand` の有効値は `stellar_aiz` / `rollup`、デフォルトは `stellar_aiz`。
`--template` は省略可。省略時は brand から自動解決される
（`assets/<brand>/customer-profile-template.pptx`、未存在なら stellar_aiz にフォールバック）。
オーケストレーター（market-overview-agent / company-deepdive-agent 等）から呼ぶ場合、
parent は scope.json の `brand` フィールドを読んで子 fill 呼び出しに `--brand` で渡す。

### 出力確認

```bash
python -m markitdown {{OUTPUT_DIR}}/CustomerProfile_output.pptx
```

---

## デザイン仕様

### フォントサイズ一覧

| 要素 | サイズ | 備考 |
|---|---|---|
| メインメッセージ | 28pt | テンプレート準拠、Bold |
| チャートタイトル | テンプレート準拠 | |
| セクションタイトル | 14pt | Bold、下線付き |
| 企業概要ラベル | 14pt | Bold、「•」付き |
| 企業概要値 | 14pt | Regular |
| データラベル | 11pt | |
| 凡例・単位表記 | 12pt | |
| 軸（年ラベル） | 11pt | 縦書き |
| CAGR数値 | 16pt | Bold、楕円内 |
| 出典 | 10pt | グレー(#666666) |

### 色

| 要素 | カラーコード |
|---|---|
| テキスト | #333333 |
| 棒グラフ（売上高） | #4E79A7 |
| 折れ線・マーカー（営業利益率） | #003366 |
| 営業利益率データラベル | #FFFFFF（白） |
| 出典テキスト | #666666 |
| CAGR矢印・楕円 | #333333 |

### レイアウト定数

| 要素 | 値 |
|---|---|
| 左パネル開始X | 0.41in |
| 右パネル開始X | 6.50in |
| パネル開始Y（共通） | 1.50in |
| 左パネル幅 | 5.80in |
| 右パネル幅 | 6.40in |
| CAGR gap_above | 1.20in |

---

## 品質チェックリスト

- [ ] メインメッセージが正しく表示されているか
- [ ] 左側: 全項目がブレットポイント形式で表示されているか
- [ ] 左側: 枠線が表示されていないか
- [ ] 右側: 棒グラフと折れ線が正しく表示されているか
- [ ] 右側: データラベルが全データポイントに表示されているか
- [ ] 右側: 目盛線が表示されていないか
- [ ] 右側: 年ラベルが縦書きになっているか
- [ ] CAGR: 自動計算値が正しいか、矢印と楕円が適切な位置にあるか
- [ ] 凡例: 単位（左）と系列名（右）が分離配置されているか
- [ ] 出典が左下に表示されているか

---

## アセット

| ファイル名 | 用途 |
|---|---|
| `assets/stellar_aiz/customer-profile-template.pptx` | Stellar AIZ 用テンプレート（company-history-template ベース、16:9） |
| `assets/stellar_aiz/layout.json` | Stellar AIZ 用座標定義（panel_y / left_x / right_x 等） |
| `assets/rollup/layout.json` | Rollup 用座標定義（V1 では Stella と同値、curated テンプレ未配置時はフォールバック） |

V1 では Rollup 専用テンプレ pptx は配置せず、`brand_resolver.template_path()` のフォールバック機構経由で `assets/stellar_aiz/customer-profile-template.pptx` を流用する。
V2 で Rollup curated テンプレ（A4 横、Yu Gothic UI、褐色アクセント）を `assets/rollup/customer-profile-template.pptx` として導入する予定。

## スクリプト

| ファイル名 | 用途 |
|---|---|
| `scripts/fill_customer_profile.py` | JSONデータからPPTXを生成するスクリプト |

## 参考

| ファイル名 | 内容 |
|---|---|
| `references/sample_data.json` | サンプルJSONデータ（株式会社ジップの例） |

---

## オーケストレーター連携

### `business-deepdive-agent` から呼び出される場合の規約

| 項目 | 値 |
|---|---|
| 入力 JSON ファイル名 | `data_NN_customer_profile.json`（NN は global_slide_offset 経由で親が採番） |
| 出力 PPTX ファイル名 | `slide_NN_customer_profile.pptx`（同上） |
| 入力ディレクトリ | `{{WORK_DIR}}/company-deepdive-agent/<parent_run_id>/segments/<segment_slug>/` |
| 出力ディレクトリ | 同上 |

`business-deepdive-agent` は本スキルを **5 論点中 4 番目（顧客は誰か？）** として呼び出す。
B2B セグメントは主要法人取引先のプロファイル、B2C セグメントは代表的な顧客セグメント像を示す。
作業ディレクトリは `company-deepdive-agent` 配下のセグメント別 subdir に統一し、merge は親が担当。

単独起動時や他オーケストレーター（market-overview-agent / strategy-report-agent）から呼ばれる場合は本セクションの規約は適用されない（既存の運用に従う）。
