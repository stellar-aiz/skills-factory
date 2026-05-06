---
name: workforce-composition-pptx
description: >
  人員構成（Workforce Composition）のPowerPointスライドを生成するスキル。
  BDD（ビジネスデュー・ディリジェンス）において、対象会社の人員推移と部署別人員構成を
  左右2カラムで1枚のスライドにまとめる。
  左側: 在籍人員数の推移（ネイティブ棒グラフ: 総従業員数/入社人数/退職人数の3系列）
  右側: 部署別人員構成テーブル（部署名, 人数, 平均年齢, 平均勤続年数, 管理職数, 有資格者数＋合計行）
  全てPowerPointネイティブオブジェクトで生成するため、ユーザーがPowerPoint上で自由に編集可能。
  company-history-pptxと同じテンプレート・フォントを使用。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「人員構成」「人員推移」「Workforce Composition」「従業員構成」「在籍人員」という言葉が出た場合
  - 「人員数の推移をスライドにして」「部署別の人員構成をパワポで」「従業員の推移をスライドに」という要望
  - 「入退社の推移」「入社・退職の推移」「ヘッドカウント推移」をスライド化したいという要望
  - BDD（ビジネスDD）の文脈で対象会社の人員構成・組織体制のスライド化を求められた場合
  - ユーザーがIM（インフォメーション・メモランダム）や会社HPの人員情報を貼り付けて、人員構成のスライド化を求めた場合
  - 「部署別人数」「部署サマリー」「組織別人員」をスライドにしたいという要望
  - 「有資格者数」「管理職数」「平均年齢」「平均勤続年数」を含む人員テーブルのスライド化を求められた場合
supported_brands: [stellar_aiz, roleup]

---

# 人員構成 PowerPoint ジェネレーター

BDD（ビジネスデュー・ディリジェンス）において、対象会社の人員推移と部署別人員構成を
左右2カラムで1枚のPowerPointスライドに整理するスキル。

---

## スライド構成

| セクション | 位置 | 内容 |
|---|---|---|
| **メインメッセージ** | 最上部 | 最大70文字。対象会社の人員構成の概要を端的に伝える |
| **チャートタイトル** | メインメッセージ直下 | 「人員構成」等 |
| **在籍人員数の推移** | 左側（約45%） | ネイティブ棒グラフ（3系列: 総従業員数/入社/退職）＋凡例＋単位 |
| **部署別人員構成** | 右側（約55%） | ネイティブテーブル（部署サマリー＋合計行） |
| **出典** | 左下 | 情報ソースの記載 |

### 左側: 在籍人員数の推移

- **棒グラフ（オレンジ #ED7D31）**: 総従業員数（メイン棒）
- **棒グラフ（青 #4472C4）**: 入社人数（サブ棒）
- **棒グラフ（グレー #A5A5A5）**: 退職人数（サブ棒、負数で表示）
- **データラベル**: 各棒に数値を表示（0は非表示）
- **凡例**: カスタム凡例（■マーカー＋ラベル）
- **期数**: 2〜5期対応（デフォルト3期）
- 目盛線なし、数値軸非表示

### 右側: 部署別人員構成テーブル

- **ヘッダー行**: グレー背景（#F0F0F0）、太字
- **データ行**: 偶数行に薄いグレー背景（#FAFAFA）
- **合計行**: 濃いグレー背景（#E8E8E8）、太字、加重平均で自動計算
- **列構成**（デフォルト6列）: 部署名, 人数, 平均年齢, 平均勤続年数, 管理職数, 有資格者数
- **セル罫線**: 薄いグレー（#CCCCCC）
- **フォント**: 12pt、Meiryo UI

### メインメッセージのルール

- 最大70文字
- 対象会社の人員構成の概要を端的に伝える文
- 例: 「対象会社は設備事業を中心にXX名体制で運営しており、有資格者を多く擁する技術者集団である」
- 「〜すべき」で終えるのではなく、事実ベースで概要を記述する

---

## 入力パターンと処理フロー

### パターンA：IM・会社HP・議事録等の情報が入力された場合

**いきなりPowerPointを作成しない**

1. **Step 1: 人員情報を抽出・整理する**
2. **Step 2: Markdownでユーザーに提示**し、確認・修正を求める
3. **Step 3: ユーザーの承認後**、PowerPointを生成する

### パターンB：データが整理済みで入力された場合

1. **Step 1: 内容を確認**し、期数・部署データが揃っているか確認
2. **Step 2: 確認後、PowerPointを生成する**

---

## Step 1: 人員情報の抽出ガイドライン

### 在籍人員数の推移（左側チャート用）

各期について以下の3項目を抽出する：
- **総従業員数**: 期末時点の在籍人員数
- **入社人数**: 当期の新規入社者数（正の値）
- **退職人数**: 当期の退職者数（負の値で記載、例: -2）

### 部署別人員構成（右側テーブル用）

各部署について以下の項目を抽出する：
- **部署名**: 事業部門・管理部門等の名称
- **人数**: 当該部署の在籍人数
- **平均年齢**: 部署の平均年齢（小数1桁）
- **平均勤続年数**: 部署の平均勤続年数（小数1桁、不明の場合はnull）
- **管理職数**: 管理職（課長以上等）の人数
- **有資格者数**: 業務関連資格の保有者数

---

## Step 2: Markdownでの確認出力フォーマット

```markdown
## 人員構成 整理結果

**Main Message（最大70文字）**
対象会社は設備事業を中心に19名体制で運営しており、有資格者を多く擁する技術者集団である

**Chart Title**
人員構成

### 在籍人員数の推移
| 期 | 総従業員数 | 入社人数 | 退職人数 |
|---|---|---|---|
| 21/6期 | 21 | 1 | -2 |
| 22/6期 | 20 | 0 | -2 |
| 23/6期 | 18 | 1 | 0 |
| 24/6期 | 19 | 0 | 0 |

### 部署別人員構成（24/6期）
| 部署名 | 人数 | 平均年齢 | 平均勤続年数 | 管理職数 | 有資格者数 |
|---|---|---|---|---|---|
| 設備事業 | 12 | 42.5 | 18.3 | 3 | 5 |
| 内作・土木・諸工事 | 3 | 38.0 | 12.3 | 1 | 2 |
| 不動産事業 | 1 | 45.0 | 7.0 | 0 | 0 |
| 総務・経理 | 2 | 50.0 | - | 1 | 0 |
| **合計** | **18** | **42.7** | **16.5** | **5** | **7** |

### 出典
対象会社提供資料
```

確認メッセージ例：
> 上記の人員構成データでよろしいでしょうか？追加・修正があればお知らせください。確認後にPowerPointを生成します。

---

## Step 3: PowerPointの生成

### テンプレートの参照

テンプレートは `assets/workforce-composition-template.pptx` を使用する。

```bash
TEMPLATE="<SKILL_DIR>/assets/workforce-composition-template.pptx"
```

※ `<SKILL_DIR>` は実際のスキルインストールパスに置き換えること。

### JSONデータ仕様

`{{WORK_DIR}}/workforce_composition_data.json` に以下の形式で保存する：

```json
{
  "main_message": "対象会社は設備事業を中心に19名体制で運営しており、有資格者を多く擁する技術者集団である",
  "chart_title": "人員構成",
  "source": "出典：対象会社提供資料",
  "headcount_trend": {
    "title": "在籍人員数の推移",
    "unit": "人",
    "series_labels": {
      "total": "総従業員数",
      "new_hires": "入社人数",
      "departures": "退職人数"
    },
    "periods": [
      {"label": "21/6期", "total": 21, "new_hires": 1, "departures": -2},
      {"label": "22/6期", "total": 20, "new_hires": 0, "departures": -2},
      {"label": "23/6期", "total": 18, "new_hires": 1, "departures": 0},
      {"label": "24/6期", "total": 19, "new_hires": 0, "departures": 0}
    ]
  },
  "department_table": {
    "title": "24/6期（5月現在）の人員構成",
    "columns": ["部署名", "人数", "平均年齢", "平均勤続年数", "管理職数", "有資格者数"],
    "departments": [
      {"name": "設備事業", "headcount": 12, "avg_age": 42.5, "avg_tenure": 18.3, "managers": 3, "certified": 5},
      {"name": "内作・土木・諸工事", "headcount": 3, "avg_age": 38.0, "avg_tenure": 12.3, "managers": 1, "certified": 2},
      {"name": "不動産事業", "headcount": 1, "avg_age": 45.0, "avg_tenure": 7.0, "managers": 0, "certified": 0},
      {"name": "総務・経理", "headcount": 2, "avg_age": 50.0, "avg_tenure": null, "managers": 1, "certified": 0}
    ],
    "show_total": true
  }
}
```

### JSONフィールド仕様

| フィールド | 型 | 必須 | 説明 |
|---|---|---|---|
| `main_message` | string | 必須 | メインメッセージ。最大70文字。人員構成の概要を端的に伝える |
| `chart_title` | string | 任意 | チャートタイトル。デフォルト「人員構成」 |
| `source` | string | 任意 | 出典テキスト（左下に表示） |
| `headcount_trend.title` | string | 任意 | 左側セクションタイトル。デフォルト「在籍人員数の推移」 |
| `headcount_trend.unit` | string | 任意 | 単位。デフォルト「人」 |
| `headcount_trend.series_labels.total` | string | 任意 | 総従業員数の系列名。デフォルト「総従業員数」 |
| `headcount_trend.series_labels.new_hires` | string | 任意 | 入社人数の系列名。デフォルト「入社人数」 |
| `headcount_trend.series_labels.departures` | string | 任意 | 退職人数の系列名。デフォルト「退職人数」 |
| `headcount_trend.periods` | array | 必須 | 期別データの配列（2〜5要素） |
| `headcount_trend.periods[].label` | string | 必須 | 期のラベル（例: "21/6期"） |
| `headcount_trend.periods[].total` | number | 必須 | 総従業員数 |
| `headcount_trend.periods[].new_hires` | number | 必須 | 入社人数（正の値） |
| `headcount_trend.periods[].departures` | number | 必須 | 退職人数（負の値） |
| `department_table.title` | string | 任意 | 右側セクションタイトル。デフォルト「人員構成」 |
| `department_table.columns` | array | 任意 | テーブル列名の配列。デフォルト6列 |
| `department_table.departments` | array | 必須 | 部署データの配列 |
| `department_table.departments[].name` | string | 必須 | 部署名 |
| `department_table.departments[].headcount` | number | 必須 | 人数 |
| `department_table.departments[].avg_age` | number | 任意 | 平均年齢（null可） |
| `department_table.departments[].avg_tenure` | number | 任意 | 平均勤続年数（null可） |
| `department_table.departments[].managers` | number | 任意 | 管理職数 |
| `department_table.departments[].certified` | number | 任意 | 有資格者数 |
| `department_table.show_total` | boolean | 任意 | 合計行の表示。デフォルトtrue |

---

## スクリプト実行コマンド

```bash
pip install python-pptx -q --break-system-packages

python <SKILL_DIR>/scripts/fill_workforce_composition.py \
  --brand {stellar_aiz|roleup} \
  --data {{WORK_DIR}}/workforce_composition_data.json \
  --output {{OUTPUT_DIR}}/WorkforceComposition_output.pptx
```

`--brand` 省略時は `stellar_aiz`(16:9, Meiryo UI, V1 配色)。
`--brand roleup` 指定時は A4 横テンプレ + Yu Gothic UI 10pt + 茶系 chart_palette + Source 3 placeholder 6pt + サブヘッダ 12pt subtitle 色 (#897141)。
`--template` 明示指定で brand_resolver の解決を上書き可能。

※ `<SKILL_DIR>` は実際のスキルインストールパスに置き換えること。

### 出力確認

```bash
python -m markitdown {{OUTPUT_DIR}}/WorkforceComposition_output.pptx
```

---

## デザイン仕様

### フォントサイズ一覧

| 要素 | サイズ | 備考 |
|---|---|---|
| メインメッセージ | テンプレート準拠 | Bold |
| チャートタイトル | テンプレート準拠 | |
| セクションタイトル | 14pt | Bold、下線付き |
| 単位表記 | 11pt | |
| 凡例 | 10pt | |
| チャートデータラベル | 11pt | Bold |
| チャート軸ラベル | 11pt | |
| テーブルヘッダー | 12pt | Bold |
| テーブル本文 | 12pt | |
| 出典 | 10pt | グレー(#666666) |

### 色

| 要素 | カラーコード |
|---|---|
| テキスト | #333333 |
| 総従業員数（棒） | #ED7D31（オレンジ） |
| 入社人数（棒） | #4472C4（青） |
| 退職人数（棒） | #A5A5A5（グレー） |
| テーブルヘッダー背景 | #F0F0F0 |
| テーブル偶数行背景 | #FAFAFA |
| テーブル合計行背景 | #E8E8E8 |
| テーブル罫線 | #CCCCCC |
| 出典テキスト | #666666 |

### レイアウト定数

| 要素 | 値 |
|---|---|
| 左パネル開始X | 0.41in |
| 右パネル開始X | 6.50in |
| パネル開始Y（共通） | 1.50in |
| 左パネル幅 | 5.80in |
| 右パネル幅 | 6.40in |
| チャート高さ | 5.00in |
| テーブル最大高さ | 5.20in |

---

## 合計行の自動計算ルール

`show_total: true` の場合、合計行は以下のルールで自動計算される：

| 列 | 計算方法 |
|---|---|
| 部署名 | 「合計」固定 |
| 人数 | 全部署の単純合計 |
| 平均年齢 | 人数による加重平均（nullの部署は除外） |
| 平均勤続年数 | 人数による加重平均（nullの部署は除外） |
| 管理職数 | 全部署の単純合計 |
| 有資格者数 | 全部署の単純合計 |

---

## 品質チェックリスト

- [ ] メインメッセージが正しく表示されているか
- [ ] 左側: 棒グラフの3系列（総従業員数/入社/退職）が正しく表示されているか
- [ ] 左側: データラベルが全棒に表示されているか（0は非表示）
- [ ] 左側: 凡例の色とラベルが正しいか
- [ ] 左側: 目盛線が表示されていないか
- [ ] 右側: 全部署がテーブルに表示されているか
- [ ] 右側: 合計行が正しく計算されているか（加重平均）
- [ ] 右側: ヘッダー行が太字・グレー背景になっているか
- [ ] 右側: テーブルがスライド内に収まっているか（行数に応じた高さ調整）
- [ ] 出典が左下に表示されているか
- [ ] PPTXのmarkitdown出力でプレースホルダーが残っていないか

---

## アセット

| ファイル名 | 用途 |
|---|---|
| `assets/stellar_aiz/workforce-composition-template.pptx` | stella 16:9 テンプレート(Title 1 / Text Placeholder 2) |
| `assets/stellar_aiz/layout.json` | stella レイアウト座標 (左右パネル + chart/table 高さ) |
| `assets/roleup/workforce-composition-template.pptx` | roleup A4 横テンプレート(cp roleup template から派生、Title 1 / Text Placeholder 2 / Source 3 + 茶色ガイド矩形) |
| `assets/roleup/layout.json` | roleup レイアウト座標 (A4 横、左右 0.41in マージン) |

## スクリプト

| ファイル名 | 用途 |
|---|---|
| `scripts/fill_workforce_composition.py` | JSONデータからPPTXネイティブオブジェクト（棒グラフ＋テーブル）を生成するスクリプト |

## 参考

| ファイル名 | 内容 |
|---|---|
| `references/sample_data.json` | サンプルJSONデータ（設備事業会社の例） |
