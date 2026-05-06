---
name: current-period-forecast-pptx
description: >
  BDD（Business Due Diligence）における「今後の見立て：当期着地見込み」のPowerPointスライドを生成するスキル。
  財務モデルのExcelデータから、マネジメント計画（Base）と弊社計画（Downside case / Upside case）の
  PL項目（売上・売上原価・販管費・営業利益・EBITDA等）を比較するテーブルスライドを作成する。
  各ケースの数値・構成比と前提条件（Assumption）を1枚のスライドに整理する。
  HTMLで描画→Playwrightでスクリーンショット→PPTXに画像として挿入する方式。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「当期着地見込み」「当期見込み」「Current Period Forecast」「着地見込み」「着地予想」という言葉が出た場合
  - 「進行期見込み」「今期見込み」「当期のPL見込み」「今期の業績見込み」という言葉が出た場合
  - 「Downside / Upside」「ダウンサイド・アップサイド」「弊社計画」と「マネジメント計画」の比較スライドを求められた場合
  - 「財務モデルからスライドを作って」「Excelの財務データをスライドに」と当期着地の文脈で言われた場合
  - BDD（ビジネスDD）の文脈で対象会社の当期業績見込み・着地見込みのスライド化を求められた場合
  - 「今後の見立て」「業績見込み」「PL見込み」をDownside/Upsideで比較するスライドを作りたいという要望
  - ユーザーが財務モデルExcelをアップロードし、当期着地見込みのスライド化を求めた場合
supported_brands: [stellar_aiz, roleup]

---

# 当期着地見込み PowerPoint ジェネレーター（BDD用）

BDDにおける対象会社の当期着地見込みを、マネジメント計画（Base）と弊社計画（Downside / Upside）で比較する1枚スライドを生成するスキル。

---

## 当期着地見込みスライドとは

M&AのBDD（Business Due Diligence）で使用する、対象会社の進行期の業績見込みを1枚に集約したスライド。
マネジメントが提示した計画をベースに、弊社（コンサルティングファーム）が独自にDownside case（保守的シナリオ）とUpside case（楽観的シナリオ）を設定し、PL項目ごとに比較する。

### スライド全体の構成

| 要素 | 配置 | 説明 |
|------|------|------|
| **Main Message** | 上段左寄せ（Title 1） | 当期見込みの要約を一文で表現。**最大70文字**。必ず「〜すべき」で締める |
| **Chart Title** | 下段左寄せ（Text Placeholder 2） | スライドタイトル。**10〜25文字**。通常は「今後の見立て：当期着地見込み」 |
| **テーブル** | コンテンツエリア全体 | マネジメント計画 vs 弊社計画（Downside/Upside）の比較テーブル |
| **出典** | 左下 | 「受領資料、Q＆A、マネジメントインタビューより当社作成」（10pt） |

### テーブル構成

テーブルは以下の6列で構成される：

| 列 | 内容 | 備考 |
|----|------|------|
| 項目（調整後） | PL項目名 | 売上、売上原価、販管費、営業利益、EBITDA等 |
| マネジメント計画(Base) | マネジメント提示の数値＋構成比 | 基準となる計画値 |
| Downside case 数値 | 弊社設定の保守的数値＋構成比 | |
| Downside case 前提 | Downsideの前提条件・算定根拠 | ■ マーカー付きで記載 |
| Upside case 数値 | 弊社設定の楽観的数値＋構成比 | |
| Upside case 前提 | Upsideの前提条件・算定根拠 | ■ マーカー付きで記載 |

ヘッダーは2段構成：
- 1段目：「弊社計画」がDownside/Upsideの4列にまたがる
- 2段目：「Downside case」「Upside case」がそれぞれ数値＋前提の2列にまたがる

### 標準的なPL項目

| # | 項目 | 説明 |
|---|------|------|
| 1 | 売上 | 売上高（=100%基準） |
| 2 | 売上原価 | 売上原価（売上比%） |
| 3 | 販管費 | 販売費及び一般管理費（売上比%） |
| 4 | 営業利益 | 売上−売上原価−販管費（売上比%） |
| 5 | EBITDA | 営業利益＋減価償却費（売上比%） |

案件に応じて「経常利益」「当期純利益」「粗利」等を追加・変更してよい。

### 記述ルール

- **Main Message**: 最大70文字。当期見込みの要点を総括し、必ず「〜すべき」で終える
- **Chart Title**: 10〜25文字程度。通常は「今後の見立て：当期着地見込み」
- **数値**: 百万円単位。構成比（売上比%）をカッコ内に併記
- **前提条件**: 各ケースの算定根拠を簡潔に記載。■ マーカー付き
- **出典**: デフォルトは「受領資料、Q＆A、マネジメントインタビューより当社作成」

---

## 入力パターンと処理フロー

### パターンA：財務モデルExcelが入力された場合

**いきなりPowerPointを作成しない**

1. **Step 1: Excelから当期着地見込みのデータを抽出**する（売上〜EBITDA、マネジメント計画・Downside・Upside）
2. **Step 2: 前提条件を整理**する（各ケースの算定根拠をExcelのシート・セルコメント・Q&A回答等から特定）
3. **Step 3: Markdownでユーザーに提示**し、確認・修正を求める
4. **Step 4: ユーザーの承認後**、PowerPointを生成する

### パターンB：データが直接入力された場合

1. **Step 1: 内容を確認**し、PL項目・数値・前提が揃っているか確認
2. 不足・曖昧な点があれば修正提案を行う
3. **Step 2: 確認後、PowerPointを生成する**

### パターンC：議事録・マネジメントインタビューメモが入力された場合

1. **Step 1: 議事録から当期着地に関する数値・前提を抽出**する
2. **Step 2: Markdownでユーザーに提示**し、確認・修正を求める
3. **Step 3: ユーザーの承認後**、PowerPointを生成する

---

## Step 1: Excelからのデータ抽出ガイドライン

財務モデルExcelから以下を抽出する：

### 数値の抽出
- **マネジメント計画**: マネジメントが提示した各PL項目の計画値
- **Downside case**: 保守的シナリオの数値（原価率上昇リスク等を織り込み）
- **Upside case**: 楽観的シナリオの数値（実績トレンドに基づく上振れ等）
- **構成比**: 各項目の売上比（%）

### 前提条件の抽出
- Excelのセルコメント、別シートのメモ、Q&A回答から各ケースの算定根拠を特定
- 典型的な前提パターン：
  - 「マネジメント案を使用」（マネジメント計画をそのまま採用）
  - 「22/6期実績を参考に設定」（過去実績ベース）
  - 「原価率を○○%→○○%に調整」（リスク織り込み）
  - 「売上−売上原価−販管費」（計算項目）
  - 「営業利益＋減価償却費」（計算項目）

---

## Step 2: 当期着地見込みのMarkdown出力フォーマット

ユーザーに確認を求める際は、以下のフォーマットで出力する：

```markdown
## 当期着地見込み 整理結果

**Main Message（※ドラフト）**
〜すべき（最大70文字）

**Chart Title**
今後の見立て：当期着地見込み

**期間ラベル**
進行期（XX/X期）見込み

### テーブルデータ

| 項目 | マネジメント計画 | Downside | Downside前提 | Upside | Upside前提 |
|------|---------------|----------|-------------|--------|-----------|
| 売上 | 543 (100%) | 543 (100%) | マネジメント案を使用 | 554 (100%) | 22/6実績を参考に設定 |
| 売上原価 | 385 (70.9%) | 389 (71.6%) | 原価率上昇リスクを織り込み | 391 (70.5%) | 22/6期実績ベース |
| ... | ... | ... | ... | ... | ... |

### 出典
受領資料、Q＆A、マネジメントインタビューより当社作成
```

確認メッセージ例：
> 上記の当期着地見込み整理でよろしいでしょうか？数値や前提条件の修正があればお知らせください。確認後にPowerPointを生成します。

---

## Step 3: PowerPointの生成

### テンプレートの参照

テンプレートは brand 別に分離されている (`assets/<brand>/current-period-forecast-template.pptx`)。
`--brand` 引数を渡すか、`brand_resolver.template_path()` で自動解決される。

```bash
# stella (16:9, Meiryo UI)
TEMPLATE="<SKILL_DIR>/assets/stellar_aiz/current-period-forecast-template.pptx"

# roleup (A4 横, Yu Gothic UI, 茶系)
TEMPLATE="<SKILL_DIR>/assets/roleup/current-period-forecast-template.pptx"
```

`<SKILL_DIR>` はこのスキルがインストールされたディレクトリパスに置き換えること。

テンプレートのShape構造は `references/template-mapping.md` を参照。

### 当期着地見込みデータのJSON化

データを `{{WORK_DIR}}/forecast_data.json` に以下の形式で保存する：

```json
{
  "main_message": "マネジメント計画をベースに、売上原価のリスクを織り込んだDownside caseと実績ベースのUpside caseを設定すべき",
  "chart_title": "今後の見立て：当期着地見込み",
  "period_label": "進行期（24/6期）見込み",
  "unit_label": "単位：百万円/(%)",
  "management_plan_label": "マネジメント\n計画\n(Base)",
  "our_plan_label": "弊社計画",
  "downside_label": "Downside case",
  "upside_label": "Upside case",
  "item_header": "項目（調整後）",
  "source": "受領資料、Q＆A、マネジメントインタビューより当社作成",
  "rows": [
    {
      "item": "売上",
      "management_value": "543",
      "management_pct": "100%",
      "downside_value": "543",
      "downside_pct": "100%",
      "downside_assumption": "マネジメント案を使用",
      "upside_value": "554",
      "upside_pct": "100%",
      "upside_assumption": "22/6実績を参考に設定（22/6期4月⇒期末にかけての売上伸長率を使用）"
    },
    {
      "item": "売上原価",
      "management_value": "385",
      "management_pct": "70.9%",
      "downside_value": "389",
      "downside_pct": "71.6%",
      "downside_assumption": "マネジメント案原価率を22/6期実績まで上昇するリスクを織り込んで設定（残期間原価率：67%→70%）",
      "upside_value": "391",
      "upside_pct": "70.5%",
      "upside_assumption": "22/6期実績を参考に設定（22/6期4月⇒期末にかけての売上原価伸長率を使用）"
    }
  ]
}
```

### JSON仕様

| フィールド | 必須 | 説明 |
|-----------|------|------|
| `main_message` | ○ | Main Message（Title 1）。最大70文字。「〜すべき」で終える |
| `chart_title` | ○ | Chart Title（Text Placeholder 2）。10〜25文字 |
| `period_label` | △ | 期間ラベル（現在はHTMLテーブル内では非表示、chart_titleに含める） |
| `unit_label` | △ | 単位表示。デフォルト「単位：百万円/(%)」 |
| `management_plan_label` | △ | マネジメント計画列のヘッダー。デフォルト「マネジメント\n計画\n(Base)」。`\n`で改行 |
| `our_plan_label` | △ | 弊社計画のヘッダー。デフォルト「弊社計画」 |
| `downside_label` | △ | Downside列のヘッダー。デフォルト「Downside case」 |
| `upside_label` | △ | Upside列のヘッダー。デフォルト「Upside case」 |
| `item_header` | △ | 項目列のヘッダー。デフォルト「項目（調整後）」 |
| `source` | ○ | 出典テキスト。「出典：」プレフィックスは自動付与 |
| `rows` | ○ | PL項目の配列。各要素は下表の構造 |

### rows配列の各要素

| フィールド | 必須 | 説明 |
|-----------|------|------|
| `item` | ○ | PL項目名（売上、売上原価、販管費、営業利益、EBITDA等） |
| `management_value` | ○ | マネジメント計画の数値（文字列） |
| `management_pct` | △ | マネジメント計画の構成比（「100%」「70.9%」等） |
| `downside_value` | ○ | Downside caseの数値（文字列） |
| `downside_pct` | △ | Downside caseの構成比 |
| `downside_assumption` | ○ | Downsideの前提条件テキスト |
| `upside_value` | ○ | Upside caseの数値（文字列） |
| `upside_pct` | △ | Upside caseの構成比 |
| `upside_assumption` | ○ | Upsideの前提条件テキスト |

### スクリプト実行コマンド

```bash
apt-get install -y -qq fonts-noto-cjk 2>/dev/null
pip install python-pptx playwright Pillow -q --break-system-packages
playwright install chromium 2>/dev/null

python <SKILL_DIR>/scripts/fill_current_period_forecast.py \
  --data {{WORK_DIR}}/forecast_data.json \
  --output {{OUTPUT_DIR}}/CurrentPeriodForecast_output.pptx \
  --brand stellar_aiz
```

`--brand roleup` を渡すと A4 横 + Yu Gothic UI + 茶系のテンプレ・配色で生成される。
`--template` は任意 (省略時は brand 別 curated テンプレを自動解決)。

※ `<SKILL_DIR>` は実際のスキルインストールパスに置き換えること。

### 出力確認

```bash
python -m markitdown {{OUTPUT_DIR}}/CurrentPeriodForecast_output.pptx
```

内容が正しく反映されているか確認し、ユーザーに提示する。

---

## 品質チェックリスト

PowerPoint生成後、以下を確認：

- [ ] Main Messageが70文字以内で「〜すべき」で終わっているか
- [ ] Chart Titleが10〜25文字でスライドの文脈を示しているか
- [ ] テーブルの全行（PL項目）が表示されているか（最終行が切れていないか）
- [ ] マネジメント計画・Downside・Upsideの数値が正しく反映されているか
- [ ] 構成比（%）が各数値の下に表示されているか
- [ ] 前提条件テキストが■マーカー付きで表示されているか
- [ ] 出典が左下に10ptで表示されているか
- [ ] PPTXのmarkitdown出力でプレースホルダーが残っていないか

---

## アセット

| ファイル名 | 用途 |
|---|---|
| `assets/stellar_aiz/current-period-forecast-template.pptx` | stella (16:9, Meiryo UI) 用テンプレート (company-overview派生) |
| `assets/stellar_aiz/layout.json` | stella 用座標 |
| `assets/roleup/current-period-forecast-template.pptx` | roleup (A4 横, Yu Gothic UI, 茶系) 用 curated テンプレート |
| `assets/roleup/layout.json` | roleup 用座標 |

## スクリプト

| ファイル名 | 用途 |
|---|---|
| `scripts/fill_current_period_forecast.py` | JSONデータからHTMLで当期着地見込みテーブルを生成→スクリーンショット→PPTXに挿入するスクリプト。`--brand` で stella / roleup を切替 |
| `scripts/build_roleup_template.py` | one-shot generator: stella テンプレから roleup curated テンプレを派生生成。テンプレ更新時のみ手動実行 |

## 参考

| ファイル名 | 内容 |
|---|---|
| `references/template-mapping.md` | テンプレートのShape名とスライド各セクションのマッピング表 |
