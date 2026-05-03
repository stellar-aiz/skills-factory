---
name: company-history-pptx
description: >
  会社沿革（Company History）のPowerPointスライドを生成するスキル。
  BDD（ビジネスデュー・ディリジェンス）において、対象会社の歴史・沿革を
  「年×概要」のテーブル形式で1枚のスライドにまとめる。
  PowerPointネイティブテーブルオブジェクトで生成する方式で、
  1〜15行の可変行数に柔軟対応する。14ptフォント、左寄せ。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「会社沿革」「沿革」「Company History」「沿革スライド」「沿革をパワポに」という言葉が出た場合
  - 「会社の歴史をスライドにして」「沿革をまとめて」「対象会社の歴史をパワポで」という要望
  - 「設立から現在までの歩みをスライドに」「会社のマイルストーンをスライドにして」という要望
  - BDD（ビジネスDD）の文脈で対象会社の沿革・歴史のスライド化を求められた場合
  - ユーザーがIM（インフォメーション・メモランダム）や会社HPの沿革情報を貼り付けて、スライド化を求めた場合
  - 既に沿革が整理されたテキストが提示され、PowerPoint化を求められた場合
  - 「年表」「タイムライン」を会社の歴史の文脈でスライドにしたいという要望
---

# 会社沿革 PowerPoint ジェネレーター

BDD（ビジネスデュー・ディリジェンス）において、対象会社の沿革を「年×概要」テーブル形式で
1枚のPowerPointスライドに整理するスキル。

---

## 会社沿革スライドとは

対象会社の設立から現在までの主要な出来事を時系列でまとめた、BDDの基礎資料。
IMや会社HPの沿革情報をもとに、年と概要の2列テーブルで構成する。

| 要素 | 定義 | ポイント |
|------|------|------|
| **メインメッセージ** | スライド最上部の要約メッセージ（Main Message） | 最大70文字。沿革全体のポイントを総括し「〜すべき」で締める。左寄せ |
| **チャートタイトル** | スライドのサブタイトル（Chart Title、デフォルト: 「会社沿革」） | 10〜20文字。Main Messageを補足する文脈を示す。左寄せ |
| **年** | 出来事が起きた年（西暦） | 「1989年」「2017年」のように「年」を付ける |
| **概要** | その年に起きた主要な出来事 | 1つの年に複数イベントがある場合は「、」で結合して1行にする |

### 記述ルール

- **メインメッセージ**: 最大70文字。沿革全体から導かれる結論・示唆を一言で総括し、必ず「〜すべき」で終える。ユーザーが指定した場合はそのまま使用、指定がない場合はドラフトして確認を取る
- **チャートタイトル**: 10〜20文字程度。Main Messageのテーマを補足する文脈フレーズ（デフォルト「会社沿革」）
- **行数**: 1〜15行。15行を超える場合は重要イベントに絞って15行以内に収める
- **年**: 西暦＋「年」で統一（例: 2017年）。同一年に複数イベントがある場合は1行にまとめる
- **概要**: 各イベントは簡潔に。日付がある場合は括弧書きで補足（例: 「本社移転（2月1日）」）
- **1つの年に複数イベント**: events配列に複数文字列を入れると「、」で結合して1行で表示される
- **フォントサイズ**: 14pt固定、左寄せ

### 典型的な沿革イベントの種類

BDDの文脈で重要な沿革イベント:
- 設立・法人化（設立日、資本金、設立地）
- 組織変更（有限会社→株式会社、商号変更等）
- 増資（資本金の変遷）
- 拠点関連（工場新設・増築、本社移転、支店開設）
- 経営陣交代（代表取締役の就退任）
- 認証取得（ISO等）
- M&A・事業提携
- 上場・IPO関連
- 海外展開
- その他重要イベント（受賞、記念事業、SDGs関連等）

---

## 入力パターンと処理フロー

### パターンA：IM・会社HP・議事録が入力された場合

**いきなりPowerPointを作成しない**

1. **Step 1: 沿革情報を抽出・整理する**（後述の抽出ガイドラインに従う）
2. **Step 2: Markdownでユーザーに提示**し、確認・修正を求める
3. **Step 3: ユーザーの承認後**、PowerPointを生成する

### パターンB：沿革データが整理済みで入力された場合

1. **Step 1: 内容を確認**し、年・概要が揃っているか確認
2. 不足・曖昧な点があれば修正提案を行う
3. **Step 2: 確認後、PowerPointを生成する**

---

## Step 1: IM・会社HPから沿革情報を抽出する

### 抽出のポイント

IMや会社HPの「沿革」「History」セクションから以下を抽出する：

- **年**: 出来事が起きた年。月日がある場合は概要の括弧書きに含める
- **概要**: 各年の主要イベント。同一年に複数あれば全て拾う
- **15行制限**: 行数が15を超える場合は、BDDの観点から重要度の高いイベントを優先する

### 優先度（15行超の場合の絞り込み基準）

1. 設立・法人化（必須）
2. 増資・資本金変遷（必須）
3. 経営陣交代（重要）
4. M&A・事業再編（重要）
5. 拠点の新設・移転（重要）
6. 上場・IPO（重要）
7. 認証取得・受賞（中程度）
8. 海外展開（中程度）
9. その他（低）

---

## Step 2: 沿革のMarkdown出力フォーマット

ユーザーに確認を求める際は、以下のフォーマットで出力する：

```markdown
## 会社沿革 整理結果

| 年 | 概要 |
|---|---|
| 1989年 | 有限会社○○を設立（1月31日）、資本金500万円 |
| 1992年 | 株式会社○○に変更（6月30日）、資本金2,000万円（増資） |
| 1993年 | 第一工場および事務所完成 |
| ... | ... |
```

確認メッセージ例：
> 上記の沿革整理でよろしいでしょうか？追加・修正があればお知らせください。確認後にPowerPointを生成します。

---

## Step 3: PowerPointの生成

### テンプレートの参照

テンプレートは brand 別に配置: `assets/<brand>/company-history-template.pptx`。
`--brand` 引数で切替（デフォルト `stellar_aiz`）、`--template` 省略時は brand から自動解決
（`brand_resolver.template_path()` 経由、curated rollup テンプレ未配置時は stella にフォールバック）。

```bash
TEMPLATE="<SKILL_DIR>/assets/stellar_aiz/company-history-template.pptx"  # 明示指定する場合
```

※ `<SKILL_DIR>` は実際のスキルインストールパスに置き換えること。
テンプレートのShape構造は `references/template-mapping.md` を参照。

### 沿革データのJSON化

沿革情報を `{{WORK_DIR}}/company_history_data.json` に以下の形式で保存する：

```json
{
  "main_message": "○○社は設立以来着実に事業基盤を拡大しており、成長ポテンシャルを評価すべき",
  "chart_title": "会社沿革",
  "history": [
    {
      "year": "1989年",
      "events": [
        "有限会社○○を設立（1月31日）",
        "資本金500万円"
      ]
    },
    {
      "year": "1992年",
      "events": [
        "株式会社○○に変更（6月30日）",
        "資本金2,000万円（増資）"
      ]
    },
    {
      "year": "1993年",
      "events": ["第一工場および事務所完成"]
    }
  ]
}
```

### JSONフィールド仕様

| フィールド | 型 | 必須 | 説明 |
|---|---|---|---|
| `main_message` | string | 必須 | メインメッセージ。最大70文字、「〜すべき」で締める |
| `chart_title` | string | 任意 | チャートタイトル。デフォルト「会社沿革」。10〜20文字 |
| `history` | array | 必須 | 沿革データの配列（1〜15要素） |
| `history[].year` | string | 必須 | 年（例: "1989年"） |
| `history[].events` | array of string | 必須 | その年のイベント一覧。複数ある場合は「、」で結合される |

### スクリプト実行コマンド

```bash
pip install python-pptx -q --break-system-packages

# ブランド未指定（デフォルト = stellar_aiz）
python <SKILL_DIR>/scripts/fill_company_history.py \
  --data {{WORK_DIR}}/company_history_data.json \
  --output {{OUTPUT_DIR}}/CompanyHistory_output.pptx

# Rollup 社向け
python <SKILL_DIR>/scripts/fill_company_history.py \
  --data {{WORK_DIR}}/company_history_data.json \
  --brand rollup \
  --output {{OUTPUT_DIR}}/CompanyHistory_output.pptx
```

※ `<SKILL_DIR>` は実際のスキルインストールパスに置き換えること。
`--brand` の有効値は `stellar_aiz` / `rollup`、デフォルトは `stellar_aiz`。
オーケストレーター（business-deepdive-agent / company-deepdive-agent 等）から呼ぶ場合、
parent は scope.json の `brand` を `--brand` で渡す。

### 出力確認

```bash
python -m markitdown {{OUTPUT_DIR}}/CompanyHistory_output.pptx
```

内容が正しく反映されているか確認し、ユーザーに提示する。

---

## 品質チェックリスト

PowerPoint生成後、以下を確認：

- [ ] スライドタイトルが正しく表示されているか
- [ ] 全ての沿革行がテーブルに表示されているか（切れていないか）
- [ ] 年が時系列順（昇順）に並んでいるか
- [ ] 1つの年に複数イベントがある場合、改行で正しく表示されているか
- [ ] フォントサイズが適切か（15行でもテーブルがスライド内に収まっているか）
- [ ] PPTXのmarkitdown出力でプレースホルダーが残っていないか

---

## アセット

| ファイル名 | 用途 |
|---|---|
| `assets/stellar_aiz/company-history-template.pptx` | Stellar AIZ 用テンプレート（16:9、Shape 構造は references/template-mapping.md 参照） |
| `assets/stellar_aiz/layout.json` | Stellar AIZ 用 layout（テンプレ駆動のため最小、コメントのみ） |
| `assets/rollup/layout.json` | Rollup 用 layout（V1 placeholder、テンプレは stella にフォールバック） |

V1 では Rollup 専用テンプレ pptx は配置せず、`brand_resolver.template_path()` のフォールバック経由で stella テンプレを流用する。
V2 で Rollup curated テンプレ（A4 横、Yu Gothic UI、褐色アクセント）を `assets/rollup/company-history-template.pptx` として導入する予定。

## スクリプト

| ファイル名 | 用途 |
|---|---|
| `scripts/fill_company_history.py` | JSONデータからPPTXネイティブテーブルを動的生成するスクリプト |

## 参考

| ファイル名 | 内容 |
|---|---|
| `references/template-mapping.md` | テンプレートのShape名と沿革スライド各セクションのマッピング表 |
