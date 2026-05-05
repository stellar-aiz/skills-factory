---
name: strategy-report-agent
description: >
  戦略コンサルティングファームが作成する企業調査レポート（BDD・競合分析・M&Aターゲット評価・
  新規参入調査・投資検討・経営企画向けベンチマーク等）のPowerPointデッキを生成する
  オーケストレータースキル。本スキル自体はスクリプトを持たず、Web検索と複数の既存スキルを
  呼び出してデッキ全体を組み立てる役割に特化する。
  v5.0では「事業環境の理解 → 対象会社の戦い方の理解」の4セクション構成へ思想転換。
  公開情報のみで書けない「戦略的示唆・推奨アクション」を廃止し、
  代わりに「今後検証すべき論点」でレポートを着地させる知的誠実性を重視した構成へ進化。
  デッキの深度（基本/標準/拡張/カスタム）を対話でユーザーに選択してもらい、
  対象会社の概要理解・外部環境・戦い方・検証論点の各セクションのスライドを順次生成する。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「競合調査」「競合分析」「Competitor Analysis」「Competitive Landscape」という言葉が出た場合
  - 「企業調査レポート」「企業分析レポート」「Strategy Research Report」の作成要望
  - 「BDD」「ビジネスDD」「ビジネスデューデリジェンス」でレポートを作りたい場合
  - 「M&Aターゲット評価」「買収候補分析」「投資対象レポート」の作成要望
  - 「新規参入調査」「市場参入レポート」「Entry Strategy Report」の要望
  - 「対象会社の競合を調べてスライドにして」「競合を分析して資料化して」という要望
  - ユーザーが対象会社名のみを伝えて「レポート作って」「分析して資料化して」と求めた場合
  - 戦略コンサルティング品質の企業調査PPTXを求められた場合
---

# 戦略コンサル向け企業調査レポート オーケストレーター v5.1

戦略コンサルティングファーム（BCG / McKinsey / ベイン / 日系コンサル 等）が作成する
**企業調査PowerPointレポート**を自動生成するオーケストレータースキル。

## v5.1 の追加要素（v5.0 からの差分）

情報信頼性とビジュアル品質を**人手レビューに依存せず**担保するため、2段のレビュアーを組み込んだ:

1. **ファクトチェック（Step 2.5）**: JSON組成後・スライド生成前に、`fact-check-reviewer` で
   Web取得情報を再検索で裏取り。疑わしい主張（数値・シェア・日付・固有名詞）をフラグ
2. **ビジュアル品質レビュー（Step 最終+1）**: マージ後、`visual-quality-reviewer` で
   デッキ全ページをLibreOffice経由で画像化 → 文字溢れ・重なり・密度過多等を Claude 自身が
   multimodal で目視判定 → `high` 重大度は自動で該当スライドを再生成し再マージ（最大2ラウンド）

v5.0 の思想（事業環境の理解→対象会社の戦い方）はそのまま維持。上記は**品質担保の付加レイヤー**。

## 対応する調査タイプ

| 調査タイプ | 典型的な用途 | 推奨デッキ深度 |
|---|---|---|
| **BDD（ビジネスDD）** | M&A時のターゲット事業性評価（フェーズ1：公開情報） | 標準 or 拡張 |
| **競合分析** | 経営企画・中計策定のための競合ベンチマーク | 標準 or 拡張 |
| **M&Aターゲット評価** | 買収候補の魅力度・リスク評価 | 拡張 |
| **新規参入調査** | 新事業・新市場への参入可能性検討 | 拡張 |
| **投資検討** | VC・PE・事業会社の投資判断 | 標準 or 拡張 |
| **経営企画ベンチマーク** | 自社の相対ポジション把握 | 標準 |

---

## v5.0の思想転換（最重要）

### 従来（v4.x）の思想と問題点

v4.xは「**事実 → 示唆 → アクション**」という戦略示唆を含む構成だったが、本スキルは
**Web情報（公開情報）のみ**で作成されるため、以下の問題があった:

- 内部の経営判断・KPI・リソース配分の情報がないと、まともな「〜すべき」は提言できない
- 書いたとしても「DX推進すべき」「海外展開すべき」等の一般論に陥りがち
- 戦略コンサルの実務でも、公開情報フェーズで提言は書かない（IM・マネジメントインタビューを経て初めて書く）

### v5.0の新しい思想

**「事業環境の理解 → 対象会社の戦い方の理解 → 今後検証すべき論点の提示」**

- **Web情報で書けること**に徹底的に注力する（無理に提言を書かない）
- 対象会社の概要・事業環境・戦い方を**事実記述ベース**で深く理解する
- 公開情報で見えなかった点・仮説にとどまる点は「**今後検証すべき論点**」として率直に提示
- 「〜すべき」と断定せず、「〜という構造になっている」「〜を確認する必要がある」と記述する
- これは戦略コンサルの **BDDフェーズ1（公開情報調査）の正しい姿** に沿った設計

| v4.x | v5.0 |
|---|---|
| 事実 → 示唆 → アクション | **事業環境の理解 → 対象会社の戦い方の理解** |
| recommendation-action を結論に | **issue-risk-list（検証論点一覧）で着地** |
| 「〜すべき」で締める | 「〜という構造である」「〜を確認する必要がある」で記述する |
| value-chain-pptx → 戦略示唆 | value-chain-pptx → **業界の利益構造（外部環境）** |
| value-chain-matrix → （未組込） | **value-chain-matrix → 対象会社のポジショニング記述（戦い方）** |
| business-model → （未組込） | **business-model → 対象会社の取引構造理解（概要）** |

---

## デッキ構成（v5.0: 4セクション構成）

### セクション設計

| Section | テーマ | 色 | 含まれるスライド |
|---|---|---|---|
| **Section 1** | 対象会社の概要理解 | 紺 | 対象会社プロファイル / 事業ポートフォリオ / ビジネスモデル |
| **Section 2** | 外部環境（業界・市場・競合） | 青 | PEST / Five Forces / 業界バリューチェーン・利益プール / 市場規模 / 市場シェア / ポジショニングマップ / 競合比較サマリー / 財務ベンチマーク / 各競合プロファイル |
| **Section 3** | 対象会社の戦い方 | 緑 | SWOT / バリューチェーン・ポジショニング・マトリクス / 成長ドライバー分析 |
| **Section 4** | 今後検証すべき論点 | オレンジ | 検証論点一覧 |

### Section 2内の論理的グルーピング（並び順で表現）

Section 2は10スライド前後と厚いため、**並び順で以下の論理構造を表現**する（中扉は追加しない）:

```
[マクロ]     PEST
[業界構造]   Five Forces → 業界バリューチェーン・利益プール
[市場]       市場規模 → 市場シェア → ポジショニングマップ
[競合]       競合比較サマリー → 財務ベンチマーク → 各競合プロファイル
```

この並び順でマクロから徐々にズームイン（マクロ→業界→市場→競合）する読み手の思考に沿う。

---

### 🏃 基本版（7-8枚 / Quick）

最小構成。時間制約や情報量が限定的な場合に使用。中扉なし。

| # | スライド | スキル |
|---|---|---|
| 1 | エグゼクティブサマリー | `executive-summary-pptx` |
| 2 | 目次 | `table-of-contents-pptx` |
| 3 | 対象会社プロファイル | `customer-profile-pptx` |
| 4 | 事業ポートフォリオ | `business-portfolio-pptx` |
| 5 | ビジネスモデル | `business-model-pptx` ⭐ |
| 6 | 市場環境（市場規模） | `market-environment-pptx` |
| 7 | 競合比較サマリー | `competitor-summary-pptx` |
| 8 | **今後検証すべき論点** | `issue-risk-list-pptx` ⭐ |
| (末尾) | データアベイラビリティ（オプション） | `data-availability-pptx` |

### 📊 標準版（13-15枚 / Standard）

実務で使われる定番構成。**推奨**。中扉なし（セクション境界は目次で示す）。

| # | スライド | スキル |
|---|---|---|
| 1 | エグゼクティブサマリー | `executive-summary-pptx` |
| 2 | 目次 | `table-of-contents-pptx` |
| 3 | 対象会社プロファイル | `customer-profile-pptx` |
| 4 | 事業ポートフォリオ | `business-portfolio-pptx` |
| 5 | ビジネスモデル | `business-model-pptx` ⭐ |
| 6 | マクロ環境 PEST分析 | `pest-analysis-pptx` |
| 7 | 業界構造 Five Forces | `five-forces-pptx` |
| 8 | 市場環境(市場規模) | `market-environment-pptx` |
| 9 | 市場シェア分析 | `market-share-pptx` |
| 10 | 競合比較サマリー | `competitor-summary-pptx` |
| 11 | 財務ベンチマーク | `financial-benchmark-pptx` |
| 12 | 対象会社 SWOT | `swot-pptx` |
| 13 | **バリューチェーン・ポジショニング・マトリクス** | `value-chain-matrix-pptx` ⭐ |
| 14 | 成長ドライバー分析 | `growth-driver-pptx` |
| 15 | **今後検証すべき論点** | `issue-risk-list-pptx` ⭐ |
| (末尾) | データアベイラビリティ（オプション） | `data-availability-pptx` |

### 🎯 拡張版（20-25枚 / Comprehensive）

戦略コンサル品質。マクロ→業界→市場→競合→自社の戦い方→論点の完全な俯瞰分析。
中扉4枚（各セクション冒頭）で構造を明示。

| # | スライド | スキル |
|---|---|---|
| 1 | エグゼクティブサマリー | `executive-summary-pptx` |
| 2 | 目次 | `table-of-contents-pptx` |
| 3 | (中扉) 対象会社の概要理解 | `section-divider-pptx` |
| 4 | 対象会社プロファイル | `customer-profile-pptx` |
| 5 | 事業ポートフォリオ | `business-portfolio-pptx` |
| 6 | ビジネスモデル | `business-model-pptx` ⭐ |
| 7 | (中扉) 外部環境（業界・市場・競合） | `section-divider-pptx` |
| 8 | マクロ環境 PEST分析 | `pest-analysis-pptx` |
| 9 | 業界構造 Five Forces | `five-forces-pptx` |
| 10 | **業界バリューチェーン・利益プール** | `value-chain-pptx` ⭐ |
| 11 | 市場環境(市場規模) | `market-environment-pptx` |
| 12 | 市場シェア分析 | `market-share-pptx` |
| 13 | ポジショニングマップ | `positioning-map-pptx` |
| 14 | 競合比較サマリー | `competitor-summary-pptx` |
| 15 | 財務ベンチマーク | `financial-benchmark-pptx` |
| 16〜 | 各競合プロファイル | `customer-profile-pptx` or `company-overview-pptx-v2` |
| ## | (中扉) 対象会社の戦い方 | `section-divider-pptx` |
| ## | 対象会社 SWOT | `swot-pptx` |
| ## | **バリューチェーン・ポジショニング・マトリクス** | `value-chain-matrix-pptx` ⭐ |
| ## | 成長ドライバー分析 | `growth-driver-pptx` |
| ## | (中扉) 今後検証すべき論点 | `section-divider-pptx` |
| ## | **今後検証すべき論点** | `issue-risk-list-pptx` ⭐ |
| (末尾) | データアベイラビリティ | `data-availability-pptx` |

⭐ = **v5.0で組み込まれた新スキル（`business-model-pptx` / `value-chain-matrix-pptx` / `issue-risk-list-pptx`）**

**最終結合**: 全モード共通で `merge-pptxv2` で1つのPPTXにまとめる

---

## 処理フロー（パターン別）

### パターンA：対象会社名のみ入力された場合

**いきなりPowerPointを作成しない。必ず以下の順で進行する。**

1. **Step 0**: 調査タイプとデッキ深度をユーザーに確認
2. **Step 1**: Web検索で情報収集
3. **Step 2**: データアベイラビリティ整理
4. **Step 3**: Markdownで全情報をユーザーに提示 → 承認
5. **Step 4**: 競合ごとの詳細形式をユーザーに確認（財務データ欠損時のみ）
6. **Step 5**: エグゼクティブサマリーのKey Findings 3-5個を整理（事実記述ベース）
7. **Step 6**: 検証論点 3-7個を整理（v5.0で変更：推奨アクション → 検証論点）
8. **Step 7〜N**: 選択されたモードに応じて各スライド生成
9. **Step 最終**: `merge-pptxv2` で結合

### パターンB：対象会社名＋競合リストが指定された場合

- Step 0 は深度確認のみ、競合企業の特定はスキップ

### パターンC：全情報が整理済みで入力された場合

- Step 0-4をスキップし、Step 7以降に進む

---

## Step 0: 調査タイプとデッキ深度の確認

<!-- source: skills/_common/prompts/step0_brand_clarification.md (manual sync until D2) -->

### Step 0.0-pre: ブランド確認（必須）

本デッキの**出力ブランド**（クライアント別 PPTX フォーマット）を確定し、後続 Step で参照する `scope.brand` 文字列として会話メモリに保持する。共通原則・AskUserQuestion テンプレ・自由記述ハンドリング・unsupported skill fallback の詳細仕様は `skills/_common/prompts/step0_brand_clarification.md` を正本とする。

実装パターン（agnostic、`_discover_brands()` で動的取得）:

```python
import json, os, sys
sys.path.insert(0, os.path.join("{{SKILL_DIR}}", "..", "_common", "lib"))
from brand_resolver import _discover_brands, _BRANDS_DIR

discovered = _discover_brands()  # 例: ('roleup', 'stellar_aiz')
options = []
for brand_id in discovered:
    with open(os.path.join(_BRANDS_DIR, brand_id, "theme.json")) as f:
        theme_data = json.load(f)
    label = theme_data.get("label", brand_id)
    if brand_id == "stellar_aiz":
        label += " (Recommended)"
    options.append({"label": label, "description": f"id={brand_id}"})

AskUserQuestion(
    question="このデッキはどのクライアント・ブランドのフォーマットで出力しますか？",
    header="ブランド", options=options, multiSelect=False,
)
# 確定値は scope.brand に文字列保存（既定 "stellar_aiz"）。
# 「Other」で _discover_brands に含まれない id を入力された場合は AskUserQuestion を再実行。
```

<!-- source: skills/_common/prompts/step0_scope_clarification.md (manual sync until D2) -->

共通原則・Step 0.5（事前スコーピング Web 検索）・`included_business_models` / `excluded_segments` フィールドの定義は `skills/_common/prompts/step0_scope_clarification.md` を正本とする。本 SKILL.md には strategy-report-agent 固有の質問項目と Step 0.5 の運用上の注意のみ記載する。

### Step 0.0: 固有質問の確定

```python
questions = [
    {
        "question": "どの調査タイプですか？",
        "options": [
            "A. 競合分析レポート（経営企画・中計用）",
            "B. BDD（ビジネスデューデリジェンス・M&A向け）",
            "C. 新規参入調査",
            "D. 投資検討レポート（VC・PE・事業会社）",
            "E. M&Aターゲット評価"
        ],
        "type": "single_select"
    },
    {
        "question": "どのデッキ深度にしますか？",
        "options": [
            "A. 基本版（7-8枚） - 短時間で要点のみ",
            "B. 標準版（13-15枚） - 実務で使われる定番（推奨）",
            "C. 拡張版（20-25枚） - 完全な俯瞰分析（中扉あり）",
            "D. カスタム - 含めるスライドを個別に選ぶ"
        ],
        "type": "single_select"
    },
    {
        "question": "データアベイラビリティスライドの配置は？",
        "options": [
            "A. デッキ末尾（Appendix）- 推奨",
            "B. エグゼクティブサマリー直後（前提を最初に示す）",
            "C. 配置しない"
        ],
        "type": "single_select"
    }
]
```

### カスタムモードの場合（multi-select）

```python
questions = [{
    "question": "含めるスライドを選んでください（エグサマ・目次・対象会社プロファイル・検証論点は必須）",
    "options": [
        "事業ポートフォリオ",
        "ビジネスモデル",
        "マクロ環境 PEST分析",
        "業界構造 Five Forces",
        "業界バリューチェーン・利益プール",
        "市場規模推移",
        "市場シェア分析",
        "ポジショニングマップ",
        "競合比較サマリー",
        "財務ベンチマーク",
        "各競合プロファイル",
        "対象会社 SWOT",
        "バリューチェーン・ポジショニング・マトリクス",
        "成長ドライバー分析",
        "中扉スライド（セクション区切り）",
        "データアベイラビリティ"
    ],
    "type": "multi_select"
}]
```

### Step 0.5: 事前スコーピング Web 検索（必須）

調査タイプ・デッキ深度の確認後、Step 1 の Web 検索に進む前に **対象会社が属する業界の構造ザックリ把握用の Web 検索を 1〜2 件** 走らせる。
目的は「市場シェア・競合比較で並べる対象が、収益構造の異なる事業モデルを混在させていないか」を Step 1 の本格的な情報収集前に検知すること。

検索クエリ例:
- `<対象会社の業界> 業界構造 / バリューチェーン / プレイヤー類型`
- `<対象会社の業界> 競合 定義 / 市場区分`

異種事業モデル（タクシー事業者 vs 配車アプリ、半導体装置 vs IDM vs ファブレス vs ファウンドリ等）が併存する業界では、`AskUserQuestion` で境界確認:

```
「<対象会社の業界>」には収益構造の異なる事業モデルが併存しています。本レポートではどの層を競合として扱いますか？
A. 対象会社と同じ事業モデルのみ（推奨）
B. 異なる事業モデルも含める（業界全体の競争環境として描く場合）
C. その他（自由記述）
```

異種併存の典型例（共通プロンプト参照）: タクシー / 半導体 / 教育 / 飲食 / 物流 / 金融。
ユーザーが冒頭で対象会社と競合リストを明示している場合（パターンB）は Step 0.5 をスキップしてよいが、競合リストに事業モデル混在がないかは確認すること。

### Step 0 出力（scope の保存）

調査タイプ・デッキ深度・データアベイラビリティ位置・事業モデル境界を会話メモリに保持し、後続 Step で参照する。market-overview-agent と異なり scope.json への永続化は必須ではないが、`included_business_models` / `excluded_segments` の決定は必ず Step 1 の Web 検索クエリに反映する責務がある（fill_*.py は scope を読まない。`skills/_common/references/orchestrator_contract.md` 参照）。

`excluded_segments` が空配列でない場合は、Step 2 のデータアベイラビリティと Step 5 の Key Findings 冒頭で「本レポートでは <excluded_segments> を対象外として除外している」旨を明記する。

---

## Step 1: Web検索で情報収集

各社・各分析項目について以下を収集:

| 項目 | 用途スキル | セクション | 情報源の優先度 |
|---|---|---|---|
| 事業内容・財務基本情報 | customer-profile / company-overview-v2 | S1 | 公式HP > IR > 業界レポート |
| セグメント別売上 | business-portfolio | S1 | 有報・決算短信 |
| **取引構造（サプライヤー・顧客）** | **business-model** ⭐ | **S1** | **公式HP・有報・IR資料・業界レポート** |
| マクロ要因 | pest-analysis | S2 | 政府統計・シンクタンク |
| 業界5競争要因 | five-forces | S2 | 業界レポート・有報のリスク情報 |
| **バリューチェーン・利益プール** | **value-chain** | **S2** | **業界レポート・有報の原価率内訳** |
| 市場規模・シェア | market-environment / market-share | S2 | 矢野経済・富士経済・業界統計 |
| 2軸ポジショニング属性 | positioning-map | S2 | 各社HP・業界レポート |
| 財務指標 | financial-benchmark | S2 | 有報・決算短信・SPEEDA |
| 強み・弱み・機会・脅威 | swot | S3 | HP・IR・業界レポート・ニュース |
| **対象会社のバリューチェーン上の工程・提携・競合関係** | **value-chain-matrix** ⭐ | **S3** | **公式HP・IR・業界レポート・ニュース** |
| 売上/利益要因分解 | growth-driver | S3 | 決算説明会資料・IR Q&A |
| **公開情報で見えない・仮説レベルに留まる事項** | **issue-risk-list** ⭐ | **S4** | **データアベイラビリティの裏返しとして整理** |

---

## Step 2: データアベイラビリティを整理

カテゴリ×項目×ステータス（✓取得済/△一部取得/✗未取得）×データソースで整理。
`data-availability-pptx` のJSON形式に整理しておく。

**重要**: データアベイラビリティの ✗/△ 項目は、そのまま Step 6 の**検証論点の種**になる。

---

## Step 2.5: ファクトチェック（v5.1新規）

<!-- source: skills/_common/prompts/step2_5_factcheck_invocation.md (manual sync until D2) -->

スライド生成前に、Web取得情報の真偽を `fact-check-reviewer` スキルで裏取りする。

### Step 2.5-a: スコープをユーザーに選ばせる

`AskUserQuestion` で以下から選択させる:

| 選択肢 | 内容 |
|---|---|
| **high_risk**（推奨） | 数値・シェア・市場規模・日付・固有名詞のみを検証。時間とAPIコストを抑える |
| **all** | 上記＋テキスト主張も全件検証。網羅性最大だが時間がかかる |
| **skip** | ファクトチェックを省略して Step 3 へ進む |

### Step 2.5-b: fact-check-reviewer を起動（skip 以外の場合）

Step 1 で整理した情報を一旦 `{{WORK_DIR}}/data_*.json` として書き出してから起動する。
本スキルの SKILL.md に従って裏取りを実行し、`{{WORK_DIR}}/fact_check_report.json` を得る。

### Step 2.5-c: フラグ項目を Step 3 の Markdown に統合

`fact_check_report.json` の `flags[]` のうち `severity=high` と `medium` を、次の Step 3 で
ユーザーに提示する Markdown に**「要確認項目」セクション**として差し込む。ユーザーは
調査結果の確認と併せて、ファクトフラグの扱い（JSON修正 / ソース追加 / スキップ）を決める。

`overall_verdict=pass` の場合はフラグ提示を省略し、Step 3 の末尾に「ファクトチェック結果: 問題なし」の一文のみ添える。

---

## Step 3: Markdownでユーザーに確認

調査結果と推奨スライド構成を提示し、承認を得る。

---

## Step 4: 競合の詳細形式確認（必要時のみ）

財務データ欠損のある競合のみ:
- A. customer-profile（業績チャート付き、欠損年度N/A）
- B. company-overview-v2（HP画像付き、チャートなし）

---

## Step 5: エグゼクティブサマリーのKey Findings整理

**v5.0のトーン**: 「〜すべき」ではなく、**事実記述ベース**で Key Findings を書く。

5 Findings パターン推奨:
1. **対象会社の概要**: 事業規模・収益構造・顧客構造（例: 「売上XXX億円、B2Bが8割、上位顧客3社で売上の60%を占める」）
2. **外部環境**: マクロ/業界/市場の構造的特徴（例: 「業界はFive Forcesで買い手交渉力が強く、利益プールは川下に偏る」）
3. **競合ポジション**: 相対ポジション（例: 「国内シェア3位、上位2社とは規模で2倍差」）
4. **対象会社の戦い方**: 工程ポジション・強み（例: 「川中工程に特化し、設計・製造を内製化。上流R&Dは大学提携」）
5. **検証論点**: 公開情報で確認できなかった主要論点（例: 「収益構造の持続性は顧客依存度の高さが論点として残る」）

---

## Step 6: 検証論点整理（v5.0新規・旧「推奨アクション」を置換）

`issue-risk-list-pptx`（**組織スキル**: `/mnt/skills/organization/issue-risk-list-pptx`）用に、3-7個の検証論点を整理する。

### 論点の組み立て方（重要）

**データアベイラビリティで ✗/△ の項目を論点の起点にする。**

| 列 | 内容 | 例 |
|---|---|---|
| # | 通し番号 | 1 |
| カテゴリ | 論点の領域 | 収益構造 / 顧客基盤 / オペレーション / 組織・人材 / 競争優位 / M&A関連 |
| 論点 | Web情報で見えなかった問い | 上位顧客3社への依存度は持続可能か |
| 仮説 | Web情報から見えた兆候 | 過去5年間顧客上位は固定。長期契約の可能性あり |
| 確認方法 | IM・マネジメントインタビューで聞くこと | 顧客別売上推移、契約更新条件、代替可能性 |
| 優先度 | 高/中/低 | 高 |

### トーン

- 「〜を確認する必要がある」「〜は論点として残る」「〜の実態把握が必要」
- 「〜すべき」は絶対に使わない（公開情報では断定できない）

### v5.0で `recommendation-action-pptx` を使わない理由

- Web情報のみで「推奨アクション」を書くと一般論に陥る
- 戦略コンサルのBDDフェーズ1でも、公開情報だけで提言はしない（マネジメントインタビューが前提）
- 代わりに「検証論点」で着地させることで、知的誠実性を保つ

---

## Step 7〜N: スライド生成

選択されたモード・スライドに応じて、以下のスキルを順次呼び出す。

### ⚠️ 最重要：ファイル名の番号は「最終並び順」と一致させること

**過去の失敗パターン（絶対に繰り返さない）:**

ファイル名の番号（例 `slide_22_*.pptx`）を「生成した順番のID」として扱い、マージ時に数字順でソートしてそのまま並べた結果、**中扉がセクションの末尾寄りに配置される**というデッキ全体が破綻するミスが発生した。

**徹底すべきルール:**

1. **ファイル名の番号 = デッキ上の最終的な通し番号** として最初から割り振る
2. **生成する順序とファイル名の番号は同じである必要はない**（生成は任意の順でOK、ただし番号は最終順序）
3. スライドを生成する前に、**デッキ全体の通し番号→ファイル名の対応表**を作成し、それに従ってファイル名を決める
4. 途中でスライドを追加・削除する場合、影響を受けるファイルを**全てリネーム**して番号を詰めること（歯抜けや逆転を絶対に残さない）

**番号割当表の作成例（生成前に必須）:**

```
01 → slide_01_exec_summary.pptx
02 → slide_02_toc.pptx
03 → slide_03_section1_target.pptx    [中扉: Section 1冒頭]
04 → slide_04_target_profile.pptx
05 → slide_05_target_portfolio.pptx
06 → slide_06_target_business_model.pptx    [ビジネスモデル ⭐v5.0]
07 → slide_07_section2_external.pptx   [中扉: Section 2冒頭]
...
NN → slide_NN_section3_strategy.pptx   [中扉: Section 3冒頭]
NN → slide_NN_target_swot.pptx
NN → slide_NN_value_chain_matrix.pptx  [⭐v5.0]
NN → slide_NN_growth_driver.pptx
NN → slide_NN_section4_issues.pptx     [中扉: Section 4冒頭]
NN → slide_NN_issue_risk_list.pptx     [⭐v5.0]
```

### 中扉（section-divider）配置の絶対ルール

**中扉は、そのセクションが指すコンテンツ群の「最初」に配置する。末尾ではない。**

正しい例:
```
[中扉: Section 2 外部環境] ← セクション冒頭
  PEST分析
  Five Forces
  業界バリューチェーン・利益プール
  市場規模
  市場シェア
  ポジショニングマップ
  競合比較サマリー
  財務ベンチマーク
  各競合プロファイル
[中扉: Section 3 対象会社の戦い方] ← 次セクション冒頭
  ...
```

誤った例（過去に発生したバグ）:
```
PEST分析
Five Forces
...
競合比較サマリー
[中扉: Section 2] ← 末尾にあるのは論理的に誤り
  ...
```

中扉を生成するタイミングは任意だが、**ファイル名の番号は必ずセクションの最初のコンテンツスライドの直前の番号にすること**。

### スライド生成開始前: brand fallback バッファ初期化（必須）

Step 0.0-pre で確定した `scope.brand` を使い、未対応 fill 検出用の warning バッファを初期化する。各 fill 起動前に `resolve_fill_brand_with_warning()` を呼び、未対応スキルでは `stellar_aiz` に fallback + warning を buffer に蓄積する（`skills/_common/lib/orchestrator_helpers.py` 参照）。

```python
import os, sys, subprocess
sys.path.insert(0, os.path.join("{{SKILL_DIR}}", "..", "_common", "lib"))
from orchestrator_helpers import (
    resolve_fill_brand_with_warning,
    append_brand_warnings_to_merge_file,
)

scope_brand = "stellar_aiz"  # Step 0.0-pre で確定した値（会話メモリから）
brand_warnings: list = []
```

### 共通パターン

各 fill 起動前に `resolve_fill_brand_with_warning(skill_dir, scope_brand, brand_warnings)` で fill に渡す brand を確定する。supported なら `scope_brand` がそのまま、未対応なら `stellar_aiz` が返り `brand_warnings` に `brand_fallback` エントリが追記される。

```python
skill_dir = "<SKILL_DIR>"
fill_brand = resolve_fill_brand_with_warning(skill_dir, scope_brand, brand_warnings)
subprocess.run([
    "python", os.path.join(skill_dir, "scripts", "fill_[skill].py"),
    "--brand", fill_brand,
    "--data", "{{WORK_DIR}}/[skill]_data.json",
    "--template", os.path.join(skill_dir, "assets", "[skill]-template.pptx"),
    "--output", "{{WORK_DIR}}/slide_NN_[name].pptx",
], check=True)
```

bash で直接書く場合（既存スキル踏襲、warning fallback は使わない場合）:

```bash
pip install python-pptx -q --break-system-packages

cat > {{WORK_DIR}}/[skill]_data.json <<'EOF'
{ ... }
EOF

python <SKILL_DIR>/scripts/fill_[skill].py \
  --brand stellar_aiz \
  --data {{WORK_DIR}}/[skill]_data.json \
  --template <SKILL_DIR>/assets/[skill]-template.pptx \
  --output {{WORK_DIR}}/slide_NN_[name].pptx
```

### 推奨スライド順序（拡張版・最大25枚想定）

| # | 出力ファイル | スキル |
|---|---|---|
| 01 | `slide_01_exec_summary.pptx` | executive-summary-pptx |
| 02 | `slide_02_toc.pptx` | table-of-contents-pptx |
| 03 | `slide_03_section1_target.pptx` | section-divider-pptx (Section 01: 対象会社の概要理解) |
| 04 | `slide_04_target_profile.pptx` | customer-profile-pptx |
| 05 | `slide_05_target_portfolio.pptx` | business-portfolio-pptx |
| 06 | `slide_06_target_business_model.pptx` | **business-model-pptx** ⭐ (組織スキル) |
| 07 | `slide_07_section2_external.pptx` | section-divider-pptx (Section 02: 外部環境) |
| 08 | `slide_08_pest.pptx` | pest-analysis-pptx |
| 09 | `slide_09_five_forces.pptx` | **five-forces-pptx** (組織スキル) |
| 10 | `slide_10_value_chain.pptx` | **value-chain-pptx** (業界の利益プール分析) |
| 11 | `slide_11_market_size.pptx` | market-environment-pptx |
| 12 | `slide_12_market_share.pptx` | market-share-pptx |
| 13 | `slide_13_positioning.pptx` | positioning-map-pptx |
| 14 | `slide_14_competitor_summary.pptx` | competitor-summary-pptx |
| 15 | `slide_15_fin_benchmark.pptx` | financial-benchmark-pptx |
| 16+ | `slide_NN_competitor_X.pptx` | customer-profile or overview-v2 |
| ## | `slide_NN_section3_strategy.pptx` | section-divider-pptx (Section 03: 対象会社の戦い方) |
| ## | `slide_NN_target_swot.pptx` | swot-pptx |
| ## | `slide_NN_value_chain_matrix.pptx` | **value-chain-matrix-pptx** ⭐ |
| ## | `slide_NN_growth_driver.pptx` | growth-driver-pptx |
| ## | `slide_NN_section4_issues.pptx` | section-divider-pptx (Section 04: 今後検証すべき論点) |
| ## | `slide_NN_issue_risk_list.pptx` | **issue-risk-list-pptx** ⭐ (組織スキル) |
| 最終 | `slide_ZZ_data_avail.pptx` | data-availability-pptx |

### JSON生成の共通原則

- **対象会社名の一貫性**: 全スキルで同じ表記
- **年度の一貫性**: 全スライドで同じ基準年（実行年の前年）
- **色の一貫性**:
  - 対象会社の強調色: #E15759（赤系）に統一
  - 中扉と目次のセクション色: 同じセクション番号で同じ色
- **エグゼクティブサマリー**: 他の全スライドの事実を統合した概観
- **検証論点**: データアベイラビリティと連動。見えなかった点を率直に列挙

### main_message 共通ルール

<!-- source: skills/_common/prompts/main_message_principles.md (manual sync until D2) -->

#### 基本ルール

##### ルール1-A: 長さは **65 文字以内**（厳格）

- 句読点・記号・スペースを含めて 65 文字以内
- テンプレート最上部のメッセージ枠が固定幅のため、超えた場合は要約や段落分けではなく **書き直し**
- 65 文字を 1 字でも超えた状態で `fill_*.py` に渡すと ValueError で hard-fail する

##### ルール1-B: トーンは **事実記述ベース**（「〜すべき」禁止）

- 公開情報のみで断定できないアクションや戦略示唆は書かない
- 不明な点は「〜は公開情報からは確定できず追加調査が必要」と率直に書く（検証論点として明示）

**例**:
- ✗ 「対象会社は海外展開を加速すべき」（公開情報では断定不可）
- ✓ 「対象会社は国内売上比率が 90% と高く、海外展開の実績は限定的である」（事実記述）
- ✓ 「対象会社の海外展開方針は Web 情報では限定的、マネジメントインタビューで確認が必要」（検証論点）

#### 65 字オーバー時の短縮原則 4 つ

LLM が初稿で 65 字を超えた場合、以下の順で書き直す:

1. **主語は 1 つだけ** — 「市場は」「主要プレイヤーは」「対象会社は」のいずれか 1 つに絞る
2. **修飾語を削除** — 「主要な」「重要な」「大きな」「急速な」等の主観的な修飾語を落とす
3. **数値は 1 つだけ残す** — CAGR と シェアを両方載せず、より重要な 1 つを選ぶ
4. **結論を述べる、根拠は本文スライドに任せる** — 「〜だから〜である」の前段を切り、結論部のみ残す

**例（市場概要系）**:
- ✗（85字）「国内 HR Tech クラウド市場は CAGR 約 32% で急拡大しており、勝つには『データ蓄積』『日本仕様適合』『統合エコシステム』の 3 つの KBF を抑える必要がある」
- ✓（44字）「国内 HR Tech クラウド市場は CAGR 約 32% で急拡大、3 つの KBF が勝ち筋を決める」

**例（対象会社系）**:
- ✗（72字）「対象会社は売上 500 億円規模で国内シェア 15% を持ち、北米進出を 2024 年から開始したが現状は売上比率 5% に留まる」
- ✓（38字）「対象会社は国内シェア 15% で 2 位、海外売上比率は 5% に留まる」

---

## Step 最終-1: マージ順序の照合チェック（必須）

<!-- source: skills/_common/references/orchestrator_contract.md (manual sync until D2) -->
<!-- merge_order.json の正規スキーマ（category 値域含む）は上記参照 -->

**マージコマンドを組み立てる前に、必ず以下の照合を行うこと。** 過去、このステップを飛ばしてファイル名の数字順にそのままマージした結果、中扉が末尾に流れるミスが発生した。

### チェック手順

1. **ファイル一覧を取得**: `ls {{WORK_DIR}}/*.pptx | sort` で全スライドをリストアップ
2. **目次（TOC）スライドの章構成を再確認**: `table-of-contents-pptx`で定義した `sections` 配列と、中扉の `section_number` が一致しているか
3. **中扉配置ルールの検証**: 各中扉ファイル（`slide_NN_sectionX_*.pptx`）が、その配下コンテンツ（同じセクション番号のスライド群）の**最小番号の直前**に位置しているか
4. **ファイル番号の歯抜け・逆転を確認**: 01, 02, 03... と連続しているか

### 照合表のフォーマット（マージ前にMarkdownで明示）

マージコマンドを打つ前に、以下の形式で照合表を出力し、**自己チェックしてから**実行すること。

```markdown
## マージ順序照合表

| 通し番号 | ファイル名 | 種別 | 所属セクション | 備考 |
|---|---|---|---|---|
| 01 | slide_01_exec_summary.pptx | エグサマ | 冒頭 | - |
| 02 | slide_02_toc.pptx | 目次 | 冒頭 | TOCの section[0] = "対象会社の概要理解" |
| 03 | slide_03_section1_target.pptx | **中扉** | Section 1冒頭 ✓ | section_number=1 |
| 04 | slide_04_target_profile.pptx | コンテンツ | Section 1 | - |
| 05 | slide_05_target_portfolio.pptx | コンテンツ | Section 1 | - |
| 06 | slide_06_target_business_model.pptx | コンテンツ | Section 1 | ⭐v5.0 |
| 07 | slide_07_section2_external.pptx | **中扉** | Section 2冒頭 ✓ | section_number=2 |
| ... | ... | ... | ... | ... |
| NN | slide_NN_section3_strategy.pptx | **中扉** | Section 3冒頭 ✓ | section_number=3 |
| NN | slide_NN_target_swot.pptx | コンテンツ | Section 3 | - |
| NN | slide_NN_value_chain_matrix.pptx | コンテンツ | Section 3 | ⭐v5.0 |
| NN | slide_NN_growth_driver.pptx | コンテンツ | Section 3 | - |
| NN | slide_NN_section4_issues.pptx | **中扉** | Section 4冒頭 ✓ | section_number=4 |
| NN | slide_NN_issue_risk_list.pptx | コンテンツ | Section 4 | ⭐v5.0 |
```

### セルフチェック項目（マージ実行前にすべてYESを確認）

- [ ] 全ての中扉が「そのセクションのコンテンツより前」に位置しているか
- [ ] TOCの `sections[i].title` と、対応する中扉の `title` が一致しているか
- [ ] TOCの `sections[i].page` と、対応する中扉の通し番号が一致しているか
- [ ] ファイル番号に歯抜け・重複・逆転がないか
- [ ] エグゼクティブサマリーが通し番号1番に配置されているか
- [ ] **検証論点（issue-risk-list）がデッキ末尾（データアベイラビリティの直前）に配置されているか** ⭐v5.0

**一つでもNOがあれば、マージ実行前にファイルをリネームして修正すること。**

### merge_order.json の併記出力（v5.1 新規）

Step 最終+1 のビジュアルレビュアーが参照するため、照合表を**機械可読な JSON 形式でも出力**する:

```json
{
  "entries": [
    {"slide_number": 1, "file_name": "slide_01_exec_summary.pptx",
     "skill_name": "executive-summary-pptx", "data_file": "data_01_exec.json"},
    {"slide_number": 2, "file_name": "slide_02_toc.pptx",
     "skill_name": "table-of-contents-pptx", "data_file": "data_02_toc.json"},
    ...
  ]
}
```

保存先: `{{WORK_DIR}}/merge_order.json`

`data_file` は、そのスライドの生成に使った JSON ファイル名（`{{WORK_DIR}}/data_NN_*.json`）。
中扉や TOC など data_file を持たないスライドは `null` にする。

---

## Step 最終: 結合（merge-pptxv2）

全スライドを1ファイルに結合する。

### ⚠️ マージコマンド組み立ての原則

- **`ls *.pptx | sort` の結果をそのまま引数に流すのは絶対にNG**（ファイル名の番号が最終順序と一致していれば結果的に同じになるが、それを暗黙の前提にしない）
- **必ず Step 最終-1 で作成した「マージ順序照合表」の通し番号順にファイル名を並べる**
- マージコマンドを組み立てたら、**実行前にもう一度「通し番号 ↔ 引数の位置」が一致しているかを目視確認**する

```bash
pip install lxml --break-system-packages -q

python <merge-pptxv2_DIR>/scripts/merge_pptx_v2.py \
  {{OUTPUT_DIR}}/StrategyReport_[対象会社名].pptx \
  {{WORK_DIR}}/slide_01_exec_summary.pptx \
  {{WORK_DIR}}/slide_02_toc.pptx \
  {{WORK_DIR}}/slide_03_section1_target.pptx \
  ... \
  {{WORK_DIR}}/slide_ZZ_data_avail.pptx
```

出力ファイル名: `StrategyReport_[対象会社名].pptx`

### merge 完了後: brand_warnings を merge_warnings.json に追記（必須）

merge-pptxv2 は `merge_warnings.json` を `"w"` モードで上書きするため、Step 5（共通パターン）中に蓄積した `brand_warnings` は merge 完了後にここで追記する。

```python
append_brand_warnings_to_merge_file(
    "{{OUTPUT_DIR}}/merge_warnings.json", brand_warnings,
)
# brand_warnings が空なら no-op（既存ファイルは触らない）。
# 末尾の最終ユーザー伝達でも warning 件数 + 内訳を必ず提示する。
```

### マージ後の最終検証

`merge_pptx_v2.py` の出力で各スライド番号の Main Message とshape数が表示される。以下を確認:

- 1枚目のMain Messageがエグゼクティブサマリーの内容になっているか
- 中扉（shape数が少ない=8前後、タイトル＋サブタイトル＋トピックリストのみ）が**そのセクションのコンテンツの直前**に出現しているか
  - shape数の推移が「コンテンツ（多）→ 中扉（少=8前後）→ コンテンツ（多）」の谷になっているのが正常
  - 「コンテンツ（多）→ コンテンツ（多）→ 中扉（少）」は**中扉が末尾に流れている誤りのサイン**
- 最後のスライドが検証論点またはデータアベイラビリティになっているか

**不整合があれば、ファイルをリネームして必ずもう一度マージし直す。**

---

## Step 最終+1: ビジュアル品質レビュー（v5.1新規）

<!-- source: skills/_common/prompts/step_final_visual_review_loop.md (manual sync until D2) -->

マージ完了後、`visual-quality-reviewer` を起動してデッキ全体をページ画像化 → 目視レビューする。

### 起動

本スキルの SKILL.md に従い、以下を入力として渡す:

- `merged_pptx`: `{{OUTPUT_DIR}}/StrategyReport_[対象会社名].pptx`
- `merge_order`: `{{WORK_DIR}}/merge_order.json`
- `data_dir`: `{{WORK_DIR}}`

<!-- @if:claude_code -->
結果は `{{FACTORY_ROOT}}/work/visual-quality-reviewer/visual_review_report.json` に出力される。
<!-- @endif -->

### レビュー結果の分岐（オーケストレーターの自動処理）

| `overall_verdict` | 処理 |
|---|---|
| `pass` | 終了。完成デッキをユーザーに提示 |
| `needs_fixes` かつ **`severity=high` が1件以上** | **自動修正ループへ**（下記） |
| `needs_fixes` かつ `severity=high` が0件 | ユーザーに差分レポートを Markdown で提示し、手動修正 or 許容を選ばせる |
| `reject` | LibreOffice レンダリング失敗を疑いユーザーに報告して停止 |

### 自動修正ループ（最大2ラウンド）

`severity=high` の各 issue について:

1. `issues[i].skill_name` と `issues[i].data_file` から、該当スライド生成に使った JSON ファイルを特定
2. `issues[i].regeneration_hint` に従って **`data_NN_*.json` を修正**（例: bullets を短縮、項目数を減らす）
3. 該当スキル（例: `pest-analysis-pptx`）の `fill_*.py` を**同じ `slide_NN_*.pptx` ファイル名で再実行** → 既存スライドを上書き
4. 全修正完了後、**`merge-pptxv2` を再実行**して最新デッキを再生成
5. 再度 `visual-quality-reviewer` を起動

**2ラウンド終了時点で `high` が残存する場合**:
- ユーザーに残存 issue を提示し、手動修正か許容の判断を仰ぐ
- 無限ループには絶対に入らない（カウンタを必ず持つ）

### ユーザーへの最終出力

- `overall_verdict=pass` 時: 「ビジュアル品質レビュー: 問題なし」のみ
- 自動修正でpassに到達した時: 「自動修正 N 件を適用しました（詳細: `visual_review_report.json`）」
- 手動対応が必要な時: レポートと共に次アクションを明示

---

## 依存スキル一覧

### コアスキル（全モードで必須）

| スキル名 | 配置 | 役割 |
|---|---|---|
| `executive-summary-pptx` | user | デッキ冒頭のサマリースライド |
| `table-of-contents-pptx` | user | デッキ目次 |
| `customer-profile-pptx` | user | 対象会社・競合の詳細プロファイル |
| `business-portfolio-pptx` | user | 事業セグメント別売上構成 |
| `business-model-pptx` ⭐ | **organization** | **取引構造図（サプライヤー・自社・顧客）** |
| `market-environment-pptx` | user | マーケット規模推移 |
| `competitor-summary-pptx` | user | 競合比較サマリー |
| `issue-risk-list-pptx` ⭐ | **organization** | **今後検証すべき論点一覧** |
| `company-overview-pptx-v2` | user | 非上場競合等のプロファイル |
| `merge-pptxv2` | user | 全スライドの結合 |

### 拡張スキル（標準・拡張モードで選択的に使用）

| スキル名 | 配置 | モード | 役割 |
|---|---|---|---|
| `pest-analysis-pptx` | user | 標準・拡張 | マクロ環境分析 |
| `five-forces-pptx` | **organization** | 標準・拡張 | 業界5競争要因分析 |
| `value-chain-pptx` | user | 拡張 | 業界バリューチェーン・利益プール分析 |
| `market-share-pptx` | user | 標準・拡張 | 市場シェア分析 |
| `positioning-map-pptx` | user | 拡張 | 2軸ポジショニングマップ |
| `financial-benchmark-pptx` | user | 標準・拡張 | 複数財務指標の競合比較 |
| `swot-pptx` | user | 標準・拡張 | SWOT分析 |
| `value-chain-matrix-pptx` ⭐ | user | 標準・拡張 | **バリューチェーン・ポジショニング・マトリクス（自社工程／他社工程／プレーヤー／競合・協力）** |
| `growth-driver-pptx` | user | 標準・拡張 | 売上/利益要因分解 |
| `section-divider-pptx` | user | 拡張 | 中扉（セクション区切り） |
| `data-availability-pptx` | user | 全モード（オプション） | 調査の網羅度・制約事項 |

⭐ = v5.0で新規組み込み / 役割変更
**organization** = 組織スキル（`/mnt/skills/organization/` 配下）。依存するため事前にインストール済みである必要がある。

### 品質レビュー系スキル（v5.1新規・全モードで必須）

| スキル名 | 配置 | 呼び出し位置 | 役割 |
|---|---|---|---|
| `fact-check-reviewer` ⭐ | user | Step 2.5 | Web取得情報を再検索で裏取りし、疑わしい主張をフラグ |
| `visual-quality-reviewer` ⭐ | user | Step 最終+1 | マージ後デッキをページ画像化し、文字溢れ・重なり・密度過多等を目視判定 |

### v5.0で削除されたスキル依存

| スキル名 | v4.x | v5.0 | 理由 |
|---|---|---|---|
| `recommendation-action-pptx` | 必須 | **削除** | 公開情報のみでは一般論に陥るため。代わりに `issue-risk-list-pptx` で検証論点として着地 |

---

## 中扉と目次の連動ロジック

`table-of-contents-pptx` と `section-divider-pptx` をペアで使う場合:

| セクション # | 目次の `sections[]` | 中扉の `section_number` | 共通の色 |
|---|---|---|---|
| 1 | 対象会社の概要理解 | 1 | 紺 |
| 2 | 外部環境（業界・市場・競合） | 2 | 青 |
| 3 | 対象会社の戦い方 | 3 | 緑 |
| 4 | 今後検証すべき論点 | 4 | オレンジ |

---

## 品質チェックリスト

### デッキ全体
- [ ] エグゼクティブサマリーが冒頭に配置されているか
- [ ] 目次がエグサマ直後に配置されているか
- [ ] 拡張版で中扉が各セクション開始時に配置されているか（4セクション分）
- [ ] 検証論点（issue-risk-list）がレポート結論として末尾近くに配置されているか
- [ ] データアベイラビリティが配置されているか（オプションで冒頭 or 末尾）
- [ ] 全スライドが結合済みで1ファイルになっているか

### 中扉配置の厳格チェック（過去バグ再発防止）
- [ ] **全ての中扉が、そのセクションのコンテンツスライド「より前」に位置しているか**（末尾寄りになっていないか）
- [ ] **ファイル名の番号とデッキ最終順序が一致しているか**（`slide_20_section4_*.pptx` は通し番号20番目にあるべき）
- [ ] マージ後のshape数推移で「コンテンツ（多）→ 中扉（少=8前後）→ コンテンツ（多）」の谷が、各セクション境界で正しく出現しているか
- [ ] 目次の `sections[i].page` が、対応する中扉の通し番号と一致しているか

### 内容の一貫性（v5.0の思想準拠）
- [ ] **メインメッセージが「〜すべき」で締められていないか**（事実記述ベースになっているか）
- [ ] エグゼクティブサマリーのKey Findingsが事実記述ベースで、他スライドの内容と整合しているか
- [ ] 検証論点（issue-risk-list）がデータアベイラビリティの ✗/△ 項目と連動しているか
- [ ] 目次の各セクションと、中扉の section_number / 色が一致しているか
- [ ] **business-model（S1）と value-chain-matrix（S3）の間で対象会社の事業構造が一貫しているか**

### v5.0モード別追加チェック

**基本版（7-8枚）:**
- [ ] ビジネスモデル（business-model）が表示されているか
- [ ] 検証論点（issue-risk-list）が表示されているか

**標準版（13-15枚）:**
- [ ] Section 1の3スライド（プロファイル/ポートフォリオ/ビジネスモデル）が揃っているか
- [ ] バリューチェーン・ポジショニング・マトリクス（value-chain-matrix）が表示されているか
- [ ] 成長ドライバー分析が表示されているか
- [ ] 検証論点（issue-risk-list）が表示されているか

**拡張版（20-25枚）:**
- [ ] 中扉が4セクション分（S1対象会社概要・S2外部環境・S3戦い方・S4検証論点）配置されているか
- [ ] 業界バリューチェーン・利益プール（value-chain）がS2に配置されているか
- [ ] Section 2の並び順が「マクロ→業界→市場→競合」のズームインになっているか
- [ ] SWOT・value-chain-matrix・成長ドライバーがSection 3（対象会社の戦い方）に集約されているか

---

## 注意事項

- **v5.0の最大の特徴は「知的誠実性」**
  - 公開情報で書けないことは書かない
  - 「〜すべき」と断定せず、「〜という構造である」「〜を確認する必要がある」で記述
  - 検証論点で率直に未解明点を示す
- **事業環境の理解 → 対象会社の戦い方の理解 の論理展開**
  - Section 1（対象会社概要）→ Section 2（外部環境）→ Section 3（対象会社の戦い方）→ Section 4（検証論点）
  - S3 は S2（業界の構造）を踏まえた上で「その環境で対象会社がどう戦っているか」を記述する
- **Section 2 の並び順はマクロ→業界→市場→競合のズームイン**
  - 中扉で細分化せず、並び順で論理グループを表現
- **拡張モードの制約**:
  - `business-portfolio-pptx`: 上場企業でセグメント開示がある場合
  - `business-model-pptx`: 公式HP・有報の事業の内容セクションから取引構造が特定できる場合（組織スキル）
  - `pest-analysis-pptx`: マクロレベルで書く（特定企業固有の話は避ける）
  - `five-forces-pptx`: 業界の構造分析、対象会社視点ではなく業界視点（組織スキル）
  - `value-chain-pptx`: 業界全体のバリューチェーンを5-7段階に分解できる場合（業界の利益構造）
  - `value-chain-matrix-pptx`: 対象会社が各工程で何をやっているか、誰と競合・協業しているかが特定できる場合
  - `growth-driver-pptx`: 決算説明会資料で要因分解が開示されている場合
  - `issue-risk-list-pptx`: 常に必須（組織スキル）
- **エグゼクティブサマリーは最後に書く**: 全スライド生成後、各スライドの事実を統合してFindings作成
- **検証論点も最後に**: 全分析が出揃ってから、公開情報で見えなかった点を論点として列挙

---

## 使用例

### 例1: BDDフェーズ1レポート（標準版）

```
User: トヨタ自動車のBDDレポートを作って
Assistant:
1. [Step 0] 調査タイプ=B (BDD), デッキ深度=B (標準版), データアベイラビリティ=A (末尾)
2. [Step 1-4] Web検索・データ整理・確認・競合形式確認
3. [Step 5] Key Findings 5個を事実記述ベースで整理
4. [Step 6] 検証論点 5-7個を整理（マネジメントインタビューで聞くべきこと）
5. [Step 7-N] 標準版15枚 + データアベイラビリティ1枚 = 16枚を生成
6. [Step 最終] merge-pptxv2 で結合
7. ファイル出力: StrategyReport_トヨタ自動車.pptx
```

### 例2: 新規参入調査（拡張版）

```
User: EV市場への新規参入を検討している。拡張版で調査レポート作って。
Assistant:
1. [Step 0] 調査タイプ=C (新規参入), デッキ深度=C (拡張版), データアベイラビリティ=B (冒頭配置)
2. [Step 1-6] 情報収集・整理・Findings/検証論点作成
3. [Step 7-最終] 拡張版22-25枚 (エグサマ + 目次 + 4中扉 + 各セクション + データアベイラビリティ) を生成・結合
```

### 例3: M&Aターゲット評価（カスタム）

```
User: B社の買収検討。ビジネスモデルとバリューチェーン・マトリクス、検証論点中心で。
Assistant:
1. [Step 0] 調査タイプ=E (M&A), デッキ深度=D (カスタム)
   Custom選択内容: 事業ポートフォリオ・ビジネスモデル・財務ベンチマーク・
                バリューチェーン・マトリクス・成長ドライバー・検証論点・データアベイラビリティ
2. [Step 1-6] 情報収集・整理
3. [Step 7-最終] カスタムデッキ10-12枚を生成・結合
```

---

## 参考

| ファイル名 | 内容 |
|---|---|
| `references/CHANGELOG.md` | バージョン履歴 |

---

## バージョン履歴（抜粋、詳細は CHANGELOG 参照）

- **v5.1** (2026年4月): **品質レビュー2段組み込み**
  - `fact-check-reviewer` を Step 2.5 に追加（情報の真偽を裏取り、疑わしい主張をフラグ）
  - `visual-quality-reviewer` を Step 最終+1 に追加（マージ後デッキを画像化し目視レビュー）
  - `high` 重大度のビジュアル不備は**自動再生成＋再マージ**（最大2ラウンド）
  - Step 最終-1 で `merge_order.json` を併記出力（ビジュアルレビュアーが参照）
- **v5.0** (2026年4月): **思想転換＋4セクション構成再編**
  - **思想転換**: 「事実→示唆→アクション」から「事業環境の理解→対象会社の戦い方の理解」へ
  - `recommendation-action-pptx` を依存から削除（公開情報では一般論に陥るため）
  - `issue-risk-list-pptx`（組織スキル）を「今後検証すべき論点」として新規組み込み
  - `business-model-pptx`（組織スキル）を Section 1 に新規組み込み
  - `value-chain-matrix-pptx` を Section 3（対象会社の戦い方）に新規組み込み
  - `value-chain-pptx` を Section 4（戦略示唆）→ Section 2（外部環境：業界の利益プール）に移動
  - SWOT を Section 1 → Section 3（対象会社の戦い方）に移動
  - セクション構成: 4セクション（対象会社概要 / 外部環境 / 対象会社の戦い方 / 検証論点）
  - メインメッセージのトーンを「〜すべき」から「〜という構造である」「〜を確認する必要がある」へ
- **v4.1** (2026年4月): 中扉配置バグ防止の強化
- **v4.0** (2026年4月): 戦略示唆セクション追加（v5.0で思想転換のため廃止）
- **v3.0** (2026年4月): リブランド `competitor-analyst-agent` → `strategy-report-agent`
- **v2.0**: 6つのスキル追加
- **v1.0**: 基本4枚構成
