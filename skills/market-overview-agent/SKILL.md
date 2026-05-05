---
name: market-overview-agent
description: >
  「XX市場を調べて」というユーザー依頼を受けて、MBB 戦略コンサルタント品質の市場分析
  PowerPoint デッキ（10〜12枚）と、出典・自信度を記録したファクトチェック Markdown
  レポートの2点セットを納品するオーケストレーター型エージェントスキル。
  本スキル自体はスクリプトを持たず、Web検索＋複数の既存 PPTX スキル
  （market-environment / market-share / positioning-map / market-kbf / pest-analysis /
  competitor-summary / executive-summary / table-of-contents / section-divider /
  data-availability）と品質レビューア（fact-check-reviewer / visual-quality-reviewer）、
  結合スキル（merge-pptxv2）を順次呼び出してデッキ全体を組み立てる。
  公開情報で確定できない論点は「検証すべき論点」として明示し、知的誠実性を保つ。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「XX市場を調べて」「XX業界を調査して」「市場分析を作って」「市場概要をパワポに」という依頼
  - 「市場規模・シェア・KBF・PEST・競合戦略」を統合した1本のレポートが求められた場合
  - 「Market Overview」「市場サマリー」「市場リサーチ」「業界調査」「市場レビュー」という言葉
  - 単一の PPTX スキル（market-environment-pptx 等）の起動で済まない、横断的な市場調査依頼

  以下の場合は別スキルを使う:
  - 「対象会社 1 社の深掘り」 → company-deepdive-agent
  - 「単一の事業セグメントだけ深掘り」 → business-deepdive-agent
  - 単一の PPTX スライドだけ作りたい → 該当する個別 PPTX スキル（market-environment-pptx / market-share-pptx 等）
---

# Market Overview Agent

「XX市場を調べて」依頼から、市場分析 PowerPoint（標準10〜12枚）＋ファクトチェックMDレポートの
2点セットを生成するオーケストレーター。

**設計原則**:
- 各論点に1枚のスライド（情報密度を1論点1スライドで担保）
- 公開情報で確定できないものは断定せず「検証すべき論点」として残す
- 数値・固有名詞は必ず Web 検索で裏取り（fact-check-reviewer）
- 生成済みデッキは PNG 化して目視レビュー（visual-quality-reviewer）し、不備は自動修正ループで処理
- 最終納品物は **PPTX + FactCheck_Report.md** の2つ

---

## 🚨 絶対ルール（v0.1 改善版）

以下は本オーケストレーターが呼び出す**全 PPTX スキル横断**の絶対ルール。逸脱した時点で
visual review が `needs_fixes` になるため、JSON 生成段階で必ず守ること。

### ルール1: `main_message` の取り扱い

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

### ルール2: 詳細テキストは可能な限り簡潔に

- `executive-summary-pptx` の `findings[].detail`: **120文字以内**
- `competitor-summary-pptx` の `target_company.*` / `competitors[].*` の各値: **30文字以内**
- `market-kbf-pptx` の `kbf_list[].description`: **80〜120文字**、`player_examples[].example`: **40〜80文字**
- `pest-analysis-pptx` の `pest.<axis>.items[].text`: **40文字以内**

### ルール3: category などの分類名は **重複禁止**（findings 5件で）

- `executive-summary-pptx` の `findings[].category` は5件のうち重複させない
- 「市場」が2件出るなら「市場規模・成長」と「競争構造」のように粒度を変える
- 推奨カテゴリ: `市場規模`, `競争構造`, `プレイヤー`, `成功要因`, `マクロ環境`, `結論` 等

### ルール4: 競合社数統一の徹底（v0.2 で可変化）

- `market-share-pptx` / `positioning-map-pptx` / `competitor-summary-pptx` / `market-kbf-pptx`（player_examples）すべてで **`scope.json` の `max_competitors`（既定 5、上限 5）と同数の同じ社**を採用する
- 同一実行内での社名表記も統一（例: 「ヒューマンテクノロジーズ」と「KING OF TIME」を混在させない）
- `scope.json.kbf_count` も同様にスキル横断で一致させる（market-kbf-pptx の `kbf_list` の長さ）

---

## 答える論点（市場ベース・5本柱＋PEST）

| # | 論点 | 担当スキル |
|---|------|------------|
| 1 | 市場の過去X年成長率＆今後の成長率 | `market-environment-pptx` |
| 2 | 市場のキービジネスファクター（KBF） | `market-kbf-pptx` |
| 3 | 各社の市場におけるシェア | `market-share-pptx` |
| 4 | プレイヤーの位置付け（ポジショニング） | `positioning-map-pptx` |
| 5 | 各社の戦略比較（一覧） | `competitor-summary-pptx` |
| + | PEST環境 | `pest-analysis-pptx` |

論点の拡充がスキルの成長軸。今後の拡張で `five-forces-pptx` / `value-chain-pptx` 等を
組み込んだ拡張モード（v0.2 以降）を予定。

---

## 標準デッキ構成（10〜12枚・3セクション）

| # | 出力ファイル | スキル | セクション |
|---|---|---|---|
| 01 | `slide_01_exec_summary.pptx` | `executive-summary-pptx` | 冒頭 |
| 02 | `slide_02_toc.pptx` | `table-of-contents-pptx` | 冒頭 |
| 03 | `slide_03_section1_market_size.pptx` | `section-divider-pptx`（Section 1: 市場規模・成長）| Section 1 冒頭 |
| 04 | `slide_04_market_environment.pptx` | `market-environment-pptx` | Section 1 |
| 05 | `slide_05_section2_competition.pptx` | `section-divider-pptx`（Section 2: 競争構造）| Section 2 冒頭 |
| 06 | `slide_06_market_share.pptx` | `market-share-pptx` | Section 2 |
| 07 | `slide_07_positioning.pptx` | `positioning-map-pptx` | Section 2 |
| 08 | `slide_08_competitor_summary.pptx` | `competitor-summary-pptx` | Section 2 |
| 09 | `slide_09_section3_success_factors.pptx` | `section-divider-pptx`（Section 3: 成功要因と外部環境）| Section 3 冒頭 |
| 10 | `slide_10_market_kbf.pptx` | `market-kbf-pptx` ⭐新規 | Section 3 |
| 11 | `slide_11_pest.pptx` | `pest-analysis-pptx` | Section 3 |
| 12 | `slide_12_data_availability.pptx` | `data-availability-pptx` | 末尾 |

⭐ = v0.1 で新規開発したスキル

---

## 処理フロー

```
Step 0: 市場スコープ確認（AskUserQuestion）
  ↓
Step 1: Web検索による論点別情報収集 → data_NN_*.json
  ↓
Step 2: データアベイラビリティ整理（Markdown）
  ↓
Step 2.5: fact-check-reviewer（high_risk/all/skip）
  ↓
Step 3: Markdownでユーザーに確認（要確認項目を統合）
  ↓
Step 4: Key Findings 整理（Executive Summary 用）
  ↓
Step 5: スライド生成（slide_NN_*.pptx を順次作成）
  ↓
Step 6: merge_order.json + マージ順序照合表
  ↓
Step 7: merge-pptxv2 で結合 → MarketOverview_<market_name>_<date>.pptx
  ↓
Step 8: visual-quality-reviewer ＋ 自動修正ループ（最大2ラウンド）
  ↓
Step 9: FactCheck_Report.md 生成（最終納品物）
  ↓
Step 10: ユーザーへ提示（PPTX + MDの2ファイル）
```

---

## 進捗トラッキング規約（全 Step 横断、必須）

<!-- source: skills/_common/prompts/step_state_tracking.md (manual sync until D2) -->

各 Step の開始/終了で `TaskCreate` / `TaskUpdate` を呼び、`task_state.json` を更新する。詳細規約は `skills/_common/prompts/step_state_tracking.md` を正本とする。

- **subject フォーマット**: `market-overview: Step <N> - <topic>`(例: `market-overview: Step 1 - Web検索 (5論点+PEST)`)。サブ番号 (Step 0.5 / 2.5 / 6-a / 8-b 等) も `Step 0.5` / `Step 6-a` のように subject に含める
- **task_state.json 配置**: `{{WORK_DIR}}/<run_id>/task_state.json`(scope.json と同じディレクトリ)
- **step_id**: `step_0` / `step_0_5` / `step_2_5` / `step_6_a` のように `.` と `-` を `_` に置換
- **開始時**: `TaskCreate` で task を起こす → `TaskUpdate(in_progress)` → `task_state.json.steps[]` に append
- **終了時**: `TaskUpdate(completed)` → `task_state.json` の該当 entry を `completed` + `completed_at` に更新
- **失敗・再試行時**: `TaskUpdate(completed)` を呼ばない。`task_state.json` の `retry_count` のみインクリメント。Step 8 の visual-review 自動修正ループでは別途カウンタ（最大 2 ラウンド）を保持する

`tools/hooks/check_task_progression.py` が `fill_*.py` / `merge_pptx_v2.py` 起動前にこのファイルを参照し、Step ordering inversion を検出して exit 2 でブロックする。

---

## ハーネスレバー利用規約（参照）

<!-- source: skills/_common/references/harness_levers.md (manual sync until D2) -->

本オーケストレーターは Claude Code ハーネス機構を以下のとおり活用する。詳細規約は `skills/_common/references/harness_levers.md` を参照。

| レバー | 適用箇所 |
|---|---|
| ① hooks (`tools/hooks/*.py`) | `check_merge_order_exists`(Step 7 直前) / `validate_pptx_after_fill`(Step 5 / 7) / `check_task_progression`(全 Step) / `load_session_context` が `.claude/settings.json` 経由で発火 |
| ② subagent (`.claude/agents/research-subagent.md`) | Step 1 で 5 論点 + PEST を並列起動。25-40 件の Web 検索結果（生 HTML）を親 context に積まず要約 JSON のみ受け取る（market-overview の 12箇条 #3 改善の核心）|
| ③ TaskCreate / AskUserQuestion | 各 Step で TaskCreate（上記）。Step 0（スコープ確認）/ Step 0.5（事業モデル境界）/ Step 2.5（factcheck スコープ）/ Step 3（Markdown 承認）/ Step 8 残存 issue 判断で AskUserQuestion 必須 |

---

## Step 0: 市場スコープ確認

**進捗**: 開始時 `TaskCreate(subject="market-overview: Step 0 - スコープ確認")` → Step 0.0-pre / 0.0 / 0.5 完了後にまとめて `TaskUpdate(completed)` + `task_state.json` 更新。**`AskUserQuestion` 必須**(自由対話での確定は禁止、Step 0.0-pre / 0.0 / 0.5 全てで必須)。

<!-- source: skills/_common/prompts/step0_brand_clarification.md (manual sync until D2) -->

### Step 0.0-pre: ブランド確認（必須）

本デッキの**出力ブランド**（クライアント別 PPTX フォーマット）を `scope.json.brand` に保存する。共通原則・AskUserQuestion テンプレ・自由記述ハンドリング・unsupported skill fallback の詳細仕様は `skills/_common/prompts/step0_brand_clarification.md` を正本とする。

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
# 確定値は scope.json.brand に文字列保存（既定 "stellar_aiz"）。
# 「Other」で _discover_brands に含まれない id を入力された場合は AskUserQuestion を再実行。
```

<!-- source: skills/_common/prompts/step0_scope_clarification.md (manual sync until D2) -->

共通原則・Step 0.5（事前スコーピング Web 検索）・`included_business_models` / `excluded_segments` フィールドの定義は `skills/_common/prompts/step0_scope_clarification.md` を正本とする。本 SKILL.md には market-overview-agent 固有の質問項目と Step 0.5 の運用上の注意のみ記載する。

### Step 0.0: 固有質問の確定

`AskUserQuestion` で以下を確定する。すべて単一選択（必要なら「Other」で自由記述）。

| 質問 | 選択肢例 | 既定値 |
|---|---|---|
| 地理スコープ | 国内 / グローバル / アジア / 北米・欧州 | — |
| セグメント粒度 | 業界全体ザックリ / セグメント細分化（例：BtoBのみ・BtoCのみ）| — |
| 分析年数 | 過去3年＋今後3年 / 過去5年＋今後5年 / 過去10年＋今後10年 | — |
| 主要競合の上限（max_competitors） | 3 / 5（範囲: `references/deck_skeleton_standard.json` の `limits.max_competitors` を参照、上限 5） | 5 |
| KBFの数（kbf_count） | 2 / 3 / 4 / 5（範囲: 同上 `limits.kbf_count`） | 3 |
| 強調したい会社（highlight_company） | 強調しない / 自由記述で会社名を入力 | **強調しない** |

**v0.2 追加**: `max_competitors` と `kbf_count` は `references/deck_skeleton_standard.json` の `limits` セクションで `min/max/default` が定義されている。ユーザーが明示しない場合は default を採用。範囲外の値は AskUserQuestion で再確認すること。

**強調会社（highlight_company）の運用**:
- 既定値は `null`（強調しない）。クライアントが特定の会社にフォーカスしたい場合のみ会社名を文字列で受け取る
- 指定された会社名は P6 (market-share) / P7 (positioning-map) / P8 (competitor-summary) の 3 スライドで横断的に強調される（イエロー背景・赤バブル等）
- 指定なしの場合は P6/P7/P8 すべてで「強調列なし／全プレイヤー均等表示」になる
- 後続 Step 1 で確定する `players[].name` と完全一致させる必要があるため、Step 1 で実プレイヤー名が揃った段階で `scope.json.highlight_company` の表記を再確認・正規化すること（不一致なら強調が当たらない）
- 本フィールドは Step 2.5 の `fact-check-reviewer` に渡す `target_company`（検索精度向上用）とは**別物**。両者を混同しないこと

### Step 0.5: 事前スコーピング Web 検索（必須）

`market_name` が確定したら、上記の固有質問に進む前に **市場構造ザックリ把握用の Web 検索を 1〜2 件** 走らせる。
目的は「同一業界内に収益構造の異なる事業モデルが併存していないか」を検知し、`included_business_models` / `excluded_segments` をユーザーに確認すること。

検索クエリ例:
- `<market_name> 業界構造 / バリューチェーン / プレイヤー類型`
- `<market_name> 市場規模 定義 / 統計対象範囲`

異種事業モデルが検知された場合は `AskUserQuestion` で境界確認:

```
「<market_name>」には収益構造の異なる事業モデルが併存しています。どの層を調査対象に含めますか？
A. <事業モデル1>のみ（例: タクシー事業者 = 営業収入ベース）
B. <事業モデル2>のみ（例: 配車アプリ事業者 = 配車手数料ベース）
C. 両方含める（シェア表は別レイヤーで分けて表示）
D. その他（自由記述）
```

異種併存の典型例（共通プロンプト参照）: タクシー / 半導体 / 教育 / 飲食 / 物流 / 金融。
ユーザーが冒頭で `included_business_models` を明示している場合（例:「タクシー事業者の市場を調べて」）は Step 0.5 をスキップしてよいが、`scope.json` の該当フィールドは判明している値で埋めること。

### scope.json の保存

確定したスコープは `{{WORK_DIR}}/<run_id>/scope.json` に保存：

```json
{
  "market_name": "国内タクシー市場",
  "geography": "国内",
  "segment": "業界全体",
  "analysis_years": { "past": 5, "future": 5 },
  "max_competitors": 5,
  "kbf_count": 3,
  "highlight_company": null,
  "included_business_models": ["タクシー事業者"],
  "excluded_segments": ["配車アプリ事業者"],
  "brand": "stellar_aiz",
  "brand_label": "Stellar AIZ（既定）",
  "run_id": "2026-04-27_taxi_industry_operators",
  "started_at": "2026-04-27T10:00:00+09:00"
}
```

`highlight_company` は文字列（強調する会社名）または `null`（強調なし、デフォルト）。

**重要**:
- `scope.json` の `max_competitors` / `kbf_count` は後続 Step（Web 検索・データ生成・fill_*.py 入力）で参照される。market-share / positioning-map / competitor-summary / market-kbf-pptx の 4 スキル間で一貫した値を使うこと。
- `included_business_models` / `excluded_segments` の境界尊重責務はオーケストレーター（本スキル）にある。Step 1 の Web 検索クエリ・data_06_market_share.json の母集団・data_08_competitor_summary.json の比較対象は `included_business_models` の範囲内に限定する。fill_*.py は scope.json を読まない（単体起動互換性のため。`skills/_common/references/orchestrator_contract.md` 参照）。
- `excluded_segments` が空配列でない場合は、Step 2 の data_12_data_availability.json と最終 FactCheck_Report.md の冒頭で「本レポートでは <excluded_segments> を対象外として除外している」旨を明記する。

`{{WORK_DIR}}/<run_id>/` を以下「作業ディレクトリ」と呼ぶ。

---

## Step 1: research-subagent 経由で論点別情報収集

**進捗**: 開始時 `TaskCreate(subject="market-overview: Step 1 - Web検索 (5論点+PEST)")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。

5 論点 + PEST のそれぞれについて、`research-subagent`(`.claude/agents/research-subagent.md`) を **Agent ツール経由で並列起動** して情報収集する。各 subagent は 5-8 件の Web 検索 + 必要に応じた fetch を実施し、`output_schema` に沿った要約済み JSON のみ親に返却する。**生 HTML / 検索結果テキストは subagent 自身の context 内で完結し、親 context には流入しない**(market-overview は 25-40 件の Web 検索を伴うため、本対策が 12箇条 #3「コンテキスト制御」改善の核心)。

### subagent 呼び出しパターン（論点ごと、並列起動可）

```python
import json
from skills._common.lib.parse_subagent_return import parse_subagent_return
result = Agent(
    subagent_type="research-subagent",
    description=f"{market_name} の<論点名>調査",
    prompt=json.dumps({
        "topic_id": "data_<NN>_<topic>",  # 論点別の data ファイル名と対応
        "topic_description": "<論点の自然文要約>",
        "output_schema": {<該当 PPTX スキルの JSON schema>},
        "parent_context": {
            "industry": market_name,
            "geography": scope.geography,
            "scope_constraints": {
                "included_business_models": scope.included_business_models,
                "excluded_segments": scope.excluded_segments
            },
            "target_company": scope.highlight_company  # 強調会社が指定されていれば
        },
        "search_budget": {"min_searches": 5, "max_searches": 8}
    })
)
# subagent return は parse_subagent_return() 経由で dict 化する（必須）。
# 直接 json.loads(result) しないこと: subagent が稀に前置き文・code fence・末尾
# Sources を混入させるため（ISSUE-009）。helper はそれらを吸収する。
parsed = parse_subagent_return(result)
# parsed["data"] を {{WORK_DIR}}/<run_id>/data_<NN>_<topic>.json に Write で書き出す
# parsed["open_questions"] を data_12_data_availability.json と FactCheck_Report.md に転記
```

### 検索の優先順位

| 論点 | 優先ソース |
|------|-----------|
| 市場規模・成長率 | 矢野経済 / 富士経済 / IDC / 業界団体 / 政府統計 |
| KBF | 業界紙 / コンサルレポート / 専門家インタビュー記事 / 各社IR |
| シェア | 各社IR（年次報告書）/ 業界統計 / 経済紙 |
| ポジショニング | 各社HP / IR / プレスリリース |
| 戦略比較 | 中期経営計画 / 統合報告書 / 決算説明会資料 |
| PEST | 政府統計 / シンクタンク / メディア / 業界団体 |

検索キーワードのテンプレは `prompts/step1_research_template.md` 参照。

### data_NN_*.json の命名

各 JSON は以下の固定命名で作業ディレクトリに保存:

| ファイル名 | 内容 | 対応スキル |
|---|---|---|
| `data_01_exec_summary.json` | 5 Key Findings（Step 4で生成）| executive-summary-pptx |
| `data_02_toc.json` | 目次（3セクション分）| table-of-contents-pptx |
| `data_03_section1.json` | Section 1中扉 | section-divider-pptx |
| `data_04_market_environment.json` | 市場規模推移 | market-environment-pptx |
| `data_05_section2.json` | Section 2中扉 | section-divider-pptx |
| `data_06_market_share.json` | 市場シェア | market-share-pptx |
| `data_07_positioning.json` | ポジショニング | positioning-map-pptx |
| `data_08_competitor_summary.json` | 戦略比較 | competitor-summary-pptx |
| `data_09_section3.json` | Section 3中扉 | section-divider-pptx |
| `data_10_market_kbf.json` | KBF×`kbf_count`（既定 3） | market-kbf-pptx |
| `data_11_pest.json` | PEST | pest-analysis-pptx |
| `data_12_data_availability.json` | 取得状況 | data-availability-pptx |

### 競合の絞り込み（max_competitors 連動）

ユーザーが `scope.json.max_competitors` を超える数を挙げた場合、または Web 検索で同数を超えてヒットした場合は、以下の優先順位で
**上位 `max_competitors` 社に絞り込む**:

1. シェア順位（市場の主要プレイヤー）
2. 売上規模
3. メディア・業界レポートでの言及頻度

絞り込み根拠は `data_06_market_share.json` の `notes` フィールドに明記する（「対象市場
の主要 N 社に絞り込み。N+1 番手以下のプレイヤーは追加調査時に拡張」等）。

### max_competitors 上限の統一適用

`market-share-pptx` / `positioning-map-pptx` / `competitor-summary-pptx` /
`market-kbf-pptx`（player_examples）すべてで **`scope.json.max_competitors` と同数の同じ社**を採用する（読み手が一貫したストーリーで追えるようにするため）。

### highlight_company の伝搬（P6/P7/P8 横断）

`scope.json.highlight_company` の値に応じて、Step 1 のデータ生成時に各 JSON へ次のとおり伝搬する。**3 スライドで同じ社を強調することで読み手のストーリーが揃う**。

| scope.highlight_company | data_06 (market-share) | data_07 (positioning) | data_08 (competitor-summary) |
|---|---|---|---|
| `null`（強調なし・デフォルト） | `target_company` フィールド省略（または `null`） | 同左 | `target_company` オブジェクト省略 → 全社フラット表示 |
| `"○○社"`（強調指定あり） | `target_company: "○○社"` | `target_company: "○○社"` | `target_company: {"name": "○○社", <comparison_items の各 key>: ...}` オブジェクト形式 |

**運用ルール**:
1. `scope.highlight_company` の文字列は `data_06.players[].name`・`data_07.players[].name`・`data_08.competitors[].name` のいずれかと**完全一致**させる（不一致なら強調されない）
2. P8 では `target_company` がオブジェクトなので、強調する会社の `comparison_items` 各キーの値を必ず埋める。一方で `competitors[]` からはその会社を**除外**する（重複表示防止）
3. P6/P7/P8 は単独実行可能なスキル。`scope.json` を直接読まないため、本オーケストレーターが値を JSON に埋め込む責務を負う
4. ユーザーが「強調を後から変更したい」と言った場合は、Step 1 の data_06/07/08 JSON を 3 つとも書き換えるだけで済む（再 Web 検索は不要）

---

## Step 2: データアベイラビリティ整理

**進捗**: 開始時 `TaskCreate(subject="market-overview: Step 2 - データアベイラビリティ")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。

`data-availability-pptx` のJSON 形式（カテゴリ×項目×ステータス×データソース）に整理。
ステータスは `✓取得済 / △一部取得 / ✗未取得` の3値。

**重要**: ✗/△ の項目は、そのまま FactCheck_Report.md の `data_gaps` セクションへも転記する。

`{{WORK_DIR}}/<run_id>/data_12_data_availability.json` に保存。

---

## Step 2.5: ファクトチェック

**進捗**: 開始時 `TaskCreate(subject="market-overview: Step 2.5 - ファクトチェック")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。**スコープ選択の `AskUserQuestion` 必須**(high_risk / all / skip)。

<!-- source: skills/_common/prompts/step2_5_factcheck_invocation.md (manual sync until D2) -->

`fact-check-reviewer` スキルを呼び出して、Web 取得情報の真偽を裏取り。

### Step 2.5-a: スコープをユーザーに選ばせる

`AskUserQuestion` で以下から選択：

| 選択肢 | 内容 |
|---|---|
| **high_risk**（推奨）| 数値・シェア・市場規模・日付・固有名詞のみを検証 |
| **all** | 上記＋テキスト主張も全件検証（時間がかかる）|
| **skip** | 省略して Step 3 へ |

### Step 2.5-b: fact-check-reviewer 起動

入力:
- `data_dir`: `{{WORK_DIR}}/<run_id>/`
- `scope`: ユーザー選択値
- `target_company`: 検索精度向上用の代表プレイヤー名。主要競合（`scope.json.max_competitors` 社）の最大シェアプレイヤーを推奨。**スライド上の「強調会社」（`scope.json.highlight_company`）とは無関係**で、本パラメータは fact-check-reviewer の Web 検索クエリ精度向上のみに使う。`scope.json.highlight_company` が指定されていればそれを優先採用しても良いが、未指定でも本パラメータは最大シェアプレイヤーで埋めて構わない

<!-- @if:claude_code -->
出力: `{{FACTORY_ROOT}}/work/fact-check-reviewer/fact_check_report.json`
<!-- @endif -->

### Step 2.5-c: フラグ項目の取り扱い

`fact_check_report.json` の `flags[]` を以下に分配:

- `severity=high` または `medium` → **Step 3 の Markdown に「要確認項目」として差し込む**
- 全件 → Step 9 で `FactCheck_Report.md` に転記（最終納品物）

`overall_verdict=pass` の場合はフラグ提示を省略。

---

## Step 3: Markdownでユーザーに確認

**進捗**: 開始時 `TaskCreate(subject="market-overview: Step 3 - Markdown 承認")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。**承認の `AskUserQuestion` 必須**。

調査結果＋推奨スライド構成＋ファクトチェック要確認項目を1つの Markdown でユーザーに提示し、
承認を得る。

### Markdown テンプレート

```markdown
## 市場調査結果サマリー: {market_name}

### 1. 調査スコープ
- 地理: {geography}
- セグメント: {segment}
- 分析年数: 過去{N}年 + 今後{M}年
- 主要競合: {N}社

### 2. データアベイラビリティ
| カテゴリ | 項目 | ステータス | データソース |
| ... |

### 3. ファクトチェック要確認項目（severity=high/medium）
| # | データファイル | 主張 | 検証結果 | severity |

### 4. 推奨デッキ構成（標準10〜12枚）
（標準デッキ構成テーブル）

### 5. 確認事項
- 上記スコープと推奨構成でデッキ生成を進めてよいか
- ファクトチェック要確認項目について、どう対応するか
  - JSON修正（数値の置き換え・削除）
  - 注釈付与（出典情報を脚注に明記）
  - スキップ（そのまま進める）
```

ユーザー承認後、必要な JSON 修正を行ってから Step 4 へ進む。

---

## Step 4: Key Findings 整理（Executive Summary 用）

**進捗**: 開始時 `TaskCreate(subject="market-overview: Step 4 - Key Findings 整理")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。

5 Findings パターンで `data_01_exec_summary.json` を生成:

1. **市場規模と成長性**: 「市場規模は{X}億円、CAGRは過去{N}年で{Y}%、今後{M}年は{Z}%が見込まれる」
2. **競争構造**: 「上位{N}社で{X}%のシェアを占め、{N}位に対象会社が位置する。XX社が継続的に首位」
3. **プレイヤー特性**: 「{特性Aの企業群}と{特性Bの企業群}に二分されており、ポジションの違いが利益率の差につながる」
4. **成功要因（KBF）**: 「市場で勝つには「{KBF1}」「{KBF2}」「{KBF3}」が鍵。先行プレイヤーが先行優位を持つ」
5. **マクロ環境（PEST）**: 「{P/E/S/T のうち最も影響度の高い1〜2要因}が市場形成の主たるドライバー」

**トーン**:
- 事実記述ベース（「〜である」「〜と見られる」）
- 「〜すべき」は禁止（公開情報では断定不可）
- 不明な点は「〜は公開情報からは確定できず、追加調査が必要」と明示

---

## Step 5: スライド生成

**進捗**: 開始時 `TaskCreate(subject="market-overview: Step 5 - スライド生成 (10-12枚)")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。本 Step 中は `validate_pptx_after_fill.py` hook が各 fill_*.py 実行後に PPTX 整合性を自動検証する。

`references/deck_skeleton_standard.json` の順序通りに slide_NN_*.pptx を生成。

### ⚠️ 最重要: ファイル名の番号 = 最終通し番号

strategy-report-agent v5.1 の規約と同じ。番号と最終並び順を一致させる。

### Step 5 開始前: brand fallback バッファ初期化（必須）

scope.json から brand を読み出し、未対応 fill 検出用の warning バッファを初期化する。各 fill 起動前に `resolve_fill_brand_with_warning()` を呼び、未対応スキルでは `stellar_aiz` に fallback + warning を buffer に蓄積する（`skills/_common/lib/orchestrator_helpers.py` 参照）。

```python
import json, os, sys, subprocess
sys.path.insert(0, os.path.join("{{SKILL_DIR}}", "..", "_common", "lib"))
from orchestrator_helpers import (
    resolve_fill_brand_with_warning,
    append_brand_warnings_to_merge_file,
)

with open("{{WORK_DIR}}/<run_id>/scope.json", encoding="utf-8") as f:
    scope = json.load(f)
scope_brand = scope.get("brand", "stellar_aiz")
brand_warnings: list = []  # Step 7 後に merge_warnings.json へ append する
```

### 共通実行パターン

各 fill 起動前に `resolve_fill_brand_with_warning(skill_dir, scope_brand, brand_warnings)` で fill に渡す brand を確定する。supported なら `scope_brand` がそのまま、未対応なら `stellar_aiz` が返り `brand_warnings` に `brand_fallback` エントリが追記される。

```python
skill_dir = os.path.join("{{SKILL_DIR}}", "<dependency_skill>")
fill_brand = resolve_fill_brand_with_warning(skill_dir, scope_brand, brand_warnings)
subprocess.run([
    "python", os.path.join(skill_dir, "scripts", "fill_<name>.py"),
    "--brand", fill_brand,
    "--data", "{{WORK_DIR}}/<run_id>/data_NN_<name>.json",
    "--template", os.path.join(skill_dir, "assets", "<template>.pptx"),
    "--output", "{{WORK_DIR}}/<run_id>/slide_NN_<name>.pptx",
], check=True)
```

bash で直接書く場合（既存スキル踏襲、warning fallback は使わない場合）:

```bash
pip install python-pptx -q --break-system-packages

python {{SKILL_DIR}}/<dependency_skill>/scripts/fill_<name>.py \
  --brand "$(jq -r '.brand // "stellar_aiz"' {{WORK_DIR}}/<run_id>/scope.json)" \
  --data {{WORK_DIR}}/<run_id>/data_NN_<name>.json \
  --template {{SKILL_DIR}}/<dependency_skill>/assets/<template>.pptx \
  --output {{WORK_DIR}}/<run_id>/slide_NN_<name>.pptx
```

### 中扉（section-divider）配置の絶対ルール

中扉は、そのセクションの**最初**に配置（末尾ではない）。  
本デッキは3セクション構成のため、中扉は3つ（slide_03 / slide_05 / slide_09）。

### 色とセクション番号の連動

| セクション# | 中扉 section_number | 色 | TOC sections[i].title |
|---|---|---|---|
| 1 | 1 | 紺 | 市場規模・成長 |
| 2 | 2 | 青 | 競争構造 |
| 3 | 3 | 緑 | 成功要因と外部環境 |

---

## Step 6: マージ順序照合表 + merge_order.json

**進捗**: 開始時 `TaskCreate(subject="market-overview: Step 6 - merge_order.json 構築")` → Step 6-a / 6-b 両方完了後に `TaskUpdate(completed)` + `task_state.json` 更新。

<!-- source: skills/_common/references/orchestrator_contract.md (manual sync until D2) -->
<!-- merge_order.json / merge_warnings.json / regeneration_hint の正規スキーマは上記参照 -->


### Step 6-a: 照合表（Markdown）の出力＆セルフチェック

```markdown
## マージ順序照合表（Market Overview 標準版）

| 通し番号 | ファイル名 | 種別 | セクション | 備考 |
|---|---|---|---|---|
| 01 | slide_01_exec_summary.pptx | エグサマ | 冒頭 | - |
| 02 | slide_02_toc.pptx | 目次 | 冒頭 | sections=3 |
| 03 | slide_03_section1_market_size.pptx | **中扉** | Section 1冒頭 ✓ | section_number=1 |
| 04 | slide_04_market_environment.pptx | コンテンツ | Section 1 | - |
| 05 | slide_05_section2_competition.pptx | **中扉** | Section 2冒頭 ✓ | section_number=2 |
| 06 | slide_06_market_share.pptx | コンテンツ | Section 2 | - |
| 07 | slide_07_positioning.pptx | コンテンツ | Section 2 | - |
| 08 | slide_08_competitor_summary.pptx | コンテンツ | Section 2 | - |
| 09 | slide_09_section3_success_factors.pptx | **中扉** | Section 3冒頭 ✓ | section_number=3 |
| 10 | slide_10_market_kbf.pptx | コンテンツ | Section 3 | ⭐新規 |
| 11 | slide_11_pest.pptx | コンテンツ | Section 3 | - |
| 12 | slide_12_data_availability.pptx | コンテンツ | 末尾 | - |
```

セルフチェック項目（マージ実行前にすべてYES）：

- [ ] 全ての中扉が「そのセクションのコンテンツより前」に位置している
- [ ] TOCの `sections[i].title` と、対応する中扉の `title` が一致
- [ ] TOCの `sections[i].page` と、対応する中扉の通し番号が一致
- [ ] ファイル番号に歯抜け・重複・逆転がない
- [ ] エグゼクティブサマリーが通し番号01に配置
- [ ] データアベイラビリティが末尾に配置

### Step 6-b: merge_order.json 出力（merge-pptxv2 / visual-quality-reviewer 用）

`{{WORK_DIR}}/<run_id>/merge_order.json` に保存：

```json
{
  "entries": [
    {"slide_number": 1, "file_name": "slide_01_exec_summary.pptx",
     "skill_name": "executive-summary-pptx", "data_file": "data_01_exec_summary.json",
     "category": "header"},
    {"slide_number": 2, "file_name": "slide_02_toc.pptx",
     "skill_name": "table-of-contents-pptx", "data_file": "data_02_toc.json",
     "category": "header"},
    {"slide_number": 3, "file_name": "slide_03_section1_market_size.pptx",
     "skill_name": "section-divider-pptx", "data_file": "data_03_section1.json",
     "category": "section_divider"},
    {"slide_number": 4, "file_name": "slide_04_market_environment.pptx",
     "skill_name": "market-environment-pptx", "data_file": "data_04_market_environment.json",
     "category": "content"},
    {"slide_number": 5, "file_name": "slide_05_section2_competition.pptx",
     "skill_name": "section-divider-pptx", "data_file": "data_05_section2.json",
     "category": "section_divider"},
    {"slide_number": 6, "file_name": "slide_06_market_share.pptx",
     "skill_name": "market-share-pptx", "data_file": "data_06_market_share.json",
     "category": "content"},
    {"slide_number": 7, "file_name": "slide_07_positioning.pptx",
     "skill_name": "positioning-map-pptx", "data_file": "data_07_positioning.json",
     "category": "content"},
    {"slide_number": 8, "file_name": "slide_08_competitor_summary.pptx",
     "skill_name": "competitor-summary-pptx", "data_file": "data_08_competitor_summary.json",
     "category": "content"},
    {"slide_number": 9, "file_name": "slide_09_section3_success_factors.pptx",
     "skill_name": "section-divider-pptx", "data_file": "data_09_section3.json",
     "category": "section_divider"},
    {"slide_number": 10, "file_name": "slide_10_market_kbf.pptx",
     "skill_name": "market-kbf-pptx", "data_file": "data_10_market_kbf.json",
     "category": "content"},
    {"slide_number": 11, "file_name": "slide_11_pest.pptx",
     "skill_name": "pest-analysis-pptx", "data_file": "data_11_pest.json",
     "category": "content"},
    {"slide_number": 12, "file_name": "slide_12_data_availability.pptx",
     "skill_name": "data-availability-pptx", "data_file": "data_12_data_availability.json",
     "category": "footer"}
  ]
}
```

#### `category` フィールド規約

| 値 | 用途 | Market Overview 標準デッキでの該当 |
|---|---|---|
| `header` | セクション開始前の冒頭スライド | slide_01 (exec summary), slide_02 (TOC) |
| `content` | 通常のコンテンツスライド | slide_04, 06, 07, 08, 10, 11 |
| `section_divider` | 中扉 | slide_03, 05, 09 |
| `footer` | 末尾の付録的スライド | slide_12 (data availability) |

正規スキーマは `references/deck_skeleton_standard.json`。`merge-pptxv2` の
`--merge-order` 検証では `section_divider` の直後が必ず `content` であることを assert する
（違反は `merge_warnings.json` に記録され、マージは継続）。

---

## Step 7: 結合（merge-pptxv2）

**進捗**: 開始時 `TaskCreate(subject="market-overview: Step 7 - merge-pptxv2 結合")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。本 Step は `check_merge_order_exists.py` hook が `merge_order.json` の存在を assert（無ければ exit 2 でブロック）し、`validate_pptx_after_fill.py` hook が結合後 PPTX を自動検証する。

通し番号順に引数を並べてマージ。**`ls *.pptx | sort` の出力をそのまま流すのは禁止**。

```bash
pip install lxml --break-system-packages -q

python {{SKILL_DIR}}/merge-pptxv2/scripts/merge_pptx_v2.py \
  --merge-order {{WORK_DIR}}/<run_id>/merge_order.json \
  {{OUTPUT_DIR}}/MarketOverview_<market_name>_<date>.pptx \
  {{WORK_DIR}}/<run_id>/slide_01_exec_summary.pptx \
  {{WORK_DIR}}/<run_id>/slide_02_toc.pptx \
  {{WORK_DIR}}/<run_id>/slide_03_section1_market_size.pptx \
  {{WORK_DIR}}/<run_id>/slide_04_market_environment.pptx \
  {{WORK_DIR}}/<run_id>/slide_05_section2_competition.pptx \
  {{WORK_DIR}}/<run_id>/slide_06_market_share.pptx \
  {{WORK_DIR}}/<run_id>/slide_07_positioning.pptx \
  {{WORK_DIR}}/<run_id>/slide_08_competitor_summary.pptx \
  {{WORK_DIR}}/<run_id>/slide_09_section3_success_factors.pptx \
  {{WORK_DIR}}/<run_id>/slide_10_market_kbf.pptx \
  {{WORK_DIR}}/<run_id>/slide_11_pest.pptx \
  {{WORK_DIR}}/<run_id>/slide_12_data_availability.pptx
```

`<market_name>` はスネークケース化（例: `HRTech`）、`<date>` は実行日（YYYY-MM-DD）。
`--merge-order` を指定すると `section_divider` 位置検証が走り、結果は出力 PPTX と同じ
ディレクトリの `merge_warnings.json` に保存される（違反ゼロでも空配列で出力される）。

### merge 完了後: brand_warnings を merge_warnings.json に追記（必須）

merge-pptxv2 は `merge_warnings.json` を `"w"` モードで上書きするため、Step 5 中に蓄積した `brand_warnings` は merge 完了後にここで追記する。

```python
append_brand_warnings_to_merge_file(
    "{{OUTPUT_DIR}}/merge_warnings.json", brand_warnings,
)
# brand_warnings が空なら no-op（既存ファイルは触らない）。
# 末尾の Step 8（残存 issue 判断）でユーザーに warning 件数 + 内訳を提示すること。
```

### マージ後の最終検証

merge-pptxv2 の出力ログで、各スライド番号の Main Message と shape数が表示される。
中扉のshape数は8前後（タイトル＋サブタイトル＋トピックリストのみ）と少なくなり、
**「コンテンツ（多）→ 中扉（少）→ コンテンツ（多）」の谷**が3箇所（slide_03/05/09）に
出現するのが正常。

加えて、`merge_warnings.json` を確認し、`section_divider_position` 違反が 0 件であることを
チェック。違反があれば該当 `slide_index` の前後をデータで見直す（中扉の直後に header/footer/
別の中扉が混入していないか、merge_order.json の `category` 設定ミスではないか）。

---

## Step 8: ビジュアル品質レビュー＋自動修正ループ

**進捗**: 開始時 `TaskCreate(subject="market-overview: Step 8 - Visual Review + 自動修正")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。`severity=high` 残存時は **`AskUserQuestion`** で手動修正 / 許容を選ばせる必須。自動修正ループは別カウンタ（最大 2 ラウンド）。

<!-- source: skills/_common/prompts/step_final_visual_review_loop.md (manual sync until D2) -->

### Step 8-a: visual-quality-reviewer 起動

入力:
- `merged_pptx`: `{{OUTPUT_DIR}}/MarketOverview_<market_name>_<date>.pptx`
- `merge_order`: `{{WORK_DIR}}/<run_id>/merge_order.json`
- `data_dir`: `{{WORK_DIR}}/<run_id>/`

<!-- @if:claude_code -->
出力: `{{FACTORY_ROOT}}/work/visual-quality-reviewer/visual_review_report.json`
<!-- @endif -->

### Step 8-b: レビュー結果の分岐

| `overall_verdict` | 処理 |
|---|---|
| `pass` | Step 9 へ進む |
| `needs_fixes` かつ `severity=high` ≥1件 | **自動修正ループ**（最大2ラウンド）|
| `needs_fixes` かつ `severity=high` =0件 | ユーザーに差分提示、手動修正 or 許容を選ばせる |
| `reject` | LibreOffice レンダリング失敗を疑い、ユーザーに報告して停止 |

### 自動修正ループ

`severity=high` の各 issue について:
1. `issues[i].skill_name` と `issues[i].data_file` から、該当スライドの JSON を特定
2. `issues[i].regeneration_hint` に従って `data_NN_*.json` を修正
3. 該当スキルの `fill_*.py` を**同じ slide_NN ファイル名で再実行** → 上書き
4. 全修正完了後、`merge-pptxv2` を再実行
5. 再度 `visual-quality-reviewer` を起動

**2ラウンド終了時点で `high` が残れば、ユーザーに残存 issue を提示して判断を仰ぐ**。
カウンタを必ず持って無限ループを防止すること。

---

## Step 9: FactCheck_Report.md 生成（最終納品物）

**進捗**: 開始時 `TaskCreate(subject="market-overview: Step 9 - FactCheck_Report.md 生成")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。

`fact_check_report.json` を Markdown に整形して、`{{OUTPUT_DIR}}/FactCheck_Report_<market_name>_<date>.md`
に保存する。テンプレートは `prompts/step9_factcheck_md_template.md` を参照。

### MD レポートの構成（必須セクション）

```markdown
# ファクトチェックレポート: {market_name}

**実行日**: {YYYY-MM-DD}  
**スコープ**: high_risk / all / skip  
**総合判定**: pass / needs_user_review / skipped  
**検証件数**: {N} 件 / フラグ件数: {M} 件（high={a} / medium={b} / low={c}）

---

## 1. クレーム別検証結果（スライド番号順）

| スライド | データファイル | 主張 | 検証結果 | severity | 出典URL（裏取り）| 補足 |
|---|---|---|---|---|---|---|
| 4 | data_04_market_environment.json | 市場規模1.2兆円（2024年）| confirmed | low | https://... | - |
| 6 | data_06_market_share.json | A社シェア25.3% | discrepancy | high | https://... | IRでは22.8%。要修正 or 注釈 |
| ... |

## 2. 自信度サマリー

- **High confidence**（複数ソース確認）: {X} 件
- **Medium confidence**（単一ソース確認）: {Y} 件
- **Low confidence**（裏取り不可・推定）: {Z} 件

## 3. データギャップ（Web 公開情報では取得不可）

データアベイラビリティ ✗/△ 項目を転記:

| カテゴリ | 項目 | 理由 | 推奨対応 |
|---|---|---|---|
| 市場規模 | 2026年予測値 | 公的統計が2024年までしか公開されていない | 業界団体・矢野経済の予測レポートの追加調達 |

## 4. 検索クエリログ

検索したキーワードと取得した出典URLの一覧:

| クレームID | クエリ | 取得URL | 取得日 |
|---|---|---|---|
| c001 | "国内HR Tech市場 規模 2024" | https://... | 2026-04-26 |
```

### MD レポートの位置付け

- **クライアント納品物**として PPTX と並列に位置付ける
- 各クレームに `severity`（高/中/低）と検証結果（confirmed / single_source / discrepancy / not_found / stale）を必ず記録
- 出典URLは裏取りに使った検索結果のURLを記録（一次情報を辿れるように）

---

## Step 10: ユーザーへ提示

**進捗**: 開始時 `TaskCreate(subject="market-overview: Step 10 - 最終納品")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。

最終的に2つのファイルを提示する：

1. `{{OUTPUT_DIR}}/MarketOverview_<market_name>_<date>.pptx`（PPTX デッキ 12枚）
2. `{{OUTPUT_DIR}}/FactCheck_Report_<market_name>_<date>.md`（ファクトチェックレポート）

提示メッセージ例：

```
✅ 市場調査レポートの生成が完了しました。

📊 PPTXデッキ: MarketOverview_HRTech_2026-04-26.pptx（12枚）
🔎 ファクトチェック: FactCheck_Report_HRTech_2026-04-26.md
   - 検証件数: 28件 / フラグ: 3件（high=0, medium=2, low=1）
   - 総合判定: pass

【次に検証すべき論点（公開情報で確定できなかった項目）】
- 2026年以降の市場規模予測（公的統計未整備、業界団体レポート要確認）
- 上位 N 社以外の新興プレイヤーのシェア（決算非開示）

ご確認ください。
```

---

## 依存スキル一覧

### コアスキル（必須）

| スキル名 | 配置 | 役割 |
|---|---|---|
| `executive-summary-pptx` | user | デッキ冒頭の Key Findings サマリー |
| `table-of-contents-pptx` | user | 目次（3セクション）|
| `section-divider-pptx` | user | 中扉（3枚: Section 1/2/3 冒頭）|
| `market-environment-pptx` | user | 市場規模推移＋CAGR（論点1）|
| `market-share-pptx` | user | 市場シェア（論点3）|
| `positioning-map-pptx` | user | プレイヤー位置付け（論点4）|
| `competitor-summary-pptx` | user | 各社戦略比較（論点5）|
| `market-kbf-pptx` ⭐ | user | KBF×`kbf_count`（既定 3）のテーブル（論点2、v0.1新規）|
| `pest-analysis-pptx` | user | PEST環境（追加論点）|
| `data-availability-pptx` | user | データ取得状況・調査制約 |
| `merge-pptxv2` | user | 全スライドの結合 |

### 品質レビュー系スキル（必須）

| スキル名 | 配置 | 呼び出し位置 | 役割 |
|---|---|---|---|
| `fact-check-reviewer` | user | Step 2.5 | Web取得情報の裏取り |
| `visual-quality-reviewer` | user | Step 8 | マージ後デッキの目視レビュー＋自動修正 |

⭐ = v0.1 で新規開発

---

## 品質チェックリスト

### デッキ全体

- [ ] エグゼクティブサマリーが冒頭（slide_01）に配置されている
- [ ] 目次（slide_02）の `sections[]` が3要素（市場規模・成長 / 競争構造 / 成功要因と外部環境）
- [ ] 中扉が3枚（slide_03/05/09）配置され、それぞれ section_number = 1/2/3
- [ ] データアベイラビリティが末尾（slide_12）に配置されている

### 中扉配置（過去バグ再発防止）

- [ ] 全ての中扉が「そのセクションのコンテンツより前」に位置している
- [ ] ファイル名の番号と最終順序が一致している
- [ ] マージ後の shape 数推移で「コンテンツ→中扉→コンテンツ」の谷が3箇所（03/05/09）に正しく出現している
- [ ] 目次の `sections[i].page` が、対応する中扉の通し番号（3/5/9）と一致している

### 内容の一貫性

- [ ] 主要競合（`scope.json.max_competitors` 社）が全PPTXスキル（market-share / positioning-map / competitor-summary / market-kbf）で一致している
- [ ] 全スライドのメインメッセージが事実記述ベース（「〜すべき」が含まれていない）
- [ ] エグゼクティブサマリーの 5 Findings が他スライドの内容と整合している
- [ ] FactCheck_Report.md にすべての data_NN_*.json のクレームが記録されている
- [ ] FactCheck_Report.md の `data_gaps` セクションが、データアベイラビリティの ✗/△ 項目と一致している

### 納品物

- [ ] `{{OUTPUT_DIR}}/MarketOverview_<market_name>_<date>.pptx` が生成された
- [ ] `{{OUTPUT_DIR}}/FactCheck_Report_<market_name>_<date>.md` が生成された
- [ ] PowerPoint で開いた際に修復ダイアログが出ない

---

## 注意事項

- **公開情報主義**: Web情報・ユーザーアップロード情報のみで分析する。推定や断定は避ける。
- **競合社数上限**: `scope.json.max_competitors`（既定 5、上限 5、2026-04-29 に 10→5 撤回）に従い、ポジショニング・シェア・競合比較・KBF 事例で**同じ社**を採用する（一貫性）。`kbf_count` も同様にスキル横断で揃える。
- **検証論点の明示**: 公開情報で確定できないものは `data_gaps` または FactCheck_Report.md の追加調査推奨セクションへ明示する。
- **Web 検索深度**: 1論点あたり 5〜8 コール程度を目安。深度の議論は v0.2 以降。
- **無限ループ防止**: 自動修正ループは最大2ラウンド。3ラウンド目に入る前に必ずユーザーに判断を仰ぐ。

---

## アセット

| ファイル | 内容 |
|---|---|
| `prompts/step0_scope_clarification.md` | Step 0 の AskUserQuestion 雛形 |
| `prompts/step1_research_template.md` | Step 1 の論点別検索キーワード |
| `prompts/step9_factcheck_md_template.md` | FactCheck_Report.md の雛形 |
| `references/deck_skeleton_standard.json` | 標準デッキ12枚の順序とスキル割当 |
