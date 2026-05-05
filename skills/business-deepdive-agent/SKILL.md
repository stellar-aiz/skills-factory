---
name: business-deepdive-agent
description: >
  対象会社の特定の事業セグメント（有報の報告セグメント単位）について
  「事業の概要・ビジネスモデル・差別化・顧客・顧客成長」の 5 論点を深掘りし、
  5 枚の PowerPoint スライドを生成するオーケストレータースキル。

  単独でも起動可能だが、主には company-deepdive-agent から各セグメントごとに
  並列起動される。出力 PPTX 群は親オーケストレータが merge-pptxv2 で他セグメント分や
  会社レベル分と結合する設計（本スキルは結合まで行わない）。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「○○社の○○事業の深掘り」「事業セグメント分析」「事業の戦略を透視」という言葉が出た場合
  - company-deepdive-agent から呼び出された場合
  - 多角化企業の特定セグメントの戦略を、会社全体ではなく事業単位で深く調べたい場合
  - 顧客市場・差別化・ビジネスモデルを 1 つの事業セグメントに絞って整理したい場合

  以下の場合は別スキルを使う:
  - 「会社全体（複数事業横断）を深掘り」 → company-deepdive-agent
  - 「市場全体の調査」 → market-overview-agent
  - 単一の PPTX スライドだけ作りたい → 該当する個別 PPTX スキル（business-overview-pptx 等）
---

# 事業セグメント深掘りオーケストレーター

ISSUE-004（v0.3）における新規オーケストレーター。`company-deepdive-agent` の下流で
各事業セグメントを深掘りし、5 枚の個別 PPTX を返却する。

## 設計原則

- **対象は単一の事業セグメント** のみ。複数事業の場合は親が本スキルを N 回呼ぶ
- **結合は本スキルでは行わない**: 5 枚の個別 PPTX を出力し、merge は親（`company-deepdive-agent`）が担当
- **fact-check / visual-review は本スキルでは呼ばない**: 親オーケストレーターが統合デッキに対して一括実施する（重複コール削減）
- 公開情報のみ。取れない項目は data-availability に「✗」記録（親オーケストレータに引き渡す）

---

## 進捗トラッキング規約（全 Step 横断、必須）

<!-- source: skills/_common/prompts/step_state_tracking.md (manual sync until D2) -->

各 Step の開始/終了で `TaskCreate` / `TaskUpdate` を呼び、`task_state.json` を更新する。詳細規約は `skills/_common/prompts/step_state_tracking.md` を正本とする。

- **subject フォーマット**: `business-deepdive: Step <N> - <topic>`(例: `business-deepdive: Step 1 - Web検索 (5論点)`)
- **task_state.json 配置**: `{{WORK_DIR}}/company-deepdive-agent/<parent_run_id>/segments/<segment_slug>/task_state.json`(scope.json と同じディレクトリ。単独起動でも同様の構造を擬似生成)
- **開始時**: `TaskCreate` で task を起こす → `TaskUpdate(in_progress)` → `task_state.json.steps[]` に append
- **終了時**: `TaskUpdate(completed)` → `task_state.json` の該当 entry を `completed` + `completed_at` に更新
- **失敗・再試行時**: `TaskUpdate(completed)` を呼ばない。`task_state.json` の `retry_count` のみインクリメント。2 回失敗したらユーザーに判断を仰ぐ

`tools/hooks/check_task_progression.py` が `fill_*.py` / `merge_pptx_v2.py` 起動前にこのファイルを参照し、Step ordering inversion（前 Step が `completed` でないまま次 Step に進んだ状態）を検出して exit 2 でブロックする。

---

## ハーネスレバー利用規約（参照）

<!-- source: skills/_common/references/harness_levers.md (manual sync until D2) -->

本オーケストレーターは Claude Code ハーネス機構を以下のとおり活用する。詳細規約は `skills/_common/references/harness_levers.md` を参照。

| レバー | 適用箇所 |
|---|---|
| ① hooks (`tools/hooks/*.py`) | `check_merge_order_exists` / `validate_pptx_after_fill` / `check_task_progression` / `load_session_context` が `.claude/settings.json` 経由で発火（親 `company-deepdive-agent` 配下で merge / validate 実施時にも有効）|
| ② subagent (`.claude/agents/research-subagent.md`) | Step 1 で 5 論点を並列 Web 検索する際に **必ず** Agent ツール経由で呼ぶ。生 HTML / 検索結果を親 context に積まない |
| ③ TaskCreate / AskUserQuestion | 各 Step の開始/終了で TaskCreate（上記規約）。Step 0（単独起動時の対話）と Step 3（5 論点 Markdown 承認）で AskUserQuestion 必須 |

---

## 事業レベル 5 論点 → PPTX マッピング

| # | 論点 | PPTX スキル |
|---|---|---|
| 1 | 事業の概要は？ | `business-overview-pptx` |
| 2 | その事業のビジネスモデルは？ | `business-model-pptx` |
| 3 | その事業の差別化ポイントは？ | `value-chain-matrix-pptx`（バリューチェーン上のポジショニング） |
| 4 | その事業の顧客は誰か？ | `customer-profile-pptx`（顧客企業 / 顧客セグメントのプロファイル）|
| 5 | その事業の顧客は成長するか？ | `market-environment-pptx`（**顧客側市場**の規模・成長率推移）|

「顧客は誰か」「顧客は成長するか」は **顧客側** の情報。本スキルの対象事業の顧客
（B2B なら主要取引先・B2C なら顧客セグメント）を扱う。

---

## 出力スライド構成（5 枚）

```
B-01 事業の概要 (business-overview-pptx)
B-02 ビジネスモデル (business-model-pptx)
B-03 差別化（バリューチェーン上のポジション）(value-chain-matrix-pptx)
B-04 主要顧客プロファイル (customer-profile-pptx)
B-05 顧客市場の成長性 (market-environment-pptx)
```

ファイル番号は親 `company-deepdive-agent` のグローバル通し番号で書き換えられる
（B-01..B-05 は本スキル内での暫定番号）。

エグゼクティブサマリーは Phase 4 の `company-deepdive-agent` で
「会社レベル+全セグメント」の文脈で生成する設計のため、本スキルでは生成しない。

---

## 作業ディレクトリ規約

本スキルは独自の work dir を持たず、**親 `company-deepdive-agent` 配下のセグメント別 subdir** に
全ファイルを書き込む。これにより `comparison-synthesis-agent` や `merge-pptxv2` が一元的に参照可能。

```
{{WORK_DIR}}/company-deepdive-agent/<parent_run_id>/segments/<segment_slug>/
├── data_<NN>_business_overview.json
├── data_<NN+1>_business_model.json
├── data_<NN+2>_value_chain_matrix.json
├── data_<NN+3>_customer_profile.json
├── data_<NN+4>_market_environment.json
├── slide_<NN>_business_overview.pptx
├── slide_<NN+1>_business_model.pptx
├── slide_<NN+2>_value_chain_matrix.pptx
├── slide_<NN+3>_customer_profile.pptx
├── slide_<NN+4>_market_environment.pptx
├── segment_data_availability.json
└── segment_summary.json
```

`<NN>` = `global_slide_offset + 1`。`<segment_slug>` は親が決定（事業名 → URL-safe ASCII）。

**単独起動の場合**も同じ構造を擬似生成する（後述）。

---

## Step 構造

### Step 0: 引数受領 / 単独起動時の対話

**進捗**: 開始時 `TaskCreate(subject="business-deepdive: Step 0 - 引数受領")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。単独起動の場合は `AskUserQuestion` 必須。

<!-- source: skills/_common/prompts/step0_brand_clarification.md (manual sync until D2) -->

#### 内部呼び出しの場合（推奨）

`company-deepdive-agent` から以下のパラメータを JSON で受け取る。**`brand` フィールドは必須**（親 agent の `scope.json.brand` を転写）。受領した brand は子側で AskUserQuestion で再質問せず、そのまま fill 起動時の `--brand` 引数に伝播する:

```json
{
  "parent_company_name": "第一交通産業株式会社",
  "segment_name": "タクシー事業",
  "segment_slug": "taxi",
  "parent_run_id": "2026-04-29_daiichikoutsu",
  "global_slide_offset": 11,
  "is_listed": true,
  "industry": "陸運業",
  "analysis_years": 7,
  "brand": "stellar_aiz"
}
```

`brand` フィールドが欠落している場合（古い親 agent からの呼び出し等）は `stellar_aiz` を既定として処理する。

#### 単独起動の場合

AskUserQuestion で以下を聞く（**ブランドは先頭で確定する**。詳細仕様は `skills/_common/prompts/step0_brand_clarification.md` を正本とする）:

1. **`brand`** — 出力 PPTX フォーマット（`_discover_brands()` で動的取得した選択肢から、既定 `stellar_aiz`）
2. `parent_company_name` — 対象会社の正式名（例: 「第一交通産業株式会社」）
3. `segment_name` — 深掘り対象のセグメント名（例: 「タクシー事業」）
4. `industry` — 業種（任意、検索クエリ精度向上用）
5. `analysis_years` — 顧客市場分析の年数（任意、default 7 年）

ブランド質問の実装パターン（agnostic、`_discover_brands()` で動的取得）:

```python
import json, os, sys
sys.path.insert(0, os.path.join("{{SKILL_DIR}}", "..", "_common", "lib"))
from brand_resolver import _discover_brands, _BRANDS_DIR

discovered = _discover_brands()
options = []
for brand_id in discovered:
    with open(os.path.join(_BRANDS_DIR, brand_id, "theme.json")) as f:
        theme_data = json.load(f)
    label = theme_data.get("label", brand_id)
    if brand_id == "stellar_aiz":
        label += " (Recommended)"
    options.append({"label": label, "description": f"id={brand_id}"})
# AskUserQuestion(question="...", header="ブランド", options=options, multiSelect=False)
```

`parent_run_id` は `YYYY-MM-DD_<parent_company_slug>` 形式で自動生成（`<parent_company_slug>` は会社名を ASCII slug 化した値）。
`segment_slug` は `segment_name` を ASCII 化した値（例: タクシー事業 → `taxi`）を生成し、ユーザーに確認して必要なら修正させる。
`global_slide_offset = 0`（単独起動なので NN = 1..5）。

### Step 1: research-subagent 経由でセグメント別 5 論点情報収集

**進捗**: 開始時 `TaskCreate(subject="business-deepdive: Step 1 - Web検索 (5論点)")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。

5 論点それぞれについて、`research-subagent`(`.claude/agents/research-subagent.md`) を **Agent ツール経由で起動** して情報収集する。各 subagent は 5-8 件の Web 検索 + 必要に応じた fetch を実施し、`output_schema` に沿った要約済み JSON のみ親に返却する。**生 HTML / 検索結果テキストは subagent 自身の context 内で完結し、親 context には流入しない**。

#### subagent 呼び出しパターン（論点ごと）

```python
import json
from skills._common.lib.parse_subagent_return import parse_subagent_return
result = Agent(
    subagent_type="research-subagent",
    description=f"{segment_name} の<論点名>調査",
    prompt=json.dumps({
        "topic_id": "data_<NN>_business_overview",  # 論点別の data ファイル名と対応
        "topic_description": "<論点の自然文要約>",
        "output_schema": {<該当 PPTX スキルの JSON schema>},
        "parent_context": {
            "industry": industry,
            "target_company": parent_company_name,
            "scope_constraints": {}  # 業務上の境界がある場合のみ
        },
        "search_budget": {"min_searches": 5, "max_searches": 8}
    })
)
# subagent return は parse_subagent_return() 経由で dict 化する（必須）。
# 直接 json.loads(result) しないこと: subagent が稀に前置き文・code fence・末尾
# Sources を混入させるため（ISSUE-009）。helper はそれらを吸収する。
parsed = parse_subagent_return(result)
# parsed["data"] を {{work_dir}}/data_<NN>_<topic>.json に Write で書き出す
# parsed["open_questions"] / parsed["sources_summary"] は segment_data_availability.json と segment_summary.json に転記
```

`prompts/step1_research_template.md` のクエリテンプレートは subagent 起動時の `topic_description` 構築のヒントとして親が参照する（subagent は受け取った `topic_description` を起点に検索する）。

| 論点 | 優先ソース | subagent への topic_description で強調すべき点 |
|---|---|---|
| 事業の概要 | 有報「事業の状況」/ セグメント情報 / 公式 HP の事業説明ページ | セグメントの売上・主要製品・主要数字 |
| ビジネスモデル | 統合報告書 / セグメント別事業説明 / IR Day 資料 | サプライヤー・顧客の取引フロー、価値交換の構造 |
| 差別化 | 業界レポート / IR Q&A / メディアインタビュー（**自社が属する業界**のバリューチェーン） | 自社業界（顧客業界ではない）のバリューチェーン上のポジション |
| 顧客 | 有報「主要販売先」/ 業界レポート / 顧客側 IR 資料(**顧客の業種 `customer_industry` を確定**)| 顧客企業のプロファイル + 顧客の業種を必ず特定 |
| 顧客の成長 | **`customer_industry` の業界レポート**（矢野経済・富士経済・官公庁統計等）。**自社の事業市場ではない**| 論点 4 で確定した `customer_industry` の市場規模・成長率 |

#### 順序厳守: 論点 4 → 論点 5

論点 5（顧客の成長）は、論点 4（顧客）で確定した **`customer_industry`** を入力に取る。論点 4 を飛ばして論点 5 を先に subagent 起動すると、自社の事業市場と顧客市場を取り違える典型的な混同（後述「論点間整合性ルール」）が発生する。

そのため subagent 起動も **論点 1-4 を完了してから論点 5 を起動** する逐次パターンを必須とする（論点 1-3 は並列起動可、論点 4 完了 → `customer_industry` 確定 → 論点 5 の `parent_context.industry` をその値に切り替えて起動）。

```
論点 1-3 並列 subagent → 結果 JSON を data_<NN>_*.json に書き出し
   ↓
論点 4 subagent → 顧客企業/セグメント特定 → customer_industry = ?（業種を確定）
   ↓
論点 5 subagent: parent_context.industry = customer_industry
```

### Step 2: data-availability セグメント単位記録

**進捗**: 開始時 `TaskCreate(subject="business-deepdive: Step 2 - data-availability")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。

各論点について「取得済(✓) / 一部取得(△) / 未取得(✗)」を `segment_data_availability.json` に記録:

```json
{
  "segment_name": "タクシー事業",
  "parent_company_name": "第一交通産業株式会社",
  "segment_slug": "taxi",
  "items": [
    {"topic": "business_overview", "status": "obtained", "source": "有報 第65期 セグメント情報"},
    {"topic": "business_model", "status": "partial", "source": "公式HP"},
    {"topic": "value_chain_matrix", "status": "partial", "source": "業界レポート 推測込み"},
    {"topic": "customer_profile", "status": "obtained", "source": "有報 主要販売先"},
    {"topic": "market_environment", "status": "obtained", "source": "国土交通省 タクシー事業概況"}
  ]
}
```

親 `company-deepdive-agent` が全社統合の data-availability スライドに転記する。

### Step 3: 5 つの data_NN_*.json を作成 → ユーザー承認

**進捗**: 開始時 `TaskCreate(subject="business-deepdive: Step 3 - data JSON 作成 + 承認")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。**ユーザー承認の `AskUserQuestion` 必須**（自由対話での承認は禁止）。

各論点について、対応 PPTX スキルの SKILL.md に従って `data_NN_<topic>.json` を作成。
ファイル名:

| 論点 | data ファイル名 |
|---|---|
| 1 事業の概要 | `data_<NN>_business_overview.json` |
| 2 ビジネスモデル | `data_<NN+1>_business_model.json` |
| 3 差別化 | `data_<NN+2>_value_chain_matrix.json` |
| 4 顧客 | `data_<NN+3>_customer_profile.json` |
| 5 顧客成長 | `data_<NN+4>_market_environment.json` |

`<NN>` = `global_slide_offset + 1`。

5 枚分の `main_message` / 主要内容を Markdown でユーザーに提示し、承認を得る。

#### Step 3 出力フォーマット（論点間整合性確認表を必須）

ユーザー承認 Markdown には、5 論点の内容に加えて以下の整合性確認表を**必ず含める**。LLM が JSON を組み立てる前に明示的に点検する強制プロセス。

```markdown
## 論点間整合性確認（必須）

| # | 確認項目 | 期待される値 | 実態 | OK? |
|---|---|---|---|---|
| 1 | 全論点の主語 = 論点 1 の `segment_name` | <セグメント名、例：施設運営事業> | <各論点の主語> | ✅/❌ |
| 2 | 論点 4 顧客の業種 = 論点 5 顧客市場の業種 | <顧客業種、例：不動産業> | <論点 5 で扱う市場の業種> | ✅/❌ |
| 3 | 論点 3 業界 = 自社の業界 | <自社業界、例：ビルメンテナンス業> | <論点 3 で扱う業界> | ✅/❌ |
```

❌ がある場合は JSON を組み立てる**前**に修正する。詳細は後述「論点間整合性ルール」を参照。

### Step 4: 5 つの fill_*.py を順次実行

**進捗**: 開始時 `TaskCreate(subject="business-deepdive: Step 4 - PPTX 生成 (5枚)")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。本 Step 中は `validate_pptx_after_fill.py` hook が各 fill_*.py 実行後に PPTX 整合性を自動検証する（壊れていれば exit 2 で停止）。

#### Step 4 開始前: brand fallback バッファ初期化（必須）

Step 0 で受領した（または単独起動時に確定した）`brand` を使い、5 つの fill 起動それぞれで `resolve_fill_brand_with_warning()` を呼ぶ。**本スキルは merge を実施しない**ため、`brand_warnings` は `segment_summary.json` に含めて親（company-deepdive-agent）に返却し、親が merge 完了後に `merge_warnings.json` へ追記する責務を持つ。

```python
import os, sys, subprocess
sys.path.insert(0, os.path.join("{{SKILL_DIR}}", "..", "_common", "lib"))
from orchestrator_helpers import resolve_fill_brand_with_warning

scope_brand = "stellar_aiz"  # Step 0 で受領した brand（既定 stellar_aiz）
brand_warnings: list = []

# 5 fill それぞれに対して同じパターン:
# skill_dir = os.path.join("{{SKILL_DIR}}", "<skill-name>-pptx")
# fill_brand = resolve_fill_brand_with_warning(skill_dir, scope_brand, brand_warnings)
# subprocess.run(["python", os.path.join(skill_dir, "scripts", "fill_<name>.py"),
#                 "--brand", fill_brand, "--data", "...", "--output", "..."], check=True)
```

承認後、5 つの PPTX を生成（すべての起動で `--brand <fill_brand>` を渡す）:

```bash
# 1. 事業の概要
python ~/.claude/skills/business-overview-pptx/scripts/fill_business_overview.py \
  --brand stellar_aiz \
  --data <work_dir>/data_<NN>_business_overview.json \
  --template ~/.claude/skills/business-overview-pptx/assets/business-overview-template.pptx \
  --output <work_dir>/slide_<NN>_business_overview.pptx

# 2. ビジネスモデル
python ~/.claude/skills/business-model-pptx/scripts/fill_business_model.py \
  --brand stellar_aiz \
  --data <work_dir>/data_<NN+1>_business_model.json \
  --template ~/.claude/skills/business-model-pptx/assets/business-model-template.pptx \
  --output <work_dir>/slide_<NN+1>_business_model.pptx

# 3. 差別化（バリューチェーン上のポジション）
python ~/.claude/skills/value-chain-matrix-pptx/scripts/fill_value_chain_matrix.py \
  --brand stellar_aiz \
  --data <work_dir>/data_<NN+2>_value_chain_matrix.json \
  --template ~/.claude/skills/value-chain-matrix-pptx/assets/value-chain-matrix-template.pptx \
  --output <work_dir>/slide_<NN+2>_value_chain_matrix.pptx

# 4. 主要顧客プロファイル
python ~/.claude/skills/customer-profile-pptx/scripts/fill_customer_profile.py \
  --brand stellar_aiz \
  --data <work_dir>/data_<NN+3>_customer_profile.json \
  --template ~/.claude/skills/customer-profile-pptx/assets/customer-profile-template.pptx \
  --output <work_dir>/slide_<NN+3>_customer_profile.pptx

# 5. 顧客市場の成長性
python ~/.claude/skills/market-environment-pptx/scripts/fill_market_environment.py \
  --brand stellar_aiz \
  --data <work_dir>/data_<NN+4>_market_environment.json \
  --template ~/.claude/skills/market-environment-pptx/assets/market-environment-template.pptx \
  --output <work_dir>/slide_<NN+4>_market_environment.pptx
```

`<work_dir>` = `{{WORK_DIR}}/company-deepdive-agent/<parent_run_id>/segments/<segment_slug>/`

各 fill_*.py が成功したか確認（exit code 0、出力 PPTX 存在）。`main_message` 65 字超過時の hard-fail は本スキルの想定エラーとして再生成ループで対処する（`prompts/main_message_principles.md` を参照して書き直し → 該当 fill_*.py を再実行）。

### Step 5: segment_summary.json を出力

**進捗**: 開始時 `TaskCreate(subject="business-deepdive: Step 5 - segment_summary")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。

```json
{
  "segment_name": "タクシー事業",
  "parent_company_name": "第一交通産業株式会社",
  "segment_slug": "taxi",
  "parent_run_id": "2026-04-29_daiichikoutsu",
  "global_slide_offset": 0,
  "slide_files": [
    "slide_01_business_overview.pptx",
    "slide_02_business_model.pptx",
    "slide_03_value_chain_matrix.pptx",
    "slide_04_customer_profile.pptx",
    "slide_05_market_environment.pptx"
  ],
  "key_findings": [
    {"category": "差別化", "finding": "..."},
    {"category": "顧客", "finding": "..."}
  ],
  "open_questions": [
    "公開情報では確認できず、業界ヒアリング推奨の論点 1",
    "..."
  ],
  "data_gaps": [
    "セグメント別営業利益率の詳細内訳",
    "..."
  ],
  "brand": "stellar_aiz",
  "brand_warnings": []
}
```

`open_questions` は親 `comparison-synthesis-agent` で全社統合の検証論点に集約される。

`brand` は Step 0 で受領した値をそのまま転記。`brand_warnings` は Step 4 で蓄積した未対応 fill 検出ログ（空配列なら 0 件）。**親 `company-deepdive-agent` は本セグメントの `brand_warnings` を全セグメント分まとめて受け取り、merge 完了後に `merge_warnings.json` へ追記する責務を持つ**。子側（本スキル）は `merge_warnings.json` を直接書かない。

### Step 6: 終了

**進捗**: 開始時 `TaskCreate(subject="business-deepdive: Step 6 - 終了処理")` → 完了時 `TaskUpdate(completed)` + `task_state.json` 更新。

#### 単独起動の場合
- ユーザーに 5 PPTX のフルパスを提示して終了
- merge は実施しない（必要なら別途 merge-pptxv2 を起動）

#### 内部呼び出しの場合
- 親 `company-deepdive-agent` に以下を返却:
  - 5 枚の `slide_*.pptx`（フルパス）
  - `segment_data_availability.json`（フルパス）
  - `segment_summary.json`（フルパス）
- 親が merge_order.json を組み、merge-pptxv2 と visual-quality-reviewer を呼ぶ

---

## main_message 共通ルール（5 PPTX 全てに適用、厳守）

<!-- source: skills/_common/prompts/main_message_principles.md (manual sync until D2) -->

### ルール1: 長さは **65 文字以内**（厳格）

- 句読点・記号・スペースを含めて 65 文字以内
- テンプレート最上部のメッセージ枠が固定幅のため、超えた場合は要約や段落分けではなく **書き直し**
- 65 文字を 1 字でも超えた状態で `fill_*.py` に渡すと ValueError で hard-fail する

### ルール2: トーンは **事実記述ベース**（「〜すべき」禁止）

- 公開情報のみで断定できないアクションや戦略示唆は書かない
- 不明な点は「〜は公開情報からは確定できず追加調査が必要」と率直に書く（検証論点として明示）

**例**:
- ✗ 「対象会社は海外展開を加速すべき」（公開情報では断定不可）
- ✓ 「対象会社は国内売上比率が 90% と高く、海外展開の実績は限定的である」（事実記述）
- ✓ 「対象会社の海外展開方針は Web 情報では限定的、マネジメントインタビューで確認が必要」（検証論点）

### 65 字オーバー時の短縮原則 4 つ

1. **主語は 1 つだけ** — 「市場は」「主要プレイヤーは」「対象会社は」のいずれか 1 つに絞る
2. **修飾語を削除** — 「主要な」「重要な」「大きな」「急速な」等の主観的な修飾語を落とす
3. **数値は 1 つだけ残す** — CAGR と シェアを両方載せず、より重要な 1 つを選ぶ
4. **結論を述べる、根拠は本文スライドに任せる** — 「〜だから〜である」の前段を切り、結論部のみ残す

注: `business-model-pptx` / `value-chain-matrix-pptx` / `customer-profile-pptx` の SKILL.md は「〜すべきで締める / 70 字」と書かれているが、本オーケストレーターから呼ぶ場合は **本ルールで上書き**（65 字・事実記述）。

---

## 論点間整合性ルール（5 論点を 1 デッキに統合する際、厳守）

<!-- source: skills/_common/prompts/cross_topic_consistency_check.md (manual sync until D2) -->

複数の論点で 1 つのデッキを構成するため、各論点が独立して LLM に処理されると**論点間の意味的整合性**が崩れやすい。既存の品質ゲート（main_message ルール / fill_*.py validate / visual-quality-reviewer）はすべて**単一スライドの自己整合性**しか見ないため、論点間の混同は構造的に検出されない。

### 絶対ルール

#### ルール1: 全論点で扱う主語は同じ事業セグメント名

論点 1〜5 すべての主語は `segment_name`（例: 「施設運営事業」）。論点ごとに「全社」「グループ」「他事業」に揺れない。

#### ルール2: 「顧客市場」は**顧客が属する業種の市場**であり、自社の事業市場ではない

これは特に **B2B 受託業（自分自身がサービス提供者の場合）** で混同が起きやすい:

| 業態 | 自社の事業市場（NG） | 顧客市場（OK） |
|---|---|---|
| ビルメンテナンス業 | ✗ ビルメンテナンス市場 | ✓ 不動産業 / オフィスビル賃貸市場 |
| 半導体製造装置メーカー | ✗ 半導体製造装置市場 | ✓ 半導体メーカーの設備投資市場 |
| HR Tech SaaS | ✗ HR Tech 市場 | ✓ 企業の人事関連支出市場 |
| 建設コンサルタント | ✗ 建設コンサル市場 | ✓ 建設投資市場 / インフラ整備需要 |
| 物流アウトソース | ✗ 物流アウトソース市場 | ✓ 顧客業種（製造業 / 小売業 等）の出荷量・売上 |

判別基準: **「自社が消えても伸びる市場」が顧客市場**。自社の事業市場（受託金額の総和）は自社が属する受託業界の指標であり、論点 5 の対象ではない。

#### ルール3: 業界バリューチェーン（差別化）は自社が属する業界

論点 3（差別化＝バリューチェーン上のポジショニング）で扱う業界は、**自社が属する業界**であり、顧客の業界ではない。論点 5（顧客市場）と業界が異なることに注意:
- ビルメンテ業の場合: 論点 3 = ビルメンテ業界バリューチェーン、論点 5 = 不動産業界市場
- 同じスライドで両方の業界を扱うことはない

### 整合性チェック表テンプレ（Step 3 で必須）

Step 3 のユーザー承認 Markdown に**必ず**以下の表を含める（再掲、上記 Step 3 セクション参照）:

| # | 確認項目 | 期待される値 | 実態 | OK? |
|---|---|---|---|---|
| 1 | 全論点の主語 = 論点 1 の `segment_name` | <セグメント名> | <各論点の主語> | ✅/❌ |
| 2 | 論点 4 顧客の業種 = 論点 5 顧客市場の業種 | <顧客業種> | <論点 5 で扱う市場の業種> | ✅/❌ |
| 3 | 論点 3 業界 = 自社の業界 | <自社業界> | <論点 3 で扱う業界> | ✅/❌ |

❌ がある場合は JSON を組み立てる前に修正する。

### アンチパターン

- ❌ 論点を独立に処理して、各論点の主語が揺れる（全社/グループ/他事業）
- ❌ B2B 受託業で論点 5（顧客市場）に自社の事業市場を入れる
- ❌ Step 3 で個別論点の内容だけ確認して、論点間整合性を確認しない
- ❌ visual-quality-reviewer の役割を「論点間整合性検査」と誤解する（visual-quality-reviewer は単一スライドの視覚品質のみ）

### 失敗事例（参考）

**Phase 3 (二幸産業 施設運営事業)**: 論点 5 に「ビルメンテナンス市場」（自社の事業市場、4-5 兆円）を入れた。設計上は「顧客 = ビルオーナー法人 = 不動産業」が属する市場（不動産業 / オフィスビル賃貸市場）を扱うべきだった。論点 4（顧客 = 東京建物・不動産業）と論理的整合性が取れない状態に。Phase 4 後追補で本ルールを導入し、構造的に防げる体制に。

---

## オーケストレーター契約（merge_order.json は親が組む）

<!-- source: skills/_common/references/orchestrator_contract.md (manual sync until D2) -->

本スキルは個別 PPTX のみ返却し、`merge_order.json` の生成は親 `company-deepdive-agent` の責務。
親が組む `merge_order.json.entries[]` の本セグメント部分は以下のような形式になる:

```json
{
  "slide_number": 11,
  "file_name": "slide_11_business_overview.pptx",
  "skill_name": "business-overview-pptx",
  "data_file": "data_11_business_overview.json",
  "category": "content"
}
```

- `category` は本セグメントの 5 枚すべて `content`（中扉は親側で別途追加）
- `data_file` / `file_name` のパスは親から見て `segments/<segment_slug>/` 相対 or 絶対パスで解決可能とする

詳細は `skills/_common/references/orchestrator_contract.md` を参照。

---

## 単独起動 vs 内部呼び出し

### 単独起動の場合

```
ユーザー: 「第一交通産業のタクシー事業を深掘りして」
→ Step 0 (AskUserQuestion で対話) → Step 1..6
```

`parent_run_id` を `YYYY-MM-DD_<parent_company_slug>` で自動生成し、
`{{WORK_DIR}}/company-deepdive-agent/<parent_run_id>/segments/<segment_slug>/` 構造を
business-deepdive-agent 自身が作る。Phase 4（company-deepdive-agent 実装後）はそのまま流用可。

### 内部呼び出しの場合（推奨）

```
company-deepdive-agent → Task tool で本スキルを起動
→ parent_run_id / segment_slug / global_slide_offset を JSON で受領 → Step 1..6
```

並列起動可（同じ会社の複数セグメントを並列で深掘り）。

---

## 注意事項

- **対象は単一の事業セグメント**: 複数事業を扱う場合は親が本スキルを複数回呼ぶ
- **顧客 = 対象事業の顧客**: 本スキルが扱う「顧客」は会社全体の顧客ではなく、対象セグメントの顧客
- **顧客の成長 = 顧客市場の成長**: market-environment-pptx を顧客側市場（例: タクシー事業なら「観光・交通需要」市場）に向けて使用
- **「〜すべき」表現禁止**: 上記 main_message ルールに従う
- **fact-check / visual-review は親が実施**: 本スキル単体では fact-check-reviewer / visual-quality-reviewer を呼ばない

---

## 依存スキル一覧

### コア（必須）

| スキル名 | 役割 |
|---|---|
| `business-overview-pptx` | 事業概要 |
| `business-model-pptx` | ビジネスモデル |
| `value-chain-matrix-pptx` | 差別化（バリューチェーン上のポジション） |
| `customer-profile-pptx` | 顧客プロファイル |
| `market-environment-pptx` | 顧客市場の成長性 |

### 任意

`section-divider-pptx`（拡張版で各セグメント冒頭に中扉を入れる場合、親 `company-deepdive-agent` が制御）

---

## アセット

| ファイル | 内容 |
|---|---|
| `prompts/step1_research_template.md` | セグメント単位の 5 論点別 Web 検索クエリテンプレート |
| `references/deck_skeleton.json` | 5 枚の標準スライド構造定義（merge_order.json の素材） |
