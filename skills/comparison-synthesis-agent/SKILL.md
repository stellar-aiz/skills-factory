---
name: comparison-synthesis-agent
description: >
  複数社（マーケット内の主要プレイヤー）の Deepdive デッキを横並びで比較し、
  全社共通の検証論点を統合する PowerPoint デッキを生成するオーケストレータースキル。
  market-overview-agent → company-deepdive-agent (社数分) で各社の深掘りデッキが
  揃った後の最終統合レイヤーを担う。

  入力は各社の `Deepdive_<社名>.pptx` 群（または対応する work dir のセグメント別データ）。
  出力は競合横並び比較サマリー + 検証論点統合の `Comparison_<業界>.pptx` デッキ。
  全社共通で取れなかった項目はデータアベイラビリティで明示的に集約する。

  本スキルは competitor-summary-pptx の SKILL.md で参照されていた
  「`competitor-analysis-pptx` オーケストレーター」の正式実装に相当する
  （v0.3 で planned-but-not-built だった参照を解消する）。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「○○社、○○社、○○社の戦略を比較」「複数社の深掘りを統合」「競合横並びサマリー」という言葉が出た場合
  - market-overview-agent → company-deepdive-agent N 社分 の流れの後段
  - BDD・経営企画・中計策定で複数社の戦略横並び比較が必要な場合
  - 各社の Deepdive デッキを既に持っており、それらを統合したい場合
---

# 競合横並び比較・検証論点統合オーケストレーター（skeleton、実装保留中）

> ⚠️ **本スキルは skeleton のまま実装保留中です（2026-04-29 判断）**
>
> ISSUE-004 の Phase 4 完了時点で、本スキルが提供する「全社横並び比較」の機能の大半が
> 既存スキル（`market-overview-agent` + `financial-benchmark-pptx` + `competitor-summary-pptx` +
> `market-share-pptx` + `positioning-map-pptx` + `market-kbf-pptx`）の組み合わせで
> 80% 以上カバーできていることを確認した。Phase 5 として独自実装する独自価値は
> 「複数社 deepdive デッキの論点間統合」「全社共通の検証論点（open_questions）の業界横断集約」の
> 2 点に絞られるが、この需要が顕在化した時点で実装着手する方針に変更。
>
> **トリガー条件が出ても自動起動はせず、ユーザー指示があった場合に実装着手から始める。**
>
> 復活させる場合の参照先:
> - `skills/_common/prompts/cross_topic_consistency_check.md`（適用予定スキルの 1 つ）
> - `skills/business-deepdive-agent/SKILL.md`（同等の Step 構成・規約パターン）
> - 関連計画: `~/.claude/plans/tidy-soaring-elephant.md`

ISSUE-004（v0.3）における新規オーケストレーター。市場 → 各社深掘りの最終統合層として skeleton 設計済。

## 設計原則

- **横並びの粒度は会社レベルと事業セグメントレベルの 2 層**
  - 会社レベル: 規模・収益性・株主構造の比較
  - 事業レベル: 同一カテゴリのセグメントを並べて比較（例: 全社のタクシー事業）
- **検証論点を全社統合で集約**: 各社 Deepdive の「open_questions」を集めて整理
- **データアベイラビリティを全社統合で表示**: どの項目が全社で取れて、どこが取れていないかを 1 枚で示す

---

## 入力

| 入力 | 形式 |
|---|---|
| 各社 Deepdive デッキ | `outputs/Deepdive_<社名>_<date>.pptx`（複数）|
| 各社 work dir | `work/company-deepdive-agent/<run_id>/`（segment_summary.json 等を読む）|
| 対象市場名 | 文字列（例: 国内タクシー市場）|

入力指定方法は以下のいずれか:
- A: ディレクトリ指定（`work/company-deepdive-agent/` 配下を自動スキャン）
- B: 明示リスト（`scope.json` に `target_companies: [...]` を渡す）

---

## 出力スライド構成（標準: 8〜12 枚）

```
01 エグゼクティブサマリー (executive-summary-pptx)
02 目次 (table-of-contents-pptx)
03 (中扉) 全社の俯瞰 (section-divider-pptx)
04 全社サマリーテーブル (competitor-summary-pptx)
05 規模・収益性ベンチマーク (financial-benchmark-pptx)
06 ポジショニングマップ (positioning-map-pptx)
07 (中扉) 事業セグメント別比較 (section-divider-pptx)
08 同一カテゴリのセグメント横並び（例: 各社のタクシー事業）(competitor-summary-pptx)
09 同上（例: 各社のバス事業）
   ...（事業カテゴリ数だけ繰り返し）
N (中扉) 全社共通の検証論点 (section-divider-pptx)
N+1 全社共通の検証論点 (issue-risk-list-pptx)
末 全社統合のデータアベイラビリティ (data-availability-pptx)
```

---

## Step 構造

### Step 0: 比較対象の確定

<!-- source: skills/_common/prompts/step0_brand_clarification.md (manual sync until D2) -->

`AskUserQuestion` で以下を確定（**ブランドは先頭で確定**。詳細仕様は `skills/_common/prompts/step0_brand_clarification.md` を正本とする）:

| # | 質問 | 選択肢 |
|---|---|---|
| 1 | **ブランド**（出力 PPTX フォーマット） | `_discover_brands()` の戻り値から動的取得（既定 `stellar_aiz`）。各社 deepdive run_id の `scope.json.brand` が割れている場合は本ステップで再確定する |
| 2 | 比較対象の会社（複数選択）| `work/company-deepdive-agent/` 配下の run_id 一覧から複数選択 |
| 3 | 対象市場名 | テキスト入力（例: 国内タクシー市場）|
| 4 | 事業セグメント横並び比較を含めるか | A. 含める（推奨）/ B. 会社レベルのみ |
| 5 | 検証論点の集約方法 | A. 各社の open_questions を集約（推奨）/ B. 全社共通の論点のみ抽出 |

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
# 推奨: 各社 deepdive run_id の scope.json.brand を読んで「全社一致なら既定値、割れていたら再確定」をデフォルト提示。
# AskUserQuestion(question="...", header="ブランド", options=options, multiSelect=False)
```

確定したブランドは本 agent の比較デッキ全体（fill スキル群への `--brand` 引数）で使う。各社 deepdive デッキを再 merge する場合は、既存の各社デッキは元の brand のまま、新規追加スライド群は確定したブランドで生成する（混在が嫌な場合はユーザーに警告して全社再生成を提案）。

### Step 1: 各社 Deepdive データの読み込み

各社の `work/company-deepdive-agent/<run_id>/` から以下を読み込み:

- `scope.json`（会社名・上場区分・セグメント一覧）
- `data_NN_*.json`（会社レベル 5 論点 + 事業セグメント 5 論点）
- `segment_summary.json`（business-deepdive-agent が出力した検証論点）
- `data_NN_data_availability.json`（データ取得状況）

### Step 2: 比較データ構造の構築

```
全社共通スコープ:
  - 会社レベル比較データ（売上・収益性・株主構造）
  - セグメントカテゴリ別グルーピング（同名 / 類似名のセグメントを束ねる）

例（タクシー業界）:
  - タクシー事業 = [日本交通・第一交通・国際自動車・大和・東京無線] (5 社)
  - バス事業 = [第一交通] (1 社、横並びにならないので除外)
  - 不動産事業 = [第一交通・大和] (2 社)
  - 介護・福祉事業 = [第一交通] (1 社、除外)
```

横並び比較は **2 社以上** の同カテゴリ事業がある場合のみ実施。

### Step 3: 比較スライド生成

#### Step 3 開始前: brand fallback バッファ初期化（必須）

Step 0 で確定した `scope_brand` を使い、未対応 fill 検出用の warning バッファを初期化する。各 fill 起動前に `resolve_fill_brand_with_warning()` を呼び、未対応スキルでは `stellar_aiz` に fallback + warning を buffer に蓄積する（`skills/_common/lib/orchestrator_helpers.py` 参照）。

```python
import os, sys, subprocess
sys.path.insert(0, os.path.join("{{SKILL_DIR}}", "..", "_common", "lib"))
from orchestrator_helpers import (
    resolve_fill_brand_with_warning,
    append_brand_warnings_to_merge_file,
)

scope_brand = "stellar_aiz"  # Step 0 で確定した値（会話メモリから）
brand_warnings: list = []  # Step 5 (merge) 後に merge_warnings.json へ append する

# 各 fill 起動例:
# skill_dir = os.path.join("{{SKILL_DIR}}", "competitor-summary-pptx")
# fill_brand = resolve_fill_brand_with_warning(skill_dir, scope_brand, brand_warnings)
# subprocess.run(["python", os.path.join(skill_dir, "scripts", "fill_competitor_summary.py"),
#                 "--brand", fill_brand, "--data", "...", "--output", "..."], check=True)
```

各論点に応じて既存 PPTX スキルを呼び出し（すべての起動で `--brand <fill_brand>` を渡す）:

| スライド | スキル |
|---|---|
| 全社サマリー | `competitor-summary-pptx`（横並びテーブル） |
| 財務ベンチマーク | `financial-benchmark-pptx`（KPI 横並び） |
| ポジショニング | `positioning-map-pptx`（2 軸プロット）|
| 事業セグメント横並び | `competitor-summary-pptx` の事業版 |
| 検証論点 | `issue-risk-list-pptx`（カテゴリ別整理） |
| データアベイラビリティ | `data-availability-pptx` |

### Step 4: 全社共通の検証論点を整理

各社の `segment_summary.json.open_questions` を集約し、以下のカテゴリに整理:

- 公開情報で確定できない打ち手仮説（経営インタビューで確認）
- 業界横断の市場動向（業界ヒアリング推奨）
- 各社固有の戦略意図（IR・統合報告書・マネジメント発言の追加調査）
- 財務詳細（セグメント別営業利益率等、決算説明会で確認）

### Step 5: merge_order.json + merge-pptxv2 で結合

会社レベル比較 + 事業セグメント別比較 + 検証論点 + データアベイラビリティを結合。

`outputs/Comparison_<業界>_<date>.pptx` に保存。

#### merge 完了後: brand_warnings を merge_warnings.json に追記（必須）

merge-pptxv2 は `merge_warnings.json` を `"w"` モードで上書きするため、Step 3 中に蓄積した `brand_warnings` は merge 完了後にここで追記する。

```python
append_brand_warnings_to_merge_file(
    "outputs/Comparison_<業界>_<date>/merge_warnings.json", brand_warnings,
)
# brand_warnings が空なら no-op（既存ファイルは触らない）。
# Step 7（ユーザーへ提示）でも warning 件数 + 内訳を必ず提示する。
```

### Step 6: visual-quality-reviewer + 自動修正ループ

最大 2 ラウンド。

### Step 7: ユーザーへ提示

- `outputs/Comparison_<業界>_<date>.pptx`（比較統合デッキ）
- `outputs/Comparison_<業界>_<date>/synthesis.md`（検証論点・データギャップの統合 Markdown）

---

## `_common/` 再利用

```
<!-- source: skills/_common/prompts/main_message_principles.md (manual sync until D2) -->
<!-- source: skills/_common/prompts/step_final_visual_review_loop.md (manual sync until D2) -->
<!-- source: skills/_common/references/orchestrator_contract.md (manual sync until D2) -->
```

Step 2.5（fact-check）は本スキルでは実施しない（各社 `company-deepdive-agent` 側で実施済を前提）。

---

## scope.json（本オーケストレータ内部）

`{{WORK_DIR}}/comparison-synthesis-agent/<run_id>/scope.json`:

```json
{
  "market_name": "国内タクシー市場",
  "target_companies": [
    {"name": "第一交通産業株式会社", "deepdive_run_id": "2026-04-29_daiichikoutsu"},
    {"name": "日本交通株式会社", "deepdive_run_id": "2026-04-29_nihonkotsu"},
    {"name": "国際自動車株式会社", "deepdive_run_id": "2026-04-29_kokusaijidosha"},
    {"name": "大和自動車交通株式会社", "deepdive_run_id": "2026-04-29_daiwa"},
    {"name": "東京無線協同組合", "deepdive_run_id": "2026-04-29_tokyomusen"}
  ],
  "include_segment_comparison": true,
  "open_questions_aggregation": "all",
  "run_id": "2026-04-29_taxi_comparison",
  "started_at": "2026-04-29T15:00:00+09:00"
}
```

---

## 依存スキル一覧

### コア（必須）

| スキル名 | 役割 |
|---|---|
| `executive-summary-pptx` | デッキ冒頭サマリー |
| `table-of-contents-pptx` | 目次 |
| `section-divider-pptx` | 中扉 |
| `competitor-summary-pptx` | 全社・事業セグメント横並びテーブル |
| `financial-benchmark-pptx` | 財務 KPI ベンチマーク |
| `positioning-map-pptx` | 2 軸ポジショニング |
| `issue-risk-list-pptx` | 全社共通の検証論点 |
| `data-availability-pptx` | 全社統合のデータ取得状況 |
| `merge-pptxv2` | 結合 |

### 品質レビュー

| スキル名 | 呼び出し位置 |
|---|---|
| `visual-quality-reviewer` | Step 6 |

---

## `competitor-analysis-pptx` 参照の解消（並行整理）

`competitor-summary-pptx/SKILL.md` には現状以下の記述がある:

> 本スキルは単体で 1 枚のサマリースライドを生成するスキルであり、
> Web 検索での競合特定や複数スライドのデッキ生成は行わない。
> それらが必要な場合は `competitor-analysis-pptx` スキル（オーケストレーター）を使うこと。

この `competitor-analysis-pptx` は実装されていない planned-but-not-built な参照。
**本スキル `comparison-synthesis-agent` の実装と同時に**、`competitor-summary-pptx/SKILL.md` の
該当箇所を `comparison-synthesis-agent` への参照に書き換える。

---

## 注意事項

- **本スキルは「比較」と「集約」のみ**: 個別社の深掘りは行わない。新規社を含めたい場合は先に `company-deepdive-agent` を実行
- **検証論点を勝手に作らない**: 入力データ（各社の open_questions）に基づく集約のみ
- **データアベイラビリティの集約原則**: 全社で取れた項目は ✓、一部でしか取れていない項目は △、全社で取れていない項目は ✗

---

## アセット（Phase 5 で作成）

| ファイル | 内容 |
|---|---|
| `references/deck_skeleton.json` | 比較スライド配列定義 |
