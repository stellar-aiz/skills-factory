# ハーネスレバー利用規約 — 各 SKILL.md 共通

> **このファイルは `skills/_common/references/harness_levers.md` です。**
> オーケストレータースキル（market-overview-agent / company-deepdive-agent / business-deepdive-agent 等）の SKILL.md 冒頭から、`<!-- source: skills/_common/references/harness_levers.md (manual sync until D2) -->` コメント付きで**手動コピペ**するか、参照リンクを張ってください。
> このファイルを変更したら `grep -rn "source: skills/_common/references/harness_levers.md" skills/*/SKILL.md` で被参照スキルを全て検出し、整合を取り直すこと（ISSUE-001 D2 で自動化検討中）。

skills_factory のオーケストレーターが Claude Code ハーネス（`.claude/settings.json` の hooks / Agent ツール / TaskCreate / AskUserQuestion）をどう使うかの **横断規約**。SKILL.md は宣言的手順書として残しつつ、ハーネス側で LLM の遵守を強制するための前提を定義する。

設計背景: `docs/harness_check/dependency_map.md` および `docs/harness_check/lever_mapping.md` で 12箇条 × 3 レバーのマッピングを整理済み。本ファイルはその「打ち手側」を SKILL.md author 視点でまとめた実装規約。

---

## レバー① hooks 連携の前提

`.claude/settings.json` で配線された hooks（`tools/hooks/*.py`）が前提条件を機械的に検証する。SKILL.md author は以下を守る:

### `check_merge_order_exists.py`(PreToolUse, Bash)

**強制内容**: `merge_pptx_v2.py` 起動時に `--merge-order <path>` 引数が必須、かつ指定 path が実在すること。

**SKILL.md 側の責務**:
- `merge-pptxv2` を呼ぶ Step で **必ず先に** `merge_order.json` を生成する（スキーマは `orchestrator_contract.md` 参照）
- merge コマンド例には `--merge-order` を必ず含める（`--merge-order` を忘れると hook で exit 2 ブロック）

### `validate_pptx_after_fill.py`(PostToolUse, Bash)

**強制内容**: `fill_*.py` または `merge_pptx_v2.py` で生成した PPTX を `tools/validate_pptx.py` で自動検証。OOXML integrity / rels / chart chains のいずれかに不整合があれば exit 2 でブロック。

**SKILL.md 側の責務**:
- 生成 PPTX の `--output` 引数を明示する（hook が path を抽出するため）
- validate 失敗時の対応手順（テンプレート再生成、データ修正等）を再生成ループに組み込む

### `load_session_context.py`(SessionStart)

**強制内容**: セッション起動時に未解決 ISSUES と直近 plan のタイトル一覧を context に注入。

**SKILL.md 側の責務**: 特になし（受動的な恩恵）。ただし ISSUES.md のフォーマット（`## ISSUE-NNN: <title>` + `**Status**: <status>`）を維持すること。

### `check_task_progression.py`(B-2-d、未実装)

**強制内容（実装後）**: Step N+1 のツール呼び出し前に Step N の TaskCreate タスクが `completed` であることを assert。違反は exit 2 でブロック。

**SKILL.md 側の責務**: `step_state_tracking.md` の規約（TaskCreate / TaskUpdate / `task_state.json` 更新）を厳守。

### hook がブロックされたときの SKILL.md author の責任

hook の stderr メッセージは LLM の context に注入される。SKILL.md には「hook がブロックを出した場合、メッセージに従って前提条件を満たしてから再実行する」旨を明示する。

---

## レバー② subagent 呼び出し規約

Web 検索 / fact-check / visual-review 等の独立フェーズは subagent (`Agent` ツール経由) に切り出す。親 orchestrator から見た subagent の使い方規約。

### subagent の責務境界

- **subagent**: Web 検索の生 HTML/MD、PNG レビュー画像 etc. の **大量の生データ** を自分の context に閉じ込め、要約済み JSON だけを親に返す
- **親**: 受け取った JSON を `data_NN_*.json` 等に保存し、後続スキルに渡す

### Agent ツール起動時のデータ受け渡し規約

```
Agent(
  subagent_type="<subagent-name>",
  description="<3-5 word summary>",
  prompt=<JSON シリアライズ済みの input。subagent SKILL.md で定義された schema に従う>
)
```

**重要**:
- `prompt` には自由文ではなく、subagent SKILL.md で定義された schema に沿った構造化入力を渡す
- subagent は **return value として要約済み JSON を返却**（親 context を生データで圧迫しない）
- 親はその JSON を直接 `data_NN_*.json` に書き出すか、追加加工してから書き出す

### subagent SKILL.md author の責務

- 入力 schema・出力 schema を冒頭に明示
- 「親に返却する JSON 以外の生データは subagent ディレクトリにのみ残す」を明記
- AskUserQuestion を呼ばない（人間との対話は親の責任）

---

## レバー③ TaskCreate / AskUserQuestion 必須地点

### TaskCreate / TaskUpdate 必須地点

**全 Step**: 各 Step の冒頭で TaskCreate、終了で TaskUpdate(completed)。例外なし。

具体的なテンプレートと task_state.json スキーマは **[`step_state_tracking.md`](../prompts/step_state_tracking.md)** 参照。

### AskUserQuestion 必須地点

以下の局面では **必ず** `AskUserQuestion` ツールを使う（自由対話のテキスト出力では不可）:

| 局面 | 理由 |
|---|---|
| Step 0: スコープ確認 | 後続 Step の真実源。曖昧さを最初に潰す |
| 異種事業モデルの境界確認（Step 0.5） | シェア表での誤解を防ぐ |
| ファクトチェック後の修正方針選択 | 数値の信頼性に関わる |
| 中間 Markdown 承認（Step 3 等） | デッキ構成の最終確定 |
| visual-review 自動修正で残存 issue があった場合 | 手動修正 / 許容の判断 |

具体的な質問テンプレートは `step0_scope_clarification.md` 参照。

### 自由対話で良い局面

- 進捗報告（「Step 3 完了しました、結果は…」）
- 補足説明・解釈
- ユーザーから明示的に質問が来た場合の回答

---

## レバー③ 補足: SKILL.md の `description` / `triggers` 精緻化規約（B-5）

LLM がどのスキルを発火させるかは SKILL.md frontmatter の `description` と本文中の triggers リストで決まる。発火条件が広すぎると関係ない依頼でもスキルがロードされ、context を圧迫する。

### `description` の書き方

| 観点 | 規約 |
|---|---|
| 長さ | **5〜15 行**(LLM が triggers 判定で読む量を抑える) |
| 1 行目 | 「<何> をする <何> スキル」の 1 文（型と機能の宣言） |
| 2-5 行目 | 主要な入出力・前提（テンプレート / 依存スキル / 出力ファイル種別） |
| 6 行目以降 | 必ず `以下のいずれかのトリガーで必ずこのスキルを使うこと:` を含む trigger リスト |

### triggers リストの書き方

| ✅ Good | ❌ Bad |
|---|---|
| 「市場規模を積み上げ棒グラフで」 | 「業界の話」 |
| 「マーケットの中で○○社を分析して」 | 「会社のことを調べて」 |
| 「BDD の財務モデルを更新して」 | 「データを更新して」 |

具体性が高い trigger ほど発火が正確になり、無関係な発火を防げる。

### 発火しない条件の明示（推奨）

トリガー条件の下に **「以下の場合は別スキルに委譲する」** リストを書くと、LLM の判断が速くなる:

```yaml
description: >
  ...
  以下のいずれかのトリガーで必ずこのスキルを使うこと:
  - 「市場規模・成長性を積み上げ棒グラフで」
  - 「市場環境分析」「Market Environment」「マーケット分析」
  以下の場合は別スキルを使う:
  - 「市場シェア」 → market-share-pptx
  - 「市場全体の調査」 → market-overview-agent
```

### orchestrator スキルの description 特例

orchestrator は「他スキルを呼ぶ orchestrator スキル」と冒頭で明示。これで LLM が「単一スライドのつもりが orchestrator が起動する」誤発火を抑える。

---

## アンチパターン

- ❌ hook の stderr を context で受け取りつつ無視する（修正なしに同じツールを再実行 → 同じ block を繰り返す）
- ❌ subagent に「ユーザーに XX を聞いて」と頼む（subagent は AskUserQuestion を持たない設計）
- ❌ subagent から大量の生データ（HTML / 画像 base64 等）を return value で返す（親 context 圧迫の温床）
- ❌ Step を `completed` にせず次の Step に進む（hooks の進捗チェックでブロック、または `task_state.json` の整合性が崩れる）
- ❌ `description` を 30 行以上書く（context 圧迫、triggers が埋もれる）
- ❌ trigger に「相談」「議論」など曖昧語を入れる（誤発火頻発）
- ❌ 同じ trigger phrase を複数スキルが宣言する（LLM の発火判定が不安定）

---

## 関連ドキュメント

- `tools/hooks/README.md` — hooks 入出力 contract と実装規約
- `skills/_common/prompts/step_state_tracking.md` — TaskCreate / TaskUpdate / task_state.json の標準パターン
- `skills/_common/prompts/step0_scope_clarification.md` — Step 0 のスコープ確認
- `skills/_common/prompts/step_final_visual_review_loop.md` — Step 最終のビジュアル品質レビュー
- `skills/_common/references/orchestrator_contract.md` — merge_order.json / merge_warnings.json / regeneration_hint
- `docs/harness_check/dependency_map.md` — 12箇条 × 3 層 ヒートマップ
- `docs/harness_check/lever_mapping.md` — 12箇条 × 3 レバー 打ち手マトリクス
- `docs/harness_check/settings_design.md` — `.claude/settings.json` 設計メモ
- `~/.claude/plans/md-llm-melodic-twilight.md` — 全体計画
