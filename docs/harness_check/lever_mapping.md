# 12箇条 × ハーネスレバー マッピング（Phase A.5-2）

**目的**: 12箇条の各原則について、skills_factory で打つべき具体的な「ハーネス設定」を、3レバーに分解して整理する。Phase B のタスク一覧の素材になる。

**読み方**:
- 🟢 = 既に対応済（現状で十分）
- 🟡 = 部分対応（Phase B で改善）
- 🔴 = 未対応（Phase B で新規対応）

**3 レバー**:
- **①** hooks + settings.json + description 精緻化
- **②** subagent 分割
- **③** AskUserQuestion / TaskCreate 構造化運用

各セルには「**ルール（SKILL.md側）** / **仕組み（ハーネス側）**」の 2 段で記述。

---

## サマリーマトリクス（Status のみ）

| # | 原則 | レバー① hooks/settings | レバー② subagent | レバー③ AUQ/Task |
|---|---|---|---|---|
| 1 | 自然言語→ツール呼び出し | 🟡 | 🔴 | 🟢 |
| 2 | プロンプトをコード管理 | 🟢 | 🟢 | N/A |
| 3 | コンテキスト制御 | 🔴 | 🔴 | 🟡 |
| 4 | ツール=構造化出力 | 🟡 | 🟢 | N/A |
| 5 | 状態統合 | 🟡 | 🟢 | 🔴 |
| 6 | 開始・停止・再開 | 🟡 | 🟢 | 🔴 |
| 7 | 人間とのやり取り | N/A | 🟢 | 🟢 |
| 8 | 制御フロー | 🔴 | 🟡 | 🔴 |
| 9 | エラー圧縮 | 🟡 | 🟢 | N/A |
| 10 | 責務サイズ（≤10ステップ）| 🟡 | 🟡 | 🟡 |
| 11 | どこからでも起動 | 🟡 | N/A | N/A |
| 12 | ステートレス・リデューサ | 🟡 | 🟢 | 🔴 |

**集計**: 🟢=15 / 🟡=14 / 🔴=8 / N/A=10  
**Phase B 完了目標**: 🔴 を 0 に、🟡 を半数以上 🟢 に。#11 だけは本フェーズで触らない（レバー④ 範囲外）。

---

## 詳細マッピング

### #1 自然言語→ツール呼び出し変換

> 自然言語のまま処理せず、構造化された操作（API/関数）に落とす

| レバー | Status | ルール（SKILL.md側） | 仕組み（ハーネス側） |
|---|---|---|---|
| ① | 🟡 | 各 Step の出力 JSON のスキーマを SKILL.md に明示。fill_*.py の hard-fail 制約を SKILL.md でも説明 | PreToolUse hook で JSON validate（jsonschema 利用、現状は fill_*.py 内部のみ）。fill_*.py 直前で外側からも検証 |
| ② | 🔴 | research-subagent は「自然文クエリ → 構造化 JSON 検索結果」を返却すると subagent SKILL.md に明記 | Agent ツール経由で呼び出し、戻り値は要約済み JSON のみ（生 HTML/MD は親 context に入れない） |
| ③ | 🟢 | Step 0 / 2.5 / 3 の対話地点で AskUserQuestion を使う旨を SKILL.md に明示 | AskUserQuestion ツール（既存）。自由対話を構造化選択肢に強制変換 |

### #2 プロンプトを自分で管理する

> プロンプトをブラックボックスにせず、コードとして管理

| レバー | Status | ルール | 仕組み |
|---|---|---|---|
| ① | 🟢 | SKILL.md / `_common/prompts/*.md` / settings.json を全て git 管理 | build_skill.py の {{VAR}} / @if/@endif 機構で profile 別ビルド（既存） |
| ② | 🟢 | subagent 定義ファイル `agents/<name>.md` も git 管理 | 既存 build パイプラインで配布 |
| ③ | N/A | — | — |

**所見**: skills_factory は #2 についてはほぼ完璧。ISSUE-001 (`@import` 自動化) のみ将来課題。

### #3 コンテキストウィンドウを自分で制御する

> モデルに渡す情報を取捨選択・設計

| レバー | Status | ルール | 仕組み |
|---|---|---|---|
| ① | 🔴 | SKILL.md description を絞って発火条件を狭くする規約。triggers の書き方規約（`harness_levers.md` で定義） | description の精緻化 + SessionStart hook で「現在 run_id の進捗だけ」を context に注入。793 行の SKILL.md がフルロードされない設計 |
| ② | 🔴 | research-subagent / factcheck-subagent / visual-quality-reviewer 等で web 検索 25-40 件・PNG レビュー画像を subagent context に閉じ込める。subagent SKILL.md に「親に返却するのは要約 JSON のみ」と明記 | Agent ツール（subagent_type 指定で context 隔離）。親 context に大量データを積まない |
| ③ | 🟡 | 各 Step を TaskCreate で起こし、進捗を会話履歴ではなく TaskList で持つ | TaskList が進捗の真実源。会話履歴に依存しないので、長い orchestrator でも context が圧迫されない |

### #4 ツールは構造化出力として扱う

> ツール呼び出しを特別視せず、LLMが返す構造化出力として設計

| レバー | Status | ルール | 仕組み |
|---|---|---|---|
| ① | 🟡 | 入出力スキーマを SKILL.md に明示。fill_*.py の引数仕様を CLI 仕様として記載 | PreToolUse / PostToolUse hook で I/O スキーマを assert（現状は fill_*.py 内部のみ） |
| ② | 🟢 | subagent は単一目的に集中（research / factcheck / visual review）と subagent SKILL.md に明記 | agents/<name>.md でスコープを宣言 |
| ③ | N/A | — | — |

### #5 実行状態と業務状態を統合する

> 進行・待機・再試行・業務状態を、履歴・イベント・コンテキストから推論できる形に

| レバー | Status | ルール | 仕組み |
|---|---|---|---|
| ① | 🟡 | scope.json + run_id 規約（既存 `orchestrator_contract.md`） | SessionStart hook で `{{WORK_DIR}}/<run_id>/` 配下の状態（scope.json / data_NN_*.json / merge_order.json の存在）を context に提示 |
| ② | 🟢 | subagent は自分の出力ファイルを subagent ディレクトリに残す | 既存 work/ ディレクトリ規約 |
| ③ | 🔴 | 各 Step を TaskCreate で起こし、Step 完了時に TaskUpdate(completed) を必須化。各 Step の `subject` 命名規約を `step_state_tracking.md` で統一 | TaskList が state machine の真実源。「いま何 Step まで終わってるか」がツールで取得できる |

### #6 開始・停止・再開をシンプルに扱える

> 長時間タスクを前提に、ユーザー・アプリ・外部イベントから開始/停止/再開可能に

| レバー | Status | ルール | 仕組み |
|---|---|---|---|
| ① | 🟡 | SKILL.md に「再開時の参照ファイル」を明記（scope.json / merge_order.json / fact_check_report.json） | SessionStart hook で run_id 配下の中間状態を読み、「Step 5 まで完了、Step 6 から再開可能」を context に表示 |
| ② | 🟢 | subagent は単発実行で完結。冪等性をsubagent SKILL.md で保証 | 同じ JSON 入力で再呼び出し可能（subagent は state を持たない） |
| ③ | 🔴 | TaskList から再開地点が一目でわかる規約 | TaskList で resume 地点を機械的に判定。LLM が会話履歴を辿る必要なし |

### #7 人間とのやり取りもツールとして扱う

> 承認・確認・追加質問を構造化された処理ステップとして組み込む

| レバー | Status | ルール | 仕組み |
|---|---|---|---|
| ① | N/A | — | — |
| ② | 🟢 | subagent は AskUserQuestion を使わない（親が user との対話を担当） | subagent SKILL.md でその旨明記 |
| ③ | 🟢 | Step 0 / 2.5 / 3 の対話ポイントを SKILL.md に明示 | AskUserQuestion ツール（既存）。Markdown承認も Step として明示 |

**所見**: #7 は既に良い。skills_factory の orchestrator は構造化された対話を組み込んでいる。

### #8 制御フローは自分で持つ

> 分岐・ループ・リトライ・承認・中断・再開は、LLM任せにせずアプリケーション側で制御

| レバー | Status | ルール | 仕組み |
|---|---|---|---|
| ① | 🔴 | SKILL.md に「Step N の前提条件」（必須ファイル / 必須 TaskUpdate）を明示 | PreToolUse hook で前提条件を assert。例: merge-pptxv2 起動前に `merge_order.json` 存在チェック、Step N+1 のツール起動前に Step N の TaskUpdate(completed) チェック |
| ② | 🟡 | subagent 呼び出し順序を親 SKILL.md に明示 | Agent ツールで明示的に順次起動。subagent 内のループ・分岐は subagent 側の責務に閉じる |
| ③ | 🔴 | Step 完了時に TaskUpdate(completed) を必須化 | hook で TaskUpdate されない遷移を検知して警告 |

### #9 エラーはコンテキストに圧縮して渡す

> エラーをLLMが判断できる形でコンテキストに戻し、次の試行で自己修復

| レバー | Status | ルール | 仕組み |
|---|---|---|---|
| ① | 🟡 | SKILL.md に typical error → recovery action 対応表（main_message 65字超過 → 短縮原則 4 つ等） | PostToolUse hook でエラー時のヒント（regeneration_hint）を構造化フォーマットで context に返す |
| ② | 🟢 | subagent 内のエラーは要約して親に返す。詳細ログは subagent ディレクトリに残す | Agent ツール（returns single message）。生エラーログは親 context に積まない |
| ③ | N/A | — | — |

### #10 小さく責務の明確なエージェントに分ける

> 巨大な万能エージェントを避け、3〜10ステップ程度で完結する小さな責務に

| レバー | Status | ルール | 仕組み |
|---|---|---|---|
| ① | 🟡 | orchestrator の Step 数上限を `harness_levers.md` で規約化（10 Step 以内推奨） | description / triggers の精緻化で発火条件を絞る。将来は linter で行数 / Step 数を自動チェック（本フェーズ範囲外） |
| ② | 🟡 | 「責務 1 つにつき 1 subagent」ルール。`research` / `factcheck` / `visual-review` / `summary` 等の境界を明示 | Agent ツールで分割。market-overview-agent (10+ Step) → research-subagent + 親 (5 Step) |
| ③ | 🟡 | Step 数を TaskCreate で可視化 | TaskList で過大な Step 数を発見。Step 8+ になったら orchestrator の見直し候補 |

### #11 どこからでも起動できるようにする

> チャットUIに依存せず、Slack、メール、Webhook、cron、業務イベントから起動可能に

| レバー | Status | ルール | 仕組み |
|---|---|---|---|
| ① | 🟡 | SKILL.md に「単独起動 / 内部呼び出し」両対応の引数仕様を記載（既存 orchestrator は対応済） | description で trigger phrase を整理（slash command 化の前段） |
| ② | N/A | — | — |
| ③ | N/A | — | — |

**所見**: #11 はレバー④（slash / cron / SDK）が本丸。本フェーズでは触らない（4/30 ユーザー判断で除外）。

### #12 エージェントをステートレスなリデューサとして設計する

> 「現在の状態 + 入力 → 次の状態」を返す関数として設計し、再実行・テスト・デバッグしやすく

| レバー | Status | ルール | 仕組み |
|---|---|---|---|
| ① | 🟡 | scope.json / run_state.json 規約を SKILL.md / `orchestrator_contract.md` で統一 | SessionStart hook で state file 読込→ context へ |
| ② | 🟢 | subagent は (input → output) の純関数として設計、内部 state を持たない | Agent ツール内で完結 |
| ③ | 🔴 | 各 Step で TaskCreate / TaskUpdate。タスク状態を外部化し、LLM が「前回の会話で覚えているはず」を回避 | TaskList が state machine の真実源 |

---

## Phase B タスク導出

このマトリクスから、Phase B で実施すべき具体タスクが導出される:

### レバー① 関連（hooks / settings.json）→ Phase B-1, B-2 で実装
- PreToolUse: `check_merge_order_exists.py`（#8 の 🔴）
- PreToolUse: `check_task_progression.py`（#5 #8 #12 の 🔴 を一気に解消）
- PostToolUse: `validate_pptx_after_fill.py`（#1 #4 #9 の 🟡 改善）
- SessionStart: `load_session_context.py`（#3 #5 #6 #12 の 🟡 改善）
- description / triggers 規約 → `harness_levers.md`（#3 #10 の 🔴 / 🟡 改善）

### レバー② 関連（subagent）→ Phase B-3 で実装
- `research-subagent` 試作（#3 #4 #10 の 🔴 / 🟡 改善）

### レバー③ 関連（AskUserQuestion / TaskCreate）→ Phase B-4 / B-6 で実装
- `step_state_tracking.md` 規約（#5 #6 #8 #12 の 🔴 を一気に解消）
- 既存 orchestrator 3 本に TaskCreate / AskUserQuestion マーカー追加

### 共通基盤
- `harness_levers.md`（規約のハブ、Phase B-4 で作成）
- `.claude/settings.json`（hooks 配線、Phase B-1 で作成）

---

## Phase B 完了時の期待 Status

| # | 原則 | ① | ② | ③ |
|---|---|---|---|---|
| 1 | ツール呼び出し | 🟢 (PreToolUse) | 🟢 (research-subagent) | 🟢 |
| 2 | プロンプト管理 | 🟢 | 🟢 | N/A |
| 3 | コンテキスト制御 | 🟡 (descr 精緻化) | 🟢 (subagent) | 🟢 (TaskCreate) |
| 4 | 構造化出力 | 🟢 (hook validation) | 🟢 | N/A |
| 5 | 状態統合 | 🟢 (SessionStart) | 🟢 | 🟢 (TaskCreate) |
| 6 | 開始停止再開 | 🟢 (SessionStart) | 🟢 | 🟢 (TaskList resume) |
| 7 | 人間との対話 | N/A | 🟢 | 🟢 |
| 8 | 制御フロー | 🟢 (PreToolUse assert) | 🟡 | 🟢 (TaskUpdate) |
| 9 | エラー圧縮 | 🟢 (PostToolUse) | 🟢 | N/A |
| 10 | 責務サイズ | 🟡 (規約) | 🟡 (部分分割) | 🟡 |
| 11 | どこからでも起動 | 🟡 | N/A | N/A |
| 12 | ステートレス | 🟡 (state file) | 🟢 | 🟢 (TaskCreate) |

**目標**: 🔴 を 0、🟢 を現在 15 → 24 へ。#10 / #11 は本フェーズで満点にしない（subagent 完全分割と起動経路拡張は別フェーズ）。

---

## Phase B 完了時の **実 Status**(2026-05-02)

α 検証（smoke test + build_skill.py check + 規約整備）まで完了した時点の Status。**E2E 未確認（β / γ）のため、実装済でも実動作確認は次セッション以降**。

| # | 原則 | ① | ② | ③ | 備考 |
|---|---|---|---|---|---|
| 1 | ツール呼び出し | 🟡 ¹ | 🟡 ² | 🟢 | ¹ hook 実装済 unit test PASS、E2E 未 / ² subagent 配置済、E2E 未 |
| 2 | プロンプト管理 | 🟢 | 🟢 | N/A | 元から達成 |
| 3 | コンテキスト制御 | 🟡 ³ | 🟡 ² | 🟡 ⁴ | ³ description 規約は 5-15 行で書いたが、実 LLM の発火抑制効果は未測定 / ⁴ TaskCreate 規約は SKILL.md に書いたが LLM 遵守は未測定 |
| 4 | 構造化出力 | 🟡 ¹ | 🟡 ² | N/A | hook validation の実発火は β で確認 |
| 5 | 状態統合 | 🟡 ⁵ | 🟢 | 🟡 ⁴ | ⁵ SessionStart hook は smoke test で stdout 出力確認済、実 LLM context 注入の効果は未測定 |
| 6 | 開始停止再開 | 🟡 ⁵ | 🟢 | 🟡 ⁴ | TaskList による resume 動作は実運用で確認要 |
| 7 | 人間との対話 | N/A | 🟢 | 🟢 | 元から達成、subagent SKILL.md で AskUserQuestion 不可を明示済 |
| 8 | 制御フロー | 🟡 ¹ | 🟡 ² | 🟡 ⁴ | check_task_progression.py の inversion 検出は unit test PASS、実 LLM が TaskCreate を呼ぶかの確認が β / γ のキーポイント |
| 9 | エラー圧縮 | 🟡 ¹ | 🟢 | N/A | validate_pptx_after_fill.py の実発火は β で確認 |
| 10 | 責務サイズ | 🟡 (規約) | 🟡 (部分分割) | 🟡 | 元計画通り、本フェーズで満点目指さず |
| 11 | どこからでも起動 | 🟡 | N/A | N/A | レバー④ 対応の別フェーズ（本フェーズ範囲外） |
| 12 | ステートレス | 🟡 ⁵ | 🟢 | 🟡 ⁴ | task_state.json スキーマ確定済、運用は β / γ で確認 |

### 凡例

- 🟢: 実装 + 実動作確認済（smoke test or 元からの実績）
- 🟡: 実装済だが E2E 未確認 / 部分達成
- 🔴: 未対応 / 期待外

### 集計
- 🟢: 6（元から達成のセル）
- 🟡: 22（実装済 E2E 未）
- 🔴: 0
- N/A: 8

### 次セッションでの目標
β / γ E2E 実施で 🟡 → 🟢 への昇格を目指す。少なくとも `#1 ① / #4 ① / #8 ① / #9 ①` の hooks 系と `#1 ② / #3 ② / #4 ②` の subagent 系は 🟢 に上がる見込み。詳細は `docs/harness_check/handoff.md` Section 2 / `~/.claude/plans/harness-check-verification-next-session.md` 参照。

---

## Phase B 完了時の **最終実 Status**(2026-05-03 / β + γ E2E 完了後)

β（business-deepdive-agent / 二幸産業 × 施設運営事業）と γ（market-overview-agent / 国内タクシー市場 Step 1 限定 Before/After 比較）の E2E 実測結果に基づく最終 Status。

| # | 原則 | ① | ② | ③ | β/γ で確認した動作 |
|---|---|---|---|---|---|
| 1 | ツール呼び出し | 🟢 | 🟢 | 🟢 | β で hook 発火確認、γ で subagent 経由の構造化 JSON 返却確認 |
| 2 | プロンプト管理 | 🟢 | 🟢 | N/A | 元から達成 |
| 3 | コンテキスト制御 | 🟡 ⁶ | 🟡 ⁷ | 🟢 | ⁶ description 精緻化は実装済だが LLM 発火抑制の定量測定なし / ⁷ γ で 38-44 件の生検索結果隔離確認も、return value wrapper bloat で削減効果は理論値 50% → 実測 20% |
| 4 | 構造化出力 | 🟢 | 🟢 | N/A | β / γ で fill_*.py / subagent JSON return が機能 |
| 5 | 状態統合 | 🟢 | 🟢 | 🟢 | β で task_state.json 7 step 全記録確認 |
| 6 | 開始停止再開 | 🟡 ⁸ | 🟢 | 🟢 | ⁸ resume シナリオは β / γ で発生せず（一気通貫実行）、運用での確認待ち |
| 7 | 人間との対話 | N/A | 🟢 | 🟢 | 元から達成 |
| 8 | 制御フロー | 🟢 | 🟡 | 🟢 | β で AskUserQuestion / TaskUpdate 確認、subagent 順序起動も論点 4→5 で機能 |
| 9 | エラー圧縮 | 🟡 ⁹ | 🟢 | N/A | ⁹ β / γ で実エラー発生せず PostToolUse hook 発火なし、unit test 範囲のみ |
| 10 | 責務サイズ | 🟡 (規約) | 🟡 (部分分割) | 🟡 | 元計画通り、本フェーズで満点目指さず |
| 11 | どこからでも起動 | 🟡 | N/A | N/A | レバー④ 対応の別フェーズ（本フェーズ範囲外） |
| 12 | ステートレス | 🟢 | 🟢 | 🟢 | β で scope.json + task_state.json + data_NN_*.json の状態外部化を全 step で確認 |

### 集計
- 🟢: 18（元から達成 6 + β/γ で昇格 12）
- 🟡: 7（実装済だが効果未測定 or 想定通り未達）
- 🔴: 0
- N/A: 11

🔴 ゼロ達成。期待 Status 24 グリーンに対し実測 18 グリーンで、**75% 達成**。

### γ で発見した重要事象（残課題）

**subagent return value の wrapper bloat 問題**（#3 ② が 🟢 に届かない理由）:

| 観察 | 件数 | 効果 |
|---|---|---|
| (A) 強化版 research-subagent.md でも、6 subagent 全てが JSON 外混入を再発 | 6/6 | preamble + code fence + Sources trailing |
| 二重 JSON 出力（最深刻、`data_06`） | 1/6 | +5k tokens の重複 |
| 結果として After Step 1 削減効果 | -20.4%(理論値 -50%) | wrapper bloat が削減量を半減 |

**subagent 機能自体は動作している**(38-44 件の生検索結果は完全隔離) が、return value 規約遵守の弱さで実効性に課題あり。**新規 ISSUE 候補**(ISSUES.md 参照): (B) 親側 JSON 抽出 helper の実装を v0.4 で検討。

### 凡例

- 🟢: 実装 + β / γ E2E で実動作確認済
- 🟡: 実装済だが効果未測定 / 部分達成
- 🔴: 未対応 / 期待外
