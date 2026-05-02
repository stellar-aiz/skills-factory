# `.claude/settings.json` 設計メモ

skills_factory リポジトリ内の `.claude/settings.json` の設計方針と、グローバル設定とのマージ規則を整理する。

## ファイル配置と優先順位

Claude Code は以下の順で設定をマージする（後優先 = 上書き）:

1. `~/.claude/settings.json` — グローバル（全プロジェクト共通）
2. **`.claude/settings.json` — プロジェクト共通（git 管理、本ファイルが対象）**
3. `.claude/settings.local.json` — プロジェクト個人（gitignore、個人専用）
4. （CLI 引数等）

skills_factory リポを clone した全員に適用したい設定は **2** に書く。個人専用は **3**（gitignore 推奨）。

## 衝突管理

### `permissions.allow`

配列はマージされる（重複は実害なし）。本リポの `.claude/settings.json` には skills_factory の routine 操作だけを書き、汎用的な開発作業（git 操作、エディタ起動 等）はグローバル側に任せる。

### `env`

同名キーはプロジェクト側が優先。本リポでは `FACTORY_ROOT` のみ定義。

### `hooks`

**重要**: 同じイベント（例 `PreToolUse`）に対して、グローバルとプロジェクトの両方に hook が定義されていると **両方発火**する。マージではなく加算挙動。

skills_factory リポの hooks は全て `tools/hooks/*.py` を呼ぶ形式に統一しているため、グローバル側に hook がなければ衝突しない。本リポ adoption 時には clone 者のグローバル `~/.claude/settings.json` に hooks があるか念のため確認。

調査時点（Phase A.5 / 2026-04-30）では nakamaru 個人のグローバル設定に hooks 定義は無し。skills_factory のチームメンバー全員も同様であることを前提とする。

## 含めている内容（4/30 時点）

### `env`

| キー | 値 | 用途 |
|---|---|---|
| `FACTORY_ROOT` | `$CLAUDE_PROJECT_DIR` | scripts と SKILL.md から参照されるプロジェクト ルート。`profiles/claude_code.json` の `FACTORY_ROOT=AUTO` と等価な値を Claude Code セッションに渡す |

### `permissions.allow`

skills_factory で頻出する Bash コマンドを許可リストに登録し、ユーザーの permission prompt を削減。具体的には:

- `tools/*.py` の各種ビルド・検証スクリプト
- `fill_*.py` / `merge_pptx_v2.py` の起動
- `pip install python-pptx` / `pip install lxml`（PPTX 処理に必須）
- `python -m markitdown`（PPTX 確認用）
- `mkdir -p outputs/*` / `work/*` / `docs/*`

破壊的操作（`rm`, `git push -f`, etc.）は **意図的に含めない**。これらはユーザー承認を毎回取る運用を維持。

### `hooks`

| イベント | matcher | スクリプト | 状態 |
|---|---|---|---|
| `SessionStart` | (なし) | `tools/hooks/load_session_context.py` | 実装済 (B-2) |
| `PreToolUse` | `Bash` | `tools/hooks/check_merge_order_exists.py` | 実装済 (B-2) |
| `PreToolUse` | `Bash` | `tools/hooks/check_task_progression.py` | 実装済 (B-2-d) |
| `PostToolUse` | `Bash` | `tools/hooks/validate_pptx_after_fill.py` | 実装済 (B-2) |

詳細は `tools/hooks/README.md` 参照。`tools/hooks/_test_hooks.py` で 26 ユニットテストが PASS。

## 採用していない設定（意図的）

| キー | 採用しない理由 |
|---|---|
| `model` | グローバル設定 / `/fast` トグル / `/model` コマンドでセッション単位に設定する方が柔軟 |
| `theme` | 個人差が大きい、global で設定すべき |
| `statusLine` | global で設定済（npx ccstatusline）|
| `hooks.Stop` | 現状で具体的な用途なし。E2E 検証で必要性が出てきたら追加 |
| `hooks.UserPromptSubmit` | プロンプト前処理は CLAUDE.md の SessionStart で十分 |

## Phase B 進行に伴う改訂予定

- ✅ B-2 で hook スクリプトを実体化、状態欄を「実装済」に更新（2026-05-01）
- ✅ B-2-d で `check_task_progression.py` を実装、task_state.json ベースの Step ordering 検出に確定（2026-05-02）
- ✅ B-3 で `.claude/agents/research-subagent.md` を導入。tools whitelist が `WebSearch / WebFetch / Read` のみで read-only のため `permissions` への追加は不要（2026-05-01）
- B-6 で 既存 orchestrator 3 本を改訂したら、それらが期待する `FACTORY_ROOT` 等を再確認（次フェーズ）

## 参考

- Claude Code 公式 hooks 仕様（4/30 時点で claude-code-guide で確認済）
- `tools/hooks/README.md` — hook 実装規約
- `docs/harness_check/lever_mapping.md` — 12箇条 × hooks のマッピング
- `~/.claude/plans/md-llm-melodic-twilight.md` — 全体計画
