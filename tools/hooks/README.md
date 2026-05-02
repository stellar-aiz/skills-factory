# tools/hooks/ — Claude Code hooks 実装

skills_factory プロジェクト用の Claude Code hook スクリプト群。`.claude/settings.json` の `hooks` セクションから呼び出される。

## 各 hook の責務

| ファイル | イベント | matcher | 責務 |
|---|---|---|---|
| `load_session_context.py` | `SessionStart` | (なし) | ISSUES.md と直近 plan を stdout に出力し context に注入 |
| `check_merge_order_exists.py` | `PreToolUse` | `Bash` | `merge_pptx_v2.py` 起動前に `--merge-order` 引数のファイル存在を assert |
| `check_task_progression.py` | `PreToolUse` | `Bash` | `fill_*.py` / `merge_pptx_v2.py` 起動前に `task_state.json` の Step ordering inversion を検出。違反時 exit 2 でブロック |
| `validate_pptx_after_fill.py` | `PostToolUse` | `Bash` | `fill_*.py` / `merge_pptx_v2.py` 実行後に出力 PPTX を `tools/validate_pptx.py` で自動検証 |

すべて Phase B-2 で実装済み。`tools/hooks/_test_hooks.py` で **26 ユニットテスト** が PASS することを確認している（PASS/FAIL の expected behavior、エッジケース、引用符パターン、state inversion 検出、backward compat 等）。

## 入出力 contract

### 入力（stdin）

Claude Code は hook 起動時に stdin に JSON を流す。共通フィールド:

```json
{
  "session_id": "abc123",
  "transcript_path": "/path/to/transcript.jsonl",
  "cwd": "/current/working/dir",
  "permission_mode": "default",
  "hook_event_name": "PreToolUse"
}
```

`PreToolUse` / `PostToolUse` の場合は追加で:

```json
{
  "tool_name": "Bash",
  "tool_use_id": "toolu_01abc",
  "tool_input": {
    "command": "python3 ...",
    "timeout": 120000,
    "run_in_background": false
  }
}
```

`SessionStart` の場合:

```json
{
  "source": "startup|resume|clear|compact",
  "model": "claude-sonnet-4-6"
}
```

各 hook スクリプトは冒頭で `event = json.load(sys.stdin)` で読む。複数イベントを 1 スクリプトで処理する場合は `event["hook_event_name"]` で分岐する。

### 出力（exit code + stderr/stdout）

| Exit Code | 意味 |
|---|---|
| `0` | OK、ツール実行を通す |
| `2` | **ブロック**。stderr の内容を Claude に context として注入する |
| その他 | Non-blocking error（実行は続行、ユーザーに通知） |

`SessionStart` の特殊挙動: exit 0 でも **stdout に書いた内容が context に直接追加される**。他のイベントでは stdout は debug log のみ。構造化 context が必要な場合は exit 0 で stdout に JSON `{"additionalContext": "..."}` を書く。

## 実装規約（B-2 で守ること）

1. **冪等性**: 同じ stdin に対して常に同じ exit code / stderr を返す
2. **無限ループ禁止**: hook 内部で Claude Code のツール（Bash, Skill, Agent 等）を呼ばない。subprocess での `python3 tools/validate_pptx.py` のような **直接 Python 起動**は OK
3. **タイムアウト遵守**: settings.json の `timeout` 内に必ず終了する。重い処理は別プロセス化
4. **stderr に必ず文脈を書く**: exit 2 でブロックする際は「何がダメで、どう直すか」を必ず明示
5. **ファイルベース状態共有**: TaskList 等の Claude ツールは hook から呼べないため、必要な状態は work/ 配下のファイルに書き出す規約を併用（B-4 で `step_state_tracking.md` 整備）
6. **小さいユニット**: 1 hook = 1 検証目的。複合検証は別スクリプトに分ける

## デバッグ

hook が黙って失敗するとデバッグが難しいため、各スクリプトは:

- `os.environ["SKILLS_FACTORY_HOOK_DEBUG"]` を見て、設定されていれば debug ログを stderr に書く
- 重大な内部エラーは exit code 1 で返し、stderr にトレースを書く（Claude Code がユーザーに通知）

## 関連ドキュメント

- `.claude/settings.json` — hook 配線
- `docs/harness_check/settings_design.md` — settings.json のグローバル/プロジェクトマージ規則
- `docs/harness_check/lever_mapping.md` — 12箇条 × hooks の打ち手マッピング
- `~/.claude/plans/md-llm-melodic-twilight.md` — 全体計画
