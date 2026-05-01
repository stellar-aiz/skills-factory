#!/usr/bin/env python3
"""PreToolUse hook (Bash matcher) — TODO: implement in Phase B-2.

予定: orchestrator が次の Step に進む前に、前 Step の TaskCreate タスクが completed
になっているかを検証する。Claude Code の TaskList ツールは hook から直接呼べないため、
タスク状態をファイル（例: work/<run_id>/task_state.json）に書き出させる規約を併用する。
規約は skills/_common/prompts/step_state_tracking.md を参照（B-4 で作成予定）。

今はスタブで全 Bash 呼び出しを素通りさせる（exit 0）。

設計仕様: tools/hooks/README.md を参照。
"""
import sys

sys.exit(0)
