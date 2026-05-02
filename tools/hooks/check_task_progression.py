#!/usr/bin/env python3
"""PreToolUse hook (Bash matcher) — task_state.json の Step 順序整合性を assert.

step_state_tracking.md(skills/_common/prompts/) で定義された task_state.json を
{{FACTORY_ROOT}}/work/<orchestrator>/<run_id>/task_state.json から探し、
**Step ordering inversion**（最後の step より前に completed でない step がある状態）
を検出してブロックする。

検出ロジック:
  steps[:-1] のうち status != "completed" のものがあれば違反。
  → 「前 Step を completed にせずに次 Step に進んだ」典型ケースを捕捉。

発火条件: command が fill_*.py / merge_pptx_v2.py を含むときだけチェック。
無関係な Bash 呼び出し（ls / git status 等）には影響しない。

Exit code:
  0 = 素通り or 整合性 OK
  2 = 順序違反検出 (stderr に違反内容 + state file パス)
"""
from __future__ import annotations

import glob
import json
import os
import re
import sys
from pathlib import Path


def _is_step_transition(cmd: str) -> bool:
    """command が Step 遷移を伴うツール起動なら True。"""
    return bool(re.search(r"(fill_\w+\.py|merge_pptx_v2\.py)", cmd))


def _find_active_task_state(factory_root: str) -> str | None:
    """{{FACTORY_ROOT}}/work/*/*/task_state.json のうち最も最近更新されたものを返す。"""
    pattern = str(Path(factory_root) / "work" / "*" / "*" / "task_state.json")
    candidates = glob.glob(pattern)
    if not candidates:
        return None
    return max(candidates, key=os.path.getmtime)


def _check_state_inversion(state: dict) -> str | None:
    """順序違反があればエラーメッセージ、無ければ None を返す。"""
    steps = state.get("steps", [])
    if len(steps) < 2:
        return None  # 0 or 1 step では順序違反は定義されない
    for i, step in enumerate(steps[:-1]):
        if step.get("status") != "completed":
            return (
                f"Step ordering violation in {state.get('orchestrator', '?')} "
                f"(run_id={state.get('run_id', '?')}):\n"
                f"  steps[{i}] '{step.get('name', '?')}' "
                f"(step_id={step.get('step_id', '?')}) is "
                f"'{step.get('status', '?')}', but a later step exists.\n"
                f"  → 前 Step を completed にせずに次 Step に進むのは "
                f"step_state_tracking.md 規約違反です。\n"
                f"  該当 step を completed にしてから再実行してください。"
            )
    return None


def main() -> int:
    try:
        event = json.load(sys.stdin)
    except (json.JSONDecodeError, ValueError):
        return 0

    if event.get("hook_event_name") != "PreToolUse":
        return 0
    if event.get("tool_name") != "Bash":
        return 0

    cmd = event.get("tool_input", {}).get("command", "")
    if not _is_step_transition(cmd):
        return 0

    factory_root = os.environ.get("FACTORY_ROOT") or event.get("cwd") or os.getcwd()
    state_path = _find_active_task_state(factory_root)
    if not state_path:
        return 0  # state file 不在は backward compat（未対応 orchestrator は素通り）

    try:
        with open(state_path, encoding="utf-8") as f:
            state = json.load(f)
    except (OSError, json.JSONDecodeError):
        return 0  # state file が壊れている場合は素通り（hook の robustness 優先）

    err = _check_state_inversion(state)
    if err:
        print(err, file=sys.stderr)
        print(f"  state file: {state_path}", file=sys.stderr)
        return 2

    return 0


if __name__ == "__main__":
    sys.exit(main())
