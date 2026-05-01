#!/usr/bin/env python3
"""PreToolUse hook (Bash matcher) — merge_pptx_v2.py 起動前に merge_order.json 存在を assert.

stdin に Claude Code から渡される PreToolUse イベントの JSON を受ける。
command が merge_pptx_v2.py の起動なら --merge-order の引数を抽出し、指定された
JSON ファイルが存在することを確認する。存在しない場合は exit 2 でブロックする。

Exit code:
  0 = 素通り or 検証 PASS
  2 = ブロック (stderr に「merge_order.json が見つからない」旨)
"""
from __future__ import annotations

import json
import os
import re
import sys


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
    if "merge_pptx_v2.py" not in cmd:
        return 0

    m = re.search(r"--merge-order\s+(['\"]?)([^\s'\"]+)\1", cmd)
    if not m:
        print(
            "merge_pptx_v2.py 起動時は --merge-order 必須です（skills_factory 規約）。\n"
            "Step 6 で merge_order.json を作成してから再起動してください。",
            file=sys.stderr,
        )
        return 2

    path = m.group(2)
    cwd = event.get("cwd") or os.getcwd()
    if not os.path.isabs(path):
        path = os.path.join(cwd, path)

    if not os.path.exists(path):
        print(
            f"merge_order.json が見つかりません: {path}\n"
            "Step 6 で merge_order.json を作成してから merge-pptxv2 を起動してください。\n"
            "スキーマは skills/_common/references/orchestrator_contract.md を参照。",
            file=sys.stderr,
        )
        return 2

    return 0


if __name__ == "__main__":
    sys.exit(main())
