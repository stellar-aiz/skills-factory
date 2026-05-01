#!/usr/bin/env python3
"""PostToolUse hook (Bash matcher) — fill_*.py / merge_pptx_v2.py 後の PPTX 自動検証.

stdin に Claude Code から渡される PostToolUse イベントの JSON を受ける。
command が fill_*.py または merge_pptx_v2.py の起動なら、出力 PPTX を
tools/validate_pptx.py で検証し、失敗していれば exit 2 でブロックする。

Exit code:
  0 = 素通り or 検証 PASS
  2 = 検証 FAIL (stderr に validate_pptx の出力を流す)
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys


def _extract_pptx_path(cmd: str) -> str | None:
    """command 文字列から検証対象の PPTX パスを抽出。"""
    if "fill_" in cmd and ".py" in cmd:
        m = re.search(r"--output\s+(['\"]?)([^\s'\"]+)\1", cmd)
        if m:
            return m.group(2)
    if "merge_pptx_v2.py" in cmd:
        # merge_pptx_v2.py の最初の位置引数が出力先 PPTX
        m = re.search(r"merge_pptx_v2\.py\s+(?:--\S+\s+\S+\s+)*?(['\"]?)([^\s'\"]+\.pptx)\1", cmd)
        if m:
            return m.group(2)
    return None


def main() -> int:
    try:
        event = json.load(sys.stdin)
    except (json.JSONDecodeError, ValueError):
        return 0

    if event.get("hook_event_name") != "PostToolUse":
        return 0
    if event.get("tool_name") != "Bash":
        return 0

    cmd = event.get("tool_input", {}).get("command", "")
    if not re.search(r"(fill_\w+\.py|merge_pptx_v2\.py)", cmd):
        return 0

    pptx = _extract_pptx_path(cmd)
    if not pptx:
        return 0

    cwd = event.get("cwd") or os.getcwd()
    if not os.path.isabs(pptx):
        pptx = os.path.join(cwd, pptx)
    if not os.path.exists(pptx):
        # fill_*.py が失敗して出力がない場合。ここでは追加エラーを出さない。
        return 0

    factory_root = os.environ.get("FACTORY_ROOT") or cwd
    validate_script = os.path.join(factory_root, "tools", "validate_pptx.py")
    if not os.path.exists(validate_script):
        return 0

    try:
        result = subprocess.run(
            ["python3", validate_script, pptx],
            capture_output=True,
            text=True,
            timeout=25,
        )
    except subprocess.TimeoutExpired:
        print(f"⚠️ validate_pptx.py が timeout (25s): {pptx}", file=sys.stderr)
        return 2

    if result.returncode != 0:
        print(f"⚠️ 生成 PPTX に問題が検出されました: {pptx}", file=sys.stderr)
        if result.stdout:
            print(result.stdout, file=sys.stderr)
        if result.stderr:
            print(result.stderr, file=sys.stderr)
        return 2

    return 0


if __name__ == "__main__":
    sys.exit(main())
