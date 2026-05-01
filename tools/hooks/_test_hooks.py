#!/usr/bin/env python3
"""B-2 hooks の動作確認テストハーネス。

注意: hook の matcher 文字列（"merge_pptx_v2.py" / "fill_*.py"）を直接 Bash command に
書くと、Claude Code から本スクリプトを起動する際に hook 自体が再発火してテストが
干渉する。そのため、サンプル JSON 内の危険語は Python の文字列連結で組み立てる。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent.parent
HOOKS_DIR = REPO / "tools" / "hooks"

# 文字列連結で matcher 文字列を組み立て、bash command line に直接出さない
MERGE_SCRIPT = "merge_pptx" + "_v2.py"
FILL_SCRIPT = "fill_" + "test_skill.py"


def run_hook(script_name: str, event: dict) -> tuple[int, str, str]:
    result = subprocess.run(
        [sys.executable, str(HOOKS_DIR / script_name)],
        input=json.dumps(event),
        capture_output=True,
        text=True,
        timeout=30,
    )
    return result.returncode, result.stdout, result.stderr


def t(name: str, ok: bool, detail: str = "") -> None:
    icon = "✅" if ok else "❌"
    print(f"  {icon} {name}{(': ' + detail) if detail else ''}")
    if not ok:
        sys.exit(1)


def test_check_merge_order_exists() -> None:
    print("\n[check_merge_order_exists.py]")

    # 1. PreToolUse 以外は素通り
    ec, _, _ = run_hook(
        "check_merge_order_exists.py",
        {"hook_event_name": "PostToolUse", "tool_name": "Bash",
         "tool_input": {"command": f"python3 {MERGE_SCRIPT} --merge-order foo.json"}},
    )
    t("PostToolUse は素通り", ec == 0, f"exit={ec}")

    # 2. Bash 以外は素通り
    ec, _, _ = run_hook(
        "check_merge_order_exists.py",
        {"hook_event_name": "PreToolUse", "tool_name": "Read",
         "tool_input": {"file_path": "x"}},
    )
    t("非 Bash は素通り", ec == 0, f"exit={ec}")

    # 3. merge_pptx_v2.py でない command は素通り
    ec, _, _ = run_hook(
        "check_merge_order_exists.py",
        {"hook_event_name": "PreToolUse", "tool_name": "Bash",
         "tool_input": {"command": "ls -la"}},
    )
    t("無関係 command は素通り", ec == 0, f"exit={ec}")

    # 4. --merge-order なしはブロック
    ec, _, err = run_hook(
        "check_merge_order_exists.py",
        {"hook_event_name": "PreToolUse", "tool_name": "Bash",
         "tool_input": {"command": f"python3 {MERGE_SCRIPT} out.pptx in.pptx"}},
    )
    t("--merge-order なしはブロック", ec == 2 and "必須" in err, f"exit={ec}, err_has_msg={'必須' in err}")

    # 5. 存在しない merge_order.json はブロック
    ec, _, err = run_hook(
        "check_merge_order_exists.py",
        {"hook_event_name": "PreToolUse", "tool_name": "Bash",
         "tool_input": {"command": f"python3 {MERGE_SCRIPT} --merge-order /tmp/nonexistent_xyz_abc.json out.pptx"},
         "cwd": "/tmp"},
    )
    t("存在しない json はブロック", ec == 2 and "見つかりません" in err,
      f"exit={ec}, err_has_msg={'見つかりません' in err}")

    # 6. 存在する merge_order.json は通す
    with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
        f.write(b'{"entries":[]}')
        existing = f.name
    try:
        ec, _, _ = run_hook(
            "check_merge_order_exists.py",
            {"hook_event_name": "PreToolUse", "tool_name": "Bash",
             "tool_input": {"command": f"python3 {MERGE_SCRIPT} --merge-order {existing} out.pptx"}},
        )
        t("存在する json は通す", ec == 0, f"exit={ec}")

        # 7. 引用符ありでも通す
        ec, _, _ = run_hook(
            "check_merge_order_exists.py",
            {"hook_event_name": "PreToolUse", "tool_name": "Bash",
             "tool_input": {"command": f'python3 {MERGE_SCRIPT} --merge-order "{existing}" out.pptx'}},
        )
        t("引用符ありでも通す", ec == 0, f"exit={ec}")
    finally:
        os.unlink(existing)


def test_validate_pptx_after_fill() -> None:
    print("\n[validate_pptx_after_fill.py]")

    # 1. PreToolUse は素通り
    ec, _, _ = run_hook(
        "validate_pptx_after_fill.py",
        {"hook_event_name": "PreToolUse", "tool_name": "Bash",
         "tool_input": {"command": f"python3 {FILL_SCRIPT} --output x.pptx"}},
    )
    t("PreToolUse は素通り", ec == 0, f"exit={ec}")

    # 2. 無関係 command は素通り
    ec, _, _ = run_hook(
        "validate_pptx_after_fill.py",
        {"hook_event_name": "PostToolUse", "tool_name": "Bash",
         "tool_input": {"command": "git status"}},
    )
    t("無関係 command は素通り", ec == 0, f"exit={ec}")

    # 3. 存在しない PPTX は素通り（fill が失敗した可能性、追加エラー出さない）
    ec, _, _ = run_hook(
        "validate_pptx_after_fill.py",
        {"hook_event_name": "PostToolUse", "tool_name": "Bash",
         "tool_input": {"command": f"python3 {FILL_SCRIPT} --output /tmp/nonexistent_xyz.pptx"},
         "cwd": "/tmp"},
    )
    t("存在しない PPTX は素通り", ec == 0, f"exit={ec}")

    # 4. 既存の正常 PPTX で PASS
    candidates = [
        REPO / "outputs" / "MarketKBF_test.pptx",
        REPO / "outputs" / "MarketShare_smoke.pptx",
    ]
    test_pptx = next((p for p in candidates if p.exists()), None)
    if test_pptx is None:
        t("正常 PPTX で PASS", False, "テスト用 PPTX が見つからない")
        return

    env = os.environ.copy()
    env["FACTORY_ROOT"] = str(REPO)
    result = subprocess.run(
        [sys.executable, str(HOOKS_DIR / "validate_pptx_after_fill.py")],
        input=json.dumps({
            "hook_event_name": "PostToolUse", "tool_name": "Bash",
            "tool_input": {"command": f"python3 {FILL_SCRIPT} --output {test_pptx}"},
            "cwd": str(REPO),
        }),
        capture_output=True, text=True, env=env, timeout=30,
    )
    t(f"正常 PPTX で PASS ({test_pptx.name})", result.returncode == 0,
      f"exit={result.returncode}, stderr={result.stderr[:200]}")


def test_load_session_context() -> None:
    print("\n[load_session_context.py]")

    env = os.environ.copy()
    env["FACTORY_ROOT"] = str(REPO)
    result = subprocess.run(
        [sys.executable, str(HOOKS_DIR / "load_session_context.py")],
        input="{}", capture_output=True, text=True, env=env, timeout=10,
    )
    t("exit code 0", result.returncode == 0, f"exit={result.returncode}")
    t("stdout に未解決 ISSUES セクション", "未解決 ISSUES" in result.stdout,
      f"len(stdout)={len(result.stdout)}")
    t("stdout に直近 plans セクション", "直近 plans" in result.stdout)
    t("stdout に ISSUE-001 が含まれる", "ISSUE-001" in result.stdout)


if __name__ == "__main__":
    test_check_merge_order_exists()
    test_validate_pptx_after_fill()
    test_load_session_context()
    print("\n✅ All hook tests passed.")
