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


def run_hook(
    script_name: str,
    event: dict,
    env_overrides: dict | None = None,
) -> tuple[int, str, str]:
    env = os.environ.copy()
    if env_overrides:
        env.update(env_overrides)
    result = subprocess.run(
        [sys.executable, str(HOOKS_DIR / script_name)],
        input=json.dumps(event),
        capture_output=True,
        text=True,
        timeout=30,
        env=env,
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


def _make_state_file(tmp: Path, orchestrator: str, run_id: str, steps: list) -> Path:
    work_dir = tmp / "work" / orchestrator / run_id
    work_dir.mkdir(parents=True, exist_ok=True)
    state_path = work_dir / "task_state.json"
    state_path.write_text(json.dumps({
        "run_id": run_id,
        "orchestrator": orchestrator,
        "started_at": "2026-05-02T10:00:00+09:00",
        "steps": steps,
    }), encoding="utf-8")
    return state_path


def test_check_task_progression() -> None:
    print("\n[check_task_progression.py]")

    bash_event = {
        "hook_event_name": "PreToolUse",
        "tool_name": "Bash",
        "tool_input": {"command": f"python3 {FILL_SCRIPT} --output x.pptx"},
    }

    # 1. PreToolUse でない → 素通り
    ec, _, _ = run_hook("check_task_progression.py",
                        {**bash_event, "hook_event_name": "PostToolUse"})
    t("PostToolUse は素通り", ec == 0, f"exit={ec}")

    # 2. 非 Bash → 素通り
    ec, _, _ = run_hook("check_task_progression.py",
                        {**bash_event, "tool_name": "Read"})
    t("非 Bash は素通り", ec == 0, f"exit={ec}")

    # 3. 無関係 command (fill_*/merge_pptx_v2 でない) → 素通り
    ec, _, _ = run_hook("check_task_progression.py",
                        {**bash_event, "tool_input": {"command": "ls -la"}})
    t("無関係 command は素通り", ec == 0, f"exit={ec}")

    with tempfile.TemporaryDirectory() as tmp_str:
        tmp = Path(tmp_str)

        # 4. state file 不在 → 素通り（backward compat）
        ec, _, _ = run_hook("check_task_progression.py", bash_event,
                            env_overrides={"FACTORY_ROOT": str(tmp)})
        t("state file 不在で素通り (backward compat)", ec == 0, f"exit={ec}")

        # 5. 単一 step (順序違反不能) → 素通り
        _make_state_file(tmp, "test-orchestrator", "test_run_5", [
            {"step_id": "step_1", "name": "Step 1", "status": "in_progress"},
        ])
        ec, _, _ = run_hook("check_task_progression.py", bash_event,
                            env_overrides={"FACTORY_ROOT": str(tmp)})
        t("単一 step は素通り", ec == 0, f"exit={ec}")

        # 6. 全 step completed (最後の step も含む) → 素通り
        _make_state_file(tmp, "test-orchestrator", "test_run_6", [
            {"step_id": "step_1", "name": "Step 1", "status": "completed"},
            {"step_id": "step_2", "name": "Step 2", "status": "completed"},
        ])
        ec, _, _ = run_hook("check_task_progression.py", bash_event,
                            env_overrides={"FACTORY_ROOT": str(tmp)})
        t("全 step completed で素通り", ec == 0, f"exit={ec}")

        # 7. 末尾 in_progress、それ以前 completed → 素通り（正常進行）
        _make_state_file(tmp, "test-orchestrator", "test_run_7", [
            {"step_id": "step_1", "name": "Step 1", "status": "completed"},
            {"step_id": "step_2", "name": "Step 2", "status": "in_progress"},
        ])
        ec, _, _ = run_hook("check_task_progression.py", bash_event,
                            env_overrides={"FACTORY_ROOT": str(tmp)})
        t("正常進行は素通り (last=in_progress)", ec == 0, f"exit={ec}")

        # 8. State inversion: 前 step が in_progress のまま次に進んだ → ブロック
        _make_state_file(tmp, "test-orchestrator", "test_run_8", [
            {"step_id": "step_1", "name": "Step 1: Web検索", "status": "in_progress"},
            {"step_id": "step_2", "name": "Step 2: 整理", "status": "in_progress"},
        ])
        ec, _, err = run_hook("check_task_progression.py", bash_event,
                              env_overrides={"FACTORY_ROOT": str(tmp)})
        t("state inversion (前 step in_progress) をブロック",
          ec == 2 and "規約違反" in err and "step_1" in err,
          f"exit={ec}")

        # 9. State inversion: 前 step が failed → ブロック
        _make_state_file(tmp, "test-orchestrator", "test_run_9", [
            {"step_id": "step_1", "name": "Step 1", "status": "failed"},
            {"step_id": "step_2", "name": "Step 2", "status": "in_progress"},
        ])
        ec, _, err = run_hook("check_task_progression.py", bash_event,
                              env_overrides={"FACTORY_ROOT": str(tmp)})
        t("state inversion (前 step failed) をブロック",
          ec == 2 and "failed" in err, f"exit={ec}")

        # 10. State inversion: 前 step が pending → ブロック
        _make_state_file(tmp, "test-orchestrator", "test_run_10", [
            {"step_id": "step_1", "name": "Step 1", "status": "pending"},
            {"step_id": "step_2", "name": "Step 2", "status": "in_progress"},
        ])
        ec, _, err = run_hook("check_task_progression.py", bash_event,
                              env_overrides={"FACTORY_ROOT": str(tmp)})
        t("state inversion (前 step pending) をブロック",
          ec == 2 and "pending" in err, f"exit={ec}")

        # 11. 壊れた state file → 素通り（robustness）
        broken_dir = tmp / "work" / "broken-orch" / "broken_run"
        broken_dir.mkdir(parents=True)
        (broken_dir / "task_state.json").write_text("{ this is not valid json")
        ec, _, _ = run_hook("check_task_progression.py", bash_event,
                            env_overrides={"FACTORY_ROOT": str(tmp)})
        # 11 では他の正常 state file (steps[1] inversion なし) が複数存在し、
        # 「最も最近更新された」が壊れた file になる可能性。
        # その場合は exit 0 (素通り)、別の file が拾われたら exit によらず robust 確認のみ
        t("壊れた state file は exception せず終了", ec in (0, 2),
          f"exit={ec} (どちらでも robustness OK)")


if __name__ == "__main__":
    test_check_merge_order_exists()
    test_validate_pptx_after_fill()
    test_load_session_context()
    test_check_task_progression()
    print("\n✅ All hook tests passed.")
