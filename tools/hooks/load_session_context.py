#!/usr/bin/env python3
"""SessionStart hook — 未解決 ISSUES と直近 plan のサマリーを stdout に出す.

Claude Code の SessionStart 特殊挙動（exit 0 の stdout が context に直接追加される）を
利用して、セッション開始時に skills_factory の最新状態を LLM に注入する。

注意: 全文 dump はしない（ISSUES.md は 280 行、plan は数百行になる）。タイトル一覧と
ファイル名だけ出力し、LLM は必要に応じて Read で全文を取得する。

Exit code:
  0 = 常に。例外発生時もサイレントに 0 で返す（hook がセッション起動を妨げないため）
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path


def _open_issues(issues_path: Path) -> list[str]:
    if not issues_path.exists():
        return []
    text = issues_path.read_text(encoding="utf-8")
    out: list[str] = []
    pattern = re.compile(
        r"^## (ISSUE-\d+):\s*(.+?)\n.*?\*\*Status\*\*:\s*([^/]+?)(?:\s*/|\n)",
        re.MULTILINE | re.DOTALL,
    )
    for m in pattern.finditer(text):
        issue_id = m.group(1)
        title = m.group(2).strip()
        status = m.group(3).strip()
        if "保留" in status or "進行中" in status:
            out.append(f"- {issue_id} [{status}]: {title}")
    return out


def _recent_plans(plans_dir: Path, limit: int = 3) -> list[str]:
    if not plans_dir.exists():
        return []
    plans = sorted(
        (p for p in plans_dir.glob("*.md") if p.is_file()),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return [p.name for p in plans[:limit]]


def main() -> int:
    try:
        factory_root = Path(os.environ.get("FACTORY_ROOT") or os.getcwd())
        issues = _open_issues(factory_root / "ISSUES.md")
        plans = _recent_plans(Path.home() / ".claude" / "plans")

        # 何も出すべき情報がなければサイレントに終了（context を汚さない）
        if not issues and not plans:
            return 0

        print("# skills_factory セッション開始時の自動コンテキスト")
        print()

        if issues:
            print("## 未解決 ISSUES（保留 / 進行中）")
            for line in issues:
                print(line)
            print()

        if plans:
            print("## 直近 plans (~/.claude/plans/)")
            for name in plans:
                print(f"- {name}")
            print()

        print("詳細は `Read ISSUES.md` または `Read ~/.claude/plans/<name>` で参照。")
    except Exception:
        # robust: hook がセッション起動を妨げないよう、例外時は素通り
        return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
