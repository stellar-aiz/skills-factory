"""Brand compliance checker for fill_*.py outputs.

Static checks against `theme.json` rules for a given (skill, brand) pair.
Designed to catch regressions earlier than visual review (e.g. overlapping
textboxes, forgotten template artefacts, allowed font-size set violations).

Usage
-----
  # Single pptx
  python tools/check_brand_compliance.py \\
      --pptx outputs/v2_phase4_align_v5/cp_roleup.pptx \\
      --skill customer-profile-pptx \\
      --brand roleup

  # Multiple pptx (each must specify --skill matching its content)
  python tools/check_brand_compliance.py \\
      --pptx outputs/v2_phase4_align_v5/cp_roleup.pptx --skill customer-profile-pptx \\
      --pptx outputs/v2_phase4_align_v5/me_roleup.pptx --skill market-environment-pptx \\
      --brand roleup

  # JSON output (default: human-readable text)
  python tools/check_brand_compliance.py --pptx X.pptx --skill Y --brand roleup --format json

Exit code: 0 if all PASS, 1 if any rule fails (severity=error).

Available profiles
------------------
  pilot3 × roleup        — full coverage (C1/C2/C4/C5/C6/C7/C8/C10/C11/C12)
  pilot3 × stellar_aiz   — skeleton (returns warning until ISSUE-010 populates)
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from pptx import Presentation

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "skills" / "_common" / "lib"))

from brand_compliance_rules import (  # noqa: E402
    CheckContext,
    CheckResult,
    load_theme,
    run_profile,
)


def parse_args(argv=None):
    p = argparse.ArgumentParser(
        description="Static brand compliance checker for fill_*.py outputs.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--pptx", action="append", required=True,
                   help="path to pptx (repeatable, paired with --skill)")
    p.add_argument("--skill", action="append", required=True,
                   help="skill id (repeatable, must match --pptx count). "
                        "e.g. customer-profile-pptx / market-environment-pptx / company-history-pptx")
    p.add_argument("--brand", required=True,
                   choices=["roleup", "stellar_aiz"],
                   help="brand id")
    p.add_argument("--format", choices=["text", "json"], default="text",
                   help="output format (default: text)")
    return p.parse_args(argv)


def check_one(pptx_path: str, skill_id: str, brand: str) -> dict:
    theme = load_theme(brand)
    ctx = CheckContext(pptx_path=pptx_path, skill_id=skill_id, brand=brand, theme=theme)
    prs = Presentation(pptx_path)
    results = run_profile(prs, ctx)
    return {
        "pptx": pptx_path,
        "skill": skill_id,
        "brand": brand,
        "results": [
            {
                "rule_id": r.rule_id,
                "passed": r.passed,
                "severity": r.severity,
                "message": r.message,
                "details": r.details,
            }
            for r in results
        ],
    }


def format_text(report: dict) -> str:
    lines = []
    lines.append(f"━━━ {report['pptx']} ━━━")
    lines.append(f"  skill={report['skill']}, brand={report['brand']}")
    for r in report["results"]:
        if r["passed"]:
            mark = "✅"
        elif r["severity"] == "warning":
            mark = "⚠️ "
        else:
            mark = "❌"
        lines.append(f"  {mark} [{r['rule_id']}] {r['message']}")
        if not r["passed"] and r["details"]:
            details_str = json.dumps(r["details"], ensure_ascii=False, indent=4)
            indented = "\n".join("    " + ln for ln in details_str.splitlines())
            lines.append(indented)
    return "\n".join(lines)


def has_failures(report: dict) -> bool:
    """True if any rule failed with severity='error'."""
    return any(
        not r["passed"] and r["severity"] == "error"
        for r in report["results"]
    )


def main(argv=None):
    args = parse_args(argv)
    if len(args.pptx) != len(args.skill):
        print(f"ERROR: --pptx count ({len(args.pptx)}) != --skill count ({len(args.skill)})",
              file=sys.stderr)
        sys.exit(2)

    reports = []
    any_failure = False
    for pptx_path, skill_id in zip(args.pptx, args.skill):
        try:
            report = check_one(pptx_path, skill_id, args.brand)
        except Exception as e:
            report = {
                "pptx": pptx_path,
                "skill": skill_id,
                "brand": args.brand,
                "error": f"{type(e).__name__}: {e}",
                "results": [],
            }
            any_failure = True
        reports.append(report)
        if has_failures(report):
            any_failure = True

    if args.format == "json":
        print(json.dumps(reports, ensure_ascii=False, indent=2))
    else:
        for r in reports:
            if "error" in r:
                print(f"━━━ {r['pptx']} ━━━")
                print(f"  ❌ EXCEPTION: {r['error']}")
            else:
                print(format_text(r))
            print()
        # 集計
        n_total = sum(len(r["results"]) for r in reports if "results" in r)
        n_pass = sum(
            1 for r in reports for x in r.get("results", [])
            if x["passed"]
        )
        n_fail = sum(
            1 for r in reports for x in r.get("results", [])
            if not x["passed"] and x["severity"] == "error"
        )
        n_warn = sum(
            1 for r in reports for x in r.get("results", [])
            if not x["passed"] and x["severity"] == "warning"
        )
        print(f"━━━ summary ━━━")
        print(f"  pptx: {len(reports)} / total checks: {n_total} / "
              f"pass: {n_pass} / fail: {n_fail} / warn: {n_warn}")

    sys.exit(1 if any_failure else 0)


if __name__ == "__main__":
    main()
