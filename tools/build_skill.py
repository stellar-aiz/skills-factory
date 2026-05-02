#!/usr/bin/env python3
"""Build skills from canonical source for claude_code / claude_ai targets.

Canonical skills live under ``skills/<name>/`` or ``skills/<category>/<name>/``
with ``{{VAR}}`` placeholders and optional ``@if:<profile>`` blocks. Profiles
are loaded from ``profiles/*.json``.

Commands:
  list
  check             <skill> [--profile <p>] [--strict]
  build             <skill> --profile <p> --out <dir>
  install           <skill> [--profile claude_code]
  package           <skill> [--profile claude_ai]
  install-all
  package-all       [--profile claude_ai] [--strict]
  install-category  <category> [--profile claude_code] [--strict]
  package-category  <category> [--profile claude_ai] [--strict]
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SKILLS_DIR = REPO / "skills"
PROFILES_DIR = REPO / "profiles"
DIST_DIR = REPO / "dist"
CLAUDE_CODE_SKILLS = Path.home() / ".claude" / "skills"

TEXT_EXTS = {".md", ".py", ".mjs", ".js", ".sh", ".json", ".txt", ".yaml", ".yml"}
VAR_RE = re.compile(r"\{\{([A-Z_][A-Z0-9_]*)\}\}")
IF_RE = re.compile(r"^\s*(?:<!--\s*|#\s*)@if:([a-z_][a-z0-9_]*)\s*(?:-->)?\s*$")
ENDIF_RE = re.compile(r"^\s*(?:<!--\s*|#\s*)@endif\s*(?:-->)?\s*$")


AUTO_RESOLVERS = {
    "FACTORY_ROOT": lambda: str(REPO),
}


def load_profile(name: str, skill_name: str) -> dict:
    path = PROFILES_DIR / f"{name}.json"
    if not path.exists():
        raise SystemExit(f"profile not found: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    vars_ = dict(data["vars"])
    vars_["SKILL_NAME"] = skill_name
    for k, v in list(vars_.items()):
        if v == "AUTO":
            if k not in AUTO_RESOLVERS:
                raise SystemExit(f"profile {name}: AUTO not supported for {k}")
            vars_[k] = AUTO_RESOLVERS[k]()
    for _ in range(3):
        for k, v in list(vars_.items()):
            if isinstance(v, str):
                vars_[k] = VAR_RE.sub(
                    lambda m: vars_.get(m.group(1), m.group(0)), v
                )
    return {"name": data["name"], "vars": vars_}


def filter_sections(text: str, profile: str, *, path: Path) -> str:
    out: list[str] = []
    stack: list[tuple[str, int]] = []
    for i, line in enumerate(text.splitlines(keepends=True), 1):
        m_if = IF_RE.match(line)
        if m_if:
            stack.append((m_if.group(1), i))
            continue
        if ENDIF_RE.match(line):
            if not stack:
                raise SystemExit(f"{path}:{i}: @endif without matching @if")
            stack.pop()
            continue
        if all(target == profile for target, _ in stack):
            out.append(line)
    if stack:
        target, line_no = stack[-1]
        raise SystemExit(f"{path}:{line_no}: unclosed @if:{target}")
    return "".join(out)


def substitute_vars(
    text: str, vars_: dict, *, path: Path, missing: list[tuple[Path, str]]
) -> str:
    def repl(m: re.Match) -> str:
        key = m.group(1)
        if key in vars_:
            return vars_[key]
        missing.append((path, key))
        return m.group(0)

    return VAR_RE.sub(repl, text)


def render_file(
    src: Path,
    dst: Path,
    profile: dict,
    *,
    missing: list[tuple[Path, str]],
) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    if src.suffix.lower() not in TEXT_EXTS:
        shutil.copy2(src, dst)
        return
    text = src.read_text(encoding="utf-8")
    text = filter_sections(text, profile["name"], path=src)
    text = substitute_vars(text, profile["vars"], path=src, missing=missing)
    dst.write_text(text, encoding="utf-8")


def iter_skill_files(src_root: Path):
    for src in src_root.rglob("*"):
        if src.is_dir():
            continue
        if src.name == ".DS_Store":
            continue
        yield src


def is_skill_root(d: Path) -> bool:
    """Return True if ``d`` is the shallowest directory containing a SKILL.md.

    This prevents false matches from nested SKILL.md (e.g. inside a skill's
    references/ or examples/ subdirectory).
    """
    if not (d / "SKILL.md").exists():
        return False
    p = d.parent
    while p != SKILLS_DIR and p != p.parent:
        if (p / "SKILL.md").exists():
            return False
        p = p.parent
    return True


def _all_skill_root_dirs() -> list[Path]:
    """Return all skill root directories (may include duplicate basenames)."""
    if not SKILLS_DIR.exists():
        return []
    out: list[Path] = []
    for skill_md in sorted(SKILLS_DIR.rglob("SKILL.md")):
        d = skill_md.parent
        if is_skill_root(d):
            out.append(d)
    return out


def iter_skill_dirs():
    """Yield unique skill root directories under SKILLS_DIR.

    Raises SystemExit if two skills share the same basename (name collision).
    """
    by_name: dict[str, list[Path]] = {}
    for d in _all_skill_root_dirs():
        by_name.setdefault(d.name, []).append(d)
    collisions = {n: ds for n, ds in by_name.items() if len(ds) > 1}
    if collisions:
        lines = []
        for name, dirs in sorted(collisions.items()):
            paths = ", ".join(str(d.relative_to(SKILLS_DIR)) for d in dirs)
            lines.append(f"  {name}: {paths}")
        raise SystemExit("skill name collisions:\n" + "\n".join(lines))
    for name in sorted(by_name):
        yield by_name[name][0]


def find_skill_dir(skill_name: str) -> Path:
    """Resolve a skill name to its source directory (flat or nested)."""
    for d in iter_skill_dirs():
        if d.name == skill_name:
            return d
    raise SystemExit(f"skill not found: {skill_name}")


def iter_category_skills(category: str):
    """Yield skill root directories under ``skills/<category>/``."""
    cat_dir = SKILLS_DIR / category
    if not cat_dir.is_dir():
        raise SystemExit(f"category not found: {cat_dir}")
    found_any = False
    for d in iter_skill_dirs():
        if cat_dir in d.parents:
            found_any = True
            yield d
    if not found_any:
        raise SystemExit(f"no skills found under category: {category}")


def build(
    skill: str,
    profile_name: str,
    out_dir: Path,
    *,
    strict: bool,
) -> list[tuple[Path, str]]:
    src_root = find_skill_dir(skill)
    profile = load_profile(profile_name, skill)
    if out_dir.exists():
        shutil.rmtree(out_dir)
    missing: list[tuple[Path, str]] = []
    for src in iter_skill_files(src_root):
        rel = src.relative_to(src_root)
        render_file(src, out_dir / rel, profile, missing=missing)
    if missing:
        report = "\n".join(
            f"  {p.relative_to(src_root)}: {{{{{k}}}}}" for p, k in missing
        )
        msg = f"unresolved variables:\n{report}"
        if strict:
            shutil.rmtree(out_dir, ignore_errors=True)
            raise SystemExit(msg)
        print(f"warning: {msg}", file=sys.stderr)
    return missing


def cmd_list(_: argparse.Namespace) -> None:
    for d in iter_skill_dirs():
        rel = d.relative_to(SKILLS_DIR)
        if rel.parent == Path("."):
            print(d.name)
        else:
            print(f"{rel}  ({d.name})")


def cmd_check(args: argparse.Namespace) -> None:
    profiles = [args.profile] if args.profile else ["claude_code", "claude_ai"]
    for profile_name in profiles:
        out_dir = DIST_DIR / f"_check_{args.skill}_{profile_name}"
        try:
            missing = build(args.skill, profile_name, out_dir, strict=args.strict)
        finally:
            shutil.rmtree(out_dir, ignore_errors=True)
        status = "ok" if not missing else f"{len(missing)} unresolved"
        print(f"[{profile_name}] {args.skill}: {status}")


def cmd_build(args: argparse.Namespace) -> None:
    out_dir = Path(args.out).expanduser().resolve()
    build(args.skill, args.profile, out_dir, strict=args.strict)
    print(f"built -> {out_dir}")


def _sync_common_assets() -> None:
    """Copy skills/_common/{lib,brands}/ to ~/.claude/skills/_common/.

    Brand-aware fill_*.py scripts import from `_common/lib/brand_resolver.py`
    and read `_common/brands/<id>/theme.json` at runtime. Per-skill install
    alone does not bring `_common/` along, so we sync those subtrees here.
    Other `_common/` subdirs (prompts/, references/, styles/) are
    documentation/source-only and not synced.
    """
    src_common = SKILLS_DIR / "_common"
    if not src_common.exists():
        return
    dst_common = CLAUDE_CODE_SKILLS / "_common"
    for sub in ("lib", "brands"):
        src_sub = src_common / sub
        if not src_sub.exists():
            continue
        dst_sub = dst_common / sub
        if dst_sub.exists():
            shutil.rmtree(dst_sub)
        shutil.copytree(src_sub, dst_sub)


def cmd_install(args: argparse.Namespace) -> None:
    target = CLAUDE_CODE_SKILLS / args.skill
    target.parent.mkdir(parents=True, exist_ok=True)
    build(args.skill, args.profile, target, strict=args.strict)
    _sync_common_assets()
    print(f"installed -> {target}")


def cmd_package(args: argparse.Namespace) -> None:
    DIST_DIR.mkdir(exist_ok=True)
    stage = DIST_DIR / f"_stage_{args.skill}"
    build(args.skill, args.profile, stage, strict=args.strict)
    zip_path = DIST_DIR / f"{args.skill}.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in sorted(stage.rglob("*")):
            if p.is_file():
                zf.write(p, p.relative_to(stage))
    shutil.rmtree(stage)
    print(f"packaged -> {zip_path}")


def cmd_install_all(args: argparse.Namespace) -> None:
    for d in iter_skill_dirs():
        ns = argparse.Namespace(
            skill=d.name, profile=args.profile, strict=args.strict
        )
        cmd_install(ns)


def cmd_package_all(args: argparse.Namespace) -> None:
    for d in iter_skill_dirs():
        ns = argparse.Namespace(
            skill=d.name, profile=args.profile, strict=args.strict
        )
        cmd_package(ns)


def cmd_install_category(args: argparse.Namespace) -> None:
    for d in iter_category_skills(args.category):
        ns = argparse.Namespace(
            skill=d.name, profile=args.profile, strict=args.strict
        )
        cmd_install(ns)


def cmd_package_category(args: argparse.Namespace) -> None:
    for d in iter_category_skills(args.category):
        ns = argparse.Namespace(
            skill=d.name, profile=args.profile, strict=args.strict
        )
        cmd_package(ns)


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(prog="build_skill.py")
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("list").set_defaults(func=cmd_list)

    p_check = sub.add_parser("check")
    p_check.add_argument("skill")
    p_check.add_argument("--profile")
    p_check.add_argument("--strict", action="store_true")
    p_check.set_defaults(func=cmd_check)

    p_build = sub.add_parser("build")
    p_build.add_argument("skill")
    p_build.add_argument("--profile", required=True)
    p_build.add_argument("--out", required=True)
    p_build.add_argument("--strict", action="store_true")
    p_build.set_defaults(func=cmd_build)

    p_install = sub.add_parser("install")
    p_install.add_argument("skill")
    p_install.add_argument("--profile", default="claude_code")
    p_install.add_argument("--strict", action="store_true")
    p_install.set_defaults(func=cmd_install)

    p_package = sub.add_parser("package")
    p_package.add_argument("skill")
    p_package.add_argument("--profile", default="claude_ai")
    p_package.add_argument("--strict", action="store_true")
    p_package.set_defaults(func=cmd_package)

    p_all = sub.add_parser("install-all")
    p_all.add_argument("--profile", default="claude_code")
    p_all.add_argument("--strict", action="store_true")
    p_all.set_defaults(func=cmd_install_all)

    p_pkg_all = sub.add_parser("package-all")
    p_pkg_all.add_argument("--profile", default="claude_ai")
    p_pkg_all.add_argument("--strict", action="store_true")
    p_pkg_all.set_defaults(func=cmd_package_all)

    p_cat = sub.add_parser("install-category")
    p_cat.add_argument("category", help="e.g. bdd")
    p_cat.add_argument("--profile", default="claude_code")
    p_cat.add_argument("--strict", action="store_true")
    p_cat.set_defaults(func=cmd_install_category)

    p_pkg_cat = sub.add_parser("package-category")
    p_pkg_cat.add_argument("category", help="e.g. bdd")
    p_pkg_cat.add_argument("--profile", default="claude_ai")
    p_pkg_cat.add_argument("--strict", action="store_true")
    p_pkg_cat.set_defaults(func=cmd_package_category)

    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
