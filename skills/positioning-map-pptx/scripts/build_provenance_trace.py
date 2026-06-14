"""build_provenance_trace.py — ポジショニングマップの出典・根拠トレース

設計思想（3層モデル）:
  処理は 3 層に分かれる。WEB を引くのは ① と ②.5 の LLM 作業のみ。本スクリプト自体は WEB を引かない。
    ① 調査・記録層（LLM / 将来は research-subagent）:
       Web を読み x/y/軸を判断し、**取得・判断したその場で** provenance を記録する。
       硬い数値（売上・シェア・海外比率 等）は actual_value/actual_unit/actual_metric に構造化する。
    ② レンダリング・検証層（本スクリプト・決定論的・Web なし）:
       ①が記録した provenance を検証し、人が読める表に起こすだけ。新たに作らない。
    ②.5 WEB 検証層（verify-prep[本スクリプト] → LLM が記録URLを fetch 照合 → render --verification[本スクリプト]）:
       verify-prep が「裏取りすべき硬い事実」の worklist を決定論的に抽出し、LLM が記録 source URL を
       WebFetch して値を照合（不足/到達不可なら WebSearch にフォールバック）、結果を render が表へ統合する。

  本スクリプトは fill_positioning_map.py（=data→PPTX レンダラー）とは完全に独立。
  fill は provenance を一切知らない。トレースは必ず本スクリプトを別途実行する。

provenance サイドカー `positioning_map_provenance.json`:
  json_path をキーに {value, actual_value, actual_unit, actual_metric, rationale, source_name, source, confidence}。
    rationale     : 根拠（必須・空欄NG）
    source_name   : 人が読む出典名（必須・空欄NG）例「Tesla IR 2024 決算説明資料」
    source        : 実際に辿れるロケータ＝URL or 社内資料ファイル名（必須・空欄NG）
    confidence    : 実測 / 推定 / 仮置（必須・空欄NG）
    value         : 任意（本体値との drift 検知用）。size は正規化相対値なので検証には使えない点に注意
    actual_value  : 任意（検証対象の生数値。例 "3382"）。WEB検証の worklist 採録キー
    actual_unit   : 任意（生数値の単位。例 "億円" / "%"）
    actual_metric : 任意（生数値が何の指標か。例 "連結売上(2024年12月期)" / "海外売上比率"）

対象 path（事実・配置系のみ）:
    $.players[i].x / $.players[i].y / $.players[i].size（本体に size があるときのみ）
    $.x_axis / $.y_axis（軸選択の根拠）

3 モード:
  --mode skeleton    : data から対象 path を列挙し空欄の provenance 雛形を生成（追補・--force で全上書き）
  --mode verify-prep : provenance から「confidence∈{実測,推定} かつ actual_value 非空」の硬い事実を抽出し
                       WEB検証 worklist を生成（追補・--force で全上書き）。LLM が検証欄を埋める。
  --mode render      : data + provenance を検証し Markdown トレース表を生成。--verification を渡すと
                       WEB検証結果も統合する。空欄/未記載/裏取り問題は ⚠NG として表示し非ゼロ終了する

Usage:
  python build_provenance_trace.py --mode skeleton \
    --data positioning_map_data.json --out positioning_map_provenance.json

  python build_provenance_trace.py --mode verify-prep \
    --data positioning_map_data.json --provenance positioning_map_provenance.json \
    --out positioning_map_verification.json

  python build_provenance_trace.py --mode render \
    --data positioning_map_data.json --provenance positioning_map_provenance.json \
    --verification positioning_map_verification.json \
    --out-md PositioningMap_output_provenance.md
"""

import argparse
import datetime
import json
import os
import sys

CONFIDENCE_VOCAB = ("実測", "推定", "仮置")
REQUIRED_FIELDS = ("rationale", "source_name", "source", "confidence")

# 硬い数値の構造化フィールド（任意・空欄でも render の NG にしない）
OPTIONAL_FACT_FIELDS = ("actual_value", "actual_unit", "actual_metric")

# ②.5 WEB 検証層
VERIFY_CONFIDENCE = ("実測", "推定")  # worklist に載せる confidence（仮置＝純定性は対象外）
VERIFY_VOCAB = (
    "confirmed",          # 記録URL（または再検索ソース）が値/根拠を裏付け
    "source_mismatch",    # 記録URLに到達したが値を裏付けない（誤帰属・捏造の疑い）
    "source_unreachable", # 記録URLにアクセス不可（認証/404/クロスホストリダイレクト）→ 再検索フォールバックへ
    "discrepancy",        # 信頼できるソースと値が食い違う
    "not_found",          # どのソースにも該当情報が見つからない
    "stale",              # 情報は存在するが、より新しい数値がある
)
PROBLEM_VERDICTS = ("source_mismatch", "discrepancy", "not_found")  # 非ゼロ終了の対象
VERIFY_RESULT_FIELDS = ("verification_result", "verified_value", "verification_note", "method")


# ──────────────────────────────────────────────
# 対象 path 算出（両モード共通）
# ──────────────────────────────────────────────
def covered_paths(data):
    """data から provenance を付けるべき対象 path を順序付きで返す。

    順序: 各 player（x, y, size?）→ x_axis → y_axis（= 社・軸の順）。
    返り値は (path, kind, body_value, label) のタプルのリスト。
      kind  : "player" | "axis"
      body_value : drift 検知用の本体値（軸は None）
      label : 表示補助（player 名 / 軸ラベル）
    """
    rows = []
    players = data.get("players", [])
    for i, p in enumerate(players):
        name = p.get("name", f"player {i}")
        rows.append((f"$.players[{i}].x", "player", p.get("x"), name))
        rows.append((f"$.players[{i}].y", "player", p.get("y"), name))
        if "size" in p:
            rows.append((f"$.players[{i}].size", "player", p.get("size"), name))
    x_axis = data.get("x_axis", {})
    y_axis = data.get("y_axis", {})
    rows.append(("$.x_axis", "axis", None, x_axis.get("label", "X軸")))
    rows.append(("$.y_axis", "axis", None, y_axis.get("label", "Y軸")))
    return rows


def _load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ──────────────────────────────────────────────
# skeleton モード
# ──────────────────────────────────────────────
def build_skeleton(data, existing=None, force=False):
    """対象 path の空欄 provenance 雛形を返す（既存は温存して追補）。

    既存エントリにも欠けている actual_* フィールドがあれば空欄で補う（手書き分は温存）。
    """
    out = {} if (force or existing is None) else dict(existing)
    added = 0
    for path, kind, body_value, _label in covered_paths(data):
        if path in out:
            # 既存（手書き分）は温存しつつ、actual_* が欠けていれば空欄で追補
            entry = out[path]
            if isinstance(entry, dict):
                for fld in OPTIONAL_FACT_FIELDS:
                    entry.setdefault(fld, "")
            continue
        entry = {}
        if kind == "player" and body_value is not None:
            entry["value"] = body_value  # 本体値を prefill（drift 検知用）
        # 硬い数値があればここに構造化（任意・空欄でも render の NG にしない）
        for fld in OPTIONAL_FACT_FIELDS:
            entry[fld] = ""
        entry["rationale"] = ""
        entry["source_name"] = ""
        entry["source"] = ""
        entry["confidence"] = ""
        out[path] = entry
        added += 1
    return out, added


def run_skeleton(args):
    data = _load_json(args.data)
    existing = None
    if args.out and os.path.exists(args.out) and not args.force:
        try:
            existing = _load_json(args.out)
        except Exception:
            existing = None
    skeleton, added = build_skeleton(data, existing=existing, force=args.force)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(skeleton, f, ensure_ascii=False, indent=2)
        f.write("\n")

    total = len(covered_paths(data))
    mode_note = "全上書き" if args.force else ("追補" if existing is not None else "新規")
    print(f"✅ skeleton ({mode_note}): {args.out}")
    print(f"  対象 path {total} 件 / 今回追加 {added} 件 / 合計 {len(skeleton)} エントリ")
    if added:
        print(f"  → rationale / source_name / source / confidence を埋めてください（空欄は render で NG）",
              file=sys.stderr)
        print(f"  → 硬い数値（売上・シェア・海外比率 等）は actual_value/actual_unit/actual_metric に構造化すると"
              f" verify-prep で WEB 裏取りできます", file=sys.stderr)
    return 0


# ──────────────────────────────────────────────
# verify-prep モード（②.5 WEB 検証の worklist 抽出・決定論的）
# ──────────────────────────────────────────────
def _nonempty(v):
    return bool(str(v).strip()) if v is not None else False


def build_verify_worklist(data, prov, existing=None, force=False):
    """provenance から「confidence∈{実測,推定} かつ actual_value 非空」の硬い事実を抽出し、
    WEB 検証 worklist を返す（既存の検証欄は温存して追補）。

    返り値: (worklist:dict, added:int, skipped_real_no_fact:list)
      skipped_real_no_fact : confidence=実測 だが actual_value 空（要注意）の path
    """
    out = {} if (force or existing is None) else dict(existing)
    added = 0
    skipped_real_no_fact = []
    for path, _kind, _body_value, _label in covered_paths(data):
        entry = prov.get(path)
        if not isinstance(entry, dict):
            continue
        confidence = str(entry.get("confidence", "")).strip()
        actual_value = entry.get("actual_value")
        if confidence not in VERIFY_CONFIDENCE:
            continue  # 仮置・空欄は対象外
        if not _nonempty(actual_value):
            if confidence == "実測":
                skipped_real_no_fact.append(path)
            continue  # 裏取りすべき生数値が構造化されていない
        if path in out:
            continue  # 既存（LLM 記入済みの検証欄）は温存
        out[path] = {
            "actual_value": actual_value,
            "actual_unit": entry.get("actual_unit", ""),
            "actual_metric": entry.get("actual_metric", ""),
            "rationale": entry.get("rationale", ""),
            "source_name": entry.get("source_name", ""),
            "source": entry.get("source", ""),
            "confidence": confidence,
            # ↓ LLM が記録URLを fetch 照合して埋める
            "verification_result": "",
            "verified_value": "",
            "verification_note": "",
            "method": "",
        }
        added += 1
    return out, added, skipped_real_no_fact


def run_verify_prep(args):
    data = _load_json(args.data)
    if not args.provenance:
        print("✗ --mode verify-prep には --provenance が必須です", file=sys.stderr)
        return 2
    prov = _load_json(args.provenance)

    existing = None
    if args.out and os.path.exists(args.out) and not args.force:
        try:
            existing = _load_json(args.out)
        except Exception:
            existing = None
    worklist, added, skipped = build_verify_worklist(data, prov, existing=existing, force=args.force)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(worklist, f, ensure_ascii=False, indent=2)
        f.write("\n")

    mode_note = "全上書き" if args.force else ("追補" if existing is not None else "新規")
    print(f"✅ verify-prep worklist ({mode_note}): {args.out}")
    print(f"  裏取り対象 {len(worklist)} 件 / 今回追加 {added} 件"
          f"（confidence∈{'/'.join(VERIFY_CONFIDENCE)} かつ actual_value 非空）")
    if added:
        print(f"  → 各エントリの source URL を WebFetch して actual_value を照合し、"
              f"verification_result（{'/'.join(VERIFY_VOCAB)}）/ verified_value / method / "
              f"verification_note を埋めてください。", file=sys.stderr)
        print(f"  → 記録URLが裏付けない/到達不可なら WebSearch にフォールバックすること。", file=sys.stderr)
    if skipped:
        print(f"  ⚠ WARNING: confidence=実測 なのに actual_value が空の path が {len(skipped)} 件"
              f"（裏取りできないので生数値を構造化してください）:", file=sys.stderr)
        for p in skipped:
            print(f"      - {p}", file=sys.stderr)
    return 0


# ──────────────────────────────────────────────
# render モード
# ──────────────────────────────────────────────
def _is_url(s):
    return isinstance(s, str) and (s.startswith("http://") or s.startswith("https://"))


def _md_cell(s):
    """Markdown セル用エスケープ（| と改行を無害化）。"""
    if s is None:
        return ""
    return str(s).replace("|", "\\|").replace("\n", " ").replace("\r", " ").strip()


def _values_differ(a, b):
    """drift 検知。数値は float 比較、それ以外は文字列比較。"""
    if a is None or b is None:
        return False
    try:
        return abs(float(a) - float(b)) > 1e-9
    except (TypeError, ValueError):
        return str(a) != str(b)


def _build_verification_section(data, prov, verification):
    """--verification 指定時の「WEB検証結果」節を組む。

    返り値: (lines:list[str], report:dict)
      report = {
        "ng": [path,...],            # verification_result 空（要記入）
        "vocab_warn": [(path,val)],  # VERIFY_VOCAB 外
        "problems": [(path,verdict)],# PROBLEM_VERDICTS（非ゼロ終了対象）
        "real_no_fact": [path,...],  # confidence=実測 だが actual_value 空
        "checked": int,
      }
    """
    report = {"ng": [], "vocab_warn": [], "problems": [], "real_no_fact": [], "checked": 0}
    lines = ["## WEB検証結果（記録URL照合）", ""]

    # confidence=実測 だが actual_value 空 → 裏取り不能の注意（worklist にも載らない）
    for path, _kind, _bv, _label in covered_paths(data):
        e = prov.get(path)
        if isinstance(e, dict) and str(e.get("confidence", "")).strip() == "実測" \
                and not _nonempty(e.get("actual_value")):
            report["real_no_fact"].append(path)

    if not verification:
        lines.append("検証 worklist が空です（confidence∈{実測,推定} かつ actual_value 非空のエントリがありません）。")
        lines.append("")
        return lines, report

    table = [
        "| 項目 (json_path) | actual_value | 指標 | 検証結果 | verified_value | 出典が裏付けたか | 方式 | 備考 |",
        "|---|---|---|---|---|---|---|---|",
    ]
    # covered_paths の順序を維持しつつ worklist にあるものを出す
    ordered = [p for p, _k, _b, _l in covered_paths(data) if p in verification]
    for path in ordered:
        v = verification.get(path, {})
        report["checked"] += 1
        actual = v.get("actual_value", "")
        unit = v.get("actual_unit", "")
        metric = v.get("actual_metric", "")
        result = str(v.get("verification_result", "")).strip()
        verified = v.get("verified_value", "")
        method = v.get("method", "")
        note = v.get("verification_note", "")

        actual_disp = f"{actual} {unit}".strip() if _nonempty(unit) else actual

        if not result:
            report["ng"].append(path)
            backed = "⚠NG(未検証)"
        else:
            if result not in VERIFY_VOCAB:
                report["vocab_warn"].append((path, result))
            if result in PROBLEM_VERDICTS:
                report["problems"].append((path, result))
                backed = f"✗ {result}"
            elif result == "confirmed":
                backed = "✓ confirmed"
            elif result == "source_unreachable":
                backed = "△ source_unreachable"
            elif result == "stale":
                backed = "△ stale"
            else:
                backed = result

        table.append(
            f"| {_md_cell(path)} | {_md_cell(actual)} | {_md_cell(metric)} | "
            f"{_md_cell(result)} | {_md_cell(verified)} | {backed} | {_md_cell(method)} | {_md_cell(note)} |"
        )
    lines.extend(table)
    lines.append("")

    if report["problems"]:
        lines.append("### ✗ 裏取り問題（非ゼロ終了）")
        lines.append("")
        for path, verdict in report["problems"]:
            lines.append(f"- `{path}` — {verdict}")
        lines.append("")
    if report["real_no_fact"]:
        lines.append("### ⚠ confidence=実測 だが actual_value 未記入（裏取り不能）")
        lines.append("")
        for path in report["real_no_fact"]:
            lines.append(f"- `{path}` — 生数値を actual_value/actual_unit/actual_metric に構造化してください")
        lines.append("")

    return lines, report


def render_trace(data, prov, date_str, verification=None):
    """検証しつつ Markdown トレースを組む。

    返り値: (markdown:str, ng_list:list, drift_list:list, vocab_warn:list, verify_report:dict|None)
    """
    rows = covered_paths(data)

    ng_list = []      # (path, [empty_field, ...] or "未記載")
    drift_list = []   # (path, prov_value, body_value)
    vocab_warn = []   # (path, confidence)

    table_lines = [
        "| 項目 (json_path) | 値 | 根拠 | 出典名 | 出典 (ロケータ) | 自信度 | 状態 |",
        "|---|---|---|---|---|---|---|",
    ]

    for path, kind, body_value, label in rows:
        entry = prov.get(path)
        # 値列: player は本体値、axis は軸ラベル
        value_disp = body_value if kind == "player" else label

        if entry is None:
            ng_list.append((path, "エントリ未記載"))
            table_lines.append(
                f"| {_md_cell(path)} | {_md_cell(value_disp)} |  |  |  |  | ⚠NG(未記載) |"
            )
            continue

        rationale = entry.get("rationale", "") or ""
        source_name = entry.get("source_name", "") or ""
        source = entry.get("source", "") or ""
        confidence = entry.get("confidence", "") or ""

        # 空欄NG検査
        empties = [fld for fld in REQUIRED_FIELDS if not str(entry.get(fld, "")).strip()]
        # drift 検査（player のみ）
        is_drift = kind == "player" and "value" in entry and _values_differ(
            entry.get("value"), body_value
        )
        if is_drift:
            drift_list.append((path, entry.get("value"), body_value))
        # 語彙検査
        if confidence and confidence not in CONFIDENCE_VOCAB:
            vocab_warn.append((path, confidence))

        # 状態
        if empties:
            status = f"⚠NG(空欄: {'/'.join(empties)})"
            ng_list.append((path, empties))
        elif is_drift:
            status = "⚠値ズレ"
        else:
            status = "✓"

        # 出典名セル: source が URL ならリンク化、ファイル名ならプレーン
        if source_name and _is_url(source):
            name_cell = f"[{_md_cell(source_name)}]({source})"
        else:
            name_cell = _md_cell(source_name)

        table_lines.append(
            f"| {_md_cell(path)} | {_md_cell(value_disp)} | {_md_cell(rationale)} | "
            f"{name_cell} | {_md_cell(source)} | {_md_cell(confidence)} | {status} |"
        )

    # ── Markdown 組み立て ──
    market = data.get("chart_title") or data.get("section_title") or "ポジショニングマップ"
    x_label = data.get("x_axis", {}).get("label", "")
    y_label = data.get("y_axis", {}).get("label", "")
    target = data.get("target_company", "")

    out = []
    out.append(f"# 出典・根拠トレース — {_md_cell(market)}")
    out.append("")
    out.append(f"- 採用2軸: **X = {_md_cell(x_label)}** × **Y = {_md_cell(y_label)}**")
    if target:
        out.append(f"- 対象会社: **{_md_cell(target)}**")
    out.append(f"- プレイヤー数: {len(data.get('players', []))}")
    out.append(f"- 生成日: {date_str}")
    out.append("")
    out.append("## 項目別トレース")
    out.append("")
    out.extend(table_lines)
    out.append("")

    # NG / ギャップ節
    out.append("## データギャップ / NG")
    out.append("")
    if ng_list:
        out.append("以下の対象 path は provenance が未記載または空欄です（**NG・要記入**）:")
        out.append("")
        for path, detail in ng_list:
            if isinstance(detail, list):
                out.append(f"- `{path}` — 空欄: {', '.join(detail)}")
            else:
                out.append(f"- `{path}` — {detail}")
    else:
        out.append("なし（全対象 path に rationale / source_name / source / confidence が記入済み）。")
    out.append("")

    if drift_list:
        out.append("### ⚠ 値ズレ（provenance.value ≠ 本体値）")
        out.append("")
        for path, pv, bv in drift_list:
            out.append(f"- `{path}` — provenance: {pv} / 本体: {bv}")
        out.append("")

    if vocab_warn:
        out.append("### ⚠ 自信度の語彙逸脱（実測/推定/仮置 以外）")
        out.append("")
        for path, conf in vocab_warn:
            out.append(f"- `{path}` — \"{conf}\"")
        out.append("")

    # WEB検証結果（--verification 指定時のみ）
    verify_report = None
    if verification is not None:
        v_lines, verify_report = _build_verification_section(data, prov, verification)
        out.extend(v_lines)

    # 位置づけ注記
    out.append("## 本表の位置づけ")
    out.append("")
    out.append("- x/y の**定性配置値そのもの**（x=9 等）は事実ではなく見立てであり、検証対象外。")
    out.append("  本表は配置の**出典・根拠のトレース**である。")
    out.append("- 一方、配置の根拠としてLLMが引いた**硬い事実**（売上・シェア・海外比率 等の actual_value）は、")
    out.append("  `verify-prep` → 記録URLの WebFetch 照合（→ render --verification）で**WEB裏取りできる**（本スキル内蔵）。")
    out.append("- より深いクロスソース検証（複数の独立ソースでの突き合わせ）が必要な場合のみ、")
    out.append("  `fact-check-reviewer` を併用すること。")
    out.append("")

    return "\n".join(out), ng_list, drift_list, vocab_warn, verify_report


def run_render(args):
    data = _load_json(args.data)
    if not args.provenance:
        print("✗ --mode render には --provenance が必須です", file=sys.stderr)
        return 2
    prov = _load_json(args.provenance)

    verification = None
    if args.verification:
        verification = _load_json(args.verification)

    date_str = args.date or datetime.date.today().isoformat()
    markdown, ng_list, drift_list, vocab_warn, verify_report = render_trace(
        data, prov, date_str, verification=verification
    )

    os.makedirs(os.path.dirname(os.path.abspath(args.out_md)), exist_ok=True)
    with open(args.out_md, "w", encoding="utf-8") as f:
        f.write(markdown)
        if not markdown.endswith("\n"):
            f.write("\n")

    print(f"✅ provenance トレース: {args.out_md}")

    # stderr に WARN / NG
    for path, pv, bv in drift_list:
        print(f"  ⚠ WARNING: 値ズレ {path} — provenance={pv} / 本体={bv}", file=sys.stderr)
    for path, conf in vocab_warn:
        print(f"  ⚠ WARNING: 自信度の語彙逸脱 {path} — \"{conf}\""
              f"（{'/'.join(CONFIDENCE_VOCAB)} のいずれかにすること）", file=sys.stderr)

    rc = 0

    if ng_list:
        print(f"  ✗ NG: provenance が空欄/未記載の対象 path が {len(ng_list)} 件あります（空欄はNG）:",
              file=sys.stderr)
        for path, detail in ng_list:
            d = ", ".join(detail) if isinstance(detail, list) else detail
            print(f"      - {path} ({d})", file=sys.stderr)
        print("    → 表は出力済み。空欄を埋めて再実行してください。", file=sys.stderr)
        rc = 1

    # ②.5 WEB検証の結果（--verification 指定時のみ）
    if verify_report is not None:
        for path in verify_report["real_no_fact"]:
            print(f"  ⚠ WARNING: confidence=実測 だが actual_value 未記入 {path}"
                  f"（裏取り不能・生数値を構造化してください）", file=sys.stderr)
        for path, val in verify_report["vocab_warn"]:
            print(f"  ⚠ WARNING: verification_result の語彙逸脱 {path} — \"{val}\""
                  f"（{'/'.join(VERIFY_VOCAB)} のいずれかにすること）", file=sys.stderr)
        if verify_report["ng"]:
            print(f"  ✗ NG: 検証 worklist で verification_result が空の path が "
                  f"{len(verify_report['ng'])} 件あります:", file=sys.stderr)
            for path in verify_report["ng"]:
                print(f"      - {path}", file=sys.stderr)
            print("    → 記録URLを WebFetch して検証結果を埋め、再実行してください。", file=sys.stderr)
            rc = 1
        if verify_report["problems"]:
            print(f"  ✗ NG: WEB裏取りで問題のある path が {len(verify_report['problems'])} 件あります"
                  f"（{'/'.join(PROBLEM_VERDICTS)}）:", file=sys.stderr)
            for path, verdict in verify_report["problems"]:
                print(f"      - {path} ({verdict})", file=sys.stderr)
            print("    → 出典の誤帰属・数値の食い違いの疑い。値か出典を修正して再実行してください。",
                  file=sys.stderr)
            rc = 1

    return rc


# ──────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="ポジショニングマップの出典・根拠トレース生成")
    ap.add_argument("--mode", choices=["skeleton", "verify-prep", "render"], default="render")
    ap.add_argument("--data", required=True, help="positioning_map_data.json（本体・値のみ）")
    ap.add_argument("--provenance", help="positioning_map_provenance.json（verify-prep / render で必須）")
    ap.add_argument("--verification", help="positioning_map_verification.json（render で任意・WEB検証結果を統合）")
    ap.add_argument("--out", help="skeleton / verify-prep の出力先 JSON")
    ap.add_argument("--out-md", dest="out_md", help="render の出力先 Markdown")
    ap.add_argument("--date", help="生成日（省略時は本日。YYYY-MM-DD）")
    ap.add_argument("--force", action="store_true", help="skeleton / verify-prep で既存を温存せず全上書き")
    args = ap.parse_args()

    if args.mode == "skeleton":
        if not args.out:
            ap.error("--mode skeleton には --out が必須です")
        return run_skeleton(args)
    elif args.mode == "verify-prep":
        if not args.out:
            ap.error("--mode verify-prep には --out が必須です")
        return run_verify_prep(args)
    else:
        if not args.out_md:
            ap.error("--mode render には --out-md が必須です")
        return run_render(args)


if __name__ == "__main__":
    sys.exit(main())
