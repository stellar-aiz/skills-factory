"""fill_positioning_map.py — ポジショニングマップスライドを生成

v2 (テンプレ固定化リファクタ):
  固定できる骨組み（マップ外枠・十字ガイドライン・軸ラベル枠・象限ラベル枠4・示唆枠5・
  セクションタイトル・出典）は assets/<brand>/positioning-map-template.pptx に **名前付き
  シェイプ** として焼き込み済み（tools/build_positioning_map_template.py で生成）。

  本スクリプトの実行時の責務は 2 つだけ:
    (1) 名前付きシェイプへのテキスト流し込み（空データの枠は非表示）
    (2) バブル（円）+ バブルラベルの描画 ← これだけが本質的に可変なので実行時描画

  バブル配置の幾何基準は MAP_FRAME シェイプの .left/.top/.width/.height を single source
  of truth として読む（layout.json のマージン定数からの算術は廃止）。

Usage:
  python fill_positioning_map.py \
    --data /path/positioning_map_data.json \
    --brand stellar_aiz | roleup \
    --output /path/PositioningMap_output.pptx
"""

import argparse
import json
import os
import sys

# brand_resolver bootstrap (brand-aware: stellar_aiz / roleup)
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(SKILL_DIR, "..", "_common", "lib"))
from brand_resolver import resolve_brand, add_brand_arg  # noqa: E402
from format_helpers import resolve_top_text, resolve_subtitle_text, require_source  # noqa: E402
from validate_fill_input import validate_fill_input  # noqa: E402

SKILL_ID = "positioning-map-pptx"

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree


def _finalize_pptx(path):
    """LibreOffice roundtrip to normalize OOXML so PowerPoint stops asking for repair.

    No-op if soffice is unavailable or the conversion fails; the original file
    is preserved. Added by tools/add_finalize_hook.py.
    """
    import os, shutil, subprocess, tempfile, glob
    candidates = [
        os.environ.get("SOFFICE_BIN"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/opt/homebrew/bin/soffice",
        "/usr/local/bin/soffice",
        "/usr/bin/soffice",
        shutil.which("soffice"),
        shutil.which("libreoffice"),
    ]
    soffice = next((c for c in candidates if c and os.path.exists(c)), None)
    if not soffice:
        return
    try:
        with tempfile.TemporaryDirectory(prefix="pptx_rt_") as tmp:
            subprocess.run(
                [soffice, f"-env:UserInstallation=file://{tmp}/prof",
                 "--headless", "--convert-to", "pptx",
                 "--outdir", tmp, str(path)],
                timeout=120, capture_output=True, check=True,
            )
            found = glob.glob(os.path.join(tmp, "*.pptx"))
            if found:
                shutil.move(found[0], str(path))
    except Exception:
        pass


# ── 名前付きシェイプ ID（tools/build_positioning_map_template.py と一致させること） ──
SHAPE_MAIN_MESSAGE = "Title 1"
SHAPE_CHART_TITLE = "Text Placeholder 2"
SHAPE_SOURCE = "Source 3"
SHAPE_MAP_FRAME = "MAP_FRAME"

IMPL_SLOTS = 5  # 示唆は固定 5 枠（tools/build_positioning_map_template.py と一致）

# ── バブル描画用の配色・フォント（_apply_theme で reassign される defaults） ──
COLOR_TEXT = RGBColor(0x33, 0x33, 0x33)
COLOR_TARGET = RGBColor(0xE1, 0x57, 0x59)
COLOR_TARGET_LINE = RGBColor(0x8B, 0x2C, 0x2E)
COLOR_BUBBLE_LINE = RGBColor(0x33, 0x33, 0x33)

CHART_PALETTE = [
    "#4E79A7", "#F28E2B", "#59A14F", "#76B7B2",
    "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F",
]
OTHER_COLOR = "#BAB0AC"
TARGET_COLOR = "#E15759"

FONT_NAME_JP = "Meiryo UI"
FONT_SIZE_BUBBLE = Pt(12)
FONT_SIZE_BUBBLE_TARGET = Pt(13)

_THEME = None


def _palette_color(index: int, total: int) -> str:
    if total <= 1:
        return CHART_PALETTE[0]
    return CHART_PALETTE[index % len(CHART_PALETTE)]


def _apply_theme(theme):
    """バブル描画用の brand-aware globals を解決済み BrandTheme から再設定する。

    骨組みの幾何・配色はテンプレに焼き込み済みなので、ここで設定するのは
    実行時描画（バブル + バブルラベル）に必要な色・フォントだけ。
    """
    global _THEME, COLOR_TEXT, COLOR_TARGET, COLOR_TARGET_LINE, COLOR_BUBBLE_LINE
    global CHART_PALETTE, OTHER_COLOR, TARGET_COLOR
    global FONT_NAME_JP, FONT_SIZE_BUBBLE, FONT_SIZE_BUBBLE_TARGET

    _THEME = theme
    COLOR_TEXT = theme.color("text")
    COLOR_BUBBLE_LINE = theme.color("text")
    FONT_NAME_JP = theme.font_ea

    if theme.id == "roleup":
        COLOR_TARGET = theme.color("highlight_target")
        COLOR_TARGET_LINE = theme.color("label_bar")
        CHART_PALETTE = list(theme.chart_palette[:8])
        OTHER_COLOR = theme.hex("highlight_other")
        TARGET_COLOR = theme.hex("highlight_target")
        # roleup C4 許容: target も BUBBLE と同サイズ（+1pt を避け太字+色で強調）
        FONT_SIZE_BUBBLE = theme.pt("font_size_subtitle_pt")          # 12pt
        FONT_SIZE_BUBBLE_TARGET = theme.pt("font_size_subtitle_pt")   # 12pt
    # stella は V1 ハードコード値を維持（回帰ゼロ）


# ──────────────────────────────────────────────
# Utility（グループ対応: 人が PowerPoint でマップ要素をグループ化しても動くよう、
# シェイプ探索・削除はグループ内へ再帰し、MAP_FRAME 幾何はグループ変換を解決して
# 絶対座標を算出する）
# ──────────────────────────────────────────────
def _iter_shapes_with_ancestors(shapes, ancestors=()):
    """全シェイプを (shape, 祖先グループのタプル) で再帰列挙する。"""
    for sh in shapes:
        yield sh, ancestors
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes_with_ancestors(sh.shapes, ancestors + (sh,))


def _find_with_ancestors(slide, name):
    for sh, anc in _iter_shapes_with_ancestors(slide.shapes):
        if sh.name == name:
            return sh, anc
    return None, None


def find_shape(slide, name, warn=True):
    sh, _ = _find_with_ancestors(slide, name)
    if sh is None and warn:
        print(f"  ⚠ WARNING: Shape '{name}' not found", file=sys.stderr)
    return sh


def _group_xfrm(grp):
    """グループの off/ext/chOff/chExt を EMU で返す（子座標→親座標変換用）。"""
    xfrm = grp._element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
    off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
    choff, chext = xfrm.find(qn("a:chOff")), xfrm.find(qn("a:chExt"))
    return (int(off.get("x")), int(off.get("y")),
            int(ext.get("cx")), int(ext.get("cy")),
            int(choff.get("x")), int(choff.get("y")),
            int(chext.get("cx")), int(chext.get("cy")))


def _abs_geometry(shape, ancestors):
    """祖先グループの変換を内側→外側に適用し、絶対座標 (x, y, w, h) を返す。"""
    x, y, w, h = shape.left, shape.top, shape.width, shape.height
    for grp in reversed(ancestors):
        ox, oy, ex, ey, cox, coy, cex, cey = _group_xfrm(grp)
        sx = ex / cex if cex else 1.0
        sy = ey / cey if cey else 1.0
        x = ox + (x - cox) * sx
        y = oy + (y - coy) * sy
        w *= sx
        h *= sy
    return int(x), int(y), int(w), int(h)


def set_textbox_text(shape, text):
    """テンプレ側 run[0] の rPr を保ったままテキストだけ差し替える。"""
    if shape is None:
        return
    tf = shape.text_frame
    para = tf.paragraphs[0]
    if para.runs:
        para.runs[0].text = text
        for run in para.runs[1:]:
            run.text = ""
    else:
        r_elem = etree.SubElement(para._p, qn("a:r"))
        etree.SubElement(r_elem, qn("a:rPr"), attrib={"lang": "ja-JP"})
        etree.SubElement(r_elem, qn("a:t")).text = text


def _silent_remove_shape(slide, shape_name: str) -> None:
    """グループ内も含めて名前一致シェイプを削除する。"""
    for s, _ in list(_iter_shapes_with_ancestors(slide.shapes)):
        if s.name == shape_name:
            s._element.getparent().remove(s._element)


def _fill_or_hide(slide, name, text) -> None:
    """テキストが非空なら名前付きシェイプに流し込み、空ならシェイプごと削除（非表示）。"""
    if text is not None and str(text).strip():
        set_textbox_text(find_shape(slide, name, warn=False), text)
    else:
        _silent_remove_shape(slide, name)


def hex_to_rgb(hex_str):
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _set_shape_transparency(shape, alpha_percent):
    sp_pr = shape.fill._xPr
    solidFill = sp_pr.find(qn("a:solidFill"))
    if solidFill is None:
        return
    srgb = solidFill.find(qn("a:srgbClr"))
    if srgb is None:
        return
    for elem in srgb.findall(qn("a:alpha")):
        srgb.remove(elem)
    alpha_val = (100 - alpha_percent) * 1000
    etree.SubElement(srgb, qn("a:alpha")).set("val", str(alpha_val))


# ──────────────────────────────────────────────
# Text fill into fixed named shapes
# ──────────────────────────────────────────────
def fill_text_shapes(slide, data):
    """軸ラベル・象限ラベル・示唆・セクションタイトルを固定枠へ流し込む。"""
    x_axis = data.get("x_axis", {})
    y_axis = data.get("y_axis", {})

    # セクションタイトル
    _fill_or_hide(slide, "SECTION_LEFT_TITLE",
                  data.get("section_title", "ポジショニングマップ"))
    _fill_or_hide(slide, "IMPL_TITLE",
                  data.get("implications_title", "ポジショニングからの示唆"))

    # X 軸（low/high には矢印プレフィックスを付与）
    _fill_or_hide(slide, "AXIS_X_LABEL", x_axis.get("label", ""))
    x_low = x_axis.get("low", "")
    x_high = x_axis.get("high", "")
    _fill_or_hide(slide, "AXIS_X_LOW", f"← {x_low}" if x_low else "")
    _fill_or_hide(slide, "AXIS_X_HIGH", f"{x_high} →" if x_high else "")

    # Y 軸ラベル
    _fill_or_hide(slide, "AXIS_Y_LABEL", y_axis.get("label", ""))

    # 象限 と Y-low/high は排他: quadrants があれば Y-low/high を消し、無ければ象限を消す
    quadrants = data.get("quadrants", {})
    if quadrants:
        _silent_remove_shape(slide, "AXIS_Y_LOW")
        _silent_remove_shape(slide, "AXIS_Y_HIGH")
        _fill_or_hide(slide, "QUAD_TL", quadrants.get("top_left", ""))
        _fill_or_hide(slide, "QUAD_TR", quadrants.get("top_right", ""))
        _fill_or_hide(slide, "QUAD_BL", quadrants.get("bottom_left", ""))
        _fill_or_hide(slide, "QUAD_BR", quadrants.get("bottom_right", ""))
    else:
        for q in ("QUAD_TL", "QUAD_TR", "QUAD_BL", "QUAD_BR"):
            _silent_remove_shape(slide, q)
        y_low = y_axis.get("low", "")
        y_high = y_axis.get("high", "")
        _fill_or_hide(slide, "AXIS_Y_HIGH", f"{y_high} →" if y_high else "")
        _fill_or_hide(slide, "AXIS_Y_LOW", f"← {y_low}" if y_low else "")

    # 示唆 5 固定枠（空枠は非表示）
    implications = data.get("implications", [])
    for i in range(IMPL_SLOTS):
        text = implications[i] if i < len(implications) else ""
        _fill_or_hide(slide, f"IMPL_{i + 1}", text)
    print(f"  ✓ 示唆: {min(len(implications), IMPL_SLOTS)}項目（固定{IMPL_SLOTS}枠）")


# ──────────────────────────────────────────────
# Bubbles (the only runtime drawing)
# ──────────────────────────────────────────────
def draw_bubbles(slide, data):
    """MAP_FRAME の幾何を基準に各プレイヤーのバブル + ラベルを描画する。"""
    frame, frame_anc = _find_with_ancestors(slide, SHAPE_MAP_FRAME)
    if frame is None:
        raise ValueError(
            f"テンプレに '{SHAPE_MAP_FRAME}' シェイプがありません。"
            "tools/build_positioning_map_template.py でテンプレを再生成してください。"
        )
    # 人がマップ要素をグループ化していてもバブルが正しく載るよう、グループ変換を
    # 解決して MAP_FRAME の絶対座標を求める（バブルは slide 直下に絶対座標で描く）。
    map_x, map_y, map_w, map_h = _abs_geometry(frame, frame_anc)

    players = data.get("players", [])
    target_company = data.get("target_company")
    x_axis = data.get("x_axis", {})
    y_axis = data.get("y_axis", {})
    x_min, x_max = x_axis.get("min", 0), x_axis.get("max", 10)
    y_min, y_max = y_axis.get("min", 0), y_axis.get("max", 10)

    sizes = [p.get("size", 1) for p in players]
    min_size = min(sizes) if sizes else 1
    max_size = max(sizes) if sizes else 1
    min_diam = Inches(0.40)
    max_diam = Inches(0.95)

    def compute_diameter(size_val):
        if max_size == min_size:
            return Inches(0.60)
        ratio = (size_val - min_size) / (max_size - min_size)
        return Emu(int(min_diam + (max_diam - min_diam) * ratio))

    is_roleup = (_THEME is not None and _THEME.id == "roleup")

    for i, p in enumerate(players):
        name = p["name"]
        x_val, y_val = p["x"], p["y"]
        size_val = p.get("size", 1)
        is_target = (name == target_company) if target_company else False

        x_ratio = (x_val - x_min) / (x_max - x_min) if x_max != x_min else 0.5
        y_ratio = (y_val - y_min) / (y_max - y_min) if y_max != y_min else 0.5

        diam = compute_diameter(size_val)
        cx = map_x + int(map_w * x_ratio)
        cy = map_y + int(map_h * (1 - y_ratio))
        bubble_x = cx - diam // 2
        bubble_y = cy - diam // 2

        if is_target:
            color = COLOR_TARGET
            line_color = COLOR_TARGET_LINE
            line_w = Pt(2.5)
        else:
            color = hex_to_rgb(_palette_color(i, len(players)))
            line_color = COLOR_BUBBLE_LINE
            line_w = Pt(1.0)

        bubble = slide.shapes.add_shape(MSO_SHAPE.OVAL, bubble_x, bubble_y, diam, diam)
        bubble.fill.solid()
        bubble.fill.fore_color.rgb = color
        _set_shape_transparency(bubble, 10 if is_target else 30)
        bubble.line.color.rgb = line_color
        bubble.line.width = line_w
        bubble.shadow.inherit = False
        bubble.text_frame.text = ""

        # ── バブルラベル ──
        label_pos = p.get("label_position", "bottom")
        if is_roleup:
            label_w, label_h = Inches(0.95), Inches(0.20)
        else:
            label_w, label_h = Inches(1.8), Inches(0.26)
        label_gap = Inches(0.03)

        if label_pos == "top":
            label_x = cx - label_w // 2
            label_y = cy - diam // 2 - label_h - label_gap
        elif label_pos == "left":
            label_x = bubble_x - label_w - label_gap
            label_y = cy - label_h // 2
        elif label_pos == "right":
            label_x = bubble_x + diam + label_gap
            label_y = cy - label_h // 2
        else:
            label_x = cx - label_w // 2
            label_y = cy + diam // 2 + label_gap

        # 境界はみ出しの自動反転（roleup は密なので上方向に微小許容）
        top_tolerance = Inches(0.30) if is_roleup else Emu(0)
        if label_y + label_h > map_y + map_h:
            label_x = cx - label_w // 2
            label_y = cy - diam // 2 - label_h - label_gap
        if label_y < map_y - top_tolerance:
            label_x = cx - label_w // 2
            label_y = cy + diam // 2 + label_gap

        label_tb = slide.shapes.add_textbox(label_x, label_y, label_w, label_h)
        ltf = label_tb.text_frame
        ltf.word_wrap = False
        ltf.margin_left = 0; ltf.margin_right = 0; ltf.margin_top = 0; ltf.margin_bottom = 0
        lp = ltf.paragraphs[0]
        if label_pos == "left":
            lp.alignment = PP_ALIGN.RIGHT
        elif label_pos == "right":
            lp.alignment = PP_ALIGN.LEFT
        else:
            lp.alignment = PP_ALIGN.CENTER
        lrun = lp.add_run()
        lrun.text = name
        lrun.font.size = FONT_SIZE_BUBBLE_TARGET if is_target else FONT_SIZE_BUBBLE
        lrun.font.bold = True if is_target else False
        lrun.font.name = FONT_NAME_JP
        lrun.font.color.rgb = COLOR_TARGET_LINE if is_target else COLOR_TEXT

    print(f"  ✓ ポジショニングマップ: {len(players)}プレイヤー配置")


# ──────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument(
        "--template", required=False, default=None,
        help="Optional explicit template path. If omitted, resolved from --brand.",
    )
    ap.add_argument("--output", required=True)
    add_brand_arg(ap)
    args = ap.parse_args()

    theme = resolve_brand(args.brand, SKILL_DIR)
    _apply_theme(theme)
    template_path = args.template or theme.template_path(SKILL_DIR, "positioning-map")
    print(f"  ✓ Brand: {theme.id} ({theme.label})")
    print(f"  ✓ Template: {template_path}")

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    # ISSUE-012: スキーマ齟齬の silent fail 防止
    validate_fill_input(
        data,
        required_top=["main_message", "players", "x_axis", "y_axis"],
        allowed_top=[
            "main_message", "chart_title", "section_title", "target_company",
            "x_axis", "y_axis", "quadrants", "players",
            "implications", "implications_title",
            "source", "source_label", "source_text",
            "title", "subtitle",
        ],
        nested_required={
            "x_axis": ["label", "low", "high"],
            "y_axis": ["label", "low", "high"],
        },
        per_item_required={"players": ["name", "x", "y"]},
        skill_name=SKILL_ID,
    )

    require_source(data, theme, skill_id=SKILL_ID)

    _mm = data.get("main_message", "")
    if len(_mm) > 65:
        raise ValueError(f"main_message は 65 字以内（受領: {len(_mm)}）: {_mm[:80]}...")

    PLAYERS_MIN, PLAYERS_MAX = 2, 5
    _players = data.get("players", [])
    if not isinstance(_players, list):
        raise ValueError("players は配列である必要があります")
    if not (PLAYERS_MIN <= len(_players) <= PLAYERS_MAX):
        raise ValueError(
            f"players の要素数は {PLAYERS_MIN}〜{PLAYERS_MAX} の範囲である必要があります"
            f"（受領: {len(_players)}、target_company を含む）"
        )

    # 示唆は固定 5 枠。超過分は silent 消失すると事故になるため WARN + truncate。
    _impl = data.get("implications", [])
    if isinstance(_impl, list) and len(_impl) > IMPL_SLOTS:
        print(f"  ⚠ WARNING: implications が {len(_impl)} 件あり固定枠 {IMPL_SLOTS} を超過。"
              f"先頭 {IMPL_SLOTS} 件のみ表示します。", file=sys.stderr)
        data["implications"] = _impl[:IMPL_SLOTS]

    prs = Presentation(template_path)
    slide = prs.slides[0]

    # 上部プレースホルダ（main_message / chart_title）
    top_text = resolve_top_text(data, theme)
    sub_text = resolve_subtitle_text(data, theme)
    set_textbox_text(find_shape(slide, SHAPE_MAIN_MESSAGE), top_text)
    set_textbox_text(find_shape(slide, SHAPE_CHART_TITLE),
                     sub_text or data.get("chart_title", "ポジショニングマップ"))

    # 固定枠へのテキスト流し込み + バブル描画
    fill_text_shapes(slide, data)
    draw_bubbles(slide, data)

    # 出典（両ブランドとも名前付き Source 3 へ。空なら非表示）
    source = data.get("source", "")
    _fill_or_hide(slide, SHAPE_SOURCE, source)
    if source:
        print(f"  ✓ Source: {source[:40]}...")

    os.makedirs(os.path.dirname(args.output), exist_ok=True)
    prs.save(args.output)
    _finalize_pptx(args.output)
    print(f"\n✅ Saved: {args.output}")


if __name__ == "__main__":
    main()
