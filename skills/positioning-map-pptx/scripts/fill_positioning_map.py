"""
fill_positioning_map.py — ポジショニングマップスライドをPPTXネイティブオブジェクトで生成

レイアウト:
  - 上部: メインメッセージ + チャートタイトル
  - 左側 (8.0in x 5.2in): 2軸ポジショニングマップ
      - X軸・Y軸ラベル、両端にlow/highラベル
      - 4象限ラベル（オプション）
      - 各プレイヤーをバブル（円）で配置
      - 対象会社は目立つ色＋太枠でハイライト
      - バブル下に企業名ラベル
  - 右側 (4.5in): 示唆（Implications）のブレット項目
  - 下部: 出典

描画方針: ネイティブチャートではなく MSO_SHAPE.OVAL と直線で手動描画
（コンサル品質の見た目を実現）

Usage:
  python fill_positioning_map.py \
    --data /home/claude/positioning_map_data.json \
    --template <path>/positioning-map-template.pptx \
    --output /mnt/user-data/outputs/PositioningMap_output.pptx
"""

import argparse
import json
import os
import sys

# brand_resolver bootstrap (passive --brand acceptance until brand-aware migration)
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(SKILL_DIR, "..", "_common", "lib"))
from brand_resolver import add_brand_arg  # noqa: E402

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
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



# ── Layout Constants (16:9) ──
SHAPE_MAIN_MESSAGE = "Title 1"
SHAPE_CHART_TITLE = "Text Placeholder 2"

PANEL_Y = Inches(1.55)
PANEL_H = Inches(5.35)

# Left panel (positioning map area)
LEFT_X = Inches(0.41)
LEFT_W = Inches(7.70)

# Right panel (implications)
RIGHT_X = Inches(8.30)
RIGHT_W = Inches(4.65)

# Map drawing area inside left panel (axis labels etc consume margins)
MAP_MARGIN_LEFT = Inches(0.75)   # for Y axis label + low/high
MAP_MARGIN_RIGHT = Inches(0.20)
MAP_MARGIN_TOP = Inches(0.70)    # section title + quadrant labels top
MAP_MARGIN_BOTTOM = Inches(0.80) # for X axis label + low/high

SOURCE_X = Inches(0.41)
SOURCE_Y = Inches(6.93)
SOURCE_W = Inches(12.50)

# ── Colors ──
COLOR_TEXT = RGBColor(0x33, 0x33, 0x33)
COLOR_SOURCE = RGBColor(0x66, 0x66, 0x66)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_FRAME = RGBColor(0x33, 0x33, 0x33)
COLOR_GRID = RGBColor(0xC0, 0xC0, 0xC0)
COLOR_AXIS_END = RGBColor(0x66, 0x66, 0x66)
COLOR_QUADRANT_LABEL = RGBColor(0x99, 0x99, 0x99)
COLOR_TARGET = RGBColor(0xE1, 0x57, 0x59)     # 対象会社の強調色（赤系）
COLOR_DEFAULT_BUBBLE = RGBColor(0x4E, 0x79, 0xA7)  # 通常色（紺系）

# ─── 共通配色（正本: skills/_common/styles/chart_palette.md） ───
# 編集時は _common/styles/chart_palette.md と他 4 スキルの fill_*.py も同期更新
# CHART_PALETTE には TARGET_COLOR(赤) と OTHER_COLOR(灰) を含めない（palette 外で固定）
# → target は palette と衝突しないため、非ターゲットバブルも配列 index 直引きで OK
#   （旧 NON_TARGET_PALETTE / _non_target_bubble_color() は廃止）
CHART_PALETTE = [
    "#4E79A7", "#F28E2B", "#59A14F", "#76B7B2",
    "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F",
]
OTHER_COLOR = "#BAB0AC"
TARGET_COLOR = "#E15759"
LABEL_BAR_COLOR = "#4E79A7"
LABEL_BG_COLOR = "#E8EEF5"


def _palette_color(index: int, total: int) -> str:
    if total <= 1:
        return CHART_PALETTE[0]
    return CHART_PALETTE[index % len(CHART_PALETTE)]


# 後方互換のためのエイリアス（既存コードからの参照用）
DEFAULT_COLORS = CHART_PALETTE

FONT_NAME_JP = "Meiryo UI"
FONT_SIZE_SECTION = Pt(16)
FONT_SIZE_AXIS_LABEL = Pt(14)
FONT_SIZE_AXIS_END = Pt(11)
FONT_SIZE_QUADRANT = Pt(12)
FONT_SIZE_BUBBLE = Pt(12)
FONT_SIZE_ITEM = Pt(13)
FONT_SIZE_SOURCE = Pt(11)


# ──────────────────────────────────────────────
# Utility
# ──────────────────────────────────────────────
def find_shape(slide, name):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    print(f"  ⚠ WARNING: Shape '{name}' not found", file=sys.stderr)
    return None


def set_textbox_text(shape, text):
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
        t_elem = etree.SubElement(r_elem, qn("a:t"))
        t_elem.text = text


def hex_to_rgb(hex_str):
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_section_title(slide, text, left, top, width):
    """セクションタイトル（下線付き、14pt Bold、中央揃え）"""
    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.30))
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = FONT_SIZE_SECTION
    run.font.bold = True
    run.font.color.rgb = COLOR_TEXT
    run.font.name = FONT_NAME_JP

    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top + Inches(0.30), width, Inches(0.02)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_TEXT
    line.line.fill.background()
    return txBox


def add_text_box(slide, text, left, top, width, height, font_size, bold=False,
                 color=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                 italic=False):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = FONT_NAME_JP
    if color is not None:
        run.font.color.rgb = color
    else:
        run.font.color.rgb = COLOR_TEXT
    return tb


def add_rotated_text_box(slide, text, left, top, width, height, font_size,
                          bold=False, color=None, align=PP_ALIGN.CENTER,
                          rotation=-90):
    """Y軸ラベル用: 回転したテキストボックス"""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tb.rotation = rotation
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.bold = bold
    run.font.name = FONT_NAME_JP
    if color is not None:
        run.font.color.rgb = color
    else:
        run.font.color.rgb = COLOR_TEXT
    return tb


# ──────────────────────────────────────────────
# Map Drawing
# ──────────────────────────────────────────────
def draw_positioning_map(slide, data, left, top, width, height):
    """
    2軸ポジショニングマップを描画する
    """
    # セクションタイトル
    section_title = data.get("section_title", "ポジショニングマップ")
    add_section_title(slide, section_title, left, top, width)

    # マップエリアの矩形を計算
    map_x = left + MAP_MARGIN_LEFT
    map_y = top + MAP_MARGIN_TOP
    map_w = width - MAP_MARGIN_LEFT - MAP_MARGIN_RIGHT
    map_h = height - MAP_MARGIN_TOP - MAP_MARGIN_BOTTOM

    # マップの外枠
    frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, map_x, map_y, map_w, map_h)
    frame.fill.solid()
    frame.fill.fore_color.rgb = RGBColor(0xFA, 0xFA, 0xFA)
    frame.line.color.rgb = COLOR_FRAME
    frame.line.width = Pt(1.0)
    frame.shadow.inherit = False
    frame.text_frame.text = ""

    # 十字ガイドライン（破線）
    # 縦線（中央）
    v_conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        map_x + map_w // 2, map_y,
        map_x + map_w // 2, map_y + map_h,
    )
    v_conn.line.color.rgb = COLOR_GRID
    v_conn.line.width = Pt(0.75)
    _set_dash_style(v_conn, "dash")

    # 横線（中央）
    h_conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        map_x, map_y + map_h // 2,
        map_x + map_w, map_y + map_h // 2,
    )
    h_conn.line.color.rgb = COLOR_GRID
    h_conn.line.width = Pt(0.75)
    _set_dash_style(h_conn, "dash")

    # ── 軸ラベル ──
    x_axis = data.get("x_axis", {})
    y_axis = data.get("y_axis", {})

    # X軸ラベル（下、中央）
    x_label = x_axis.get("label", "")
    if x_label:
        add_text_box(
            slide, x_label,
            map_x, map_y + map_h + Inches(0.32),
            map_w, Inches(0.25),
            FONT_SIZE_AXIS_LABEL, bold=True, align=PP_ALIGN.CENTER,
        )
    # X軸 low/high （左端・右端、軸線の下）
    x_low = x_axis.get("low", "")
    x_high = x_axis.get("high", "")
    if x_low:
        add_text_box(
            slide, f"← {x_low}",
            map_x, map_y + map_h + Inches(0.05),
            Inches(1.8), Inches(0.22),
            FONT_SIZE_AXIS_END, color=COLOR_AXIS_END, align=PP_ALIGN.LEFT,
        )
    if x_high:
        add_text_box(
            slide, f"{x_high} →",
            map_x + map_w - Inches(1.8), map_y + map_h + Inches(0.05),
            Inches(1.8), Inches(0.22),
            FONT_SIZE_AXIS_END, color=COLOR_AXIS_END, align=PP_ALIGN.RIGHT,
        )

    # Y軸ラベル（左、中央、縦書き）
    # NOTE: PowerPointの textbox.rotation は中心軸での回転。
    # 回転後の見た目の中心X = left + width/2 を、マップ枠 map_x より外側に来るよう調整。
    # 回転前 width=2.0 の場合、中心X = left + 1.0。
    # マップ外（map_x - Inches(0.4)）に中心が来るには left = map_x - Inches(1.4)
    y_label = y_axis.get("label", "")
    if y_label:
        # 回転テキストボックスで縦書き表示（回転後マップ外側に出るよう左にオフセット）
        add_rotated_text_box(
            slide, y_label,
            map_x - Inches(1.4), map_y + map_h // 2 - Inches(1.0),
            Inches(2.0), Inches(0.30),
            FONT_SIZE_AXIS_LABEL, bold=True, align=PP_ALIGN.CENTER,
            rotation=-90,
        )

    # ── 象限ラベル（オプション）──
    quadrants = data.get("quadrants", {})
    has_quadrants = bool(quadrants)

    # Y軸 low/high（象限ラベルがない場合のみ表示。重複回避）
    if not has_quadrants:
        y_low = y_axis.get("low", "")
        y_high = y_axis.get("high", "")
        if y_high:
            add_rotated_text_box(
                slide, f"{y_high} →",
                map_x - Inches(0.45), map_y - Inches(0.05),
                Inches(1.2), Inches(0.22),
                FONT_SIZE_AXIS_END, color=COLOR_AXIS_END, align=PP_ALIGN.LEFT,
                rotation=-90,
            )
        if y_low:
            add_rotated_text_box(
                slide, f"← {y_low}",
                map_x - Inches(0.45), map_y + map_h - Inches(1.25),
                Inches(1.2), Inches(0.22),
                FONT_SIZE_AXIS_END, color=COLOR_AXIS_END, align=PP_ALIGN.RIGHT,
                rotation=-90,
            )

    if has_quadrants:
        qlabel_w = Inches(2.2)
        qlabel_h = Inches(0.25)
        qmargin = Inches(0.12)
        # 右上
        if quadrants.get("top_right"):
            add_text_box(
                slide, quadrants["top_right"],
                map_x + map_w - qlabel_w - qmargin,
                map_y + qmargin,
                qlabel_w, qlabel_h,
                FONT_SIZE_QUADRANT, bold=False, italic=True,
                color=COLOR_QUADRANT_LABEL, align=PP_ALIGN.RIGHT,
            )
        # 左上
        if quadrants.get("top_left"):
            add_text_box(
                slide, quadrants["top_left"],
                map_x + qmargin,
                map_y + qmargin,
                qlabel_w, qlabel_h,
                FONT_SIZE_QUADRANT, bold=False, italic=True,
                color=COLOR_QUADRANT_LABEL, align=PP_ALIGN.LEFT,
            )
        # 右下
        if quadrants.get("bottom_right"):
            add_text_box(
                slide, quadrants["bottom_right"],
                map_x + map_w - qlabel_w - qmargin,
                map_y + map_h - Inches(0.33),
                qlabel_w, qlabel_h,
                FONT_SIZE_QUADRANT, bold=False, italic=True,
                color=COLOR_QUADRANT_LABEL, align=PP_ALIGN.RIGHT,
            )
        # 左下
        if quadrants.get("bottom_left"):
            add_text_box(
                slide, quadrants["bottom_left"],
                map_x + qmargin,
                map_y + map_h - Inches(0.33),
                qlabel_w, qlabel_h,
                FONT_SIZE_QUADRANT, bold=False, italic=True,
                color=COLOR_QUADRANT_LABEL, align=PP_ALIGN.LEFT,
            )

    # ── プレイヤーのバブル配置 ──
    players = data.get("players", [])
    target_company = data.get("target_company")

    x_min = x_axis.get("min", 0)
    x_max = x_axis.get("max", 10)
    y_min = y_axis.get("min", 0)
    y_max = y_axis.get("max", 10)

    # バブルサイズ計算
    # sizeが指定されている場合、min-max間でスケーリング
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

    for i, p in enumerate(players):
        name = p["name"]
        x_val = p["x"]
        y_val = p["y"]
        size_val = p.get("size", 1)
        is_target = (name == target_company) if target_company else False

        # 座標計算: X軸は左→右、Y軸は下→上
        x_ratio = (x_val - x_min) / (x_max - x_min) if x_max != x_min else 0.5
        y_ratio = (y_val - y_min) / (y_max - y_min) if y_max != y_min else 0.5

        # バブル直径
        diam = compute_diameter(size_val)

        # バブルの中心座標
        cx = map_x + int(map_w * x_ratio)
        cy = map_y + int(map_h * (1 - y_ratio))  # Y反転（上が高値）

        bubble_x = cx - diam // 2
        bubble_y = cy - diam // 2

        # 色（共通パレット使用、JSON の color は無視）
        # target は TARGET_COLOR(赤) で上書き、それ以外は配列 index 直引き
        # （CHART_PALETTE から赤を除外済みなので衝突しない、P6 と色順が一致）
        if is_target:
            color = COLOR_TARGET
            line_color = RGBColor(0x8B, 0x2C, 0x2E)  # 濃赤
            line_w = Pt(2.5)
        else:
            color = hex_to_rgb(_palette_color(i, len(players)))
            line_color = RGBColor(0x33, 0x33, 0x33)
            line_w = Pt(1.0)

        bubble = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, bubble_x, bubble_y, diam, diam
        )
        bubble.fill.solid()
        bubble.fill.fore_color.rgb = color
        # 透明度を設定（対象会社は薄く、通常バブルは中程度に透明）
        _set_shape_transparency(bubble, 10 if is_target else 30)
        bubble.line.color.rgb = line_color
        bubble.line.width = line_w
        bubble.shadow.inherit = False
        bubble.text_frame.text = ""

        # 企業名ラベルの配置
        # label_position: "bottom"(デフォルト), "top", "left", "right"
        label_pos = p.get("label_position", "bottom")
        label_w = Inches(1.8)
        label_h = Inches(0.26)
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
        else:  # bottom (default)
            label_x = cx - label_w // 2
            label_y = cy + diam // 2 + label_gap

        # マップ境界チェック - はみ出る場合は自動調整
        if label_y + label_h > map_y + map_h:
            # 下にはみ出るなら上に配置
            label_x = cx - label_w // 2
            label_y = cy - diam // 2 - label_h - label_gap
        if label_y < map_y:
            # 上にはみ出るなら下に
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
        # 対象会社はフォントサイズを1pt大きく、太字
        lrun.font.size = Pt(FONT_SIZE_BUBBLE.pt + 1) if is_target else FONT_SIZE_BUBBLE
        lrun.font.bold = True if is_target else False
        lrun.font.name = FONT_NAME_JP
        lrun.font.color.rgb = RGBColor(0x8B, 0x2C, 0x2E) if is_target else COLOR_TEXT

    print(f"  ✓ ポジショニングマップ: {len(players)}プレイヤー配置")


def _set_dash_style(shape, dash_style="dash"):
    """shape の line に dash style を設定（破線）"""
    ln = shape.line._get_or_add_ln()
    # prstDash element
    # 既存のprstDashを削除
    for elem in ln.findall(qn("a:prstDash")):
        ln.remove(elem)
    prstDash = etree.SubElement(ln, qn("a:prstDash"))
    prstDash.set("val", dash_style)


def _set_shape_transparency(shape, alpha_percent):
    """shape の塗りつぶし透明度を設定 (0=不透明, 100=完全透明)"""
    # Find solidFill element
    sp_pr = shape.fill._xPr
    solidFill = sp_pr.find(qn("a:solidFill"))
    if solidFill is None:
        return
    srgb = solidFill.find(qn("a:srgbClr"))
    if srgb is None:
        return
    # Remove existing alpha
    for elem in srgb.findall(qn("a:alpha")):
        srgb.remove(elem)
    # Add alpha (PPT uses alpha as "opacity"; 100000 = fully opaque)
    alpha_val = (100 - alpha_percent) * 1000
    alpha = etree.SubElement(srgb, qn("a:alpha"))
    alpha.set("val", str(alpha_val))


# ──────────────────────────────────────────────
# Right Panel: Implications
# ──────────────────────────────────────────────
def build_implications_panel(slide, section_title, implications, left, top, width, height):
    """示唆（implications）のブレット項目リスト"""
    add_section_title(slide, section_title, left, top, width)

    if not implications:
        return

    body_top = top + Inches(0.50)
    body_h = height - Inches(0.50)

    tb = slide.shapes.add_textbox(
        left + Inches(0.05), body_top,
        width - Inches(0.10), body_h,
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0

    # 既存段落クリア
    for p in list(tf.paragraphs):
        p._p.getparent().remove(p._p)

    # 各項目を追加
    for i, item_text in enumerate(implications):
        p_elem = etree.SubElement(tf._txBody, qn("a:p"))
        pPr = etree.SubElement(p_elem, qn("a:pPr"), attrib={
            "marL": "220000",
            "indent": "-220000",
        })
        if i > 0:
            spcBef = etree.SubElement(pPr, qn("a:spcBef"))
            etree.SubElement(spcBef, qn("a:spcPts"), attrib={"val": "600"})

        buChar = etree.SubElement(pPr, qn("a:buChar"), attrib={"char": "●"})
        buFont = etree.SubElement(pPr, qn("a:buFont"), attrib={"typeface": "Arial"})
        buClr = etree.SubElement(pPr, qn("a:buClr"))
        buClrSolid = etree.SubElement(buClr, qn("a:srgbClr"))
        buClrSolid.set("val", "2E4A6B")

        r = etree.SubElement(p_elem, qn("a:r"))
        rPr = etree.SubElement(r, qn("a:rPr"), attrib={
            "lang": "ja-JP",
            "sz": str(int(FONT_SIZE_ITEM.pt * 100)),
        })
        etree.SubElement(rPr, qn("a:latin"), attrib={"typeface": FONT_NAME_JP})
        etree.SubElement(rPr, qn("a:ea"), attrib={"typeface": FONT_NAME_JP})
        sf = etree.SubElement(rPr, qn("a:solidFill"))
        s = etree.SubElement(sf, qn("a:srgbClr"))
        s.set("val", "333333")
        t = etree.SubElement(r, qn("a:t"))
        t.text = item_text

    print(f"  ✓ 示唆パネル: {len(implications)}項目")


# ──────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--output", required=True)
    add_brand_arg(ap)  # passive: accepted but ignored until brand migration
    args = ap.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    _mm = data.get("main_message", "")
    if len(_mm) > 65:
        raise ValueError(
            f"main_message は 65 字以内（受領: {len(_mm)}）: {_mm[:80]}..."
        )

    # v0.2: players 数のバリデーション（deck_skeleton_standard.json limits.max_competitors と同期）
    PLAYERS_MIN = 2
    PLAYERS_MAX = 10
    _players = data.get("players", [])
    if not isinstance(_players, list):
        raise ValueError("players は配列である必要があります")
    if not (PLAYERS_MIN <= len(_players) <= PLAYERS_MAX):
        raise ValueError(
            f"players の要素数は {PLAYERS_MIN}〜{PLAYERS_MAX} の範囲である必要があります"
            f"（受領: {len(_players)}、target_company を含む）"
        )

    prs = Presentation(args.template)
    slide = prs.slides[0]

    set_textbox_text(find_shape(slide, SHAPE_MAIN_MESSAGE), data.get("main_message", ""))
    set_textbox_text(find_shape(slide, SHAPE_CHART_TITLE), data.get("chart_title", "ポジショニングマップ"))

    # 左: マップ
    draw_positioning_map(slide, data, LEFT_X, PANEL_Y, LEFT_W, PANEL_H)

    # 右: 示唆
    implications_title = data.get("implications_title", "ポジショニングからの示唆")
    implications = data.get("implications", [])
    build_implications_panel(
        slide, implications_title, implications,
        RIGHT_X, PANEL_Y, RIGHT_W, PANEL_H,
    )

    # 出典
    source = data.get("source", "")
    if source:
        add_text_box(
            slide, source,
            SOURCE_X, SOURCE_Y, SOURCE_W, Inches(0.30),
            FONT_SIZE_SOURCE, bold=False, color=COLOR_SOURCE,
            align=PP_ALIGN.LEFT,
        )
        print(f"  ✓ Source: {source[:40]}...")

    os.makedirs(os.path.dirname(args.output), exist_ok=True)
    prs.save(args.output)
    _finalize_pptx(args.output)
    print(f"\n✅ Saved: {args.output}")


if __name__ == "__main__":
    main()
