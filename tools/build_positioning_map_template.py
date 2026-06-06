#!/usr/bin/env python3
"""build_positioning_map_template.py — ポジショニングマップの richテンプレを 2ブランド分生成する bootstrap。

⚠️ 警告（再実行の前に必ず読む）:
  正本テンプレ (skills/positioning-map-pptx/assets/<brand>/positioning-map-template.pptx) は、
  本ビルダで初期生成したあと **人が PowerPoint でグループ化・位置微調整済み**。
  本スクリプトは同じパスへ **上書き保存** するため、**再実行すると手調整が全て失われる**。
  → ゼロからテンプレを作り直す時だけ使う bootstrap。通常運用では実行しない。
  （ランタイムの生成は scripts/fill_positioning_map.py が担い、本ビルダは使わない。
   そのため本ファイルはスキル配下ではなく tools/ に置き、install/配布 zip には含めない。）

背景:
  旧テンプレ (assets/<brand>/positioning-map-template.pptx) は Title 1 +
  Text Placeholder 2 程度の骨組みしか持たず、マップ外枠・十字ガイドライン・軸ラベル・
  象限ラベル・示唆パネル・出典まで全て fill_positioning_map.py が実行時にハードコード
  座標から描画していた。これが出力不安定の原因。

  本スクリプトは「固定できる骨組み」を **名前付きシェイプ** としてテンプレに焼き込む。
  fill 側は (1) 名前付きシェイプへのテキスト流し込み と (2) バブル+バブルラベルの
  実行時描画 のみを担うようになる。バブル配置の幾何基準は MAP_FRAME シェイプの
  .left/.top/.width/.height を single source of truth として読む。

設計:
  - 既存ブランドテンプレを **土台に追記** する（ゼロ生成しない）。背景・ロゴ・slide
    master・既存プレースホルダ (Title 1 / Text Placeholder 2 / roleup の Source 3) を保全。
  - ブランド座標差・配色・フォントは STELLAR / ROLEUP の 2 dict に集約（唯一の brand 分岐）。
  - 幾何は旧 fill のマップ算術 (panel + margin) と象限/軸ラベル座標を完全再現し、
    バブル座標の回帰ゼロを狙う。
  - stellar の配色・サイズは V1 ハードコード値（回帰ゼロ）。roleup は theme.json から解決。

命名スキーム（2ブランド共通・ASCII。fill を brand 非依存に）:
  Title 1 / Text Placeholder 2 / Source 3        … 既存保持（stellar は Source 3 を追加）
  SECTION_LEFT_TITLE / SECTION_LEFT_RULE          … 左セクションタイトル（+ stellar 下線）
  MAP_FRAME / MAP_GUIDE_V / MAP_GUIDE_H           … マップ外枠・十字ガイドライン
  AXIS_X_LABEL / AXIS_X_LOW / AXIS_X_HIGH         … X 軸ラベル
  AXIS_Y_LABEL / AXIS_Y_LOW / AXIS_Y_HIGH         … Y 軸ラベル（回転）
  QUAD_TL / QUAD_TR / QUAD_BL / QUAD_BR           … 4 象限ラベル（固定 4 枠）
  IMPL_TITLE / IMPL_1 … IMPL_5                     … 示唆タイトル + 固定 5 枠

冪等: 再実行で出力ファイルを上書きする（＝手調整を破壊する。上記警告参照）。

Usage:
  python3 tools/build_positioning_map_template.py            # 両ブランド
  python3 tools/build_positioning_map_template.py --brand roleup
"""
import argparse
import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from lxml import etree

# tools/ から REPO ルート経由で対象スキルと _common/lib を解決する。
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILL_DIR = os.path.join(REPO_ROOT, "skills", "positioning-map-pptx")
sys.path.insert(0, os.path.join(REPO_ROOT, "skills", "_common", "lib"))
from brand_resolver import resolve_brand  # noqa: E402

SKILL_NAME = "positioning-map"

# 配色・幾何の共通定数（_apply_theme で reassign されない値 = 両ブランド共通）
MAP_BG = "FAFAFA"
GRID = "C0C0C0"
QUADRANT_LABEL = "999999"

# 旧 fill のマップ内マージン（module 定数。brand 非依存だった）
MARGIN_L, MARGIN_R, MARGIN_T, MARGIN_B = 0.75, 0.20, 0.70, 0.80


def _brand_spec(brand):
    """ブランド別の座標・配色・サイズを 1 dict に解決して返す。"""
    if brand == "stellar_aiz":
        # V1 ハードコード値（回帰ゼロ）
        return {
            "panel_y": 1.55, "panel_h": 5.35,
            "left_x": 0.41, "left_w": 7.70,
            "right_x": 8.30, "right_w": 4.65,
            "source": (0.41, 6.93, 12.50, 0.30),
            "impl_pitch": 0.70,
            "section_align": "center",
            "underline": True,
            "sizes": {"section": 16, "axis_label": 14, "axis_end": 11,
                      "quadrant": 12, "item": 13, "source": 10},
            "colors": {"text": "333333", "source": "666666", "subtitle": "333333",
                       "frame": "333333", "axis_end": "666666", "bullet": "2E4A6B"},
        }
    if brand == "roleup":
        theme = resolve_brand("roleup", SKILL_DIR)

        def hx(k):
            return theme.hex_no_hash(k)
        return {
            "panel_y": 1.55, "panel_h": 5.95,
            "left_x": 0.41, "left_w": 6.80,
            "right_x": 7.31, "right_w": 3.98,
            "source": (0.41, 7.57, 10.88, 0.44),
            "impl_pitch": 0.70,
            "section_align": "left",
            "underline": False,
            "sizes": {
                "section": theme.pt_value("font_size_subtitle_pt"),       # 12
                "axis_label": theme.pt_value("font_size_key_message_pt"),  # 14
                "axis_end": theme.font_size_body_pt_value(),               # 10
                "quadrant": theme.pt_value("font_size_subtitle_pt"),       # 12
                "item": theme.font_size_body_pt_value(),                   # 10
                "source": theme.pt_value("font_size_source_pt"),          # 6
            },
            "colors": {"text": hx("text"), "source": hx("source"),
                       "subtitle": hx("subtitle"), "frame": hx("text"),
                       "axis_end": hx("source"), "bullet": hx("label_bar")},
        }
    raise ValueError(f"unsupported brand: {brand}")


def _font_ea(brand):
    return resolve_brand(brand, SKILL_DIR).font_ea


# ──────────────────────────────────────────────
# Shape builders
# ──────────────────────────────────────────────
def _styled_textbox(slide, name, x, y, w, h, sz_pt, color_hex, font_ea,
                    bold=False, italic=False, align=PP_ALIGN.LEFT,
                    anchor=MSO_ANCHOR.TOP, rotation=0, placeholder="",
                    bullet_hex=None):
    """名前付きの空テキストボックスを生成し、run[0] に rPr をプリセットする。

    fill 側の set_textbox_text は para.runs[0].text を差し替えるだけなので、
    ここで rPr（フォント・サイズ・色・bold/italic）と pPr（整列・bullet）を
    焼き込んでおけば、fill はテキストを流すだけで体裁が決まる。
    """
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = name
    if rotation:
        tb.rotation = rotation
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = anchor

    p = tf.paragraphs[0]
    pPr = p._p.get_or_add_pPr()
    pPr.set("algn", {PP_ALIGN.LEFT: "l", PP_ALIGN.CENTER: "ctr",
                     PP_ALIGN.RIGHT: "r"}[align])
    if bullet_hex is not None:
        pPr.set("marL", "220000")
        pPr.set("indent", "-220000")
        buClr = etree.SubElement(pPr, qn("a:buClr"))
        etree.SubElement(buClr, qn("a:srgbClr")).set("val", bullet_hex)
        etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
        etree.SubElement(pPr, qn("a:buChar")).set("char", "●")

    r = etree.SubElement(p._p, qn("a:r"))
    rPr = etree.SubElement(r, qn("a:rPr"), attrib={
        "lang": "ja-JP", "sz": str(int(sz_pt * 100)),
    })
    if bold:
        rPr.set("b", "1")
    if italic:
        rPr.set("i", "1")
    etree.SubElement(rPr, qn("a:latin")).set("typeface", font_ea)
    etree.SubElement(rPr, qn("a:ea")).set("typeface", font_ea)
    sf = etree.SubElement(rPr, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr")).set("val", color_hex)
    etree.SubElement(r, qn("a:t")).text = placeholder
    return tb


def _set_dash_style(shape, dash_style="dash"):
    ln = shape.line._get_or_add_ln()
    for elem in ln.findall(qn("a:prstDash")):
        ln.remove(elem)
    etree.SubElement(ln, qn("a:prstDash")).set("val", dash_style)


def _silent_remove_shape(slide, shape_name):
    for s in list(slide.shapes):
        if s.name == shape_name:
            s._element.getparent().remove(s._element)


def build(brand, prs):
    spec = _brand_spec(brand)
    font_ea = _font_ea(brand)
    c = spec["colors"]
    sz = spec["sizes"]
    slide = prs.slides[0]

    # roleup: customer-profile 由来の 2 カラム panel 背景矩形は positioning の
    # パネル幅と合わないため除去（旧 fill も実行時に消していた）。
    if brand == "roleup":
        _silent_remove_shape(slide, "正方形/長方形 1")
        _silent_remove_shape(slide, "正方形/長方形 8")

    panel_y, panel_h = spec["panel_y"], spec["panel_h"]
    left_x, left_w = spec["left_x"], spec["left_w"]
    right_x, right_w = spec["right_x"], spec["right_w"]

    # マップ描画領域（旧 fill 算術を完全再現 → バブル座標の回帰ゼロ）
    map_x = left_x + MARGIN_L
    map_y = panel_y + MARGIN_T
    map_w = left_w - MARGIN_L - MARGIN_R
    map_h = panel_h - MARGIN_T - MARGIN_B

    section_align = PP_ALIGN.LEFT if spec["section_align"] == "left" else PP_ALIGN.CENTER

    # ── 左セクションタイトル（+ stellar 下線） ──
    _styled_textbox(slide, "SECTION_LEFT_TITLE", left_x, panel_y, left_w, 0.30,
                    sz["section"], c["subtitle"], font_ea, bold=True,
                    align=section_align, placeholder="ポジショニングマップ")
    if spec["underline"]:
        rule = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(left_x), Inches(panel_y + 0.30),
            Inches(left_w), Inches(0.02))
        rule.name = "SECTION_LEFT_RULE"
        rule.fill.solid()
        rule.fill.fore_color.rgb = RGBColor.from_string(c["text"])
        rule.line.fill.background()
        rule.shadow.inherit = False

    # ── マップ外枠 ──
    frame = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(map_x), Inches(map_y),
        Inches(map_w), Inches(map_h))
    frame.name = "MAP_FRAME"
    frame.fill.solid()
    frame.fill.fore_color.rgb = RGBColor.from_string(MAP_BG)
    frame.line.color.rgb = RGBColor.from_string(c["frame"])
    frame.line.width = Pt(1.0)
    frame.shadow.inherit = False
    frame.text_frame.text = ""

    # ── 十字ガイドライン（破線） ──
    v = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(map_x + map_w / 2), Inches(map_y),
        Inches(map_x + map_w / 2), Inches(map_y + map_h))
    v.name = "MAP_GUIDE_V"
    v.line.color.rgb = RGBColor.from_string(GRID)
    v.line.width = Pt(0.75)
    _set_dash_style(v)
    h = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(map_x), Inches(map_y + map_h / 2),
        Inches(map_x + map_w), Inches(map_y + map_h / 2))
    h.name = "MAP_GUIDE_H"
    h.line.color.rgb = RGBColor.from_string(GRID)
    h.line.width = Pt(0.75)
    _set_dash_style(h)

    # ── X 軸ラベル ──
    _styled_textbox(slide, "AXIS_X_LABEL", map_x, map_y + map_h + 0.32, map_w, 0.25,
                    sz["axis_label"], c["text"], font_ea, bold=True,
                    align=PP_ALIGN.CENTER, placeholder="（X軸ラベル）")
    _styled_textbox(slide, "AXIS_X_LOW", map_x, map_y + map_h + 0.05, 1.8, 0.22,
                    sz["axis_end"], c["axis_end"], font_ea,
                    align=PP_ALIGN.LEFT, placeholder="← 低")
    _styled_textbox(slide, "AXIS_X_HIGH", map_x + map_w - 1.8, map_y + map_h + 0.05,
                    1.8, 0.22, sz["axis_end"], c["axis_end"], font_ea,
                    align=PP_ALIGN.RIGHT, placeholder="高 →")

    # ── Y 軸ラベル（回転） ──
    _styled_textbox(slide, "AXIS_Y_LABEL", map_x - 1.4, map_y + map_h / 2 - 1.0,
                    2.0, 0.30, sz["axis_label"], c["text"], font_ea, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                    rotation=-90, placeholder="（Y軸ラベル）")
    _styled_textbox(slide, "AXIS_Y_HIGH", map_x - 0.45, map_y - 0.05, 1.2, 0.22,
                    sz["axis_end"], c["axis_end"], font_ea, align=PP_ALIGN.LEFT,
                    anchor=MSO_ANCHOR.MIDDLE, rotation=-90, placeholder="高 →")
    _styled_textbox(slide, "AXIS_Y_LOW", map_x - 0.45, map_y + map_h - 1.25, 1.2, 0.22,
                    sz["axis_end"], c["axis_end"], font_ea, align=PP_ALIGN.RIGHT,
                    anchor=MSO_ANCHOR.MIDDLE, rotation=-90, placeholder="← 低")

    # ── 4 象限ラベル ──
    qw, qh, qm = 2.2, 0.25, 0.12
    is_roleup = (brand == "roleup")
    top_q_y = (map_y - qh - 0.02) if is_roleup else (map_y + qm)
    _styled_textbox(slide, "QUAD_TL", map_x + qm, top_q_y, qw, qh,
                    sz["quadrant"], QUADRANT_LABEL, font_ea, italic=True,
                    align=PP_ALIGN.LEFT, placeholder="（左上象限）")
    _styled_textbox(slide, "QUAD_TR", map_x + map_w - qw - qm, top_q_y, qw, qh,
                    sz["quadrant"], QUADRANT_LABEL, font_ea, italic=True,
                    align=PP_ALIGN.RIGHT, placeholder="（右上象限）")
    _styled_textbox(slide, "QUAD_BL", map_x + qm, map_y + map_h - 0.33, qw, qh,
                    sz["quadrant"], QUADRANT_LABEL, font_ea, italic=True,
                    align=PP_ALIGN.LEFT, placeholder="（左下象限）")
    _styled_textbox(slide, "QUAD_BR", map_x + map_w - qw - qm, map_y + map_h - 0.33,
                    qw, qh, sz["quadrant"], QUADRANT_LABEL, font_ea, italic=True,
                    align=PP_ALIGN.RIGHT, placeholder="（右下象限）")

    # ── 右パネル: 示唆タイトル + 固定 5 枠 ──
    _styled_textbox(slide, "IMPL_TITLE", right_x, panel_y, right_w, 0.30,
                    sz["section"], c["subtitle"], font_ea, bold=True,
                    align=section_align, placeholder="ポジショニングからの示唆")
    if spec["underline"]:
        rule2 = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(right_x), Inches(panel_y + 0.30),
            Inches(right_w), Inches(0.02))
        rule2.name = "SECTION_RIGHT_RULE"
        rule2.fill.solid()
        rule2.fill.fore_color.rgb = RGBColor.from_string(c["text"])
        rule2.line.fill.background()
        rule2.shadow.inherit = False

    body_top = panel_y + 0.50
    pitch = spec["impl_pitch"]
    for i in range(5):
        _styled_textbox(
            slide, f"IMPL_{i + 1}",
            right_x + 0.05, body_top + i * pitch, right_w - 0.10, pitch,
            sz["item"], c["text"], font_ea, align=PP_ALIGN.LEFT,
            placeholder=f"（示唆 {i + 1}）", bullet_hex=c["bullet"])

    # ── 出典（stellar はテンプレに無いので追加。roleup は既存 Source 3 を保持） ──
    if not any(s.name == "Source 3" for s in slide.shapes):
        sx, sy, sw, sh = spec["source"]
        _styled_textbox(slide, "Source 3", sx, sy, sw, sh,
                        sz["source"], c["source"], font_ea,
                        align=PP_ALIGN.LEFT, placeholder="出典：")


def _verify(path):
    prs = Presentation(path)
    s = prs.slides[0]
    print(f"  --- shapes in {os.path.basename(path)} ---")
    for shp in s.shapes:
        l = shp.left / 914400 if shp.left else 0
        t = shp.top / 914400 if shp.top else 0
        w = shp.width / 914400 if shp.width else 0
        h = shp.height / 914400 if shp.height else 0
        print(f"    {str(shp.shape_type):>14} | {shp.name!r:<22} | "
              f"x={l:.2f} y={t:.2f} w={w:.2f} h={h:.2f}")


def build_brand(brand):
    src = os.path.join(SKILL_DIR, "assets", brand,
                       f"{SKILL_NAME}-template.pptx")
    if not os.path.exists(src):
        print(f"ERROR: source template not found: {src}", file=sys.stderr)
        return 1
    prs = Presentation(src)
    build(brand, prs)
    prs.save(src)
    print(f"  ✓ Saved: {src}")
    _verify(src)
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--brand", choices=["stellar_aiz", "roleup"], default=None,
                    help="対象ブランド。省略時は両方を生成。")
    args = ap.parse_args()
    brands = [args.brand] if args.brand else ["stellar_aiz", "roleup"]
    rc = 0
    for b in brands:
        print(f"== build {b} ==")
        rc |= build_brand(b)
    return rc


if __name__ == "__main__":
    sys.exit(main())
