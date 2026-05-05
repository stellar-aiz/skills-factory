"""
fill_section_divider.py — 中扉（Section Divider）スライドを生成

レイアウト:
  - 上部1/3: 大きなセクション番号（薄色背景）
  - 中央: セクションタイトル（大）
  - サブタイトル（中）
  - 下部: そのセクションで扱うトピックリスト（オプション）

レポート内のセクション区切りに使う。

Usage:
  python fill_section_divider.py \
    --data /home/claude/section_divider_data.json \
    --template <path>/section-divider-pptx-template.pptx \
    --output /mnt/user-data/outputs/SectionDivider_output.pptx
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
from pptx.enum.shapes import MSO_SHAPE
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



SHAPE_MAIN_MESSAGE = "Title 1"
SHAPE_CHART_TITLE = "Text Placeholder 2"

# 中扉は全画面を活用
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.50)

COLOR_TEXT = RGBColor(0x33, 0x33, 0x33)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_SUBTEXT = RGBColor(0x66, 0x66, 0x66)

# セクション色のローテーション（TOCと統一）
SECTION_COLORS = [
    RGBColor(0x2E, 0x4A, 0x6B),   # 紺
    RGBColor(0x7B, 0x4F, 0xB0),   # 紫
    RGBColor(0x2E, 0x6F, 0xBF),   # 青
    RGBColor(0x3D, 0x8F, 0x5A),   # 緑
    RGBColor(0xDA, 0x7A, 0x2D),   # オレンジ
    RGBColor(0xC0, 0x3A, 0x3A),   # 赤
    RGBColor(0x59, 0x59, 0x59),   # グレー
]

FONT_NAME_JP = "Meiryo UI"
FONT_SIZE_BIG_NUMBER = Pt(180)   # 巨大なセクション番号
FONT_SIZE_SECTION_TITLE = Pt(36)
FONT_SIZE_SUBTITLE = Pt(18)
FONT_SIZE_TOPIC = Pt(13)


def find_shape(slide, name):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def add_text_box(slide, text, left, top, width, height, font_size, bold=False,
                 color=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                 font_name=FONT_NAME_JP):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0
    tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.bold = bold
    run.font.name = font_name
    if color is not None:
        run.font.color.rgb = color
    else:
        run.font.color.rgb = COLOR_TEXT
    return tb


def hex_to_rgb(hex_str):
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def remove_template_shapes(slide):
    """テンプレートのプレースホルダー（Title 1, Text Placeholder 2）を削除して中扉を一から組む"""
    for sh in list(slide.shapes):
        if sh.name in (SHAPE_MAIN_MESSAGE, SHAPE_CHART_TITLE):
            sp_tree = slide.shapes._spTree
            sp_tree.remove(sh._element)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--output", required=True)
    add_brand_arg(ap)  # passive: accepted but ignored until brand migration
    args = ap.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    prs = Presentation(args.template)
    slide = prs.slides[0]

    # テンプレートのプレースホルダー削除（中扉は専用デザイン）
    remove_template_shapes(slide)

    # セクション番号と色を取得
    section_number = data.get("section_number", 1)
    color_hex = data.get("color")
    if color_hex:
        accent_color = hex_to_rgb(color_hex)
    else:
        accent_color = SECTION_COLORS[(section_number - 1) % len(SECTION_COLORS)]

    # ── 左半分: 巨大なセクション番号 + アクセントカラー背景 ──
    left_panel_w = Inches(5.50)
    left_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Emu(0), Emu(0), left_panel_w, SLIDE_H,
    )
    left_bg.fill.solid()
    left_bg.fill.fore_color.rgb = accent_color
    left_bg.line.fill.background()
    left_bg.shadow.inherit = False
    left_bg.text_frame.text = ""

    # 巨大な番号（中央）
    num_tb = slide.shapes.add_textbox(
        Emu(0), Inches(1.5), left_panel_w, Inches(4.5),
    )
    ntf = num_tb.text_frame
    ntf.margin_left = 0; ntf.margin_right = 0
    ntf.margin_top = 0; ntf.margin_bottom = 0
    ntf.vertical_anchor = MSO_ANCHOR.MIDDLE

    for p in list(ntf.paragraphs):
        p._p.getparent().remove(p._p)

    p_elem = etree.SubElement(ntf._txBody, qn("a:p"))
    pPr = etree.SubElement(p_elem, qn("a:pPr"))
    pPr.set("algn", "ctr")
    r = etree.SubElement(p_elem, qn("a:r"))
    rPr = etree.SubElement(r, qn("a:rPr"), attrib={
        "lang": "en-US",
        "sz": str(int(FONT_SIZE_BIG_NUMBER.pt * 100)),
        "b": "1",
    })
    etree.SubElement(rPr, qn("a:latin"), attrib={"typeface": "Arial"})
    sf = etree.SubElement(rPr, qn("a:solidFill"))
    s = etree.SubElement(sf, qn("a:srgbClr"))
    s.set("val", "FFFFFF")
    t = etree.SubElement(r, qn("a:t"))
    t.text = f"{section_number:02d}"

    # 番号上の小ラベル "SECTION"
    label_tb = slide.shapes.add_textbox(
        Emu(0), Inches(2.6), left_panel_w, Inches(0.4),
    )
    ltf = label_tb.text_frame
    ltf.margin_left = 0; ltf.margin_right = 0
    ltf.margin_top = 0; ltf.margin_bottom = 0
    p_lbl = ltf.paragraphs[0]
    p_lbl.alignment = PP_ALIGN.CENTER
    r_lbl = p_lbl.add_run()
    r_lbl.text = "SECTION"
    r_lbl.font.size = Pt(20)
    r_lbl.font.bold = True
    r_lbl.font.name = "Arial"
    r_lbl.font.color.rgb = COLOR_WHITE

    # ── 右半分: タイトル・サブタイトル・トピック ──
    right_x = left_panel_w + Inches(0.50)
    right_w = SLIDE_W - right_x - Inches(0.50)

    # タイトル
    title = data.get("title", "セクションタイトル")
    title_tb = slide.shapes.add_textbox(
        right_x, Inches(2.30), right_w, Inches(1.20),
    )
    ttf = title_tb.text_frame
    ttf.word_wrap = True
    ttf.margin_left = 0; ttf.margin_right = 0
    ttf.margin_top = 0; ttf.margin_bottom = 0
    ttf.vertical_anchor = MSO_ANCHOR.TOP

    p_t = ttf.paragraphs[0]
    p_t.alignment = PP_ALIGN.LEFT
    r_t = p_t.add_run()
    r_t.text = title
    r_t.font.size = FONT_SIZE_SECTION_TITLE
    r_t.font.bold = True
    r_t.font.name = FONT_NAME_JP
    r_t.font.color.rgb = COLOR_TEXT

    # アクセントライン（タイトルの下）
    accent_line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        right_x, Inches(3.35),
        Inches(1.0), Emu(int(Inches(0.06))),
    )
    accent_line.fill.solid()
    accent_line.fill.fore_color.rgb = accent_color
    accent_line.line.fill.background()

    # サブタイトル
    subtitle = data.get("subtitle", "")
    if subtitle:
        sub_tb = slide.shapes.add_textbox(
            right_x, Inches(3.55), right_w, Inches(0.50),
        )
        stf = sub_tb.text_frame
        stf.word_wrap = True
        stf.margin_left = 0; stf.margin_right = 0
        stf.margin_top = 0; stf.margin_bottom = 0
        p_s = stf.paragraphs[0]
        p_s.alignment = PP_ALIGN.LEFT
        r_s = p_s.add_run()
        r_s.text = subtitle
        r_s.font.size = FONT_SIZE_SUBTITLE
        r_s.font.bold = False
        r_s.font.name = FONT_NAME_JP
        r_s.font.color.rgb = COLOR_SUBTEXT

    # トピックリスト
    topics = data.get("topics", [])
    if topics:
        topics_y = Inches(4.30)
        topic_label_tb = slide.shapes.add_textbox(
            right_x, topics_y, right_w, Inches(0.30),
        )
        tltf = topic_label_tb.text_frame
        tltf.margin_left = 0; tltf.margin_right = 0
        tltf.margin_top = 0; tltf.margin_bottom = 0
        p_tl = tltf.paragraphs[0]
        p_tl.alignment = PP_ALIGN.LEFT
        r_tl = p_tl.add_run()
        r_tl.text = "▍ このセクションで扱う内容"
        r_tl.font.size = Pt(13)
        r_tl.font.bold = True
        r_tl.font.name = FONT_NAME_JP
        r_tl.font.color.rgb = accent_color

        # トピック項目
        topic_tb = slide.shapes.add_textbox(
            right_x + Inches(0.10), topics_y + Inches(0.40),
            right_w - Inches(0.10), Inches(2.0),
        )
        topic_tf = topic_tb.text_frame
        topic_tf.word_wrap = True
        topic_tf.margin_left = 0; topic_tf.margin_right = 0
        topic_tf.margin_top = 0; topic_tf.margin_bottom = 0

        for p in list(topic_tf.paragraphs):
            p._p.getparent().remove(p._p)

        for i, topic in enumerate(topics):
            p_topic = etree.SubElement(topic_tf._txBody, qn("a:p"))
            pPr = etree.SubElement(p_topic, qn("a:pPr"), attrib={
                "marL": "200000",
                "indent": "-200000",
            })
            if i > 0:
                spcBef = etree.SubElement(pPr, qn("a:spcBef"))
                etree.SubElement(spcBef, qn("a:spcPts"), attrib={"val": "400"})

            buChar = etree.SubElement(pPr, qn("a:buChar"), attrib={"char": "▸"})
            buFont = etree.SubElement(pPr, qn("a:buFont"), attrib={"typeface": "Arial"})
            buClr = etree.SubElement(pPr, qn("a:buClr"))
            buClrSolid = etree.SubElement(buClr, qn("a:srgbClr"))
            buClrSolid.set("val", "{:02X}{:02X}{:02X}".format(accent_color[0], accent_color[1], accent_color[2]))

            r_top = etree.SubElement(p_topic, qn("a:r"))
            rPr_top = etree.SubElement(r_top, qn("a:rPr"), attrib={
                "lang": "ja-JP",
                "sz": str(int(FONT_SIZE_TOPIC.pt * 100)),
            })
            etree.SubElement(rPr_top, qn("a:latin"), attrib={"typeface": FONT_NAME_JP})
            etree.SubElement(rPr_top, qn("a:ea"), attrib={"typeface": FONT_NAME_JP})
            sf_top = etree.SubElement(rPr_top, qn("a:solidFill"))
            s_top = etree.SubElement(sf_top, qn("a:srgbClr"))
            s_top.set("val", "333333")
            t_top = etree.SubElement(r_top, qn("a:t"))
            t_top.text = topic

    print(f"  ✓ Section {section_number:02d}: {title}")

    os.makedirs(os.path.dirname(args.output), exist_ok=True)
    prs.save(args.output)
    _finalize_pptx(args.output)
    print(f"\n✅ Saved: {args.output}")


if __name__ == "__main__":
    main()
