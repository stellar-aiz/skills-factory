"""
fill_value_chain_matrix.py — バリューチェーン・ポジショニング・マトリクスを
PPTXネイティブオブジェクトで生成。

レイアウト:
  - 上部: メインメッセージ (28pt Bold) + チャートタイトル (14pt)
  - 中央: シェブロン行 + 4行のマトリクス
    * 行ラベル列（左、1.90")
    * N段階のシェブロン（5〜7段階、幅自動調整）
    * 行1,2,4: 記号セル (◎ ○ △ ー 空白)
    * 行3: プレーヤーバー（複数列またぎ）
  - 下部: 出典 (10pt グレー)

全オブジェクトはPPT上で編集可能（画像ではない）。

Usage:
  python fill_value_chain_matrix.py \\
    --data /home/claude/value_chain_matrix_data.json \\
    --template <SKILL_DIR>/assets/value-chain-matrix-template.pptx \\
    --output /mnt/user-data/outputs/ValueChainMatrix_output.pptx
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



# ── Shape名マッピング（テンプレート上） ──
SHAPE_MAIN_MESSAGE = "Title 1"
SHAPE_CHART_TITLE = "Text Placeholder 2"

# ── スライド全体 ──
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.50)

# ── レイアウト定数 ──
CONTENT_LEFT = Inches(0.41)
CONTENT_WIDTH = Inches(12.52)

ROW_LABEL_WIDTH = Inches(1.90)
STAGES_LEFT = CONTENT_LEFT + ROW_LABEL_WIDTH          # Inches(2.31)
STAGES_WIDTH = CONTENT_WIDTH - ROW_LABEL_WIDTH        # Inches(10.62)

CHEVRON_TOP = Inches(1.50)
CHEVRON_HEIGHT = Inches(0.75)

GAP = Inches(0.06)
DATA_ROW_HEIGHT = Inches(1.00)

# 4 rows start Y: chevron_top + chevron_height + gap
DATA_ROWS_TOP = CHEVRON_TOP + CHEVRON_HEIGHT + GAP   # 1.50 + 0.75 + 0.06 = 2.31
# End of 4th row: 2.31 + 4*1.00 = 6.31 (below source Y=6.85 ✓)

# Source — マスター要素（© 表記 Y=7.18、スライド番号 Y=7.18）を避けて配置
SOURCE_X = CONTENT_LEFT
SOURCE_Y = Inches(6.85)
SOURCE_W = CONTENT_WIDTH
SOURCE_H = Inches(0.28)

# ── 色 ──
COLOR_TEXT = RGBColor(0x33, 0x33, 0x33)
COLOR_NAVY = RGBColor(0x1E, 0x3A, 0x5F)           # シェブロン枠線・テキスト・記号
COLOR_CHEVRON_FILL = RGBColor(0xCE, 0xE0, 0xED)   # シェブロン塗り（薄いブルー）
COLOR_BAR_FILL = RGBColor(0xB4, 0xC4, 0xD4)       # プレーヤーバー塗り
COLOR_BAR_BORDER = RGBColor(0x5A, 0x70, 0x90)     # プレーヤーバー枠線
COLOR_SOURCE = RGBColor(0x66, 0x66, 0x66)
COLOR_SEPARATOR = RGBColor(0xBB, 0xBB, 0xBB)      # 行間セパレーターライン
COLOR_SEPARATOR_MAIN = RGBColor(0x66, 0x66, 0x66) # シェブロン下の太いライン

# ── フォント ──
FONT_NAME_JP = "Meiryo UI"
FONT_NAME_LATIN = "Arial"

# フォントサイズ: 列数に応じて自動調整
def calc_font_sizes(n_stages):
    """列数に応じてシェブロン・記号・バーのフォントサイズを自動決定"""
    if n_stages <= 5:
        return {"chevron": 14, "chevron_sub": 10, "symbol": 36, "bar": 14, "row_label": 12}
    elif n_stages == 6:
        return {"chevron": 13, "chevron_sub": 9, "symbol": 32, "bar": 13, "row_label": 12}
    else:  # 7
        return {"chevron": 12, "chevron_sub": 9, "symbol": 28, "bar": 12, "row_label": 11}


# ════════════════════════════════════════════════════════════════
# ユーティリティ
# ════════════════════════════════════════════════════════════════

def find_shape(slide, name):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    print(f"  ⚠ WARNING: Shape '{name}' not found", file=sys.stderr)
    return None


def set_placeholder_text(shape, text, font_size_pt=None, bold=None, color=None):
    """既存placeholderのテキストを置き換え（スタイルを指定可能）"""
    if shape is None:
        return
    tf = shape.text_frame
    # 既存段落を全削除し、新規段落を追加
    p_elem = tf.paragraphs[0]._p
    # 既存runを削除
    for r in p_elem.findall(qn("a:r")):
        p_elem.remove(r)
    # 新規run
    r = etree.SubElement(p_elem, qn("a:r"))
    rPr_attrs = {"lang": "ja-JP"}
    if font_size_pt is not None:
        rPr_attrs["sz"] = str(int(font_size_pt * 100))
    if bold is not None:
        rPr_attrs["b"] = "1" if bold else "0"
    rPr = etree.SubElement(r, qn("a:rPr"), attrib=rPr_attrs)
    if color is not None:
        solidFill = etree.SubElement(rPr, qn("a:solidFill"))
        etree.SubElement(solidFill, qn("a:srgbClr"),
                         attrib={"val": "{:02X}{:02X}{:02X}".format(color[0], color[1], color[2])})
    etree.SubElement(rPr, qn("a:latin"), attrib={"typeface": FONT_NAME_LATIN})
    etree.SubElement(rPr, qn("a:ea"), attrib={"typeface": FONT_NAME_JP})
    t = etree.SubElement(r, qn("a:t"))
    t.text = text


def add_textbox(slide, left, top, width, height, text, *,
                font_size_pt=12, bold=False, color=(0x33, 0x33, 0x33),
                align="center", vertical="middle", word_wrap=True,
                font_name_jp=FONT_NAME_JP, font_name_latin=FONT_NAME_LATIN):
    """指定位置にテキストボックスを追加（汎用）"""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = word_wrap
    # Margins
    tf.margin_left = Emu(36000)
    tf.margin_right = Emu(36000)
    tf.margin_top = Emu(18000)
    tf.margin_bottom = Emu(18000)
    # Vertical anchor
    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    if bodyPr is not None:
        anchor_map = {"top": "t", "middle": "ctr", "bottom": "b"}
        bodyPr.set("anchor", anchor_map.get(vertical, "ctr"))

    # Paragraphs: split on newlines
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
        p.alignment = align_map.get(align, PP_ALIGN.CENTER)
        # Remove default runs
        for r in p._p.findall(qn("a:r")):
            p._p.remove(r)
        # Set line spacing
        pPr = p._p.find(qn("a:pPr"))
        if pPr is None:
            pPr = etree.SubElement(p._p, qn("a:pPr"))
            p._p.insert(0, pPr)
        # Clear any existing lnSpc
        for x in pPr.findall(qn("a:lnSpc")):
            pPr.remove(x)
        lnSpc = etree.SubElement(pPr, qn("a:lnSpc"))
        etree.SubElement(lnSpc, qn("a:spcPct"), attrib={"val": "100000"})
        # New run
        r = etree.SubElement(p._p, qn("a:r"))
        rPr = etree.SubElement(r, qn("a:rPr"), attrib={
            "lang": "ja-JP",
            "sz": str(int(font_size_pt * 100)),
            "b": "1" if bold else "0",
        })
        solidFill = etree.SubElement(rPr, qn("a:solidFill"))
        etree.SubElement(solidFill, qn("a:srgbClr"),
                         attrib={"val": "{:02X}{:02X}{:02X}".format(color[0], color[1], color[2])})
        etree.SubElement(rPr, qn("a:latin"), attrib={"typeface": font_name_latin})
        etree.SubElement(rPr, qn("a:ea"), attrib={"typeface": font_name_jp})
        t = etree.SubElement(r, qn("a:t"))
        t.text = line
    return tb


def _add_styled_run(p_elem, text, font_size_pt, bold, text_color):
    """段落要素にスタイル付きrunを追加（共通ヘルパー）"""
    r = etree.SubElement(p_elem, qn("a:r"))
    rPr = etree.SubElement(r, qn("a:rPr"), attrib={
        "lang": "ja-JP",
        "sz": str(int(font_size_pt * 100)),
        "b": "1" if bold else "0",
    })
    solidFill = etree.SubElement(rPr, qn("a:solidFill"))
    etree.SubElement(solidFill, qn("a:srgbClr"),
                     attrib={"val": "{:02X}{:02X}{:02X}".format(text_color[0], text_color[1], text_color[2])})
    etree.SubElement(rPr, qn("a:latin"), attrib={"typeface": FONT_NAME_LATIN})
    etree.SubElement(rPr, qn("a:ea"), attrib={"typeface": FONT_NAME_JP})
    t = etree.SubElement(r, qn("a:t"))
    t.text = text


def _prepare_paragraph(p):
    """段落から既存runを削除し、行間・中央揃え設定を行う"""
    for r in p._p.findall(qn("a:r")):
        p._p.remove(r)
    pPr = p._p.find(qn("a:pPr"))
    if pPr is None:
        pPr = etree.SubElement(p._p, qn("a:pPr"))
        p._p.insert(0, pPr)
    for x in pPr.findall(qn("a:lnSpc")):
        pPr.remove(x)
    lnSpc = etree.SubElement(pPr, qn("a:lnSpc"))
    etree.SubElement(lnSpc, qn("a:spcPct"), attrib={"val": "100000"})
    pPr.set("algn", "ctr")


def _set_chevron_point_depth(shape, depth_val=25000):
    """CHEVRON形状の矢尻の鋭さを調整
    0=長方形に近い、50000=デフォルト、100000=最大。小さいほど文字領域が広い。"""
    spPr = shape._element.find(".//" + qn("p:spPr"))
    if spPr is None:
        return
    prstGeom = spPr.find(qn("a:prstGeom"))
    if prstGeom is None:
        return
    avLst = prstGeom.find(qn("a:avLst"))
    if avLst is None:
        avLst = etree.SubElement(prstGeom, qn("a:avLst"))
    for gd in avLst.findall(qn("a:gd")):
        avLst.remove(gd)
    etree.SubElement(avLst, qn("a:gd"), attrib={"name": "adj", "fmla": f"val {depth_val}"})


def add_shape_with_text(slide, shape_enum, left, top, width, height,
                        text, *, fill_color, border_color, text_color,
                        font_size_pt, bold=True, vertical="middle",
                        border_width_pt=1.0, sub_text=None, sub_font_size_pt=None,
                        margin_h_emu=9000, margin_v_emu=9000,
                        chevron_depth=None):
    """図形を追加してテキストを入れる（シェブロン、長方形 etc.）
    - text は \\n で改行可能（複数段落として描画）
    - chevron_depth: CHEVRON 形状の矢尻深さ（0〜100000、デフォルト 25000 で浅め）
    """
    sh = slide.shapes.add_shape(shape_enum, left, top, width, height)
    if shape_enum == MSO_SHAPE.CHEVRON:
        _set_chevron_point_depth(sh, chevron_depth if chevron_depth is not None else 25000)
    # Fill
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill_color
    # Border
    sh.line.color.rgb = border_color
    sh.line.width = Pt(border_width_pt)
    # Text
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(margin_h_emu)
    tf.margin_right = Emu(margin_h_emu)
    tf.margin_top = Emu(margin_v_emu)
    tf.margin_bottom = Emu(margin_v_emu)
    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    if bodyPr is not None:
        anchor_map = {"top": "t", "middle": "ctr", "bottom": "b"}
        bodyPr.set("anchor", anchor_map.get(vertical, "ctr"))

    # Main text: \n で複数段落に分割
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        _prepare_paragraph(p)
        _add_styled_run(p._p, line, font_size_pt, bold, text_color)

    # Optional sub text
    if sub_text:
        sub_sz = sub_font_size_pt or max(int(font_size_pt * 0.75), 8)
        for line in str(sub_text).split("\n"):
            p2 = tf.add_paragraph()
            _prepare_paragraph(p2)
            _add_styled_run(p2._p, line, sub_sz, False, text_color)

    return sh


def add_horizontal_line(slide, left, top, width, color, thickness_pt=0.75):
    """水平セパレーターライン"""
    sh = slide.shapes.add_connector(1, left, top, left + width, top)  # 1 = STRAIGHT
    sh.line.color.rgb = color
    sh.line.width = Pt(thickness_pt)
    return sh


# ════════════════════════════════════════════════════════════════
# マトリクス構築
# ════════════════════════════════════════════════════════════════

def build_matrix(slide, stages, rows, font_sizes):
    """シェブロン行 + 4行のマトリクスを描画"""
    n_stages = len(stages)
    col_w = STAGES_WIDTH / n_stages

    # ── シェブロン行 ──
    # 各シェブロンは 0.10" のオーバーラップを持たせる（視覚的連続性のため）
    overlap = Inches(0.12)
    for i, stage in enumerate(stages):
        cx = STAGES_LEFT + col_w * i
        # 最後のシェブロン以外はオーバーラップ分広く
        if i < n_stages - 1:
            cw = col_w + overlap
        else:
            cw = col_w
        add_shape_with_text(
            slide, MSO_SHAPE.CHEVRON,
            cx, CHEVRON_TOP, cw, CHEVRON_HEIGHT,
            stage.get("name", ""),
            fill_color=COLOR_CHEVRON_FILL,
            border_color=COLOR_NAVY,
            text_color=(COLOR_NAVY[0], COLOR_NAVY[1], COLOR_NAVY[2]),
            font_size_pt=font_sizes["chevron"],
            bold=True,
            border_width_pt=1.0,
            sub_text=stage.get("sub"),
            sub_font_size_pt=font_sizes["chevron_sub"],
        )

    # シェブロン行下の太いセパレーター
    sep_y = CHEVRON_TOP + CHEVRON_HEIGHT + Inches(0.02)
    add_horizontal_line(slide, CONTENT_LEFT, sep_y, CONTENT_WIDTH,
                        COLOR_SEPARATOR_MAIN, thickness_pt=1.5)

    # ── 4行のマトリクス ──
    if len(rows) != 4:
        print(f"  ⚠ WARNING: expected 4 rows, got {len(rows)}", file=sys.stderr)

    current_y = DATA_ROWS_TOP
    for row_idx, row in enumerate(rows):
        row_label = row.get("label", "")
        row_type = row.get("type", "symbols")

        # ─ 行ラベル（左列、中央揃え） ─
        add_textbox(
            slide, CONTENT_LEFT, current_y, ROW_LABEL_WIDTH, DATA_ROW_HEIGHT,
            row_label,
            font_size_pt=font_sizes["row_label"], bold=False,
            color=(COLOR_TEXT[0], COLOR_TEXT[1], COLOR_TEXT[2]),
            align="center", vertical="middle", word_wrap=True,
        )

        # ─ 行コンテンツ ─
        if row_type == "symbols":
            cells = row.get("cells", [])
            # セル数が段階数と合わなければ空白で埋める
            while len(cells) < n_stages:
                cells.append("")
            for c_idx in range(n_stages):
                cell_text = str(cells[c_idx] or "").strip()
                if cell_text:
                    cx = STAGES_LEFT + col_w * c_idx
                    add_textbox(
                        slide, cx, current_y, col_w, DATA_ROW_HEIGHT,
                        cell_text,
                        font_size_pt=font_sizes["symbol"], bold=True,
                        color=(COLOR_NAVY[0], COLOR_NAVY[1], COLOR_NAVY[2]),
                        align="center", vertical="middle",
                    )

        elif row_type == "bars":
            bars = row.get("bars", [])
            for bar in bars:
                s = int(bar.get("span_start", 0))
                e = int(bar.get("span_end", s))
                s = max(0, min(s, n_stages - 1))
                e = max(s, min(e, n_stages - 1))
                bar_label = bar.get("label", "")

                # バーの幅・位置（セル内パディング）
                bar_pad_x = Inches(0.08)
                bar_pad_y = Inches(0.22)
                bar_x = STAGES_LEFT + col_w * s + bar_pad_x
                bar_w = col_w * (e - s + 1) - bar_pad_x * 2
                bar_y = current_y + bar_pad_y
                bar_h = DATA_ROW_HEIGHT - bar_pad_y * 2

                add_shape_with_text(
                    slide, MSO_SHAPE.RECTANGLE,
                    bar_x, bar_y, bar_w, bar_h,
                    bar_label,
                    fill_color=COLOR_BAR_FILL,
                    border_color=COLOR_BAR_BORDER,
                    text_color=(COLOR_NAVY[0], COLOR_NAVY[1], COLOR_NAVY[2]),
                    font_size_pt=font_sizes["bar"],
                    bold=True,
                    border_width_pt=0.75,
                )

        # ─ 行間セパレーターライン（最後の行以外） ─
        if row_idx < len(rows) - 1:
            line_y = current_y + DATA_ROW_HEIGHT
            add_horizontal_line(slide, CONTENT_LEFT, line_y, CONTENT_WIDTH,
                                COLOR_SEPARATOR, thickness_pt=0.5)

        current_y = current_y + DATA_ROW_HEIGHT

    # 最終行の下にセパレーターライン
    add_horizontal_line(slide, CONTENT_LEFT, current_y, CONTENT_WIDTH,
                        COLOR_SEPARATOR_MAIN, thickness_pt=1.0)


# ════════════════════════════════════════════════════════════════
# メイン処理
# ════════════════════════════════════════════════════════════════

def fill_slide(data, template_path, output_path):
    prs = Presentation(template_path)
    slide = prs.slides[0]

    # ── メインメッセージ ──
    main_message = data.get("main_message", "")
    msg_shape = find_shape(slide, SHAPE_MAIN_MESSAGE)
    set_placeholder_text(msg_shape, main_message,
                         font_size_pt=28, bold=True,
                         color=(COLOR_TEXT[0], COLOR_TEXT[1], COLOR_TEXT[2]))
    print(f"  ✓ メインメッセージ: {main_message[:40]}...")

    # ── チャートタイトル ──
    chart_title = data.get("chart_title", "バリューチェーン・ポジショニング・マトリクス")
    title_shape = find_shape(slide, SHAPE_CHART_TITLE)
    set_placeholder_text(title_shape, chart_title,
                         font_size_pt=14, bold=False,
                         color=(COLOR_TEXT[0], COLOR_TEXT[1], COLOR_TEXT[2]))
    print(f"  ✓ チャートタイトル: {chart_title}")

    # ── マトリクス本体 ──
    stages = data.get("stages", [])
    rows = data.get("rows", [])
    if not stages:
        raise ValueError("stages は必須です")
    if len(stages) < 5 or len(stages) > 7:
        print(f"  ⚠ WARNING: stages数は5〜7推奨（現在 {len(stages)}）", file=sys.stderr)
    if not rows:
        raise ValueError("rows は必須です")

    font_sizes = calc_font_sizes(len(stages))
    print(f"  ✓ 段階数: {len(stages)} → フォントサイズ: {font_sizes}")

    build_matrix(slide, stages, rows, font_sizes)
    print(f"  ✓ マトリクス構築完了")

    # ── 出典 ──
    source = data.get("source", "")
    if source:
        add_textbox(
            slide, SOURCE_X, SOURCE_Y, SOURCE_W, SOURCE_H,
            source,
            font_size_pt=10, bold=False,
            color=(COLOR_SOURCE[0], COLOR_SOURCE[1], COLOR_SOURCE[2]),
            align="left", vertical="top",
        )
        print(f"  ✓ 出典")

    # ── 保存 ──
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    _finalize_pptx(output_path)
    print(f"\n✓ 保存: {output_path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True, help="JSON data path")
    ap.add_argument("--template", required=True, help="PPTX template path")
    ap.add_argument("--output", required=True, help="Output PPTX path")
    add_brand_arg(ap)  # passive: accepted but ignored until brand migration
    args = ap.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    print("=" * 60)
    print("バリューチェーン・ポジショニング・マトリクス生成")
    print("=" * 60)
    fill_slide(data, args.template, args.output)


if __name__ == "__main__":
    main()
