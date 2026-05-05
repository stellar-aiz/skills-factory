"""
fill_data_availability.py — データアベイラビリティ（調査の網羅度・制約）スライドを生成

レイアウト:
  - 上部: メインメッセージ + チャートタイトル
  - 左側（約8in）: カバレッジテーブル
      カテゴリ × 項目 × ステータス × データソース
      ステータス: ✓取得済 / △一部取得 / ✗未取得
  - 右側（約4.5in）: 調査制約事項（留意事項ブレット）
  - 下部: 出典・調査期間

Usage:
  python fill_data_availability.py \
    --data /home/claude/data_availability_data.json \
    --template <path>/data-availability-template.pptx \
    --output /mnt/user-data/outputs/DataAvailability_output.pptx
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



# ── Layout Constants ──
SHAPE_MAIN_MESSAGE = "Title 1"
SHAPE_CHART_TITLE = "Text Placeholder 2"

PANEL_Y = Inches(1.50)
PANEL_H = Inches(5.25)

# Left panel (coverage table)
LEFT_X = Inches(0.41)
LEFT_W = Inches(8.00)

# Right panel (constraints)
RIGHT_X = Inches(8.55)
RIGHT_W = Inches(4.40)

SOURCE_X = Inches(0.41)
SOURCE_Y = Inches(7.05)
SOURCE_W = Inches(12.50)

# ── Colors ──
COLOR_TEXT = RGBColor(0x33, 0x33, 0x33)
COLOR_SOURCE = RGBColor(0x66, 0x66, 0x66)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_HEADER_BG = RGBColor(0x2E, 0x4A, 0x6B)  # 紺
COLOR_ROW_ALT = RGBColor(0xF5, 0xF5, 0xF5)
COLOR_CATEGORY_BG = RGBColor(0xE8, 0xEE, 0xF4)  # 薄紺

# ステータス色
COLOR_COMPLETE = RGBColor(0x1B, 0x7A, 0x3B)   # 濃緑
COLOR_PARTIAL = RGBColor(0xDA, 0x7A, 0x2D)    # オレンジ
COLOR_MISSING = RGBColor(0xC0, 0x3A, 0x3A)    # 赤
COLOR_NA = RGBColor(0x99, 0x99, 0x99)         # グレー

# 制約パネル
COLOR_CONSTRAINT_BG = RGBColor(0xFF, 0xF9, 0xE8)  # 薄イエロー（注意喚起）
COLOR_CONSTRAINT_BORDER = RGBColor(0xDA, 0x7A, 0x2D)

FONT_NAME_JP = "Meiryo UI"
FONT_SIZE_SECTION = Pt(14)
FONT_SIZE_TABLE_HEADER = Pt(11)
FONT_SIZE_TABLE = Pt(10)
FONT_SIZE_CATEGORY = Pt(11)
FONT_SIZE_STATUS = Pt(11)
FONT_SIZE_ITEM = Pt(11)
FONT_SIZE_SOURCE = Pt(10)


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


def add_section_title(slide, text, left, top, width):
    """セクションタイトル（下線付き、14pt Bold）"""
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


def _get_status_info(status):
    """
    ステータス文字列 → (記号, 色, ラベル) を返す
    """
    if not status:
        return ("—", COLOR_NA, "未設定")
    s = str(status).lower()
    if s in ("complete", "done", "full", "取得済", "取得済み", "✓", "yes", "ok"):
        return ("✓", COLOR_COMPLETE, "取得済")
    elif s in ("partial", "part", "一部", "△", "partial", "limited"):
        return ("△", COLOR_PARTIAL, "一部取得")
    elif s in ("missing", "none", "no", "✗", "x", "未取得", "n/a"):
        return ("✗", COLOR_MISSING, "未取得")
    return ("—", COLOR_NA, status)


# ──────────────────────────────────────────────
# Left Panel: Coverage Table
# ──────────────────────────────────────────────
def build_coverage_table(slide, section_title, categories, left, top, width, height):
    """
    カテゴリ × 項目 × ステータス × データソース のテーブルを描画
    categories: [
      {
        "name": "対象会社",
        "items": [
          {"label": "会社概要", "status": "complete", "source": "公式HP"},
          ...
        ]
      },
      ...
    ]
    """
    add_section_title(slide, section_title, left, top, width)

    # テーブル領域
    tbl_top = top + Inches(0.50)
    tbl_h = height - Inches(0.50)

    # 全行を作成（カテゴリヘッダー + 項目）
    rows = []  # 各要素: ("category" | "item", data)
    for cat in categories:
        rows.append(("category", cat))
        for item in cat.get("items", []):
            rows.append(("item", item))

    n_rows = len(rows) + 1  # +1 for column header
    n_cols = 4  # 項目、ステータス記号、ステータスラベル、データソース

    # 行高
    row_h = Emu(int(tbl_h / n_rows))

    shape = slide.shapes.add_table(n_rows, n_cols, left, tbl_top, width, tbl_h)
    table = shape.table

    # 列幅
    col_w = [Inches(2.80), Inches(0.55), Inches(1.35), Inches(3.30)]
    remainder = width - sum(col_w, Emu(0))
    # 最後の列で残りを調整
    if remainder != Emu(0):
        col_w[-1] = Emu(col_w[-1] + remainder)
    for i, w in enumerate(col_w):
        table.columns[i].width = w

    # tblPr
    tbl_elem = shape._element.find('.//' + qn('a:tbl'))
    old_tblPr = tbl_elem.find(qn('a:tblPr'))
    if old_tblPr is not None:
        tbl_elem.remove(old_tblPr)
    tblPr = etree.SubElement(tbl_elem, qn('a:tblPr'), attrib={
        'firstRow': '1', 'bandRow': '0'
    })
    tbl_elem.insert(0, tblPr)

    for tr in tbl_elem.findall(qn('a:tr')):
        tr.set('h', str(row_h))

    # ヘッダー行
    headers = ["項目", "", "ステータス", "データソース"]
    for c_idx, h in enumerate(headers):
        cell = table.cell(0, c_idx)
        _style_cell(
            cell, h, cell_type="header",
            font_size=FONT_SIZE_TABLE_HEADER,
        )

    # データ行
    alt_counter = 0
    for r_idx, (row_type, row_data) in enumerate(rows):
        tr_idx = r_idx + 1  # +1 for header

        if row_type == "category":
            # カテゴリ行: 4列マージして1つの見出しセル
            # Mergeする前に各セルのスタイルを設定
            cat_name = row_data.get("name", "")
            for c_idx in range(n_cols):
                cell = table.cell(tr_idx, c_idx)
                if c_idx == 0:
                    _style_cell(
                        cell, cat_name, cell_type="category",
                        font_size=FONT_SIZE_CATEGORY,
                    )
                else:
                    _style_cell(
                        cell, "", cell_type="category",
                        font_size=FONT_SIZE_CATEGORY,
                    )
            # Mergeを試みる（python-pptxの merge メソッド）
            try:
                start_cell = table.cell(tr_idx, 0)
                end_cell = table.cell(tr_idx, n_cols - 1)
                start_cell.merge(end_cell)
            except Exception:
                pass
            alt_counter = 0

        else:  # item
            label = row_data.get("label", "")
            status = row_data.get("status", "")
            source = row_data.get("source", "")

            symbol, color, status_label = _get_status_info(status)

            is_alt = (alt_counter % 2 == 1)
            alt_counter += 1

            # 項目列
            _style_cell(
                table.cell(tr_idx, 0), label, cell_type="item",
                is_alt=is_alt, font_size=FONT_SIZE_TABLE,
                align="l",
            )

            # ステータス記号列
            _style_cell(
                table.cell(tr_idx, 1), symbol, cell_type="status",
                is_alt=is_alt, font_size=FONT_SIZE_STATUS,
                align="ctr", bold=True,
                text_color=color,
            )

            # ステータスラベル列
            _style_cell(
                table.cell(tr_idx, 2), status_label, cell_type="item",
                is_alt=is_alt, font_size=FONT_SIZE_TABLE,
                align="l",
                text_color=color,
                bold=True,
            )

            # データソース列
            _style_cell(
                table.cell(tr_idx, 3), source, cell_type="item",
                is_alt=is_alt, font_size=FONT_SIZE_TABLE,
                align="l",
            )

    print(f"  ✓ カバレッジテーブル: {len(categories)}カテゴリ、合計{n_rows - 1}行")


def _style_cell(cell, text, cell_type="item", is_alt=False, font_size=Pt(10),
                align="l", bold=False, text_color=None):
    """
    cell_type: "header" | "category" | "item" | "status"
    """
    tc = cell._tc

    # 背景色
    if cell_type == "header":
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_HEADER_BG
    elif cell_type == "category":
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_CATEGORY_BG
    elif is_alt:
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_ROW_ALT
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_WHITE

    # マージン
    cell.margin_left = Inches(0.08)
    cell.margin_right = Inches(0.08)
    cell.margin_top = Inches(0.04)
    cell.margin_bottom = Inches(0.04)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    tf = cell.text_frame
    tf.word_wrap = True

    # 既存段落クリア
    for p in list(tf.paragraphs):
        p._p.getparent().remove(p._p)

    p_elem = etree.SubElement(tf._txBody, qn("a:p"))
    pPr = etree.SubElement(p_elem, qn("a:pPr"))

    # 配置
    if cell_type == "header":
        pPr.set("algn", "ctr")
    elif cell_type == "category":
        pPr.set("algn", "l")
    else:
        pPr.set("algn", align)

    # テキスト色
    if text_color is not None:
        color_val = "{:02X}{:02X}{:02X}".format(
            text_color[0], text_color[1], text_color[2]
        )
    elif cell_type == "header":
        color_val = "FFFFFF"
    elif cell_type == "category":
        color_val = "2E4A6B"
    else:
        color_val = "333333"

    # Bold判定
    is_bold = bold or (cell_type in ("header", "category"))

    r_elem = etree.SubElement(p_elem, qn("a:r"))
    rPr = etree.SubElement(r_elem, qn("a:rPr"), attrib={
        "lang": "ja-JP",
        "sz": str(int(font_size.pt * 100)),
        "b": "1" if is_bold else "0",
    })
    etree.SubElement(rPr, qn("a:latin"), attrib={"typeface": FONT_NAME_JP})
    etree.SubElement(rPr, qn("a:ea"), attrib={"typeface": FONT_NAME_JP})
    sf = etree.SubElement(rPr, qn("a:solidFill"))
    s = etree.SubElement(sf, qn("a:srgbClr"))
    s.set("val", color_val)
    t_elem = etree.SubElement(r_elem, qn("a:t"))
    t_elem.text = text


# ──────────────────────────────────────────────
# Right Panel: Constraints
# ──────────────────────────────────────────────
def build_constraints_panel(slide, section_title, constraints,
                             left, top, width, height):
    """調査制約事項のブレット項目リスト"""
    add_section_title(slide, section_title, left, top, width)

    if not constraints:
        return

    # 薄イエロー背景の矩形
    panel_top = top + Inches(0.50)
    panel_h = height - Inches(0.50)
    panel_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, panel_top, width, panel_h,
    )
    panel_bg.fill.solid()
    panel_bg.fill.fore_color.rgb = COLOR_CONSTRAINT_BG
    panel_bg.line.color.rgb = COLOR_CONSTRAINT_BORDER
    panel_bg.line.width = Pt(0.75)
    panel_bg.shadow.inherit = False
    panel_bg.text_frame.text = ""

    # テキストボックス
    tb = slide.shapes.add_textbox(
        left + Inches(0.15), panel_top + Inches(0.12),
        width - Inches(0.30), panel_h - Inches(0.24),
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0
    tf.margin_top = 0; tf.margin_bottom = 0

    # 既存段落クリア
    for p in list(tf.paragraphs):
        p._p.getparent().remove(p._p)

    for i, item in enumerate(constraints):
        # itemは文字列またはdict
        if isinstance(item, dict):
            text = item.get("text", "")
        else:
            text = str(item)

        p_elem = etree.SubElement(tf._txBody, qn("a:p"))
        pPr = etree.SubElement(p_elem, qn("a:pPr"), attrib={
            "marL": "200000",
            "indent": "-200000",
        })
        if i > 0:
            spcBef = etree.SubElement(pPr, qn("a:spcBef"))
            etree.SubElement(spcBef, qn("a:spcPts"), attrib={"val": "500"})

        # Bullet
        buChar = etree.SubElement(pPr, qn("a:buChar"), attrib={"char": "⚠"})
        buFont = etree.SubElement(pPr, qn("a:buFont"), attrib={"typeface": "Arial"})
        buClr = etree.SubElement(pPr, qn("a:buClr"))
        buClrSolid = etree.SubElement(buClr, qn("a:srgbClr"))
        buClrSolid.set("val", "DA7A2D")

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
        t_elem = etree.SubElement(r, qn("a:t"))
        t_elem.text = text

    print(f"  ✓ 制約事項パネル: {len(constraints)}項目")


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

    prs = Presentation(args.template)
    slide = prs.slides[0]

    set_textbox_text(find_shape(slide, SHAPE_MAIN_MESSAGE), data.get("main_message", ""))
    set_textbox_text(find_shape(slide, SHAPE_CHART_TITLE), data.get("chart_title", "調査のデータ取得状況"))
    print(f"  ✓ Main Message & Chart Title set")

    # 左: カバレッジテーブル
    left_panel = data.get("left_panel", {})
    left_section_title = left_panel.get("section_title", "データ取得カバレッジ")
    categories = data.get("categories", [])
    if categories:
        build_coverage_table(
            slide, left_section_title, categories,
            LEFT_X, PANEL_Y, LEFT_W, PANEL_H,
        )
    else:
        print("  ⚠ WARNING: no categories specified", file=sys.stderr)

    # 右: 制約事項
    right_panel = data.get("right_panel", {})
    right_section_title = right_panel.get("section_title", "調査上の留意事項・制約")
    constraints = data.get("constraints", [])
    build_constraints_panel(
        slide, right_section_title, constraints,
        RIGHT_X, PANEL_Y, RIGHT_W, PANEL_H,
    )

    # 出典
    source = data.get("source", "")
    if source:
        add_text_box(
            slide, source,
            SOURCE_X, SOURCE_Y, SOURCE_W, Inches(0.25),
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
