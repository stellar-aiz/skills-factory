"""
fill_company_overview.py — 会社概要データをPPTXネイティブオブジェクトで生成するスクリプト

テンプレート構造（company-overview-template.pptx）:
  - Title 1            (PLACEHOLDER): Main Message（上段、太字）
  - Text Placeholder 2 (PLACEHOLDER): Chart Title（下段）
  - Overview Table     (TABLE):       会社概要テーブル（2列: ラベル, 値）
  - Photo Caption 1    (TEXT_BOX):    本社家屋キャプション
  - Photo Area 1       (AUTO_SHAPE):  本社家屋画像エリア
  - Photo Caption 2    (TEXT_BOX):    主要製品キャプション
  - Photo Area 2       (AUTO_SHAPE):  主要製品画像エリア
  - Source             (TEXT_BOX):    出典

使い方:
  python fill_company_overview.py \
    --data {{WORK_DIR}}/company_overview_data.json \
    --template {{SKILL_DIR}}/assets/company-overview-template.pptx \
    --output {{OUTPUT_DIR}}/CompanyOverview_output.pptx
"""

import argparse
import copy
import json
import os
import sys

# brand_resolver bootstrap (passive --brand acceptance until brand-aware migration)
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(SKILL_DIR, "..", "_common", "lib"))
from brand_resolver import add_brand_arg  # noqa: E402

from pptx import Presentation
from pptx.util import Emu, Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
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



# ── Shape名マッピング ──────────────────────────────────────
SHAPE_MAIN_MESSAGE = "Title 1"
SHAPE_CHART_TITLE  = "Text Placeholder 2"
SHAPE_TABLE        = "Overview Table"
SHAPE_CAPTION_1    = "Photo Caption 1"
SHAPE_PHOTO_1      = "Photo Area 1"
SHAPE_CAPTION_2    = "Photo Caption 2"
SHAPE_PHOTO_2      = "Photo Area 2"
SHAPE_SOURCE       = "Source"
# ────────────────────────────────────────────────────────────


def find_shape(slide, name):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    print(f"  ⚠ WARNING: Shape '{name}' not found", file=sys.stderr)
    return None


def set_textbox_text(shape, text, font_size=None, bold=None):
    """TextBox/Placeholderのテキストを設定。font_size(pt)/boldを指定するとrPrに明示的に書き込む。"""
    if shape is None:
        return
    tf = shape.text_frame
    para = tf.paragraphs[0]
    if para.runs:
        run = para.runs[0]
        run.text = text
        rPr = run._r.find(qn("a:rPr"))
        if rPr is None:
            rPr = etree.SubElement(run._r, qn("a:rPr"))
            run._r.insert(0, rPr)
        if font_size is not None:
            rPr.set("sz", str(int(font_size * 100)))
        if bold is not None:
            rPr.set("b", "1" if bold else "0")
        for r in para.runs[1:]:
            r.text = ""
    else:
        r_elem = etree.SubElement(para._p, qn("a:r"))
        attrs = {"lang": "ja-JP"}
        if font_size is not None:
            attrs["sz"] = str(int(font_size * 100))
        if bold is not None:
            attrs["b"] = "1" if bold else "0"
        etree.SubElement(r_elem, qn("a:rPr"), attrib=attrs)
        t_elem = etree.SubElement(r_elem, qn("a:t"))
        t_elem.text = text


def rebuild_table(slide, items):
    """
    テンプレートのテーブルを削除し、itemsに応じた行数でネイティブテーブルを再構築する。
    items: [{"label": "商号", "value": "株式会社〇〇"}, ...]
    """
    table_shape = find_shape(slide, SHAPE_TABLE)
    if table_shape is None:
        print("  ⚠ Table shape not found, cannot rebuild", file=sys.stderr)
        return

    old_table = table_shape.table

    # テンプレートからセルスタイル（tcPr）とフォントスタイル（rPr）をコピー
    label_tcPr = copy.deepcopy(old_table.cell(0, 0)._tc.find(qn("a:tcPr")))
    value_tcPr = copy.deepcopy(old_table.cell(0, 1)._tc.find(qn("a:tcPr")))
    label_rPr = None
    value_rPr = None
    for para in old_table.cell(0, 0).text_frame.paragraphs:
        for run in para.runs:
            label_rPr = copy.deepcopy(run._r.find(qn("a:rPr")))
            break
        break
    for para in old_table.cell(0, 1).text_frame.paragraphs:
        for run in para.runs:
            value_rPr = copy.deepcopy(run._r.find(qn("a:rPr")))
            break
        break

    # 位置・サイズを保存
    tbl_left   = table_shape.left
    tbl_top    = table_shape.top
    tbl_width  = table_shape.width
    tbl_height = table_shape.height

    # 列幅を保存
    old_col0_width = old_table.columns[0].width
    old_col1_width = old_table.columns[1].width

    # 既存テーブルを削除
    sp_tree = slide.shapes._spTree
    sp_tree.remove(table_shape._element)

    # 新テーブルを作成
    n_rows = len(items)
    n_cols = 2

    new_shape = slide.shapes.add_table(
        n_rows, n_cols, tbl_left, tbl_top, tbl_width, tbl_height
    )
    new_shape.name = SHAPE_TABLE
    new_table = new_shape.table

    # 列幅を復元
    new_table.columns[0].width = old_col0_width
    new_table.columns[1].width = old_col1_width

    # tblPr設定（bandRow有効）
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    tbl_elem = new_shape._element.find('.//a:tbl', ns)
    old_tblPr = tbl_elem.find('a:tblPr', ns)
    if old_tblPr is not None:
        tbl_elem.remove(old_tblPr)
    tblPr = etree.SubElement(tbl_elem, qn('a:tblPr'), attrib={'bandRow': '1'})
    tbl_elem.insert(0, tblPr)

    def apply_cell(cell, text, tcPr_tmpl, rPr_tmpl, is_even_row=True):
        """セルにテキストとスタイルを適用"""
        tc = cell._tc
        txBody = tc.find(qn("a:txBody"))
        if txBody is None:
            txBody = etree.SubElement(tc, qn("a:txBody"))
            etree.SubElement(txBody, qn("a:bodyPr"))
            etree.SubElement(txBody, qn("a:lstStyle"))

        # 既存段落を削除
        for p in txBody.findall(qn("a:p")):
            txBody.remove(p)

        # 改行対応: \n で分割して複数段落に
        lines = str(text).split("\n")
        for line in lines:
            p_elem = etree.SubElement(txBody, qn("a:p"))
            pPr = etree.SubElement(p_elem, qn("a:pPr"))
            pPr.set("algn", "l")
            r_elem = etree.SubElement(p_elem, qn("a:r"))
            if rPr_tmpl is not None:
                r_elem.append(copy.deepcopy(rPr_tmpl))
            else:
                etree.SubElement(r_elem, qn("a:rPr"), attrib={
                    "lang": "ja-JP", "sz": "1400"
                })
            t_elem = etree.SubElement(r_elem, qn("a:t"))
            t_elem.text = line

        # セルスタイル適用
        old_tc = tc.find(qn("a:tcPr"))
        if old_tc is not None:
            tc.remove(old_tc)
        if tcPr_tmpl is not None:
            new_tcPr = copy.deepcopy(tcPr_tmpl)
            # 行ごとの背景色切り替え
            bg_color = "F0F1F5" if is_even_row else "FFFFFF"
            for old_fill in new_tcPr.findall(qn("a:solidFill")):
                new_tcPr.remove(old_fill)
            sf = etree.SubElement(new_tcPr, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr"), attrib={"val": bg_color})
            tc.append(new_tcPr)

        # 垂直中央揃え
        tcPr_final = tc.find(qn("a:tcPr"))
        if tcPr_final is not None:
            tcPr_final.set("anchor", "ctr")

    # データ行を挿入
    for r_idx, item in enumerate(items):
        label = item.get("label", "")
        value = item.get("value", "")
        is_even = (r_idx % 2 == 0)

        apply_cell(new_table.cell(r_idx, 0), label, label_tcPr, label_rPr, is_even)
        apply_cell(new_table.cell(r_idx, 1), value, value_tcPr, value_rPr, is_even)

    print(f"  ✓ Table rebuilt: {n_rows} rows x {n_cols} cols")


def insert_photo(slide, photo_info, area_shape_name, caption_shape_name):
    """写真をエリアに挿入、キャプションを更新"""
    if not photo_info:
        return

    # キャプション更新
    caption = photo_info.get("caption", "")
    if caption:
        cap_shape = find_shape(slide, caption_shape_name)
        if cap_shape:
            set_textbox_text(cap_shape, caption)

    # 画像ファイルパスの取得
    img_path = photo_info.get("url") or photo_info.get("path", "")
    if not img_path or not os.path.exists(img_path):
        print(f"  ℹ No image for {caption_shape_name} (placeholder retained)")
        return

    # 画像エリアの位置・サイズを取得
    area_shape = find_shape(slide, area_shape_name)
    if area_shape is None:
        return

    left = area_shape.left
    top = area_shape.top
    width = area_shape.width
    height = area_shape.height

    # 画像を挿入（エリアを覆う）
    slide.shapes.add_picture(img_path, left, top, width, height)
    print(f"  ✓ Photo inserted: {img_path}")


def main():
    parser = argparse.ArgumentParser(description="会社概要スライド生成（ネイティブテーブル方式）")
    parser.add_argument("--data", required=True, help="JSONデータファイルパス")
    parser.add_argument("--template", required=True, help="PPTXテンプレートパス")
    parser.add_argument("--output", required=True, help="出力PPTXファイルパス")
    add_brand_arg(parser)  # passive: accepted but ignored until brand migration
    args = parser.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    print("=== 会社概要スライド生成（ネイティブテーブル方式） ===")

    prs = Presentation(args.template)
    slide = prs.slides[0]

    # Shape一覧（デバッグ用）
    for s in slide.shapes:
        print(f"  Shape: '{s.name}' type={s.shape_type}")

    # 1. Main Message（Title 1 = 上段太字、最大65文字）
    main_msg = data.get("main_message", "")
    if len(main_msg) > 65:
        print(f"  ⚠ WARNING: Main Message が{len(main_msg)}文字です（上限65文字）。スライドに収まらない可能性があります。", file=sys.stderr)
    set_textbox_text(find_shape(slide, SHAPE_MAIN_MESSAGE), main_msg, font_size=26, bold=True)
    print(f"  [Main Message] ({len(main_msg)}文字) {main_msg}")

    # 2. Chart Title（Text Placeholder 2 = 下段、10〜20文字）
    title_text = data.get("title", "対象会社概要：会社概要")
    if len(title_text) > 20:
        print(f"  ⚠ WARNING: Chart Title が{len(title_text)}文字です（推奨10〜20文字）。", file=sys.stderr)
    set_textbox_text(find_shape(slide, SHAPE_CHART_TITLE), title_text, font_size=18, bold=True)
    print(f"  [Chart Title]  ({len(title_text)}文字) {title_text}")

    # 3. 出典
    source_text = data.get("source", "")
    source_shape = find_shape(slide, SHAPE_SOURCE)
    if source_text:
        set_textbox_text(source_shape, f"出典：{source_text}")
    else:
        set_textbox_text(source_shape, "")
    print(f"  [Source]        {source_text}")

    # 4. ネイティブテーブル
    items = data.get("items", [])
    if items:
        rebuild_table(slide, items)

    # 5. 写真
    photos = data.get("photos", {})
    insert_photo(slide, photos.get("headquarters"), SHAPE_PHOTO_1, SHAPE_CAPTION_1)
    insert_photo(slide, photos.get("product"), SHAPE_PHOTO_2, SHAPE_CAPTION_2)

    # 保存
    os.makedirs(os.path.dirname(args.output) or ".", exist_ok=True)
    prs.save(args.output)
    _finalize_pptx(args.output)
    print(f"  ✓ PPTX saved: {args.output}")
    print("=== 完了 ===")


if __name__ == "__main__":
    main()
