"""
fill_table_chart.py — テーブル＋意味合いスライドをPPTXテンプレートに流し込むスクリプト

テンプレート構造（TableChart.pptx 確認済み）:
  - Title 1              (PLACEHOLDER): Main Message
  - Text Placeholder 2   (PLACEHOLDER): Chart Title
  - TextBox 6            (TEXT_BOX):    テーブルセクションラベル（動的変更可）
  - Straight Connector 26(LINE):        水平線（編集不要）
  - TextBox 29           (TEXT_BOX):    意味合い（タイトル + 1〜5個のBullet）
  - Table 10             (TABLE):       動的テーブル（行・列数はコンテンツに応じて増減）

使い方:
  python fill_table_chart.py \
    --data /home/claude/table_chart_data.json \
    --template /path/to/table-chart-template.pptx \
    --output /mnt/user-data/outputs/TableChart_output.pptx
"""

import argparse
import os
import json
import sys
import copy

# brand_resolver bootstrap (passive --brand acceptance until brand-aware migration)
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(SKILL_DIR, "..", "_common", "lib"))
from brand_resolver import add_brand_arg  # noqa: E402
from validate_fill_input import validate_fill_input  # noqa: E402
from pptx import Presentation
from pptx.util import Pt, Emu
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



# ── Shape名マッピング（確認済み）──────────────────────────────
SHAPE_MAIN_MESSAGE   = "Title 1"
SHAPE_CHART_TITLE    = "Text Placeholder 2"
SHAPE_TABLE_LABEL    = "TextBox 6"
SHAPE_IMPLICATIONS   = "TextBox 29"
SHAPE_TABLE          = "Table 10"
# ────────────────────────────────────────────────────────────


def find_shape(slide, name):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    print(f"  WARNING: Shape '{name}' not found", file=sys.stderr)
    return None


def set_placeholder_text(shape, text):
    """PlaceholderのTextFrameにテキストをセット（既存スタイルを保持）"""
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


def set_textbox_text(shape, text):
    """TextBoxの最初のrunのテキストを上書き（スタイル保持）"""
    if shape is None:
        return
    tf = shape.text_frame
    para = tf.paragraphs[0]
    if para.runs:
        para.runs[0].text = text


def fill_implications(slide, items):
    """
    TextBox 29の意味合いセクションを動的に構築する。
    テンプレートの構造:
      para[0]: タイトル「意味合い」(bold, 20pt)
      para[1..N]: Bullet付きImplication (16pt, marL=285750, indent=-285750)

    items: list of str (1〜5個)
    """
    shape = find_shape(slide, SHAPE_IMPLICATIONS)
    if shape is None:
        return

    txBody = shape._element.find(qn("p:txBody"))
    paragraphs = txBody.findall(qn("a:p"))

    if len(paragraphs) < 2:
        print("  WARNING: Implications TextBox has insufficient paragraphs", file=sys.stderr)
        return

    # 2番目の段落（最初のBullet）をテンプレートとして保存
    bullet_template = copy.deepcopy(paragraphs[1])

    # タイトル段落以外をすべて削除
    for p in paragraphs[1:]:
        txBody.remove(p)

    # items数に応じてBullet段落を追加（最大5個）
    for i, text in enumerate(items[:5]):
        new_p = copy.deepcopy(bullet_template)
        runs = new_p.findall(qn("a:r"))
        if runs:
            t_elem = runs[0].find(qn("a:t"))
            if t_elem is not None:
                t_elem.text = text
            for run in runs[1:]:
                t_elem2 = run.find(qn("a:t"))
                if t_elem2 is not None:
                    t_elem2.text = ""
        else:
            r_elem = etree.SubElement(new_p, qn("a:r"))
            rPr = etree.SubElement(r_elem, qn("a:rPr"), attrib={
                "lang": "en-JP", "sz": "1600"
            })
            etree.SubElement(rPr, qn("a:latin"), attrib={"typeface": "+mn-ea"})
            t_elem = etree.SubElement(r_elem, qn("a:t"))
            t_elem.text = text
        txBody.append(new_p)
        print(f"  [Implication {i+1}] {text[:60]}{'...' if len(text) > 60 else ''}")


def rebuild_table(slide, headers, rows, template_prs=None):
    """
    既存テーブルを削除し、コンテンツに応じた行列数でテーブルを再構築する。
    ヘッダー行・データ行のセルスタイルはテンプレートから複製する。
    """
    table_shape = find_shape(slide, SHAPE_TABLE)
    if table_shape is None:
        print("  WARNING: Table shape not found, cannot rebuild", file=sys.stderr)
        return

    old_table = table_shape.table

    # セルスタイル（tcPr）をコピー
    header_tcPr = copy.deepcopy(old_table.cell(0, 0)._tc.find(qn("a:tcPr")))
    data_tcPr   = copy.deepcopy(old_table.cell(1, 0)._tc.find(qn("a:tcPr")))

    # ── ヘッダー行の罫線を白に変更（四辺すべて） ──
    if header_tcPr is not None:
        for ln_tag in ["a:lnL", "a:lnR", "a:lnT", "a:lnB"]:
            ln = header_tcPr.find(qn(ln_tag))
            if ln is not None:
                # 既存罫線の色を白に変更
                sf = ln.find(qn("a:solidFill"))
                if sf is not None:
                    for child in list(sf):
                        sf.remove(child)
                    etree.SubElement(sf, qn("a:srgbClr"), attrib={"val": "FFFFFF"})
            else:
                # 罫線が未定義の場合、白の罫線を新規追加
                ln = etree.SubElement(header_tcPr, qn(ln_tag), attrib={
                    "w": "9525", "cap": "flat", "cmpd": "sng", "algn": "ctr"
                })
                sf = etree.SubElement(ln, qn("a:solidFill"))
                etree.SubElement(sf, qn("a:srgbClr"), attrib={"val": "FFFFFF"})
                etree.SubElement(ln, qn("a:prstDash"), attrib={"val": "solid"})
                etree.SubElement(ln, qn("a:round"))

    # run properties（フォントサイズ等）をコピー
    header_rPr = None
    data_rPr   = None
    for para in old_table.cell(0, 0).text_frame.paragraphs:
        for run in para.runs:
            header_rPr = copy.deepcopy(run._r.find(qn("a:rPr")))
            break
        break
    for para in old_table.cell(1, 0).text_frame.paragraphs:
        for run in para.runs:
            data_rPr = copy.deepcopy(run._r.find(qn("a:rPr")))
            break
        break

    # ── ヘッダー行のフォント色を白に変更 ──
    if header_rPr is not None:
        # 既存のsolidFillがあれば削除
        old_fill = header_rPr.find(qn("a:solidFill"))
        if old_fill is not None:
            header_rPr.remove(old_fill)
        # 白色のsolidFillを追加
        sf = etree.SubElement(header_rPr, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), attrib={"val": "FFFFFF"})

    # 位置・サイズ
    tbl_left  = table_shape.left
    tbl_top   = table_shape.top
    tbl_width = table_shape.width
    tbl_height = table_shape.height

    # 既存テーブルを削除
    sp_tree = slide.shapes._spTree
    sp_tree.remove(table_shape._element)

    # 新テーブルを作成
    n_cols = len(headers)
    n_rows = len(rows) + 1  # ヘッダー行 + データ行

    new_shape = slide.shapes.add_table(
        n_rows, n_cols, tbl_left, tbl_top, tbl_width, tbl_height
    )
    new_table = new_shape.table

    # 列幅を均等配分
    col_width = tbl_width // n_cols
    for col in new_table.columns:
        col.width = col_width

    # tblPrを設定
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    tbl_elem = new_shape._element.find('.//a:tbl', ns)
    old_tblPr = tbl_elem.find('a:tblPr', ns)
    if old_tblPr is not None:
        tbl_elem.remove(old_tblPr)
    tblPr = etree.SubElement(tbl_elem, qn('a:tblPr'), attrib={
        'firstRow': '1', 'bandRow': '1'
    })
    tbl_elem.insert(0, tblPr)

    def apply_cell(cell, text, tcPr_tmpl, rPr_tmpl):
        """セルにテキストとスタイルを適用"""
        tc = cell._tc
        txBody = tc.find(qn("a:txBody"))
        if txBody is None:
            txBody = etree.SubElement(tc, qn("a:txBody"))
            etree.SubElement(txBody, qn("a:bodyPr"))
            etree.SubElement(txBody, qn("a:lstStyle"))

        for p in txBody.findall(qn("a:p")):
            txBody.remove(p)

        p_elem = etree.SubElement(txBody, qn("a:p"))
        etree.SubElement(p_elem, qn("a:pPr"))
        r_elem = etree.SubElement(p_elem, qn("a:r"))
        if rPr_tmpl is not None:
            r_elem.append(copy.deepcopy(rPr_tmpl))
        else:
            etree.SubElement(r_elem, qn("a:rPr"), attrib={
                "lang": "en-GB", "sz": "1200"
            })
        t_elem = etree.SubElement(r_elem, qn("a:t"))
        t_elem.text = str(text)

        old_tc = tc.find(qn("a:tcPr"))
        if old_tc is not None:
            tc.remove(old_tc)
        if tcPr_tmpl is not None:
            tc.append(copy.deepcopy(tcPr_tmpl))

    # ヘッダー行
    for c_idx, h in enumerate(headers):
        apply_cell(new_table.cell(0, c_idx), h, header_tcPr, header_rPr)
    print(f"  [Table Header] {' | '.join(headers)}")

    # データ行
    for r_idx, row_data in enumerate(rows):
        for c_idx in range(n_cols):
            val = row_data[c_idx] if c_idx < len(row_data) else ""
            apply_cell(new_table.cell(r_idx + 1, c_idx), val, data_tcPr, data_rPr)
        print(f"  [Table Row {r_idx+1}] {' | '.join(str(v) for v in row_data[:n_cols])}")

    print(f"  [Table] {n_rows} rows x {n_cols} cols")


def main():
    parser = argparse.ArgumentParser(description="テーブルチャートデータをPPTXに流し込む")
    parser.add_argument("--data",     required=True, help="table_chart_data.json のパス")
    parser.add_argument("--template", required=True, help="table-chart-template.pptx のパス")
    parser.add_argument("--output",   required=True, help="出力PPTXのパス")
    add_brand_arg(parser)  # passive: accepted but ignored until brand migration
    args = parser.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    # ISSUE-012 (2026-05-06): スキーマ齟齬の silent fail 防止
    validate_fill_input(
        data,
        required_top=["main_message", "table"],
        allowed_top=[
            "main_message", "chart_title",
            "table", "table_label", "implications",
        ],
        skill_name="table-chart-pptx",
    )

    prs = Presentation(args.template)
    slide = prs.slides[0]

    # ── Main Message ──
    main_msg = data.get("main_message", "").strip()
    if main_msg:
        shape = find_shape(slide, SHAPE_MAIN_MESSAGE)
        set_placeholder_text(shape, main_msg)
        print(f"  [Main Message] {main_msg[:60]}{'...' if len(main_msg) > 60 else ''}")
    else:
        print("  WARNING: main_message is empty", file=sys.stderr)

    # ── Chart Title ──
    chart_title = data.get("chart_title", "").strip()
    if chart_title:
        shape = find_shape(slide, SHAPE_CHART_TITLE)
        set_placeholder_text(shape, chart_title)
        print(f"  [Chart Title]  {chart_title}")
    else:
        print("  WARNING: chart_title is empty", file=sys.stderr)

    # ── Table Label ──
    table_label = data.get("table_label", "Table").strip()
    shape = find_shape(slide, SHAPE_TABLE_LABEL)
    set_textbox_text(shape, table_label)
    print(f"  [Table Label]  {table_label}")

    # ── Table (動的行列) ──
    table_data = data.get("table", {})
    table_headers = table_data.get("headers", [])
    table_rows    = table_data.get("rows", [])
    if table_headers:
        rebuild_table(slide, table_headers, table_rows)
    else:
        print("  WARNING: table.headers is empty", file=sys.stderr)

    # ── Implications (1〜5個) ──
    implications = data.get("implications", [])
    if implications:
        fill_implications(slide, implications)
    else:
        print("  WARNING: implications is empty", file=sys.stderr)

    prs.save(args.output)

    _finalize_pptx(args.output)

    print(f"\n  Saved: {args.output}")


if __name__ == "__main__":
    main()
