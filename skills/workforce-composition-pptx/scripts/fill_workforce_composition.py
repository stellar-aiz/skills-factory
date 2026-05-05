"""
fill_workforce_composition.py — 人員構成スライドをPPTXネイティブオブジェクトで生成

テンプレート: company-history-template.pptx をベースに、
  - 既存テーブルを削除
  - 左側: 在籍人員数の推移（ネイティブ棒グラフ: 3系列）
  - 右側: 部署別人員構成テーブル（ネイティブテーブル＋合計行）
を配置する。

Usage:
  python fill_workforce_composition.py \
    --data /home/claude/workforce_composition_data.json \
    --template <path>/workforce-composition-template.pptx \
    --output /mnt/user-data/outputs/WorkforceComposition_output.pptx
"""

import argparse
import json
import math
import os
import sys

# brand_resolver bootstrap (passive --brand acceptance until brand-aware migration)
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(SKILL_DIR, "..", "_common", "lib"))
from brand_resolver import add_brand_arg  # noqa: E402

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE
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



# ── Layout Constants ──
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.50)

# Template shape names
SHAPE_MAIN_MESSAGE = "Title 1"
SHAPE_CHART_TITLE = "Text Placeholder 2"

# Common Y
PANEL_Y = Inches(1.50)

# Left panel: Headcount Trend Chart
LEFT_X = Inches(0.41)
LEFT_W = Inches(5.80)

# Right panel: Department Table
RIGHT_X = Inches(6.50)
RIGHT_W = Inches(6.40)

# Style constants
COLOR_TEXT = RGBColor(0x33, 0x33, 0x33)
COLOR_BAR_TOTAL = "ED7D31"       # Orange for total headcount
COLOR_BAR_HIRES = "4472C4"       # Blue for new hires
COLOR_BAR_DEPARTURES = "A5A5A5"  # Gray for departures
COLOR_HEADER_BG = "F0F0F0"      # Table header background
COLOR_TOTAL_BG = "E8E8E8"       # Table total row background
COLOR_EVEN_ROW = "FAFAFA"       # Table even row background

FONT_NAME_JP = "Meiryo UI"
FONT_SIZE_SECTION = Pt(14)
FONT_SIZE_TABLE_HEADER = Pt(12)
FONT_SIZE_TABLE_BODY = Pt(12)


def find_shape(slide, name):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    print(f"  ⚠ WARNING: Shape '{name}' not found", file=sys.stderr)
    return None


def set_textbox_text(shape, text):
    """TextBoxのテキストを上書き（既存スタイルを保持）"""
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


def remove_shape(slide, name):
    """名前でShapeを削除"""
    shape = find_shape(slide, name)
    if shape is not None:
        sp_tree = slide.shapes._spTree
        sp_tree.remove(shape._element)
        print(f"  ✓ Shape '{name}' removed")


def add_section_title(slide, text, left, top, width):
    """セクションタイトル（下線付き）を追加"""
    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.30))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = FONT_SIZE_SECTION
    run.font.bold = True
    run.font.color.rgb = COLOR_TEXT
    run.font.name = FONT_NAME_JP

    # 下線
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top + Inches(0.30), width, Inches(0.02)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_TEXT
    line.line.fill.background()
    return txBox


def add_unit_label(slide, text, left, top, width):
    """単位表記（左寄せ）"""
    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.22))
    tf = txBox.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.size = Pt(11)
    run.font.color.rgb = COLOR_TEXT
    run.font.name = FONT_NAME_JP


def add_source_label(slide, text):
    """スライド左下に出典テキストを追加"""
    src_left = Inches(0.41)
    src_top = Inches(7.05)
    src_w = Inches(8.00)
    src_h = Inches(0.30)

    txBox = slide.shapes.add_textbox(src_left, src_top, src_w, src_h)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    run.font.name = FONT_NAME_JP
    print(f"  ✓ 出典: {text}")


# ── Chart: Headcount Trend ──

def build_headcount_chart(slide, trend_data, left, top, width, height):
    """在籍人員数の推移チャート（ネイティブ棒グラフ3系列）を作成"""
    from pptx.chart.data import CategoryChartData

    periods = trend_data["periods"]
    labels = trend_data.get("series_labels", {})
    total_label = labels.get("total", "総従業員数")
    hires_label = labels.get("new_hires", "入社人数")
    depart_label = labels.get("departures", "退職人数")

    chart_data = CategoryChartData()
    chart_data.categories = [p["label"] for p in periods]
    chart_data.add_series(total_label, [p["total"] for p in periods])
    chart_data.add_series(hires_label, [p["new_hires"] for p in periods])
    chart_data.add_series(depart_label, [p["departures"] for p in periods])

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left, top, width, height,
        chart_data
    )
    chart = chart_frame.chart

    # === Chart styling ===
    plotArea = chart._chartSpace.chart.plotArea
    barChart = plotArea.findall(qn('c:barChart'))[0]

    # Set gap width (wider gap between groups, narrower within)
    gapWidth = barChart.find(qn('c:gapWidth'))
    if gapWidth is None:
        gapWidth = etree.SubElement(barChart, qn('c:gapWidth'))
    gapWidth.set('val', '80')

    # Series overlap
    overlap = barChart.find(qn('c:overlap'))
    if overlap is None:
        overlap = etree.SubElement(barChart, qn('c:overlap'))
    overlap.set('val', '-20')

    # Style each series
    colors = [COLOR_BAR_TOTAL, COLOR_BAR_HIRES, COLOR_BAR_DEPARTURES]
    sers = barChart.findall(qn('c:ser'))

    for i, ser in enumerate(sers):
        # Set color
        spPr = ser.find(qn('c:spPr'))
        if spPr is None:
            spPr = etree.SubElement(ser, qn('c:spPr'))
        sf = spPr.find(qn('a:solidFill'))
        if sf is None:
            sf = etree.SubElement(spPr, qn('a:solidFill'))
        for child in list(sf):
            sf.remove(child)
        etree.SubElement(sf, qn('a:srgbClr'), attrib={'val': colors[i]})

        # Add data labels
        _add_data_labels(ser, num_format='#,##0;-#,##0;""')

    # Remove gridlines
    for vax in plotArea.findall(qn('c:valAx')):
        for mg in vax.findall(qn('c:majorGridlines')):
            vax.remove(mg)
        for mg in vax.findall(qn('c:minorGridlines')):
            vax.remove(mg)
    for cax in plotArea.findall(qn('c:catAx')):
        for mg in cax.findall(qn('c:majorGridlines')):
            cax.remove(mg)

    # Hide value axis
    for vax in plotArea.findall(qn('c:valAx')):
        del_elem = vax.find(qn('c:delete'))
        if del_elem is None:
            del_elem = etree.SubElement(vax, qn('c:delete'))
        del_elem.set('val', '1')

    # Style category axis font
    for cax in plotArea.findall(qn('c:catAx')):
        delete_elem = cax.find(qn('c:delete'))
        if delete_elem is not None and delete_elem.get('val') == '1':
            continue
        txPr = cax.find(qn('c:txPr'))
        if txPr is None:
            txPr = etree.SubElement(cax, qn('c:txPr'))
        bodyPr = txPr.find(qn('a:bodyPr'))
        if bodyPr is None:
            bodyPr = etree.SubElement(txPr, qn('a:bodyPr'))
            txPr.insert(0, bodyPr)
        if txPr.find(qn('a:lstStyle')) is None:
            etree.SubElement(txPr, qn('a:lstStyle'))
        p = txPr.find(qn('a:p'))
        if p is None:
            p = etree.SubElement(txPr, qn('a:p'))
        pPr = p.find(qn('a:pPr'))
        if pPr is None:
            pPr = etree.SubElement(p, qn('a:pPr'))
        defRPr = pPr.find(qn('a:defRPr'))
        if defRPr is None:
            defRPr = etree.SubElement(pPr, qn('a:defRPr'), attrib={'sz': '1100'})
        else:
            defRPr.set('sz', '1100')
        if defRPr.find(qn('a:latin')) is None:
            etree.SubElement(defRPr, qn('a:latin'), attrib={'typeface': FONT_NAME_JP})
        if defRPr.find(qn('a:ea')) is None:
            etree.SubElement(defRPr, qn('a:ea'), attrib={'typeface': FONT_NAME_JP})

    # Chart title off, legend handled by custom legend
    chart.has_title = False
    chart.has_legend = False

    # Plot area no background
    plotArea_spPr = plotArea.find(qn('c:spPr'))
    if plotArea_spPr is None:
        plotArea_spPr = etree.SubElement(plotArea, qn('c:spPr'))
    noFill = plotArea_spPr.find(qn('a:noFill'))
    if noFill is None:
        etree.SubElement(plotArea_spPr, qn('a:noFill'))

    print(f"  ✓ 人員推移チャート生成: {len(periods)}期分")
    return chart_frame


def _add_data_labels(ser_xml, num_format='#,##0'):
    """シリーズにデータラベルを追加"""
    dLbls = ser_xml.find(qn('c:dLbls'))
    if dLbls is None:
        dLbls = etree.SubElement(ser_xml, qn('c:dLbls'))

    for child in list(dLbls):
        dLbls.remove(child)

    etree.SubElement(dLbls, qn('c:numFmt'), attrib={
        'formatCode': num_format, 'sourceLinked': '0'
    })
    etree.SubElement(dLbls, qn('c:showLegendKey'), attrib={'val': '0'})
    etree.SubElement(dLbls, qn('c:showVal'), attrib={'val': '1'})
    etree.SubElement(dLbls, qn('c:showCatName'), attrib={'val': '0'})
    etree.SubElement(dLbls, qn('c:showSerName'), attrib={'val': '0'})
    etree.SubElement(dLbls, qn('c:showPercent'), attrib={'val': '0'})
    etree.SubElement(dLbls, qn('c:showBubbleSize'), attrib={'val': '0'})
    etree.SubElement(dLbls, qn('c:dLblPos'), attrib={'val': 'outEnd'})

    # Font
    txPr = etree.SubElement(dLbls, qn('c:txPr'))
    etree.SubElement(txPr, qn('a:bodyPr'))
    etree.SubElement(txPr, qn('a:lstStyle'))
    p = etree.SubElement(txPr, qn('a:p'))
    pPr = etree.SubElement(p, qn('a:pPr'))
    defRPr = etree.SubElement(pPr, qn('a:defRPr'), attrib={'sz': '1100', 'b': '1'})
    etree.SubElement(defRPr, qn('a:latin'), attrib={'typeface': FONT_NAME_JP})
    etree.SubElement(defRPr, qn('a:ea'), attrib={'typeface': FONT_NAME_JP})
    sf = etree.SubElement(defRPr, qn('a:solidFill'))
    etree.SubElement(sf, qn('a:srgbClr'), attrib={'val': '333333'})


def add_custom_legend_chart(slide, trend_data, left, top, width):
    """カスタム凡例（■総従業員数  ■入社人数  ■退職人数）"""
    labels_cfg = trend_data.get("series_labels", {})
    items = [
        (labels_cfg.get("total", "総従業員数"), COLOR_BAR_TOTAL),
        (labels_cfg.get("new_hires", "入社人数"), COLOR_BAR_HIRES),
        (labels_cfg.get("departures", "退職人数"), COLOR_BAR_DEPARTURES),
    ]

    legend_h = Inches(0.22)
    sq_size = Inches(0.14)
    sq_y = top + (legend_h - sq_size) // 2
    cursor_x = left

    for label_text, color_hex in items:
        # Color square
        marker = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            int(cursor_x), int(sq_y), sq_size, sq_size
        )
        marker.fill.solid()
        marker.fill.fore_color.rgb = RGBColor.from_string(color_hex)
        marker.line.fill.background()

        # Label text
        text_x = cursor_x + sq_size + Inches(0.04)
        text_w = Inches(1.10)
        txBox = slide.shapes.add_textbox(int(text_x), top, text_w, legend_h)
        tf = txBox.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = label_text
        run.font.size = Pt(10)
        run.font.color.rgb = COLOR_TEXT
        run.font.name = FONT_NAME_JP

        cursor_x = text_x + text_w + Inches(0.08)

    print(f"  ✓ チャート凡例: {', '.join([i[0] for i in items])}")


# ── Table: Department Composition ──

def _style_table_cell(cell, text, bold=False, font_size=None, align='l',
                      bg_color=None, font_color='333333'):
    """テーブルセルにスタイルを適用"""
    if font_size is None:
        font_size = FONT_SIZE_TABLE_BODY

    tc = cell._tc
    txBody = tc.find(qn("a:txBody"))
    if txBody is None:
        txBody = etree.SubElement(tc, qn("a:txBody"))

    # bodyPr
    old_bodyPr = txBody.find(qn("a:bodyPr"))
    if old_bodyPr is not None:
        txBody.remove(old_bodyPr)
    bodyPr = etree.SubElement(txBody, qn("a:bodyPr"), attrib={
        "wrap": "square",
        "lIns": "36576", "rIns": "36576",
        "tIns": "18288", "bIns": "18288",
        "anchor": "ctr",
    })
    txBody.insert(0, bodyPr)

    # lstStyle
    if txBody.find(qn("a:lstStyle")) is None:
        lstStyle = etree.SubElement(txBody, qn("a:lstStyle"))
        txBody.insert(1, lstStyle)

    # Clear existing paragraphs
    for p in txBody.findall(qn("a:p")):
        txBody.remove(p)

    # New paragraph
    p_elem = etree.SubElement(txBody, qn("a:p"))
    pPr = etree.SubElement(p_elem, qn("a:pPr"))

    align_map = {'l': 'l', 'c': 'ctr', 'r': 'r'}
    pPr.set("algn", align_map.get(align, 'l'))

    # Line spacing
    lnSpc = etree.SubElement(pPr, qn("a:lnSpc"))
    etree.SubElement(lnSpc, qn("a:spcPct"), attrib={"val": "100000"})
    spcBef = etree.SubElement(pPr, qn("a:spcBef"))
    etree.SubElement(spcBef, qn("a:spcPts"), attrib={"val": "0"})
    spcAft = etree.SubElement(pPr, qn("a:spcAft"))
    etree.SubElement(spcAft, qn("a:spcPts"), attrib={"val": "0"})

    # Run
    r_elem = etree.SubElement(p_elem, qn("a:r"))
    rPr = etree.SubElement(r_elem, qn("a:rPr"), attrib={
        "lang": "ja-JP",
        "sz": str(int(font_size.pt * 100)),
        "b": "1" if bold else "0",
    })
    solidFill = etree.SubElement(rPr, qn("a:solidFill"))
    etree.SubElement(solidFill, qn("a:srgbClr"), attrib={"val": font_color})
    etree.SubElement(rPr, qn("a:latin"), attrib={"typeface": FONT_NAME_JP})
    etree.SubElement(rPr, qn("a:ea"), attrib={"typeface": FONT_NAME_JP})

    t_elem = etree.SubElement(r_elem, qn("a:t"))
    t_elem.text = str(text) if text is not None else "-"

    # Cell properties (tcPr)
    old_tcPr = tc.find(qn("a:tcPr"))
    if old_tcPr is not None:
        tc.remove(old_tcPr)
    tcPr = etree.SubElement(tc, qn("a:tcPr"), attrib={
        "marL": "36576", "marR": "36576",
        "marT": "18288", "marB": "18288",
        "anchor": "ctr",
    })

    # Borders (thin gray)
    for border_name in ["a:lnL", "a:lnR", "a:lnT", "a:lnB"]:
        ln = etree.SubElement(tcPr, qn(border_name), attrib={"w": "6350", "cmpd": "sng"})
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), attrib={"val": "CCCCCC"})

    # Background color
    if bg_color:
        fill_elem = etree.SubElement(tcPr, qn("a:solidFill"))
        etree.SubElement(fill_elem, qn("a:srgbClr"), attrib={"val": bg_color})


def build_department_table(slide, dept_data, left, top, width, max_height):
    """部署別人員構成テーブルを構築（ネイティブPPTXテーブル）"""
    columns = dept_data.get("columns", [
        "部署名", "人数", "平均年齢", "平均勤続年数", "管理職数", "有資格者数"
    ])
    departments = dept_data.get("departments", [])
    show_total = dept_data.get("show_total", True)

    n_cols = len(columns)
    n_data_rows = len(departments)
    n_rows = 1 + n_data_rows + (1 if show_total else 0)  # header + data + total

    # Dynamic row height based on available space
    row_h = min(Inches(0.35), int(max_height / n_rows))
    table_h = row_h * n_rows

    # Column width distribution
    # First column (dept name) wider, rest equal
    col0_w = int(width * 0.28)
    remaining_w = width - col0_w
    other_col_w = int(remaining_w / (n_cols - 1))
    col_widths = [col0_w] + [other_col_w] * (n_cols - 1)

    shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, table_h)
    table = shape.table

    # Disable banding
    tbl_elem = shape._element.find('.//' + qn('a:tbl'))
    old_tblPr = tbl_elem.find(qn('a:tblPr'))
    if old_tblPr is not None:
        tbl_elem.remove(old_tblPr)
    tblPr = etree.SubElement(tbl_elem, qn('a:tblPr'), attrib={
        'firstRow': '0', 'bandRow': '0'
    })
    tbl_elem.insert(0, tblPr)

    # Set column widths
    for i, w in enumerate(col_widths):
        table.columns[i].width = w

    # Set row heights
    for tr in tbl_elem.findall(qn('a:tr')):
        tr.set('h', str(row_h))

    # Header row
    for c_idx, col_name in enumerate(columns):
        align = 'c' if c_idx > 0 else 'l'
        _style_table_cell(
            table.cell(0, c_idx), col_name,
            bold=True, font_size=FONT_SIZE_TABLE_HEADER,
            align=align, bg_color=COLOR_HEADER_BG
        )

    # Column key mapping
    key_map = {
        "部署名": "name",
        "人数": "headcount",
        "平均年齢": "avg_age",
        "平均勤続年数": "avg_tenure",
        "管理職数": "managers",
        "有資格者数": "certified",
    }

    # Data rows
    for r_idx, dept in enumerate(departments):
        row_num = r_idx + 1
        bg = COLOR_EVEN_ROW if r_idx % 2 == 1 else None

        for c_idx, col_name in enumerate(columns):
            key = key_map.get(col_name, col_name)
            value = dept.get(key, "-")

            # Format values
            if value is None:
                display_val = "-"
            elif isinstance(value, float):
                display_val = f"{value:.1f}"
            elif isinstance(value, int):
                display_val = str(value)
            else:
                display_val = str(value)

            align = 'c' if c_idx > 0 else 'l'
            _style_table_cell(
                table.cell(row_num, c_idx), display_val,
                bold=False, font_size=FONT_SIZE_TABLE_BODY,
                align=align, bg_color=bg
            )

    # Total row
    if show_total and departments:
        total_row = n_rows - 1
        total_headcount = sum(d.get("headcount", 0) or 0 for d in departments)

        # Weighted average for avg_age and avg_tenure
        total_for_age = sum(
            (d.get("headcount", 0) or 0) * (d.get("avg_age", 0) or 0)
            for d in departments if d.get("avg_age") is not None
        )
        count_for_age = sum(
            d.get("headcount", 0) or 0
            for d in departments if d.get("avg_age") is not None
        )
        avg_age_total = total_for_age / count_for_age if count_for_age > 0 else None

        total_for_tenure = sum(
            (d.get("headcount", 0) or 0) * (d.get("avg_tenure", 0) or 0)
            for d in departments if d.get("avg_tenure") is not None
        )
        count_for_tenure = sum(
            d.get("headcount", 0) or 0
            for d in departments if d.get("avg_tenure") is not None
        )
        avg_tenure_total = total_for_tenure / count_for_tenure if count_for_tenure > 0 else None

        total_managers = sum(d.get("managers", 0) or 0 for d in departments)
        total_certified = sum(d.get("certified", 0) or 0 for d in departments)

        total_values = {
            "name": "合計",
            "headcount": total_headcount,
            "avg_age": avg_age_total,
            "avg_tenure": avg_tenure_total,
            "managers": total_managers,
            "certified": total_certified,
        }

        for c_idx, col_name in enumerate(columns):
            key = key_map.get(col_name, col_name)
            value = total_values.get(key, "-")

            if value is None:
                display_val = "-"
            elif isinstance(value, float):
                display_val = f"{value:.1f}"
            elif isinstance(value, int):
                display_val = str(value)
            else:
                display_val = str(value)

            align = 'c' if c_idx > 0 else 'l'
            _style_table_cell(
                table.cell(total_row, c_idx), display_val,
                bold=True, font_size=FONT_SIZE_TABLE_BODY,
                align=align, bg_color=COLOR_TOTAL_BG
            )

    print(f"  ✓ 部署テーブル: {n_data_rows}部署 + {'合計行' if show_total else 'なし'}")
    return shape


# ── Main ──

def main():
    parser = argparse.ArgumentParser(description="人員構成 PowerPoint ジェネレーター")
    parser.add_argument("--data", required=True)
    parser.add_argument("--template", required=True)
    parser.add_argument("--output", required=True)
    add_brand_arg(parser)  # passive: accepted but ignored until brand migration
    args = parser.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    prs = Presentation(args.template)
    slide = prs.slides[0]

    # 1. Main Message
    main_message = data.get("main_message", "")
    set_textbox_text(find_shape(slide, SHAPE_MAIN_MESSAGE), main_message)
    print(f"  ✓ Main Message: {main_message}")

    # 2. Chart Title
    chart_title = data.get("chart_title", "人員構成")
    set_textbox_text(find_shape(slide, SHAPE_CHART_TITLE), chart_title)
    print(f"  ✓ Chart Title: {chart_title}")

    # 3. Remove existing table
    remove_shape(slide, "Table 1")

    # 4. Left panel: Headcount Trend
    trend = data.get("headcount_trend", {})
    trend_title = trend.get("title", "在籍人員数の推移")
    unit_text = f"（単位：{trend.get('unit', '人')}）"

    # Section title
    add_section_title(slide, trend_title, LEFT_X, PANEL_Y, LEFT_W)

    # Unit label + custom legend
    legend_y = PANEL_Y + Inches(0.35)
    add_unit_label(slide, unit_text, LEFT_X, legend_y, Inches(1.50))
    add_custom_legend_chart(slide, trend, LEFT_X, legend_y, LEFT_W)

    # Chart
    chart_top = PANEL_Y + Inches(0.60)
    chart_h = Inches(5.00)
    build_headcount_chart(slide, trend, LEFT_X, chart_top, LEFT_W, chart_h)

    # 5. Right panel: Department Table
    dept = data.get("department_table", {})
    dept_title = dept.get("title", "人員構成")

    # Section title
    add_section_title(slide, dept_title, RIGHT_X, PANEL_Y, RIGHT_W)

    # Table
    table_top = PANEL_Y + Inches(0.40)
    max_table_h = Inches(5.20)
    build_department_table(slide, dept, RIGHT_X, table_top, RIGHT_W, max_table_h)

    # 6. Source
    source = data.get("source", "")
    if source:
        add_source_label(slide, source)

    # 7. Save
    os.makedirs(os.path.dirname(args.output), exist_ok=True)
    prs.save(args.output)
    _finalize_pptx(args.output)
    print(f"\n  ✅ 出力完了: {args.output}")


if __name__ == "__main__":
    main()
