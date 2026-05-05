"""
fill_customer_sales_detail.py — 主要販売先詳細テーブルをHTMLで描画→スクリーンショット→PPTXに挿入

テンプレート構造（customer-sales-detail-template.pptx）:
  - Title 1           (PLACEHOLDER): メインメッセージ（上段、太字）
  - Text Placeholder 2 (PLACEHOLDER): チャートタイトル（下段）
  - Content Area       (AUTO_SHAPE):  HTML screenshot挿入先
  - Source             (TEXT_BOX):    出典（左下）

Usage:
  python fill_customer_sales_detail.py \
    --data /home/claude/customer_sales_detail_data.json \
    --template <SKILL_DIR>/assets/customer-sales-detail-template.pptx \
    --output /mnt/user-data/outputs/CustomerSalesDetail_output.pptx
"""

import argparse

# brand_resolver bootstrap (passive --brand acceptance until brand-aware migration)
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(SKILL_DIR, "..", "_common", "lib"))
from brand_resolver import add_brand_arg  # noqa: E402
import asyncio
import json
import os
import sys
import tempfile
from html import escape as _esc

from pptx import Presentation
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



# ── Shape名マッピング ──
SHAPE_MAIN_MESSAGE = "Title 1"
SHAPE_CHART_TITLE = "Text Placeholder 2"
SHAPE_CONTENT_AREA = "Content Area"
SHAPE_SOURCE = "Source"

# ── HTMLスクリーンショット設定 ──
# Content Area: 12.52" × 5.40" → 200DPI で 2504 × 1080
VIEWPORT_WIDTH = 2504
VIEWPORT_HEIGHT = 1080
DEVICE_SCALE = 2


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


def fmt_number(val):
    """数値をカンマ区切り文字列にフォーマット。null/NoneはNA"""
    if val is None:
        return "NA"
    if isinstance(val, (int, float)):
        return f"{int(val):,}"
    return str(val)


def fmt_pct(val):
    """パーセントをフォーマット。null/NoneはNA"""
    if val is None:
        return "NA"
    if isinstance(val, (int, float)):
        return f"{val:.1f}%"
    return str(val)


def generate_html(data):
    """販売先詳細データからHTML（10列テーブル）を生成する"""
    customers = data.get("customers", [])
    unit = _esc(data.get("unit", "千円"))

    # ── フォントサイズ換算 ──
    # ビューポート 2504px = スライド 12.52" → 1pt ≈ 2.78 CSS px
    PT = 2.78

    font_header = int(11 * PT)
    font_body = int(10.5 * PT)
    pad_cell_v = int(1.8 * PT)
    pad_cell_h = int(2.5 * PT)

    # ── 列幅比率 ──
    col_widths = [
        ("3%", "center"),    # #
        ("9%", "left"),      # 企業名
        ("7%", "right"),     # 売上高
        ("5%", "right"),     # 割合
        ("22%", "left"),     # 事業内容
        ("12%", "left"),     # 本社所在地
        ("8%", "left"),      # 上場
        ("11%", "right"),    # 直近期売上高
        ("9%", "right"),     # 利益
        ("7%", "right"),     # 利益率
    ]

    headers = [
        "#",
        "企業名",
        f"売上高<br><span style='font-size:{int(font_header*0.8)}px;font-weight:400;'>({unit})</span>",
        "割合",
        "事業内容",
        "本社所在地",
        "上場",
        f"直近期売上高<br><span style='font-size:{int(font_header*0.8)}px;font-weight:400;'>({unit})</span>",
        f"利益<br><span style='font-size:{int(font_header*0.8)}px;font-weight:400;'>({unit})</span>",
        "利益率",
    ]

    # colgroup
    colgroup = "<colgroup>\n"
    for w, _ in col_widths:
        colgroup += f'  <col style="width:{w};">\n'
    colgroup += "</colgroup>"

    # ヘッダー行
    header_cells = ""
    for i, h in enumerate(headers):
        _, align = col_widths[i]
        header_cells += f'''
            <th style="padding:{pad_cell_v}px {pad_cell_h}px;font-size:{font_header}px;
                font-weight:700;color:#333;text-align:{align};
                border-bottom:2px solid #666;white-space:nowrap;">{h}</th>'''

    # データ行
    body_rows = ""
    for i, cust in enumerate(customers):
        name = cust.get("name", "")
        revenue = cust.get("revenue", 0)
        share = cust.get("share", 0)
        rank = cust.get("rank", i + 1)
        business = cust.get("business")
        headquarters = cust.get("headquarters")
        listing = cust.get("listing")
        latest_revenue = cust.get("latest_revenue")
        profit = cust.get("profit")
        profit_margin = cust.get("profit_margin")

        is_other = (name == "その他" or rank == -1)

        # ランク表示
        rank_str = str(rank) if not is_other else ""

        # 売上高・割合フォーマット
        rev_str = f"{int(revenue):,}" if isinstance(revenue, (int, float)) else str(revenue)
        share_str = f"{share:.1f}%" if isinstance(share, (int, float)) else str(share)

        # TSR列（その他は空欄）
        if is_other:
            biz_str = ""
            hq_str = ""
            list_str = ""
            lr_str = ""
            prof_str = ""
            pm_str = ""
        else:
            biz_str = _esc(business) if business else ""
            hq_str = _esc(headquarters) if headquarters else ""
            list_str = _esc(listing) if listing else ""
            lr_str = fmt_number(latest_revenue)
            prof_str = fmt_number(profit)
            pm_str = fmt_pct(profit_margin)

        # 行背景色
        if is_other:
            bg = "#F5F5F5"
        elif i % 2 == 1:
            bg = "#FAFAFA"
        else:
            bg = "#FFFFFF"

        border = "1px solid #E0E0E0"

        # 事業内容は折り返し可
        biz_style = f"white-space:normal;word-break:break-all;"

        cells = [
            (rank_str, "center", ""),
            (_esc(name), "left", "white-space:nowrap;"),
            (rev_str, "right", ""),
            (share_str, "right", ""),
            (biz_str, "left", biz_style),
            (hq_str, "left", ""),
            (list_str, "left", ""),
            (lr_str, "right", ""),
            (prof_str, "right", ""),
            (pm_str, "right", ""),
        ]

        row_cells = ""
        for j, (val, align, extra_style) in enumerate(cells):
            row_cells += f'''
                <td style="padding:{pad_cell_v}px {pad_cell_h}px;font-size:{font_body}px;
                    color:#333;text-align:{align};border-bottom:{border};{extra_style}">{val}</td>'''

        body_rows += f'''
            <tr style="background:{bg};">{row_cells}
            </tr>'''

    body_pad = int(2 * PT)
    html = f'''<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
html, body {{
    width:{VIEWPORT_WIDTH}px;
    height:{VIEWPORT_HEIGHT}px;
    font-family:'Noto Sans CJK JP','Meiryo UI','Hiragino Sans',sans-serif;
    background:#FFFFFF;
    padding:{body_pad}px;
    overflow:hidden;
}}
</style>
</head>
<body>
<table style="width:100%;border-collapse:collapse;
    font-family:'Noto Sans CJK JP','Meiryo UI','Hiragino Sans',sans-serif;
    table-layout:fixed;">
    {colgroup}
    <thead>
        <tr style="background:#F0F0F0;">{header_cells}
        </tr>
    </thead>
    <tbody>{body_rows}
    </tbody>
</table>
</body></html>'''

    return html


async def take_screenshot(html_content, output_path):
    """PlaywrightでHTMLのスクリーンショットを撮る"""
    from playwright.async_api import async_playwright

    with tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode="w", encoding="utf-8") as f:
        f.write(html_content)
        html_path = f.name

    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch()
            page = await browser.new_page(
                viewport={"width": VIEWPORT_WIDTH, "height": VIEWPORT_HEIGHT},
                device_scale_factor=DEVICE_SCALE,
            )
            await page.goto(f"file://{html_path}", wait_until="networkidle")
            await page.wait_for_timeout(500)
            await page.screenshot(path=output_path, full_page=False)
            await browser.close()
    finally:
        os.unlink(html_path)

    print(f"  ✓ Screenshot saved: {output_path}")


def insert_screenshot_into_pptx(template_path, data, screenshot_path, output_path):
    """テンプレートにデータとスクリーンショットを挿入"""
    prs = Presentation(template_path)
    slide = prs.slides[0]

    # Shape一覧を表示（デバッグ用）
    for s in slide.shapes:
        print(f"  Shape: '{s.name}' type={s.shape_type}")

    # 1. メインメッセージ（Title 1 = 上段、太字）
    main_msg = data.get("main_message", "")
    msg_shape = find_shape(slide, SHAPE_MAIN_MESSAGE)
    set_textbox_text(msg_shape, main_msg)

    # 2. チャートタイトル（Text Placeholder 2 = 下段）
    chart_title = data.get("chart_title", "主要販売先からの完成工事売上高と割合")
    title_shape = find_shape(slide, SHAPE_CHART_TITLE)
    set_textbox_text(title_shape, chart_title)

    # 3. 出典
    source_text = data.get("source", "")
    source_shape = find_shape(slide, SHAPE_SOURCE)
    if source_text:
        set_textbox_text(source_shape, f"出典：{source_text}")
    else:
        set_textbox_text(source_shape, "")

    # 4. コンテンツエリアにスクリーンショットを挿入
    content_shape = find_shape(slide, SHAPE_CONTENT_AREA)
    if content_shape and os.path.exists(screenshot_path):
        left = content_shape.left
        top = content_shape.top
        width = content_shape.width
        height = content_shape.height

        slide.shapes.add_picture(
            screenshot_path, left, top, width, height
        )
        print(f"  ✓ Screenshot inserted at ({left},{top}) size=({width},{height})")

    # 保存
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    prs.save(output_path)
    _finalize_pptx(output_path)
    print(f"  ✓ PPTX saved: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="主要販売先詳細テーブルスライド生成")
    parser.add_argument("--data", required=True, help="JSONデータファイルパス")
    parser.add_argument("--template", required=True, help="PPTXテンプレートパス")
    parser.add_argument("--output", required=True, help="出力PPTXファイルパス")
    add_brand_arg(parser)  # passive: accepted but ignored until brand migration
    args = parser.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    print("=== 主要販売先詳細テーブルスライド生成 ===")

    # Step 1: HTML生成
    print("Step 1: HTML生成...")
    html_content = generate_html(data)

    # デバッグ用: HTML保存
    debug_html_path = os.path.join(tempfile.gettempdir(), "customer_sales_detail_debug.html")
    with open(debug_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"  ✓ Debug HTML: {debug_html_path}")

    # Step 2: スクリーンショット
    print("Step 2: スクリーンショット取得...")
    screenshot_path = os.path.join(tempfile.gettempdir(), "customer_sales_detail_screenshot.png")
    asyncio.run(take_screenshot(html_content, screenshot_path))

    # Step 3: PPTX生成
    print("Step 3: PPTX生成...")
    insert_screenshot_into_pptx(args.template, data, screenshot_path, args.output)

    # クリーンアップ
    if os.path.exists(screenshot_path):
        os.unlink(screenshot_path)

    print("=== 完了 ===")


if __name__ == "__main__":
    main()
