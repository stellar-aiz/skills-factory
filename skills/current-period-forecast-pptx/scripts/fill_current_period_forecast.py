"""
fill_current_period_forecast.py — 当期着地見込みテーブルをHTMLで描画→スクリーンショット→PPTXに挿入するスクリプト

BDD（Business Due Diligence）向け。
マネジメント計画（Base）と弊社計画（Downside / Upside）の財務見込みを比較するスライドを生成する。

テンプレート構造（forecast-template.pptx = company-overview-template.pptxベース）:
  - Title 1              (PLACEHOLDER): Main Message
  - Text Placeholder 2   (PLACEHOLDER): Chart Title
  - Content Area         (AUTO_SHAPE):  HTML screenshot挿入先
  - Source               (TEXT_BOX):    出典

使い方:
  python fill_current_period_forecast.py \
    --data /home/claude/forecast_data.json \
    --template <SKILL_DIR>/assets/forecast-template.pptx \
    --output /mnt/user-data/outputs/CurrentPeriodForecast_output.pptx
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
from pptx.util import Emu, Pt
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
SHAPE_CONTENT_AREA = "Content Area"
SHAPE_SOURCE       = "Source"
# ────────────────────────────────────────────────────────────

# ── HTMLスクリーンショットの設定 ────────────────────────────
VIEWPORT_WIDTH  = 2504
VIEWPORT_HEIGHT = 1080
DEVICE_SCALE    = 2
# ────────────────────────────────────────────────────────────

PT = 2.78  # 1pt あたりの CSS px


def find_shape(slide, name):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    print(f"  WARNING: Shape '{name}' not found", file=sys.stderr)
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


def generate_html(data):
    """当期着地見込みデータからHTMLテーブルを生成する"""
    period_label = _esc(data.get("period_label", "進行期 見込み"))
    unit_label = _esc(data.get("unit_label", "単位：百万円/(%)"))
    mgmt_label = _esc(data.get("management_plan_label", "マネジメント計画\n(Base)"))
    our_plan_label = _esc(data.get("our_plan_label", "弊社計画"))
    downside_label = _esc(data.get("downside_label", "Downside case"))
    upside_label = _esc(data.get("upside_label", "Upside case"))
    item_header = _esc(data.get("item_header", "項目（調整後）"))

    rows = data.get("rows", [])
    n_rows = len(rows)

    # フォントサイズ（行数に応じて調整）
    if n_rows <= 5:
        fs_header = int(11 * PT)
        fs_value = int(12 * PT)
        fs_pct = int(8 * PT)
        fs_assumption = int(9 * PT)
        row_pad = int(3 * PT)
    elif n_rows <= 8:
        fs_header = int(10 * PT)
        fs_value = int(11 * PT)
        fs_pct = int(7.5 * PT)
        fs_assumption = int(8 * PT)
        row_pad = int(2 * PT)
    else:
        fs_header = int(9 * PT)
        fs_value = int(9 * PT)
        fs_pct = int(6.5 * PT)
        fs_assumption = int(7 * PT)
        row_pad = int(1.5 * PT)

    # 色定義
    HEADER_BG = "#D9D9D9"       # ヘッダー背景（グレー）
    HEADER_DARK_BG = "#BFBFBF"  # ヘッダー濃いグレー
    ROW_EVEN_BG = "#FFFFFF"
    ROW_ODD_BG = "#F5F5F0"
    BORDER_COLOR = "#999999"
    TEXT_COLOR = "#1A1A2E"
    ACCENT_COLOR = "#1A1A2E"

    # mgmt_label の改行対応
    mgmt_label_html = mgmt_label.replace("\n", "<br>")

    # テーブルヘッダーHTML
    header_html = f"""
    <thead>
        <tr>
            <td rowspan="2" class="cell header-cell item-col"
                style="background:{HEADER_DARK_BG};font-weight:700;font-size:{fs_header}px;
                       text-align:center;vertical-align:middle;">
                <div style="font-size:{int(7*PT)}px;color:#555;text-align:left;margin-bottom:{int(1*PT)}px;">
                    {unit_label}
                </div>
                {item_header}
            </td>
            <td rowspan="2" class="cell header-cell mgmt-col"
                style="background:{HEADER_BG};font-weight:700;font-size:{fs_header}px;
                       text-align:center;vertical-align:middle;line-height:1.4;">
                {mgmt_label_html}
            </td>
            <td colspan="4" class="cell header-cell"
                style="background:{HEADER_BG};font-weight:700;font-size:{fs_header}px;
                       text-align:center;vertical-align:middle;">
                {our_plan_label}
            </td>
        </tr>
        <tr>
            <td colspan="2" class="cell header-cell"
                style="background:{HEADER_BG};font-weight:700;font-size:{fs_header}px;
                       text-align:center;vertical-align:middle;">
                {downside_label}
            </td>
            <td colspan="2" class="cell header-cell"
                style="background:{HEADER_BG};font-weight:700;font-size:{fs_header}px;
                       text-align:center;vertical-align:middle;">
                {upside_label}
            </td>
        </tr>
    </thead>"""

    # データ行HTML生成
    body_rows = ""
    for i, row in enumerate(rows):
        bg = ROW_ODD_BG if i % 2 == 0 else ROW_EVEN_BG
        item_name = _esc(row.get("item", ""))
        
        # マネジメント計画
        mgmt_val = _esc(str(row.get("management_value", "")))
        mgmt_pct = _esc(str(row.get("management_pct", "")))
        
        # Downside
        ds_val = _esc(str(row.get("downside_value", "")))
        ds_pct = _esc(str(row.get("downside_pct", "")))
        ds_assumption = _esc(str(row.get("downside_assumption", "")))
        
        # Upside
        us_val = _esc(str(row.get("upside_value", "")))
        us_pct = _esc(str(row.get("upside_pct", "")))
        us_assumption = _esc(str(row.get("upside_assumption", "")))

        # 値＋パーセンテージのHTML
        def val_cell(val, pct):
            pct_html = f'<div style="font-size:{fs_pct}px;color:#666;">({pct})</div>' if pct else ""
            return f'<div style="font-size:{fs_value}px;font-weight:700;text-align:center;">{val}</div>{pct_html}'

        # 前提のHTML（■ マーカー付き）
        def assumption_cell(text):
            if not text:
                return ""
            return f'<div style="font-size:{fs_assumption}px;color:{TEXT_COLOR};line-height:1.4;text-align:left;">■ {text}</div>'

        body_rows += f"""
        <tr>
            <td class="cell item-col" style="background:{bg};font-weight:700;
                font-size:{fs_value}px;text-align:center;vertical-align:middle;
                padding:{row_pad}px {int(4*PT)}px;">
                {item_name}
            </td>
            <td class="cell mgmt-col" style="background:{bg};vertical-align:middle;
                padding:{row_pad}px {int(4*PT)}px;text-align:center;">
                {val_cell(mgmt_val, mgmt_pct)}
            </td>
            <td class="cell ds-val-col" style="background:{bg};vertical-align:middle;
                padding:{row_pad}px {int(4*PT)}px;text-align:center;">
                {val_cell(ds_val, ds_pct)}
            </td>
            <td class="cell ds-assumption-col" style="background:{bg};vertical-align:middle;
                padding:{row_pad}px {int(6*PT)}px;">
                {assumption_cell(ds_assumption)}
            </td>
            <td class="cell us-val-col" style="background:{bg};vertical-align:middle;
                padding:{row_pad}px {int(4*PT)}px;text-align:center;">
                {val_cell(us_val, us_pct)}
            </td>
            <td class="cell us-assumption-col" style="background:{bg};vertical-align:middle;
                padding:{row_pad}px {int(6*PT)}px;">
                {assumption_cell(us_assumption)}
            </td>
        </tr>"""

    body_pad = int(4 * PT)
    border_style = f"1px solid {BORDER_COLOR}"

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
table {{
    width:100%;
    height:100%;
    border-collapse:collapse;
    table-layout:fixed;
}}
.cell {{
    border:{border_style};
}}
.item-col {{ width:10%; }}
.mgmt-col {{ width:8%; }}
.ds-val-col {{ width:8%; }}
.ds-assumption-col {{ width:32%; }}
.us-val-col {{ width:8%; }}
.us-assumption-col {{ width:34%; }}
</style>
</head>
<body>
<table>
{header_html}
<tbody>
{body_rows}
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

    print(f"  Screenshot saved: {output_path}")


def insert_screenshot_into_pptx(template_path, data, screenshot_path, output_path):
    """テンプレートにデータとスクリーンショットを挿入"""
    prs = Presentation(template_path)
    slide = prs.slides[0]

    # Shape一覧を表示（デバッグ用）
    for s in slide.shapes:
        print(f"  Shape: '{s.name}' type={s.shape_type}")

    # 1. メインメッセージ
    main_msg = data.get("main_message", "")
    msg_shape = find_shape(slide, SHAPE_MAIN_MESSAGE)
    set_textbox_text(msg_shape, main_msg)

    # 2. Chart Title
    title_text = data.get("chart_title", "")
    title_shape = find_shape(slide, SHAPE_CHART_TITLE)
    set_textbox_text(title_shape, title_text)

    # 3. 出典
    source_text = data.get("source", "受領資料、Q＆A、マネジメントインタビューより当社作成")
    source_shape = find_shape(slide, SHAPE_SOURCE)
    if source_text:
        set_textbox_text(source_shape, f"出典：{source_text}")

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
        print(f"  Screenshot inserted at ({left},{top}) size=({width},{height})")

    # 保存
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    prs.save(output_path)
    _finalize_pptx(output_path)
    print(f"  PPTX saved: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="当期着地見込みスライド生成")
    parser.add_argument("--data", required=True, help="JSONデータファイルパス")
    parser.add_argument("--template", required=True, help="PPTXテンプレートパス")
    parser.add_argument("--output", required=True, help="出力PPTXファイルパス")
    add_brand_arg(parser)  # passive: accepted but ignored until brand migration
    args = parser.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    print("=== 当期着地見込みスライド生成 ===")

    # Step 1: HTML生成
    print("Step 1: HTML生成...")
    html_content = generate_html(data)

    # デバッグ: HTML保存
    html_debug = os.path.join(tempfile.gettempdir(), "forecast_debug.html")
    with open(html_debug, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"  Debug HTML: {html_debug}")

    # Step 2: スクリーンショット
    print("Step 2: スクリーンショット取得...")
    screenshot_path = os.path.join(tempfile.gettempdir(), "forecast_screenshot.png")
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
