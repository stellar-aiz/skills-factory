"""
fill_cost_breakdown.py — コスト内訳推移チャート（1〜2チャート並列）をPPTXネイティブオブジェクトで生成

生成するネイティブオブジェクト（すべてPowerPoint上で人間が編集可能）:
  - 積み上げ棒チャート＋折れ線（PowerPointネイティブ複合チャート）
  - 各セグメントの%ラベル（カスタムデータラベル）
  - トータルラベル: テキストボックス
  - 折れ線ラベル: テキストボックス
  - 凡例: Shape＋テキストボックス
  - チャートタイトル・単位ラベル: テキストボックス

Usage:
  python fill_cost_breakdown.py \
    --data /home/claude/cost_breakdown_data.json \
    --template <SKILL_DIR>/assets/cost-breakdown-template.pptx \
    --output /mnt/user-data/outputs/CostBreakdown_output.pptx
"""

import argparse, copy, json, math, os, sys

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


# ── Shape名マッピング（テンプレート） ──
SHAPE_MAIN_MESSAGE = "Title 1"
SHAPE_CHART_TITLE  = "Text Placeholder 2"
SHAPE_CONTENT_AREA = "Content Area"
SHAPE_SOURCE       = "Source"

# ── レイアウト定数 ──
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.50)
PANEL_Y = Inches(1.50)

# 1チャートモード
SINGLE_CHART_X = Inches(0.50)
SINGLE_CHART_W = Inches(12.30)

# 2チャートモード
LEFT_CHART_X  = Inches(0.35)
LEFT_CHART_W  = Inches(6.15)
RIGHT_CHART_X = Inches(6.75)
RIGHT_CHART_W = Inches(6.15)

# 共通チャート高さ
CHART_TITLE_H = Inches(0.30)
UNIT_LABEL_H  = Inches(0.20)
LEGEND_H      = Inches(0.50)
CHART_H       = Inches(4.20)

COLOR_TEXT   = RGBColor(0x33, 0x33, 0x33)
COLOR_SOURCE = RGBColor(0x66, 0x66, 0x66)
FONT_JP      = "Meiryo UI"

# ── デフォルトカラーパレット ──
DEFAULT_COLORS = [
    "#4E79A7", "#59A14F", "#E8923F", "#F28E2B", "#76B7B2",
    "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F", "#BAB0AC",
]


# ── ユーティリティ ──
def find_shape(slide, name):
    for s in slide.shapes:
        if s.name == name:
            return s
    print(f"  ⚠ Shape '{name}' not found", file=sys.stderr)
    return None


def set_textbox_text(shape, text):
    if shape is None:
        return
    tf = shape.text_frame
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        r = etree.SubElement(p._p, qn("a:r"))
        etree.SubElement(r, qn("a:rPr"), attrib={"lang": "ja-JP"})
        t = etree.SubElement(r, qn("a:t"))
        t.text = text


def remove_shape(slide, name):
    s = find_shape(slide, name)
    if s:
        slide.shapes._spTree.remove(s._element)
        print(f"  ✓ Removed '{name}'")


def _hex2rgb(h):
    h = h.replace("#", "")
    return RGBColor(int(h[:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_textbox(slide, left, top, width, height, text, font_size=Pt(11),
                bold=False, color=COLOR_TEXT, alignment=PP_ALIGN.LEFT, font_name=FONT_JP):
    """汎用テキストボックス追加"""
    tb = slide.shapes.add_textbox(int(left), int(top), int(width), int(height))
    tb.text_frame.word_wrap = True
    p = tb.text_frame.paragraphs[0]
    p.alignment = alignment
    r = p.add_run()
    r.text = text
    r.font.size = font_size
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = font_name
    return tb


# ── カスタムデータラベル ──
def _add_custom_pct_labels(ser_xml, shares, min_share=2.0, font_color='FFFFFF', font_size=800):
    """棒セグメントにカスタム%ラベルを追加（OOXML直接操作）"""
    old = ser_xml.find(qn('c:dLbls'))
    if old is not None:
        ser_xml.remove(old)

    dLbls = etree.SubElement(ser_xml, qn('c:dLbls'))

    for i, sv in enumerate(shares):
        dLbl = etree.SubElement(dLbls, qn('c:dLbl'))
        etree.SubElement(dLbl, qn('c:idx'), attrib={'val': str(i)})

        if sv < min_share:
            etree.SubElement(dLbl, qn('c:delete'), attrib={'val': '1'})
            continue

        # カスタムテキスト（%表示）
        tx = etree.SubElement(dLbl, qn('c:tx'))
        rich = etree.SubElement(tx, qn('c:rich'))
        etree.SubElement(rich, qn('a:bodyPr'))
        etree.SubElement(rich, qn('a:lstStyle'))
        p = etree.SubElement(rich, qn('a:p'))
        pp = etree.SubElement(p, qn('a:pPr'))
        etree.SubElement(pp, qn('a:defRPr'))
        r = etree.SubElement(p, qn('a:r'))
        rPr = etree.SubElement(r, qn('a:rPr'), attrib={
            'lang': 'ja-JP', 'sz': str(font_size), 'b': '1'
        })
        sf = etree.SubElement(rPr, qn('a:solidFill'))
        etree.SubElement(sf, qn('a:srgbClr'), attrib={'val': font_color})
        etree.SubElement(rPr, qn('a:latin'), attrib={'typeface': FONT_JP})
        etree.SubElement(rPr, qn('a:ea'), attrib={'typeface': FONT_JP})
        t = etree.SubElement(r, qn('a:t'))
        t.text = f"{sv:.1f}%"

        etree.SubElement(dLbl, qn('c:dLblPos'), attrib={'val': 'ctr'})
        for k in ['showLegendKey', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize']:
            etree.SubElement(dLbl, qn(f'c:{k}'), attrib={'val': '0'})
        etree.SubElement(dLbl, qn('c:showVal'), attrib={'val': '0'})

    # グローバル設定
    for k in ['showLegendKey', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize']:
        etree.SubElement(dLbls, qn(f'c:{k}'), attrib={'val': '0'})
    etree.SubElement(dLbls, qn('c:showVal'), attrib={'val': '0'})


def _add_line_labels(ser_xml, fmt='0.0', font_color='333333', font_size=900):
    """折れ線にデータラベル追加"""
    dl = ser_xml.find(qn('c:dLbls'))
    if dl is None:
        dl = etree.SubElement(ser_xml, qn('c:dLbls'))
    for c in list(dl):
        dl.remove(c)
    etree.SubElement(dl, qn('c:numFmt'), attrib={'formatCode': fmt + '"%"', 'sourceLinked': '0'})
    for k in ['showLegendKey', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize']:
        etree.SubElement(dl, qn(f'c:{k}'), attrib={'val': '0'})
    etree.SubElement(dl, qn('c:showVal'), attrib={'val': '1'})
    etree.SubElement(dl, qn('c:dLblPos'), attrib={'val': 't'})
    tp = etree.SubElement(dl, qn('c:txPr'))
    etree.SubElement(tp, qn('a:bodyPr'))
    etree.SubElement(tp, qn('a:lstStyle'))
    p = etree.SubElement(tp, qn('a:p'))
    pp = etree.SubElement(p, qn('a:pPr'))
    dr = etree.SubElement(pp, qn('a:defRPr'), attrib={'sz': str(font_size), 'b': '1'})
    etree.SubElement(dr, qn('a:latin'), attrib={'typeface': FONT_JP})
    etree.SubElement(dr, qn('a:ea'), attrib={'typeface': FONT_JP})
    sf = etree.SubElement(dr, qn('a:solidFill'))
    etree.SubElement(sf, qn('a:srgbClr'), attrib={'val': font_color})


# ── チャート生成 ──
def build_chart(slide, chart_cfg, chart_left, chart_top, chart_w, chart_h):
    """積み上げ棒＋折れ線の複合チャートを生成"""
    from pptx.chart.data import CategoryChartData

    data = chart_cfg.get("data", [])
    bars_cfg = chart_cfg.get("stacked_bars", [])
    line_cfg = chart_cfg.get("line", None)
    n_bars = len(bars_cfg)
    n_periods = len(data)
    has_line = line_cfg is not None

    # チャートデータ構築
    cd = CategoryChartData()
    cd.categories = [d["year"] for d in data]
    for si, sb in enumerate(bars_cfg):
        cd.add_series(sb["series_name"],
                      [d["bars"][si] if si < len(d.get("bars", [])) else 0 for d in data])
    if has_line:
        cd.add_series(line_cfg["series_name"], [d.get("line_value", 0) for d in data])

    # チャート作成（積み上げ棒）
    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED,
        int(chart_left), int(chart_top), int(chart_w), int(chart_h),
        cd
    )
    chart = cf.chart
    pa = chart._chartSpace.chart.plotArea
    bc = pa.findall(qn('c:barChart'))[0]

    # 第2軸ID（折れ線用）
    sec_v = "2094734553"
    sec_c = "2094734554"

    # ── 折れ線を分離 ──
    if has_line:
        sers = bc.findall(qn('c:ser'))
        ls = copy.deepcopy(sers[-1])
        bc.remove(sers[-1])
        lc = etree.SubElement(pa, qn('c:lineChart'))
        etree.SubElement(lc, qn('c:grouping'), attrib={'val': 'standard'})
        etree.SubElement(lc, qn('c:varyColors'), attrib={'val': '0'})
        lc.append(ls)

        # マーカー
        mk = ls.find(qn('c:marker'))
        if mk is None:
            mk = etree.SubElement(ls, qn('c:marker'))
        sym = mk.find(qn('c:symbol'))
        if sym is None:
            sym = etree.SubElement(mk, qn('c:symbol'))
        sym.set('val', 'circle')
        sz_el = mk.find(qn('c:size'))
        if sz_el is None:
            sz_el = etree.SubElement(mk, qn('c:size'))
        sz_el.set('val', str(line_cfg.get("marker_size", 7)))

        # 第2軸
        etree.SubElement(lc, qn('c:axId'), attrib={'val': sec_c})
        etree.SubElement(lc, qn('c:axId'), attrib={'val': sec_v})
        sc = etree.SubElement(pa, qn('c:catAx'))
        etree.SubElement(sc, qn('c:axId'), attrib={'val': sec_c})
        s2 = etree.SubElement(sc, qn('c:scaling'))
        etree.SubElement(s2, qn('c:orientation'), attrib={'val': 'minMax'})
        etree.SubElement(sc, qn('c:delete'), attrib={'val': '1'})
        etree.SubElement(sc, qn('c:axPos'), attrib={'val': 'b'})
        etree.SubElement(sc, qn('c:crossAx'), attrib={'val': sec_v})
        sv_ax = etree.SubElement(pa, qn('c:valAx'))
        etree.SubElement(sv_ax, qn('c:axId'), attrib={'val': sec_v})
        s3 = etree.SubElement(sv_ax, qn('c:scaling'))
        etree.SubElement(s3, qn('c:orientation'), attrib={'val': 'minMax'})
        etree.SubElement(sv_ax, qn('c:delete'), attrib={'val': '1'})
        etree.SubElement(sv_ax, qn('c:axPos'), attrib={'val': 'r'})
        etree.SubElement(sv_ax, qn('c:numFmt'), attrib={'formatCode': '0.0', 'sourceLinked': '0'})
        etree.SubElement(sv_ax, qn('c:crossAx'), attrib={'val': sec_c})
        etree.SubElement(sv_ax, qn('c:crosses'), attrib={'val': 'max'})

        # 折れ線色
        lclr = line_cfg.get("color", "#4E79A7").replace("#", "")
        lsp = ls.find(qn('c:spPr'))
        if lsp is None:
            lsp = etree.SubElement(ls, qn('c:spPr'))
        ln = lsp.find(qn('a:ln'))
        if ln is None:
            ln = etree.SubElement(lsp, qn('a:ln'))
        ln.set('w', '19050')
        lsf = ln.find(qn('a:solidFill'))
        if lsf is None:
            lsf = etree.SubElement(ln, qn('a:solidFill'))
        for c in list(lsf):
            lsf.remove(c)
        etree.SubElement(lsf, qn('a:srgbClr'), attrib={'val': lclr})
        # マーカー色
        msp = mk.find(qn('c:spPr'))
        if msp is None:
            msp = etree.SubElement(mk, qn('c:spPr'))
        msf = etree.SubElement(msp, qn('a:solidFill'))
        etree.SubElement(msf, qn('a:srgbClr'), attrib={'val': lclr})
        # 折れ線ラベル
        _add_line_labels(ls)

    # ── 棒色＋カスタム%ラベル ──
    bar_sers = bc.findall(qn('c:ser'))
    for si, bs in enumerate(bar_sers):
        clr = bars_cfg[si]["color"].replace("#", "") if si < len(bars_cfg) else "999999"
        sp = bs.find(qn('c:spPr'))
        if sp is None:
            sp = etree.SubElement(bs, qn('c:spPr'))
        sf = sp.find(qn('a:solidFill'))
        if sf is None:
            sf = etree.SubElement(sp, qn('a:solidFill'))
        for c in list(sf):
            sf.remove(c)
        etree.SubElement(sf, qn('a:srgbClr'), attrib={'val': clr})

        # この系列の各期のshare値を取得
        shares = []
        for d in data:
            s_arr = d.get("shares", [])
            shares.append(s_arr[si] if si < len(s_arr) else 0)
        _add_custom_pct_labels(bs, shares)

    # ── 棒間隔 ──
    gw = bc.find(qn('c:gapWidth'))
    if gw is None:
        gw = etree.SubElement(bc, qn('c:gapWidth'))
    gw.set('val', '80')
    ov = bc.find(qn('c:overlap'))
    if ov is None:
        ov = etree.SubElement(bc, qn('c:overlap'))
    ov.set('val', '100')

    # ── 第1数値軸を非表示 ──
    for vx in pa.findall(qn('c:valAx')):
        axId = vx.find(qn('c:axId'))
        if axId is not None and axId.get('val') == sec_v:
            continue  # 第2軸はスキップ
        de = vx.find(qn('c:delete'))
        if de is None:
            de = etree.SubElement(vx, qn('c:delete'))
        de.set('val', '1')

    # ── グリッド線削除 ──
    for vx in pa.findall(qn('c:valAx')):
        for mg in vx.findall(qn('c:majorGridlines')):
            vx.remove(mg)
        for mg in vx.findall(qn('c:minorGridlines')):
            vx.remove(mg)
    for cx in pa.findall(qn('c:catAx')):
        for mg in cx.findall(qn('c:majorGridlines')):
            cx.remove(mg)

    # ── 軸フォント（カテゴリ軸 = 年ラベル） ──
    for ax in pa.findall(qn('c:catAx')):
        de = ax.find(qn('c:delete'))
        if de is not None and de.get('val') == '1':
            continue
        tp = ax.find(qn('c:txPr'))
        if tp is None:
            tp = etree.SubElement(ax, qn('c:txPr'))
        bp = tp.find(qn('a:bodyPr'))
        if bp is None:
            bp = etree.SubElement(tp, qn('a:bodyPr'))
            tp.insert(0, bp)
        if tp.find(qn('a:lstStyle')) is None:
            etree.SubElement(tp, qn('a:lstStyle'))
        p = tp.find(qn('a:p'))
        if p is None:
            p = etree.SubElement(tp, qn('a:p'))
        pp = p.find(qn('a:pPr'))
        if pp is None:
            pp = etree.SubElement(p, qn('a:pPr'))
        dr = pp.find(qn('a:defRPr'))
        if dr is None:
            dr = etree.SubElement(pp, qn('a:defRPr'), attrib={'sz': '1000'})
        else:
            dr.set('sz', '1000')
        if dr.find(qn('a:latin')) is None:
            etree.SubElement(dr, qn('a:latin'), attrib={'typeface': FONT_JP})
        if dr.find(qn('a:ea')) is None:
            etree.SubElement(dr, qn('a:ea'), attrib={'typeface': FONT_JP})
        dsf = dr.find(qn('a:solidFill'))
        if dsf is None:
            dsf = etree.SubElement(dr, qn('a:solidFill'))
        for c in list(dsf):
            dsf.remove(c)
        etree.SubElement(dsf, qn('a:srgbClr'), attrib={'val': '333333'})

    # ── 凡例・タイトル非表示 ──
    chart.has_legend = False
    chart.has_title = False

    # ── プロットエリア背景透明 ──
    psp = pa.find(qn('c:spPr'))
    if psp is None:
        psp = etree.SubElement(pa, qn('c:spPr'))
    nf = psp.find(qn('a:noFill'))
    if nf is None:
        etree.SubElement(psp, qn('a:noFill'))

    print(f"  ✓ 複合チャート: {n_periods}期, {n_bars}棒系列" + (f" + 折れ線" if has_line else ""))
    return cf


# ── トータルラベル追加（テキストボックス） ──
def add_total_labels(slide, chart_cfg, cl, ct, cw, ch):
    """棒の上にトータルラベル（合計値）をテキストボックスで追加"""
    data = chart_cfg.get("data", [])
    if not data:
        return
    nc = len(data)

    # プロットエリア推定
    plm = cw * 0.06
    prm = cw * 0.04
    ptm = ch * 0.05
    pbm = ch * 0.14
    pl = cl + plm
    pw = cw - plm - prm
    pb = ct + ch - pbm
    ph = (ct + ch - pbm) - (ct + ptm)

    # Y軸スケール推定
    totals = [sum(d.get("bars", [])) for d in data]
    max_total = max(totals) if totals else 1
    axis_max = max_total * 1.25

    bar_w = pw / nc
    for i, d in enumerate(data):
        tl = d.get("total_label", "")
        if not tl:
            continue
        total = sum(d.get("bars", []))
        bar_cx = pl + (i + 0.5) * bar_w
        bar_top_y = pb - (total / axis_max) * ph

        lw = Inches(0.80)
        lh = Inches(0.25)
        lx = bar_cx - lw / 2
        ly = bar_top_y - lh - Inches(0.02)

        add_textbox(slide, lx, ly, lw, lh, tl,
                    font_size=Pt(10), bold=True, color=COLOR_TEXT,
                    alignment=PP_ALIGN.CENTER)
    print(f"  ✓ トータルラベル: {nc}個")


# ── 凡例 ──
def add_legend(slide, chart_cfg, left, top, max_width):
    """チャートごとの凡例を描画"""
    bars = chart_cfg.get("stacked_bars", [])
    lc = chart_cfg.get("line", None)
    lh = Inches(0.20)
    ix = left
    row_y = top
    item_count = 0

    # 折れ線凡例（先頭に表示）
    if lc:
        c = _hex2rgb(lc.get("color", "#4E79A7"))
        line_w = Inches(0.30)
        ly = int(row_y + lh / 2)
        cn = slide.shapes.add_connector(1, int(ix), ly, int(ix + line_w), ly)
        cn.line.color.rgb = c
        cn.line.width = Pt(1.5)
        cs = Inches(0.10)
        ci = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            int(ix + (line_w - cs) / 2), int(row_y + (lh - cs) / 2), int(cs), int(cs)
        )
        ci.fill.solid()
        ci.fill.fore_color.rgb = c
        ci.line.fill.background()
        tx = int(ix + line_w + Inches(0.04))
        name_w = Inches(min(len(lc["series_name"]) * 0.13, 2.0))
        tb = slide.shapes.add_textbox(tx, int(row_y), int(name_w), int(lh))
        tb.text_frame.word_wrap = False
        r = tb.text_frame.paragraphs[0].add_run()
        r.text = lc["series_name"]
        r.font.size = Pt(8)
        r.font.color.rgb = COLOR_TEXT
        r.font.name = FONT_JP
        ix = tx + name_w + Inches(0.12)
        item_count += 1

    # 棒凡例
    for sb in bars:
        c = _hex2rgb(sb["color"])
        ss = Inches(0.12)
        name_w = Inches(min(len(sb["series_name"]) * 0.11 + 0.15, 1.8))
        item_w = ss + Inches(0.04) + name_w + Inches(0.06)

        # 行折り返しチェック
        if ix + item_w > left + max_width and item_count > 0:
            row_y += lh + Inches(0.02)
            ix = left
            item_count = 0

        sq = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            int(ix), int(row_y + (lh - ss) / 2), int(ss), int(ss)
        )
        sq.fill.solid()
        sq.fill.fore_color.rgb = c
        sq.line.fill.background()
        tx = int(ix + ss + Inches(0.04))
        tb = slide.shapes.add_textbox(tx, int(row_y), int(name_w), int(lh))
        tb.text_frame.word_wrap = False
        r = tb.text_frame.paragraphs[0].add_run()
        r.text = sb["series_name"]
        r.font.size = Pt(8)
        r.font.color.rgb = COLOR_TEXT
        r.font.name = FONT_JP
        ix = tx + name_w + Inches(0.06)
        item_count += 1

    print(f"  ✓ 凡例: {len(bars)}棒" + (f" + 折れ線" if lc else ""))
    return row_y + lh


# ── 1チャート分のレンダリング ──
def render_one_chart(slide, chart_cfg, area_left, area_top, area_width):
    """1つのチャート領域をレンダリング"""
    y = area_top

    # チャートタイトル
    title = chart_cfg.get("title", "")
    if title:
        add_textbox(slide, area_left, y, area_width, CHART_TITLE_H, title,
                    font_size=Pt(12), bold=True, color=COLOR_TEXT,
                    alignment=PP_ALIGN.CENTER)
        ul = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            int(area_left + area_width * 0.10), int(y + CHART_TITLE_H),
            int(area_width * 0.80), Inches(0.012)
        )
        ul.fill.solid()
        ul.fill.fore_color.rgb = COLOR_TEXT
        ul.line.fill.background()
    y += CHART_TITLE_H + Inches(0.05)

    # 単位ラベル
    unit = chart_cfg.get("unit_label", "")
    if unit:
        add_textbox(slide, area_left, y, area_width, UNIT_LABEL_H, unit,
                    font_size=Pt(8), bold=False, color=COLOR_SOURCE)
    y += UNIT_LABEL_H + Inches(0.02)

    # 凡例
    legend_bottom = add_legend(slide, chart_cfg, area_left, y, area_width)
    y = legend_bottom + Inches(0.08)

    # 残り高さからチャート高さを計算
    available_h = Inches(6.30) - y
    chart_h = min(available_h, CHART_H)

    # チャート生成
    build_chart(slide, chart_cfg, area_left, y, area_width, chart_h)

    # トータルラベル
    add_total_labels(slide, chart_cfg, area_left, y, area_width, chart_h)

    print(f"  ✓ チャート領域完了: '{title}'")


# ── メイン ──
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--output", required=True)
    add_brand_arg(ap)  # passive: accepted but ignored until brand migration
    args = ap.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    print("=== コスト内訳推移スライド生成（ネイティブPPTX）===")
    prs = Presentation(args.template)
    slide = prs.slides[0]

    # テンプレートのテキスト設定
    set_textbox_text(find_shape(slide, SHAPE_MAIN_MESSAGE), data.get("main_message", ""))
    set_textbox_text(find_shape(slide, SHAPE_CHART_TITLE), "")
    src = data.get("source", "")
    if src:
        set_textbox_text(find_shape(slide, SHAPE_SOURCE), f"出典：{src}")

    # Content Areaプレースホルダーを削除
    remove_shape(slide, SHAPE_CONTENT_AREA)

    charts = data.get("charts", [])
    n_charts = len(charts)

    if n_charts == 0:
        print("  ⚠ chartsが空です")
    elif n_charts == 1:
        print("\n── 1チャートモード ──")
        _apply_default_colors(charts[0])
        render_one_chart(slide, charts[0], SINGLE_CHART_X, PANEL_Y, SINGLE_CHART_W)
    else:
        print("\n── 2チャートモード ──")
        _apply_default_colors(charts[0])
        _apply_default_colors(charts[1])
        print("\n[左チャート]")
        render_one_chart(slide, charts[0], LEFT_CHART_X, PANEL_Y, LEFT_CHART_W)
        print("\n[右チャート]")
        render_one_chart(slide, charts[1], RIGHT_CHART_X, PANEL_Y, RIGHT_CHART_W)

    os.makedirs(os.path.dirname(args.output) or ".", exist_ok=True)
    prs.save(args.output)
    _finalize_pptx(args.output)
    print(f"\n  ✅ 出力完了: {args.output}")


def _apply_default_colors(chart_cfg):
    """色が未指定の系列にデフォルト色を適用"""
    bars = chart_cfg.get("stacked_bars", [])
    for i, sb in enumerate(bars):
        if "color" not in sb or not sb["color"]:
            sb["color"] = DEFAULT_COLORS[i % len(DEFAULT_COLORS)]


if __name__ == "__main__":
    main()
