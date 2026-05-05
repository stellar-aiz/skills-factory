"""
fill_revenue_analysis.py - Revenue Analysis slide generator (v4)

Layout: Top = EBITDA margin line, Middle = CAGR arrow, Bottom = Grouped bar chart (Revenue + EBITDA)
Grouped bar chart uses two series: Revenue (blue) and EBITDA (orange).
"""
import argparse, json, math, os, sys

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


SHAPE_MAIN_MESSAGE = "Title 1"
SHAPE_CHART_TITLE = "Text Placeholder 2"

# Upper zone: EBITDA margin line (shapes only)
MARGIN_ZONE_TOP = Inches(1.80)
MARGIN_ZONE_BOTTOM = Inches(3.10)

# Lower zone: Native grouped bar chart
CHART_LEFT = Inches(0.80)
CHART_WIDTH = Inches(11.73)
CHART_TOP = Inches(3.20)
CHART_HEIGHT = Inches(3.70)

COLOR_TEXT = RGBColor(0x33, 0x33, 0x33)
COLOR_BAR_REV = RGBColor(0x4E, 0x79, 0xA7)   # Revenue - steel blue
COLOR_BAR_EBITDA = RGBColor(0xE1, 0x81, 0x2C)  # EBITDA - orange
COLOR_LINE = RGBColor(0x00, 0x33, 0x66)
COLOR_CAGR = RGBColor(0x33, 0x33, 0x33)
FONT = "Meiryo UI"

PLOT_L_PCT, PLOT_R_PCT, PLOT_T_PCT, PLOT_B_PCT = 0.07, 0.02, 0.04, 0.12


def find_shape(slide, name):
    for s in slide.shapes:
        if s.name == name: return s
    return None

def set_text(shape, text):
    if not shape: return
    tf = shape.text_frame; p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]: r.text = ""
    else:
        r = etree.SubElement(p._p, qn("a:r"))
        etree.SubElement(r, qn("a:rPr"), attrib={"lang": "ja-JP"})
        t = etree.SubElement(r, qn("a:t")); t.text = text

def remove_shape(slide, name):
    s = find_shape(slide, name)
    if s: slide.shapes._spTree.remove(s._element)

def nice_max(val):
    if val <= 0: return 100
    t = val * 1.20
    mag = 10 ** math.floor(math.log10(t))
    for s in [1, 2, 2.5, 5, 10]:
        c = math.ceil(t / (s * mag)) * (s * mag)
        if c >= t: return int(c) if c == int(c) else c
    return math.ceil(t / mag) * mag

def nice_tick(axis_max):
    raw = axis_max / 5
    mag = 10 ** math.floor(math.log10(raw))
    for n in [1, 2, 2.5, 5, 10]:
        c = n * mag
        if c >= raw * 0.8: return int(c) if c == int(c) else c
    return raw

def fmt_num(v):
    return f"{int(v):,}" if v == int(v) else f"{v:,.1f}"

def chart_plot_bounds():
    cl, ct, cw, ch = int(CHART_LEFT), int(CHART_TOP), int(CHART_WIDTH), int(CHART_HEIGHT)
    return (cl + cw*PLOT_L_PCT, ct + ch*PLOT_T_PCT,
            cl + cw*(1-PLOT_R_PCT), ct + ch*(1-PLOT_B_PCT))

def bar_cx(n, i):
    """X center of the i-th category group"""
    pl, _, pr, _ = chart_plot_bounds()
    w = (pr - pl) / n
    return pl + (i + 0.5) * w

def bar_top_y(val, amax):
    _, pt, _, pb = chart_plot_bounds()
    return pb - (val / amax * (pb - pt)) if amax > 0 else pb

def add_dlabels(ser, pos='outEnd', fmt='#,##0', color='333333', sz=1100, bold=False):
    dl = ser.find(qn('c:dLbls'))
    if dl is None: dl = etree.SubElement(ser, qn('c:dLbls'))
    for c in list(dl): dl.remove(c)
    etree.SubElement(dl, qn('c:numFmt'), attrib={'formatCode': fmt, 'sourceLinked': '0'})
    for tag, val in [('showLegendKey','0'),('showVal','1'),('showCatName','0'),
                     ('showSerName','0'),('showPercent','0'),('showBubbleSize','0')]:
        etree.SubElement(dl, qn(f'c:{tag}'), attrib={'val': val})
    etree.SubElement(dl, qn('c:dLblPos'), attrib={'val': pos})
    txPr = etree.SubElement(dl, qn('c:txPr'))
    etree.SubElement(txPr, qn('a:bodyPr'))
    etree.SubElement(txPr, qn('a:lstStyle'))
    p = etree.SubElement(txPr, qn('a:p'))
    pPr = etree.SubElement(p, qn('a:pPr'))
    att = {'sz': str(sz)}
    if bold: att['b'] = '1'
    dr = etree.SubElement(pPr, qn('a:defRPr'), attrib=att)
    etree.SubElement(dr, qn('a:latin'), attrib={'typeface': FONT})
    etree.SubElement(dr, qn('a:ea'), attrib={'typeface': FONT})
    sf = etree.SubElement(dr, qn('a:solidFill'))
    etree.SubElement(sf, qn('a:srgbClr'), attrib={'val': color})

def set_ser_color(ser, hex_color):
    """Set fill color of a bar chart series"""
    sp = ser.find(qn('c:spPr'))
    if sp is None: sp = etree.SubElement(ser, qn('c:spPr'))
    sf = sp.find(qn('a:solidFill'))
    if sf is None: sf = etree.SubElement(sp, qn('a:solidFill'))
    for c in list(sf): sf.remove(c)
    etree.SubElement(sf, qn('a:srgbClr'), attrib={'val': hex_color})


def build_grouped_bar_chart(slide, jd):
    """Build grouped bar chart with Revenue + EBITDA in the LOWER zone."""
    from pptx.chart.data import CategoryChartData
    data = jd["data"]
    rev_label = jd.get("bar_label", "売上高")
    ebitda_label = "EBITDA"

    cd = CategoryChartData()
    cd.categories = [d["year"] for d in data]
    cd.add_series(rev_label, [d["revenue"] for d in data])
    cd.add_series(ebitda_label, [d["ebitda"] for d in data])

    cf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                                CHART_LEFT, CHART_TOP, CHART_WIDTH, CHART_HEIGHT, cd)
    chart = cf.chart
    pa = chart._chartSpace.chart.plotArea
    bc = pa.findall(qn('c:barChart'))[0]

    # Show left Y-axis
    vax = pa.findall(qn('c:valAx'))[0]
    de = vax.find(qn('c:delete'))
    if de is None: de = etree.SubElement(vax, qn('c:delete'))
    de.set('val', '0')

    revs = [d["revenue"] for d in data]
    amax = nice_max(max(revs))
    tick = nice_tick(amax)

    sc = vax.find(qn('c:scaling'))
    if sc is None: sc = etree.SubElement(vax, qn('c:scaling'))
    for tag, val in [('c:max', str(amax)), ('c:min', '0')]:
        e = sc.find(qn(tag))
        if e is None: e = etree.SubElement(sc, qn(tag))
        e.set('val', val)
    mu = vax.find(qn('c:majorUnit'))
    if mu is None: mu = etree.SubElement(vax, qn('c:majorUnit'))
    mu.set('val', str(tick))
    nf = vax.find(qn('c:numFmt'))
    if nf is None: nf = etree.SubElement(vax, qn('c:numFmt'))
    nf.set('formatCode', '#,##0'); nf.set('sourceLinked', '0')

    # Y-axis font
    tp = vax.find(qn('c:txPr'))
    if tp is None: tp = etree.SubElement(vax, qn('c:txPr'))
    bp = tp.find(qn('a:bodyPr'))
    if bp is None: bp = etree.SubElement(tp, qn('a:bodyPr')); tp.insert(0, bp)
    if tp.find(qn('a:lstStyle')) is None: etree.SubElement(tp, qn('a:lstStyle'))
    p = tp.find(qn('a:p'))
    if p is None: p = etree.SubElement(tp, qn('a:p'))
    pp = p.find(qn('a:pPr'))
    if pp is None: pp = etree.SubElement(p, qn('a:pPr'))
    dr = pp.find(qn('a:defRPr'))
    if dr is None: dr = etree.SubElement(pp, qn('a:defRPr'), attrib={'sz':'1100'})
    else: dr.set('sz', '1100')
    if dr.find(qn('a:latin')) is None: etree.SubElement(dr, qn('a:latin'), attrib={'typeface': FONT})
    if dr.find(qn('a:ea')) is None: etree.SubElement(dr, qn('a:ea'), attrib={'typeface': FONT})
    asf = dr.find(qn('a:solidFill'))
    if asf is None: asf = etree.SubElement(dr, qn('a:solidFill'))
    for c in list(asf): asf.remove(c)
    etree.SubElement(asf, qn('a:srgbClr'), attrib={'val': '333333'})

    # Series colors
    sers = bc.findall(qn('c:ser'))
    set_ser_color(sers[0], '4E79A7')  # Revenue - blue
    set_ser_color(sers[1], 'E1812C')  # EBITDA - orange

    # Data labels
    add_dlabels(sers[0], 'outEnd', '#,##0', '333333', 1100, True)   # Revenue
    add_dlabels(sers[1], 'outEnd', '#,##0', '333333', 1000, False)  # EBITDA

    chart.has_legend = False

    # Remove gridlines
    for v in pa.findall(qn('c:valAx')):
        for mg in v.findall(qn('c:majorGridlines')): v.remove(mg)
        for mg in v.findall(qn('c:minorGridlines')): v.remove(mg)
    for c in pa.findall(qn('c:catAx')):
        for mg in c.findall(qn('c:majorGridlines')): c.remove(mg)

    # Gap between groups (narrower for grouped)
    gw = bc.find(qn('c:gapWidth'))
    if gw is None: gw = etree.SubElement(bc, qn('c:gapWidth'))
    gw.set('val', '80')

    # Cat axis font
    for ax in pa.findall(qn('c:catAx')):
        de2 = ax.find(qn('c:delete'))
        if de2 is not None and de2.get('val') == '1': continue
        tp2 = ax.find(qn('c:txPr'))
        if tp2 is None: tp2 = etree.SubElement(ax, qn('c:txPr'))
        bp2 = tp2.find(qn('a:bodyPr'))
        if bp2 is None: bp2 = etree.SubElement(tp2, qn('a:bodyPr')); tp2.insert(0, bp2)
        bp2.set('rot', '0'); bp2.set('vert', 'horz')
        if tp2.find(qn('a:lstStyle')) is None: etree.SubElement(tp2, qn('a:lstStyle'))
        p2 = tp2.find(qn('a:p'))
        if p2 is None: p2 = etree.SubElement(tp2, qn('a:p'))
        pp2 = p2.find(qn('a:pPr'))
        if pp2 is None: pp2 = etree.SubElement(p2, qn('a:pPr'))
        dr2 = pp2.find(qn('a:defRPr'))
        if dr2 is None: dr2 = etree.SubElement(pp2, qn('a:defRPr'), attrib={'sz':'1200','b':'1'})
        else: dr2.set('sz','1200'); dr2.set('b','1')
        if dr2.find(qn('a:latin')) is None: etree.SubElement(dr2, qn('a:latin'), attrib={'typeface': FONT})
        if dr2.find(qn('a:ea')) is None: etree.SubElement(dr2, qn('a:ea'), attrib={'typeface': FONT})

    chart.has_title = False
    psp = pa.find(qn('c:spPr'))
    if psp is None: psp = etree.SubElement(pa, qn('c:spPr'))
    etree.SubElement(psp, qn('a:noFill'))

    print(f"  Grouped bar chart: {len(data)} years, axis_max={amax}, tick={tick}")
    return cf, amax


def draw_margin_line(slide, jd):
    """Draw EBITDA margin line in the UPPER zone."""
    data = jd["data"]
    n = len(data)
    margins = [round(d["ebitda"]/d["revenue"]*100, 1) if d["revenue"]>0 else 0 for d in data]
    mx, mn = max(margins), min(margins)
    rng = mx - mn

    zone_top = int(MARGIN_ZONE_TOP)
    zone_bot = int(MARGIN_ZONE_BOTTOM)
    label_space = Inches(0.30)
    usable_top = zone_top + int(label_space)
    usable_h = zone_bot - usable_top

    top_pct, bot_pct = 0.98, 0.90
    if rng < 0.5: top_pct = bot_pct = 0.95

    pts = []
    for i, m in enumerate(margins):
        cx = int(bar_cx(n, i))
        if rng >= 0.5:
            norm = (m - mn) / rng
            y = int(usable_top + usable_h * (1.0 - norm))
        else:
            y = int(usable_top + usable_h * 0.4)
        pts.append((cx, y))

    for i in range(len(pts)-1):
        x1,y1 = pts[i]; x2,y2 = pts[i+1]
        cn = slide.shapes.add_connector(1, x1, y1, x2, y2)
        cn.line.color.rgb = COLOR_LINE; cn.line.width = Pt(2.0)

    ms = Inches(0.16)
    for x,y in pts:
        c = slide.shapes.add_shape(MSO_SHAPE.OVAL, x-int(ms/2), y-int(ms/2), ms, ms)
        c.fill.solid(); c.fill.fore_color.rgb = COLOR_LINE; c.line.fill.background()

    lh, lw = Inches(0.25), Inches(0.80)
    for i,(x,y) in enumerate(pts):
        tb = slide.shapes.add_textbox(x-int(lw/2), y-int(lh)-Inches(0.04), lw, lh)
        tf = tb.text_frame; tf.word_wrap = False
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = f"{margins[i]:.1f}%"
        r.font.size = Pt(12); r.font.bold = True
        r.font.color.rgb = COLOR_LINE; r.font.name = FONT

    print(f"  Margin line: {n} pts, {mn:.1f}%-{mx:.1f}%")
    return margins


def add_cagr(slide, jd, amax):
    """Draw CAGR arrow above revenue bars."""
    data = jd["data"]
    if len(data) < 2: return
    fr, lr = data[0]["revenue"], data[-1]["revenue"]
    ny, nc = len(data)-1, len(data)
    if fr > 0 and lr > 0 and ny > 0:
        cagr = (lr/fr)**(1.0/ny) - 1
        txt = f"+{cagr*100:.1f}%" if cagr >= 0 else f"{cagr*100:.1f}%"
    else: txt = "N/A"

    x1 = int(bar_cx(nc, 0)); x2 = int(bar_cx(nc, nc-1))
    gap = Inches(0.40)
    y1 = int(bar_top_y(fr, amax) - gap)
    y2 = int(bar_top_y(lr, amax) - gap)

    cn = slide.shapes.add_connector(1, x1, y1, x2, y2)
    cn.line.color.rgb = COLOR_CAGR; cn.line.width = Pt(1.5)
    sp = cn._element.find(qn('p:spPr'))
    if sp is None: sp = cn._element.find(qn('a:spPr'))
    ln = sp.find(qn('a:ln'))
    if ln is None: ln = etree.SubElement(sp, qn('a:ln'))
    etree.SubElement(ln, qn('a:tailEnd'), attrib={'type':'triangle','w':'med','len':'med'})

    tw, th = Inches(1.30), Inches(0.40)
    midx, midy = (x1+x2)/2, (y1+y2)/2
    ov = slide.shapes.add_shape(MSO_SHAPE.OVAL, int(midx-tw/2), int(midy-th/2), tw, th)
    ov.fill.solid(); ov.fill.fore_color.rgb = RGBColor(0xFF,0xFF,0xFF)
    ov.line.color.rgb = COLOR_CAGR; ov.line.width = Pt(1.0)
    tf = ov.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = txt
    r.font.size = Pt(16); r.font.bold = True
    r.font.color.rgb = COLOR_TEXT; r.font.name = FONT
    print(f"  CAGR: {txt}")


def add_legend(slide, jd, left, top, width):
    """Custom legend: ■売上高  ■EBITDA  ●━EBITDA率"""
    rev_label = jd.get("bar_label", "売上高")
    line_label = jd.get("line_label", "EBITDA率")

    lh = Inches(0.22)
    lx = left + width - Inches(5.00)  # wider for 3 items
    ss = Inches(0.14)
    sy = top + (lh - ss) / 2

    cur_x = lx

    # ■ 売上高
    bm1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(cur_x), int(sy), ss, ss)
    bm1.fill.solid(); bm1.fill.fore_color.rgb = COLOR_BAR_REV; bm1.line.fill.background()
    cur_x += ss + Inches(0.05)
    tb1 = slide.shapes.add_textbox(int(cur_x), top, Inches(0.70), lh)
    tf1 = tb1.text_frame; tf1.word_wrap = False
    p1 = tf1.paragraphs[0]; p1.alignment = PP_ALIGN.LEFT
    r1 = p1.add_run(); r1.text = rev_label; r1.font.size = Pt(12)
    r1.font.color.rgb = COLOR_TEXT; r1.font.name = FONT
    cur_x += Inches(0.85)

    # ■ EBITDA
    bm2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(cur_x), int(sy), ss, ss)
    bm2.fill.solid(); bm2.fill.fore_color.rgb = COLOR_BAR_EBITDA; bm2.line.fill.background()
    cur_x += ss + Inches(0.05)
    tb2 = slide.shapes.add_textbox(int(cur_x), top, Inches(0.80), lh)
    tf2 = tb2.text_frame; tf2.word_wrap = False
    p2 = tf2.paragraphs[0]; p2.alignment = PP_ALIGN.LEFT
    r2 = p2.add_run(); r2.text = "EBITDA"; r2.font.size = Pt(12)
    r2.font.color.rgb = COLOR_TEXT; r2.font.name = FONT
    cur_x += Inches(1.00)

    # ●━ EBITDA率
    llw = Inches(0.30); lly = top + lh / 2
    cn = slide.shapes.add_connector(1, int(cur_x), int(lly), int(cur_x + llw), int(lly))
    cn.line.color.rgb = COLOR_LINE; cn.line.width = Pt(1.5)
    cs = Inches(0.12)
    cx2 = cur_x + (llw - cs) / 2; cy2 = top + (lh - cs) / 2
    cc = slide.shapes.add_shape(MSO_SHAPE.OVAL, int(cx2), int(cy2), cs, cs)
    cc.fill.solid(); cc.fill.fore_color.rgb = COLOR_LINE; cc.line.fill.background()
    cur_x += llw + Inches(0.05)
    tb3 = slide.shapes.add_textbox(int(cur_x), top, Inches(1.20), lh)
    tf3 = tb3.text_frame; tf3.word_wrap = False
    p3 = tf3.paragraphs[0]; p3.alignment = PP_ALIGN.LEFT
    r3 = p3.add_run(); r3.text = line_label; r3.font.size = Pt(12)
    r3.font.color.rgb = COLOR_TEXT; r3.font.name = FONT


def add_unit(slide, text, left, top):
    tb = slide.shapes.add_textbox(left, top, Inches(3.0), Inches(0.22))
    tf = tb.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = text; r.font.size = Pt(12)
    r.font.color.rgb = COLOR_TEXT; r.font.name = FONT

def add_source(slide, text):
    tb = slide.shapes.add_textbox(Inches(0.41), Inches(7.05), Inches(8.0), Inches(0.30))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = text; r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(0x66,0x66,0x66); r.font.name = FONT


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--output", required=True)
    add_brand_arg(ap)  # passive: accepted but ignored until brand migration
    args = ap.parse_args()
    with open(args.data, "r", encoding="utf-8") as f: jd = json.load(f)

    prs = Presentation(args.template)
    slide = prs.slides[0]

    set_text(find_shape(slide, SHAPE_MAIN_MESSAGE), jd.get("main_message",""))
    print(f"  Main Message: {jd.get('main_message','')}")
    set_text(find_shape(slide, SHAPE_CHART_TITLE), jd.get("chart_title","売上分析ー売上高・EBITDAの推移"))
    print(f"  Chart Title: {jd.get('chart_title','')}")
    remove_shape(slide, "Table 1")

    lt = MARGIN_ZONE_TOP - Inches(0.25)
    ul = jd.get("unit_label","（単位：百万円、%）")
    if ul: add_unit(slide, ul, CHART_LEFT, lt)
    add_legend(slide, jd, CHART_LEFT, lt, CHART_WIDTH)

    margins = draw_margin_line(slide, jd)
    cf, amax = build_grouped_bar_chart(slide, jd)
    add_cagr(slide, jd, amax)

    src = jd.get("source","")
    if src: add_source(slide, src)

    outdir = os.path.dirname(args.output)
    if outdir: os.makedirs(outdir, exist_ok=True)
    prs.save(args.output)
    _finalize_pptx(args.output)
    print(f"\n  Output: {args.output}")

if __name__ == "__main__":
    main()
