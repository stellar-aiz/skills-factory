"""
fill_sga_breakdown.py - SGA Breakdown Slide (Native PPTX)
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


SHAPE_MAIN_MSG = "Title 1"
SHAPE_CHART_TITLE = "Text Placeholder 2"
SHAPE_CONTENT = "Content Area"
SHAPE_SOURCE = "Source"

PANEL_Y = Inches(1.50)
CHART_X = Inches(0.50); CHART_W = Inches(8.50); CHART_H = Inches(5.00)
CHART_Y = PANEL_Y + Inches(0.55)
LEGEND_X = Inches(9.30); LEGEND_W = Inches(3.80)

COLOR_TEXT = RGBColor(0x33,0x33,0x33)
COLOR_SRC = RGBColor(0x66,0x66,0x66)
FONT = "Meiryo UI"

def find_shape(slide, name):
    for s in slide.shapes:
        if s.name == name: return s
    return None

def set_text(shape, text):
    if not shape: return
    p = shape.text_frame.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]: r.text = ""
    else:
        r = etree.SubElement(p._p, qn("a:r"))
        etree.SubElement(r, qn("a:rPr"), attrib={"lang":"ja-JP"})
        t = etree.SubElement(r, qn("a:t")); t.text = text

def rm_shape(slide, name):
    s = find_shape(slide, name)
    if s: slide.shapes._spTree.remove(s._element)

def hex2rgb(h):
    h = h.replace("#","")
    return RGBColor(int(h[:2],16), int(h[2:4],16), int(h[4:6],16))

def build_chart(slide, data, left, top, w, h):
    from pptx.chart.data import CategoryChartData
    cats = data.get("categories", [])
    periods = data.get("periods", [])
    line_cfg = data.get("line")
    threshold = data.get("label_threshold_pct", 5.0)
    n_cats = len(cats); n_per = len(periods)

    cd = CategoryChartData()
    cd.categories = [p["label"] for p in periods]
    for ci, cat in enumerate(cats):
        vals = [p["values"][ci] if ci < len(p.get("values",[])) else 0 for p in periods]
        cd.add_series(cat["name"], vals)
    has_line = line_cfg is not None
    if has_line:
        cd.add_series(line_cfg["series_name"], [p.get("line_value",0) for p in periods])

    cf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, left, top, w, h, cd)
    chart = cf.chart
    pa = chart._chartSpace.chart.plotArea
    bc = pa.findall(qn('c:barChart'))[0]

    # Split line to lineChart
    if has_line:
        sers = bc.findall(qn('c:ser'))
        ls = copy.deepcopy(sers[-1]); bc.remove(sers[-1])
        lc = etree.SubElement(pa, qn('c:lineChart'))
        etree.SubElement(lc, qn('c:grouping'), attrib={'val':'standard'})
        etree.SubElement(lc, qn('c:varyColors'), attrib={'val':'0'})
        lc.append(ls)
        # Marker
        mk = ls.find(qn('c:marker'))
        if mk is None: mk = etree.SubElement(ls, qn('c:marker'))
        sym = mk.find(qn('c:symbol'))
        if sym is None: sym = etree.SubElement(mk, qn('c:symbol'))
        sym.set('val','circle')
        sz = mk.find(qn('c:size'))
        if sz is None: sz = etree.SubElement(mk, qn('c:size'))
        sz.set('val','8')
        # Secondary axis
        sec_v="2094734553"; sec_c="2094734554"
        etree.SubElement(lc, qn('c:axId'), attrib={'val':sec_c})
        etree.SubElement(lc, qn('c:axId'), attrib={'val':sec_v})
        sc = etree.SubElement(pa, qn('c:catAx'))
        etree.SubElement(sc, qn('c:axId'), attrib={'val':sec_c})
        s2 = etree.SubElement(sc, qn('c:scaling'))
        etree.SubElement(s2, qn('c:orientation'), attrib={'val':'minMax'})
        etree.SubElement(sc, qn('c:delete'), attrib={'val':'1'})
        etree.SubElement(sc, qn('c:axPos'), attrib={'val':'b'})
        etree.SubElement(sc, qn('c:crossAx'), attrib={'val':sec_v})
        sv = etree.SubElement(pa, qn('c:valAx'))
        etree.SubElement(sv, qn('c:axId'), attrib={'val':sec_v})
        s3 = etree.SubElement(sv, qn('c:scaling'))
        etree.SubElement(s3, qn('c:orientation'), attrib={'val':'minMax'})
        etree.SubElement(sv, qn('c:delete'), attrib={'val':'1'})
        etree.SubElement(sv, qn('c:axPos'), attrib={'val':'r'})
        etree.SubElement(sv, qn('c:numFmt'), attrib={'formatCode':'0.0"%"','sourceLinked':'0'})
        etree.SubElement(sv, qn('c:crossAx'), attrib={'val':sec_c})
        etree.SubElement(sv, qn('c:crosses'), attrib={'val':'max'})
        # Line color
        lclr = line_cfg.get("color","#666666").replace("#","")
        lsp = ls.find(qn('c:spPr'))
        if lsp is None: lsp = etree.SubElement(ls, qn('c:spPr'))
        ln = lsp.find(qn('a:ln'))
        if ln is None: ln = etree.SubElement(lsp, qn('a:ln'))
        ln.set('w','19050')
        lsf = ln.find(qn('a:solidFill'))
        if lsf is None: lsf = etree.SubElement(ln, qn('a:solidFill'))
        for c in list(lsf): lsf.remove(c)
        etree.SubElement(lsf, qn('a:srgbClr'), attrib={'val':lclr})
        msp = mk.find(qn('c:spPr'))
        if msp is None: msp = etree.SubElement(mk, qn('c:spPr'))
        msf = etree.SubElement(msp, qn('a:solidFill'))
        etree.SubElement(msf, qn('a:srgbClr'), attrib={'val':lclr})
        # Line data labels
        _add_line_dlbls(ls)

    # Bar colors + custom labels
    bar_sers = bc.findall(qn('c:ser'))
    for si, bs in enumerate(bar_sers):
        clr = cats[si]["color"].replace("#","") if si < n_cats else "999999"
        sp = bs.find(qn('c:spPr'))
        if sp is None: sp = etree.SubElement(bs, qn('c:spPr'))
        sf = sp.find(qn('a:solidFill'))
        if sf is None: sf = etree.SubElement(sp, qn('a:solidFill'))
        for c in list(sf): sf.remove(c)
        etree.SubElement(sf, qn('a:srgbClr'), attrib={'val':clr})
        _add_bar_dlbls(bs, si, periods, threshold)

    # Gap/Overlap
    gw = bc.find(qn('c:gapWidth'))
    if gw is None: gw = etree.SubElement(bc, qn('c:gapWidth'))
    gw.set('val','80')
    ov = bc.find(qn('c:overlap'))
    if ov is None: ov = etree.SubElement(bc, qn('c:overlap'))
    ov.set('val','100')

    # Remove gridlines
    for vx in pa.findall(qn('c:valAx')):
        for mg in vx.findall(qn('c:majorGridlines')): vx.remove(mg)
        for mg in vx.findall(qn('c:minorGridlines')): vx.remove(mg)
    for cx in pa.findall(qn('c:catAx')):
        for mg in cx.findall(qn('c:majorGridlines')): cx.remove(mg)

    # Axis font
    for ax in pa.findall(qn('c:catAx')) + pa.findall(qn('c:valAx')):
        de = ax.find(qn('c:delete'))
        if de is not None and de.get('val')=='1': continue
        tp = ax.find(qn('c:txPr'))
        if tp is None: tp = etree.SubElement(ax, qn('c:txPr'))
        bp = tp.find(qn('a:bodyPr'))
        if bp is None: bp = etree.SubElement(tp, qn('a:bodyPr')); tp.insert(0, bp)
        if tp.find(qn('a:lstStyle')) is None: etree.SubElement(tp, qn('a:lstStyle'))
        p = tp.find(qn('a:p'))
        if p is None: p = etree.SubElement(tp, qn('a:p'))
        pp = p.find(qn('a:pPr'))
        if pp is None: pp = etree.SubElement(p, qn('a:pPr'))
        dr = pp.find(qn('a:defRPr'))
        if dr is None: dr = etree.SubElement(pp, qn('a:defRPr'), attrib={'sz':'1000'})
        else: dr.set('sz','1000')
        if dr.find(qn('a:latin')) is None: etree.SubElement(dr, qn('a:latin'), attrib={'typeface':FONT})
        if dr.find(qn('a:ea')) is None: etree.SubElement(dr, qn('a:ea'), attrib={'typeface':FONT})

    chart.has_legend = False; chart.has_title = False
    psp = pa.find(qn('c:spPr'))
    if psp is None: psp = etree.SubElement(pa, qn('c:spPr'))
    if psp.find(qn('a:noFill')) is None: etree.SubElement(psp, qn('a:noFill'))
    print(f"  Chart: {n_per} periods, {n_cats} series" + (" + line" if has_line else ""))
    return cf

def _add_bar_dlbls(ser, si, periods, threshold):
    old = ser.find(qn('c:dLbls'))
    if old is not None: ser.remove(old)
    dlbls = etree.SubElement(ser, qn('c:dLbls'))
    for pi, per in enumerate(periods):
        vals = per.get("values",[])
        val = vals[si] if si < len(vals) else 0
        total = per.get("total",0)
        pct = (val/total*100) if total > 0 else 0
        dlbl = etree.SubElement(dlbls, qn('c:dLbl'))
        etree.SubElement(dlbl, qn('c:idx'), attrib={'val':str(pi)})
        if val == 0 or pct < threshold:
            etree.SubElement(dlbl, qn('c:delete'), attrib={'val':'1'})
        else:
            txt = f"{val}\n({pct:.1f}%)"
            tx = etree.SubElement(dlbl, qn('c:tx'))
            rich = etree.SubElement(tx, qn('c:rich'))
            etree.SubElement(rich, qn('a:bodyPr'))
            etree.SubElement(rich, qn('a:lstStyle'))
            p = etree.SubElement(rich, qn('a:p'))
            pp = etree.SubElement(p, qn('a:pPr'), attrib={'algn':'ctr'})
            dr = etree.SubElement(pp, qn('a:defRPr'), attrib={'sz':'800','b':'1'})
            sf = etree.SubElement(dr, qn('a:solidFill'))
            etree.SubElement(sf, qn('a:srgbClr'), attrib={'val':'FFFFFF'})
            etree.SubElement(dr, qn('a:latin'), attrib={'typeface':FONT})
            etree.SubElement(dr, qn('a:ea'), attrib={'typeface':FONT})
            r = etree.SubElement(p, qn('a:r'))
            rPr = etree.SubElement(r, qn('a:rPr'), attrib={'lang':'ja-JP','sz':'800','b':'1'})
            sf2 = etree.SubElement(rPr, qn('a:solidFill'))
            etree.SubElement(sf2, qn('a:srgbClr'), attrib={'val':'FFFFFF'})
            etree.SubElement(rPr, qn('a:latin'), attrib={'typeface':FONT})
            etree.SubElement(rPr, qn('a:ea'), attrib={'typeface':FONT})
            t = etree.SubElement(r, qn('a:t')); t.text = txt
            etree.SubElement(dlbl, qn('c:dLblPos'), attrib={'val':'ctr'})
            for k in ['showLegendKey','showVal','showCatName','showSerName','showPercent']:
                etree.SubElement(dlbl, qn(f'c:{k}'), attrib={'val':'0'})
    for k in ['showLegendKey','showVal','showCatName','showSerName','showPercent']:
        etree.SubElement(dlbls, qn(f'c:{k}'), attrib={'val':'0'})

def _add_line_dlbls(ser):
    dl = ser.find(qn('c:dLbls'))
    if dl is None: dl = etree.SubElement(ser, qn('c:dLbls'))
    for c in list(dl): dl.remove(c)
    etree.SubElement(dl, qn('c:numFmt'), attrib={'formatCode':'0.0"%"','sourceLinked':'0'})
    for k in ['showLegendKey','showCatName','showSerName','showPercent','showBubbleSize']:
        etree.SubElement(dl, qn(f'c:{k}'), attrib={'val':'0'})
    etree.SubElement(dl, qn('c:showVal'), attrib={'val':'1'})
    etree.SubElement(dl, qn('c:dLblPos'), attrib={'val':'t'})
    tp = etree.SubElement(dl, qn('c:txPr'))
    etree.SubElement(tp, qn('a:bodyPr')); etree.SubElement(tp, qn('a:lstStyle'))
    p = etree.SubElement(tp, qn('a:p'))
    pp = etree.SubElement(p, qn('a:pPr'))
    dr = etree.SubElement(pp, qn('a:defRPr'), attrib={'sz':'1000','b':'1'})
    etree.SubElement(dr, qn('a:latin'), attrib={'typeface':FONT})
    etree.SubElement(dr, qn('a:ea'), attrib={'typeface':FONT})
    sf = etree.SubElement(dr, qn('a:solidFill'))
    etree.SubElement(sf, qn('a:srgbClr'), attrib={'val':'333333'})

def add_total_labels(slide, data, cl, ct, cw, ch):
    periods = data.get("periods",[])
    nc = len(periods)
    if nc == 0: return
    plm=cw*0.06; prm=cw*0.04; ptm=ch*0.05; pbm=ch*0.14
    pl=cl+plm; pw=cw-plm-prm; pb=ct+ch-pbm; ph=pb-(ct+ptm); caw=pw/nc
    max_t = max(p.get("total",0) for p in periods)
    ax_max = max_t * 1.20
    for pi, per in enumerate(periods):
        t = per.get("total",0)
        cx = pl + (pi+0.5)*caw
        ty = pb - (t/ax_max)*ph
        lw=Inches(0.80); lh=Inches(0.25)
        tb = slide.shapes.add_textbox(int(cx-lw/2), int(ty-lh-Emu(20000)), lw, lh)
        tb.text_frame.word_wrap = False
        p = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = str(t)
        r.font.size=Pt(10); r.font.bold=True; r.font.color.rgb=COLOR_TEXT; r.font.name=FONT
    print(f"  Total labels: {nc}")

def add_trend_arrow(slide, data, cl, ct, cw, ch):
    if not data.get("trend_arrow",False): return
    periods = data.get("periods",[]); line_cfg = data.get("line")
    if not line_cfg or len(periods)<2: return
    nc = len(periods)
    plm=cw*0.06; prm=cw*0.04; ptm=ch*0.05; pbm=ch*0.14
    pl=cl+plm; pw=cw-plm-prm; pb=ct+ch-pbm; ph=pb-(ct+ptm); caw=pw/nc
    max_t = max(p.get("total",0) for p in periods)
    ax_max = max_t * 1.20
    gap = max(ph*0.18, Inches(0.50))
    for i in range(nc-1):
        ti = periods[i].get("total",0); tj = periods[i+1].get("total",0)
        lv = periods[i+1].get("line_value",0)
        cx_i = pl+(i+0.5)*caw; cx_j = pl+(i+1.5)*caw
        bti = pb-(ti/ax_max)*ph; btj = pb-(tj/ax_max)*ph
        ay_i = int(bti-gap); ay_j = int(btj-gap)
        mn = int(ct-Inches(0.20)); ay_i=max(ay_i,mn); ay_j=max(ay_j,mn)
        conn = slide.shapes.add_connector(1, int(cx_i), ay_i, int(cx_j), ay_j)
        conn.line.color.rgb = RGBColor(0x99,0x99,0x99); conn.line.width = Pt(2.0)
        sp = conn._element.find(qn('p:spPr'))
        if sp is None: sp = conn._element.find(qn('a:spPr'))
        if sp is not None:
            ln = sp.find(qn('a:ln'))
            if ln is None: ln = etree.SubElement(sp, qn('a:ln'))
            etree.SubElement(ln, qn('a:tailEnd'), attrib={'type':'triangle','w':'med','len':'med'})
        mx=(cx_i+cx_j)/2; my=(ay_i+ay_j)/2
        ow=Inches(0.90); oh=Inches(0.28)
        tb = slide.shapes.add_textbox(int(mx-ow/2), int(my-oh/2), ow, oh)
        tb.text_frame.word_wrap = False
        p = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = f"{lv:.1f}%"
        r.font.size=Pt(11); r.font.bold=True; r.font.color.rgb=COLOR_TEXT; r.font.name=FONT
    print(f"  Trend arrows: {nc-1}")

def add_legend(slide, data, left, top, w):
    cats = data.get("categories",[]); line_cfg = data.get("line")
    periods = data.get("periods",[])

    # 使える縦幅: top ~ 出典の上(6.30in)
    LEGEND_BOTTOM = Inches(6.30)
    avail_h = LEGEND_BOTTOM - top

    # アクティブ科目（値>0のもの）
    active = []
    for ci, cat in enumerate(cats):
        if any((p.get("values",[])[ci] if ci<len(p.get("values",[])) else 0)>0 for p in periods):
            active.append(cat)
    n_active = len(active)

    # 折れ線凡例の高さ確保
    line_header_h = Inches(0.45) if line_cfg else 0

    # 2列配置の行数
    n_rows = math.ceil(n_active / 2)
    bar_area_h = avail_h - line_header_h
    # 行高さを動的計算（最大0.50in、最小0.28in）
    if n_rows > 0:
        lh = min(Inches(0.50), max(Inches(0.28), int(bar_area_h / n_rows)))
    else:
        lh = Inches(0.40)

    # フォントサイズ・マーカーサイズを行高さに応じて決定
    if lh >= Inches(0.40):
        font_sz = Pt(12); sq = Inches(0.18); line_font_sz = Pt(12)
    elif lh >= Inches(0.32):
        font_sz = Pt(11); sq = Inches(0.16); line_font_sz = Pt(11)
    else:
        font_sz = Pt(10); sq = Inches(0.14); line_font_sz = Pt(10)

    # コンテンツ全体の高さを計算し、縦中央に配置
    content_h = line_header_h + n_rows * lh
    y_offset = (avail_h - content_h) / 2
    y = top + y_offset

    cw = w / 2

    # 折れ線凡例
    if line_cfg:
        lclr = hex2rgb(line_cfg.get("color","#666666"))
        lw=Inches(0.40); ly=int(y+line_header_h/2)
        cn = slide.shapes.add_connector(1, int(left), ly, int(left+lw), ly)
        cn.line.color.rgb=lclr; cn.line.width=Pt(2.0)
        mk_sz = Inches(0.12)
        mk = slide.shapes.add_shape(MSO_SHAPE.OVAL,
            int(left+(lw-mk_sz)/2), int(ly-mk_sz/2), mk_sz, mk_sz)
        mk.fill.solid(); mk.fill.fore_color.rgb=lclr; mk.line.fill.background()
        tx=int(left+lw+Inches(0.08))
        tb = slide.shapes.add_textbox(tx, int(y), int(w-lw-Inches(0.08)), int(line_header_h))
        tb.text_frame.word_wrap = False
        tb.text_frame.paragraphs[0].space_before = Pt(0)
        r = tb.text_frame.paragraphs[0].add_run()
        r.text=line_cfg["series_name"]; r.font.size=line_font_sz
        r.font.color.rgb=COLOR_TEXT; r.font.name=FONT
        y += line_header_h

    # 棒グラフ凡例（2列）
    for i, cat in enumerate(active):
        col=i%2; row=i//2
        ix = left + col*cw; iy = y + row*lh
        sqy = iy + (lh-sq)/2
        s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(ix), int(sqy), sq, sq)
        s.fill.solid(); s.fill.fore_color.rgb = hex2rgb(cat["color"]); s.line.fill.background()
        tx = int(ix+sq+Inches(0.06))
        tb = slide.shapes.add_textbox(tx, int(iy), int(cw-sq-Inches(0.06)), int(lh))
        tb.text_frame.word_wrap = False
        r = tb.text_frame.paragraphs[0].add_run()
        r.text=cat["name"]; r.font.size=font_sz; r.font.color.rgb=COLOR_TEXT; r.font.name=FONT
    print(f"  Legend: {n_active} items (2-col, {font_sz.pt}pt, centered)")

def add_unit(slide, text, left, top):
    tb = slide.shapes.add_textbox(left, top, Inches(3.0), Inches(0.25))
    p = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text=text; r.font.size=Pt(10); r.font.color.rgb=COLOR_SRC; r.font.name=FONT

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--output", required=True)
    add_brand_arg(ap)  # passive: accepted but ignored until brand migration
    args = ap.parse_args()
    with open(args.data,"r",encoding="utf-8") as f: data = json.load(f)
    print("=== SGA Breakdown Slide ===")
    prs = Presentation(args.template); slide = prs.slides[0]
    set_text(find_shape(slide, SHAPE_MAIN_MSG), data.get("main_message",""))
    set_text(find_shape(slide, SHAPE_CHART_TITLE), data.get("chart_title",""))
    src = data.get("source","")
    if src: set_text(find_shape(slide, SHAPE_SOURCE), f"出典：{src}")
    rm_shape(slide, SHAPE_CONTENT)
    ul = data.get("unit_label","")
    if ul: add_unit(slide, ul, CHART_X, PANEL_Y)
    build_chart(slide, data, CHART_X, CHART_Y, CHART_W, CHART_H)
    add_total_labels(slide, data, CHART_X, CHART_Y, CHART_W, CHART_H)
    add_trend_arrow(slide, data, CHART_X, CHART_Y, CHART_W, CHART_H)
    add_legend(slide, data, LEGEND_X, PANEL_Y+Inches(0.10), LEGEND_W)
    os.makedirs(os.path.dirname(args.output) or ".", exist_ok=True)
    prs.save(args.output); print(f"  Done: {args.output}")
    _finalize_pptx(args.output)
if __name__ == "__main__": main()

