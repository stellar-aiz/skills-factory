"""
fill_valuation.py — M&A DDバリュエーション・財務サマリーをPPTXに生成

3つのchart_type:
  - football_field: 手法別バリュエーションレンジ比較
  - equity_bridge: EV→株式価値のウォーターフォール
  - financial_summary: 主要財務指標テーブル
"""

import argparse, json, sys, copy, math

# brand_resolver bootstrap (passive --brand acceptance until brand-aware migration)
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(SKILL_DIR, "..", "_common", "lib"))
from brand_resolver import add_brand_arg  # noqa: E402
from pptx import Presentation
from pptx.util import Pt, Emu, Inches
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


SHAPE_MM = "Title 1"
SHAPE_CT = "Text Placeholder 2"

# レイアウト共通
MARGIN_L = 370800
SLIDE_W  = 12192000
SLIDE_H  = 6858000
CONTENT_W = SLIDE_W - 2 * MARGIN_L
CHART_TOP = 1400000
CHART_BOTTOM = 6400000
CHART_H = CHART_BOTTOM - CHART_TOP


def find_shape(sl, nm):
    for s in sl.shapes:
        if s.name == nm: return s
    return None


def set_ph(sh, txt):
    if sh is None: return
    p = sh.text_frame.paragraphs[0]
    if p.runs:
        p.runs[0].text = txt
        for r in p.runs[1:]: r.text = ""
    else:
        r = etree.SubElement(p._p, qn("a:r"))
        etree.SubElement(r, qn("a:rPr"), attrib={"lang": "ja-JP"})
        t = etree.SubElement(r, qn("a:t")); t.text = txt


def extract_accent2(path):
    try:
        prs = Presentation(path)
        for rel in prs.slide_masters[0].part.rels.values():
            if "theme" in rel.reltype:
                te = etree.fromstring(rel.target_part.blob)
                ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
                cs = te.find(".//a:clrScheme", ns)
                if cs is not None:
                    a2 = cs.find("a:accent2", ns)
                    if a2 is not None:
                        sr = a2.find("a:srgbClr", ns)
                        if sr is not None: return sr.get("val")
    except: pass
    return "1A3C6E"


def hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def lighten(c, f):
    r, g, b = hex_to_rgb(c)
    return f"{int(r+(255-r)*f):02x}{int(g+(255-g)*f):02x}{int(b+(255-b)*f):02x}"


def add_rect(sl, l, t, w, h, fill=None, border=None, border_w=6350):
    sp = etree.SubElement(sl.shapes._spTree, qn("p:sp"))
    nv = etree.SubElement(sp, qn("p:nvSpPr"))
    etree.SubElement(nv, qn("p:cNvPr"), attrib={"id": str(id(sp)%99999+100), "name": "R"})
    etree.SubElement(nv, qn("p:cNvSpPr"))
    etree.SubElement(nv, qn("p:nvPr"))
    spr = etree.SubElement(sp, qn("p:spPr"))
    xf = etree.SubElement(spr, qn("a:xfrm"))
    etree.SubElement(xf, qn("a:off"), attrib={"x": str(int(l)), "y": str(int(t))})
    etree.SubElement(xf, qn("a:ext"), attrib={"cx": str(max(1,int(w))), "cy": str(max(1,int(h)))})
    pg = etree.SubElement(spr, qn("a:prstGeom"), attrib={"prst": "rect"})
    etree.SubElement(pg, qn("a:avLst"))
    if fill:
        sf = etree.SubElement(spr, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), attrib={"val": fill})
    else:
        etree.SubElement(spr, qn("a:noFill"))
    ln = etree.SubElement(spr, qn("a:ln"))
    if border:
        ln.set("w", str(border_w))
        sf2 = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf2, qn("a:srgbClr"), attrib={"val": border})
    else:
        etree.SubElement(ln, qn("a:noFill"))
    tb = etree.SubElement(sp, qn("p:txBody"))
    etree.SubElement(tb, qn("a:bodyPr"))
    etree.SubElement(tb, qn("a:lstStyle"))
    etree.SubElement(tb, qn("a:p"))
    return sp


def add_txt(sl, l, t, w, h, text, sz=1100, bold=False, color="1A1A1A", align="l", shrink=False, wrap="square"):
    sp = etree.SubElement(sl.shapes._spTree, qn("p:sp"))
    nv = etree.SubElement(sp, qn("p:nvSpPr"))
    etree.SubElement(nv, qn("p:cNvPr"), attrib={"id": str(id(sp)%99999+200), "name": "T"})
    etree.SubElement(nv, qn("p:cNvSpPr"), attrib={"txBox": "1"})
    etree.SubElement(nv, qn("p:nvPr"))
    spr = etree.SubElement(sp, qn("p:spPr"))
    xf = etree.SubElement(spr, qn("a:xfrm"))
    etree.SubElement(xf, qn("a:off"), attrib={"x": str(int(l)), "y": str(int(t))})
    etree.SubElement(xf, qn("a:ext"), attrib={"cx": str(max(1,int(w))), "cy": str(max(1,int(h)))})
    pg = etree.SubElement(spr, qn("a:prstGeom"), attrib={"prst": "rect"})
    etree.SubElement(pg, qn("a:avLst"))
    etree.SubElement(spr, qn("a:noFill"))
    ln = etree.SubElement(spr, qn("a:ln")); etree.SubElement(ln, qn("a:noFill"))
    tb = etree.SubElement(sp, qn("p:txBody"))
    bp = etree.SubElement(tb, qn("a:bodyPr"), attrib={
        "wrap": wrap, "lIns": "36000", "rIns": "36000",
        "tIns": "0", "bIns": "0", "anchor": "ctr"
    })
    if shrink:
        etree.SubElement(bp, qn("a:normAutofit"))
    etree.SubElement(tb, qn("a:lstStyle"))
    p = etree.SubElement(tb, qn("a:p"))
    etree.SubElement(p, qn("a:pPr"), attrib={"algn": align})
    r = etree.SubElement(p, qn("a:r"))
    rp = {"kumimoji": "1", "lang": "en-GB", "sz": str(sz)}
    if bold: rp["b"] = "1"
    rpr = etree.SubElement(r, qn("a:rPr"), attrib=rp)
    sf = etree.SubElement(rpr, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr"), attrib={"val": color})
    etree.SubElement(rpr, qn("a:latin"), attrib={"typeface": "Meiryo"})
    etree.SubElement(rpr, qn("a:ea"), attrib={"typeface": "Meiryo"})
    te = etree.SubElement(r, qn("a:t")); te.text = text


def add_line(sl, x1, y1, x2, y2, color="999999", w=9525, dash=None):
    sp = etree.SubElement(sl.shapes._spTree, qn("p:cxnSp"))
    nv = etree.SubElement(sp, qn("p:nvCxnSpPr"))
    etree.SubElement(nv, qn("p:cNvPr"), attrib={"id": str(id(sp)%99999+300), "name": "L"})
    etree.SubElement(nv, qn("p:cNvCxnSpPr"))
    etree.SubElement(nv, qn("p:nvPr"))
    spr = etree.SubElement(sp, qn("p:spPr"))
    xf = etree.SubElement(spr, qn("a:xfrm"))
    etree.SubElement(xf, qn("a:off"), attrib={"x": str(int(min(x1,x2))), "y": str(int(min(y1,y2)))})
    etree.SubElement(xf, qn("a:ext"), attrib={
        "cx": str(max(1, abs(int(x2-x1)))), "cy": str(max(1, abs(int(y2-y1))))
    })
    pg = etree.SubElement(spr, qn("a:prstGeom"), attrib={"prst": "line"})
    etree.SubElement(pg, qn("a:avLst"))
    ln = etree.SubElement(spr, qn("a:ln"), attrib={"w": str(w)})
    sf = etree.SubElement(ln, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr"), attrib={"val": color})
    if dash:
        etree.SubElement(ln, qn("a:prstDash"), attrib={"val": dash})


def fmt_num(v, unit=""):
    """数値を見やすくフォーマット"""
    if abs(v) >= 1e9:
        return f"{v/1e9:,.1f}B{unit}"
    if abs(v) >= 1e6:
        return f"{v/1e6:,.0f}M{unit}"
    if abs(v) >= 1e4:
        return f"{v:,.0f}{unit}"
    return f"{v:,.1f}{unit}"


# ══════════════════════════════════════════════════════════
# 1. FOOTBALL FIELD CHART
# ══════════════════════════════════════════════════════════
def build_football_field(sl, data, accent2):
    ff = data["football_field"]
    methods = ff["methods"]
    unit = ff.get("unit", "")
    n = len(methods)

    # 軸範囲
    all_vals = []
    for m in methods:
        all_vals.extend([m["low"], m["high"]])
    rec = ff.get("recommended_range", None)
    ax_min = ff.get("axis_min", min(all_vals) * 0.85)
    ax_max = ff.get("axis_max", max(all_vals) * 1.15)
    ax_range = ax_max - ax_min

    # レイアウト
    label_w = 3400000
    chart_l = MARGIN_L + label_w
    chart_w = CONTENT_W - label_w
    row_h = min(550000, CHART_H // (n + 2))
    bar_h = int(row_h * 0.55)
    chart_area_top = CHART_TOP + row_h  # 上にスペース

    def val_to_x(v):
        return chart_l + (v - ax_min) / ax_range * chart_w

    # 推定レンジのハイライト帯（背景に描画）
    if rec:
        rx1 = val_to_x(rec["low"])
        rx2 = val_to_x(rec["high"])
        add_rect(sl, rx1, chart_area_top - row_h // 2,
                 rx2 - rx1, row_h * n + row_h,
                 fill=lighten(accent2, 0.85))
        # レンジラベル
        add_txt(sl, rx1, CHART_TOP, rx2 - rx1, row_h * 0.7,
                f"推定レンジ: {fmt_num(rec['low'])}{unit} - {fmt_num(rec['high'])}{unit}",
                sz=1000, bold=True, color=accent2, align="ctr")

    # X軸目盛り
    n_ticks = 6
    for i in range(n_ticks + 1):
        v = ax_min + ax_range * i / n_ticks
        x = val_to_x(v)
        # 目盛り線
        add_line(sl, x, chart_area_top - row_h // 3,
                 x, chart_area_top + row_h * n, color="D0D0D0", w=6350, dash="dash")
        # 目盛りラベル
        add_txt(sl, x - 500000, chart_area_top + row_h * n + 30000,
                1000000, 280000, fmt_num(v) + unit, sz=900, color="666666", align="ctr")

    # 各手法のバー
    for i, m in enumerate(methods):
        y = chart_area_top + i * row_h
        bar_y = y + (row_h - bar_h) // 2

        # ラベル（折り返さない — 右寄せなので左側にはみ出すだけ）
        add_txt(sl, MARGIN_L, y, label_w - 100000, row_h,
                m["name"], sz=1200, bold=True, color="1A1A1A", align="r",
                shrink=True, wrap="none")

        # レンジバー
        x_low = val_to_x(m["low"])
        x_high = val_to_x(m["high"])
        bar_color = accent2 if i % 2 == 0 else lighten(accent2, 0.3)
        add_rect(sl, x_low, bar_y, x_high - x_low, bar_h, fill=bar_color)

        # ミッドポイント（あれば）
        if "mid" in m:
            x_mid = val_to_x(m["mid"])
            add_line(sl, x_mid, bar_y - 30000, x_mid, bar_y + bar_h + 30000,
                     color="FFFFFF", w=19050)

        # 値ラベル（バーの両端 — バー内側に配置）
        val_lbl_w = 800000
        bar_actual_w = x_high - x_low
        if bar_actual_w > val_lbl_w * 2:
            # バーが十分に広い → 両端をバー内側に配置
            add_txt(sl, x_low + 30000, bar_y, val_lbl_w, bar_h,
                    fmt_num(m["low"]), sz=900, color="FFFFFF", align="l")
            add_txt(sl, x_high - val_lbl_w - 30000, bar_y, val_lbl_w, bar_h,
                    fmt_num(m["high"]), sz=900, color="FFFFFF", align="r")
        else:
            # バーが狭い → 外側に配置（黒文字）
            add_txt(sl, x_low - val_lbl_w - 30000, bar_y, val_lbl_w, bar_h,
                    fmt_num(m["low"]), sz=900, color="1A1A1A", align="r")
            add_txt(sl, x_high + 30000, bar_y, val_lbl_w, bar_h,
                    fmt_num(m["high"]), sz=900, color="1A1A1A", align="l")

    # 単位ラベル
    add_txt(sl, chart_l + chart_w - 1500000,
            chart_area_top + row_h * n + 300000,
            1500000, 250000, f"（単位: {unit}）" if unit else "",
            sz=900, color="888888", align="r")

    print(f"  [Football Field] {n} methods, range {fmt_num(ax_min)}-{fmt_num(ax_max)}")


# ══════════════════════════════════════════════════════════
# 2. EQUITY VALUE BRIDGE (Waterfall)
# ══════════════════════════════════════════════════════════
def build_equity_bridge(sl, data, accent2):
    eb = data["equity_bridge"]
    items = eb["items"]
    unit = eb.get("unit", "")
    n = len(items)

    # 値の範囲を計算
    cumulative = []
    cum = 0
    for it in items:
        if it["type"] == "start":
            cum = it["value"]
            cumulative.append({"base": 0, "top": cum, "val": cum, **it})
        elif it["type"] == "total":
            cumulative.append({"base": 0, "top": it["value"], "val": it["value"], **it})
        else:
            old_cum = cum
            cum += it["value"]
            if it["value"] >= 0:
                cumulative.append({"base": old_cum, "top": cum, "val": it["value"], **it})
            else:
                cumulative.append({"base": cum, "top": old_cum, "val": it["value"], **it})

    max_val = max(c["top"] for c in cumulative) * 1.15
    min_val = min(min(c["base"] for c in cumulative), 0)

    val_range = max_val - min_val
    if val_range == 0: val_range = 1

    # レイアウト
    bar_area_top = CHART_TOP + 400000
    bar_area_h = CHART_H - 800000
    bar_gap = 120000
    total_bar_w = CONTENT_W - 200000
    bar_w = (total_bar_w - bar_gap * (n - 1)) // n
    bar_w = min(bar_w, 1600000)
    total_used = bar_w * n + bar_gap * (n - 1)
    start_x = MARGIN_L + (CONTENT_W - total_used) // 2

    def val_to_y(v):
        return bar_area_top + bar_area_h - (v - min_val) / val_range * bar_area_h

    # ゼロライン
    y_zero = val_to_y(0)
    add_line(sl, start_x - 100000, y_zero,
             start_x + total_used + 100000, y_zero, color="AAAAAA", w=9525)

    # 各バー
    for i, c in enumerate(cumulative):
        x = start_x + i * (bar_w + bar_gap)
        y_top = val_to_y(c["top"])
        y_base = val_to_y(c["base"])
        h = abs(y_base - y_top)
        y = min(y_top, y_base)

        # バーの色
        if c["type"] in ("start", "total"):
            fill = accent2
        elif c["val"] >= 0:
            fill = "2E7D32"  # green
        else:
            fill = "C62828"  # red

        add_rect(sl, x, y, bar_w, max(h, Emu(Pt(2))), fill=fill)

        # 値ラベル（バーの上または下）
        val_text = fmt_num(c["val"]) + unit
        if c["val"] < 0:
            val_text = fmt_num(c["val"]) + unit
        if c["type"] in ("start", "total"):
            val_text = fmt_num(c["top"]) + unit

        label_y = y - 280000 if c["val"] >= 0 or c["type"] in ("start", "total") else y + h + 30000
        add_txt(sl, x, label_y, bar_w, 260000,
                val_text, sz=1050, bold=True, color="1A1A1A", align="ctr")

        # 項目名ラベル（下部）
        add_txt(sl, x - 100000, bar_area_top + bar_area_h + 80000,
                bar_w + 200000, 350000,
                c["name"], sz=1050, bold=False, color="1A1A1A", align="ctr")

        # コネクター線（start/adjustからtotalの前まで）
        if i < n - 1 and c["type"] != "total":
            next_x = start_x + (i + 1) * (bar_w + bar_gap)
            conn_y = val_to_y(cumulative[i]["top"] if c["val"] >= 0 or c["type"] == "start" else cumulative[i]["base"])
            # total項目の場合はbaseから
            if cumulative[i+1]["type"] == "total":
                pass  # totalへのコネクターは省略
            else:
                add_line(sl, x + bar_w, conn_y, next_x, conn_y,
                         color="999999", w=6350, dash="dash")

    print(f"  [Equity Bridge] {n} items")


# ══════════════════════════════════════════════════════════
# 3. FINANCIAL SUMMARY TABLE
# ══════════════════════════════════════════════════════════
def build_financial_summary(sl, data, accent2):
    fs = data["financial_summary"]
    periods = fs["periods"]
    metrics = fs["metrics"]
    n_cols = len(periods) + 1  # +1 for metric name column
    n_rows = len(metrics) + 1  # +1 for header row

    # オプション: CAGR列
    has_cagr = any("cagr" in m for m in metrics)
    if has_cagr:
        n_cols += 1

    # テーブルレイアウト
    tbl_l = MARGIN_L
    tbl_t = CHART_TOP
    tbl_w = CONTENT_W
    row_h = min(420000, (CHART_H - 100000) // n_rows)
    name_col_w = int(tbl_w * 0.22)
    data_col_w = (tbl_w - name_col_w) // (n_cols - 1)

    # ヘッダー行
    add_rect(sl, tbl_l, tbl_t, tbl_w, row_h, fill=accent2)
    add_txt(sl, tbl_l, tbl_t, name_col_w, row_h,
            "項目", sz=1100, bold=True, color="FFFFFF", align="ctr")
    for j, per in enumerate(periods):
        x = tbl_l + name_col_w + j * data_col_w
        add_txt(sl, x, tbl_t, data_col_w, row_h,
                per, sz=1050, bold=True, color="FFFFFF", align="ctr")
    if has_cagr:
        x = tbl_l + name_col_w + len(periods) * data_col_w
        add_txt(sl, x, tbl_t, data_col_w, row_h,
                "CAGR", sz=1050, bold=True, color="FFFFFF", align="ctr")

    # ヘッダー下の線
    add_line(sl, tbl_l, tbl_t + row_h, tbl_l + tbl_w, tbl_t + row_h,
             color="333333", w=15875)

    # データ行
    for i, m in enumerate(metrics):
        y = tbl_t + (i + 1) * row_h
        is_section = m.get("is_section", False)
        is_total = m.get("is_total", False)

        # 交互背景
        if is_section:
            add_rect(sl, tbl_l, y, tbl_w, row_h, fill=lighten(accent2, 0.88))
        elif i % 2 == 1:
            add_rect(sl, tbl_l, y, tbl_w, row_h, fill="F5F5F5")

        # 項目名
        name_color = accent2 if is_section else "1A1A1A"
        add_txt(sl, tbl_l, y, name_col_w, row_h,
                m["name"], sz=1100,
                bold=is_section or is_total,
                color=name_color, align="l")

        # 値
        values = m.get("values", [])
        for j, v in enumerate(values):
            x = tbl_l + name_col_w + j * data_col_w
            if v is None or v == "":
                continue
            txt = str(v) if isinstance(v, str) else fmt_num(v, m.get("unit_suffix", ""))
            add_txt(sl, x, y, data_col_w, row_h,
                    txt, sz=1050, bold=is_total,
                    color="1A1A1A", align="r")

        # CAGR
        if has_cagr and "cagr" in m:
            x = tbl_l + name_col_w + len(periods) * data_col_w
            add_txt(sl, x, y, data_col_w, row_h,
                    m["cagr"], sz=1050, bold=False, color="1A1A1A", align="r")

        # 行区切り線
        add_line(sl, tbl_l, y + row_h, tbl_l + tbl_w, y + row_h,
                 color="DDDDDD", w=6350)
        if is_total:
            add_line(sl, tbl_l, y, tbl_l + tbl_w, y, color="333333", w=12700)

    print(f"  [Financial Summary] {n_rows-1} metrics x {len(periods)} periods")


# ══════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", required=True)
    parser.add_argument("--template", required=True)
    parser.add_argument("--output", required=True)
    add_brand_arg(parser)  # passive: accepted but ignored until brand migration
    args = parser.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    prs = Presentation(args.template)
    sl = prs.slides[0]
    a2 = extract_accent2(args.template)
    print(f"  Accent2: {a2}")

    mm = data.get("main_message", ""); ct = data.get("chart_title", "")
    set_ph(find_shape(sl, SHAPE_MM), mm)
    set_ph(find_shape(sl, SHAPE_CT), ct)
    print(f"  [Main Message] {mm[:60]}")
    print(f"  [Chart Title]  {ct}")

    chart_type = data.get("chart_type", "football_field")

    if chart_type == "football_field":
        build_football_field(sl, data, a2)
    elif chart_type == "equity_bridge":
        build_equity_bridge(sl, data, a2)
    elif chart_type == "financial_summary":
        build_financial_summary(sl, data, a2)
    else:
        print(f"  ERROR: Unknown chart_type '{chart_type}'", file=sys.stderr)
        sys.exit(1)

    prs.save(args.output)

    _finalize_pptx(args.output)

    print(f"\n  Saved: {args.output}")


if __name__ == "__main__":
    main()
