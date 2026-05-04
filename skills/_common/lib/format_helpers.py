"""Brand-aware formatting helpers for fill_*.py scripts (Phase 4, ISSUE-011).

Centralised conversions for:
  - cell value rendering (zero / negative / thousands separator)
  - fiscal-period strings ("YY/MM期" vs free-form)
  - paragraph line-spacing (roleup spcPts vs stella default)

All helpers accept a `BrandTheme` and fall back to neutral behaviour when the
theme does not specify a key, so they are safe to call from skills that mix
schema 1.0 and schema 2.0 brands.
"""
from __future__ import annotations

from lxml import etree
from typing import Optional

# `qn` mirrors python-pptx's helper, kept inline so this module has no
# dependency on python-pptx beyond what's already imported by callers.
def _qn(tag: str) -> str:
    prefix, suffix = tag.split(":")
    nsmap = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }
    return f"{{{nsmap[prefix]}}}{suffix}"


def format_cell_value(value, theme) -> str:
    """Render a cell value following the brand's number-format rules.

    Rules:
      - None / "" → ""
      - 0 / 0.0 → theme.zero_text() (default "0")
      - Negative number → theme.negative_format()-controlled rendering
          * "paren" → "(123,456)"
          * "triangle" → "△123,456"
          * "minus" → "-123,456"  (default for stella)
      - Positive number → "123,456"
      - String → str(value) unchanged

    The function trusts the brand_resolver schema 2.0 accessors. Themes without
    these accessors (e.g. an old plugin) fall back to "minus" + "0".
    """
    if value is None or value == "":
        return ""
    if isinstance(value, bool):  # bool is a subclass of int — treat as text
        return str(value)
    if isinstance(value, (int, float)):
        if value == 0:
            return theme.zero_text()
        if value < 0:
            fmt = theme.negative_format()
            abs_str = f"{abs(value):,.0f}"
            if fmt == "paren":
                return f"({abs_str})"
            if fmt == "triangle":
                return f"△{abs_str}"
            return f"-{abs_str}"
        return f"{value:,.0f}"
    return str(value)


def format_fiscal_period(year: int, month: int, theme) -> str:
    """Render a fiscal period label per the brand's fiscal_period_format.

    Currently only "YY/MM期" is recognised (roleup); when the theme does not
    specify a format, returns "{year}/{month:02d}" without a suffix.

    Args:
        year: full 4-digit year (e.g. 2019)
        month: fiscal-year-end month, 1-12
    """
    fmt = theme.fiscal_period_format() if hasattr(theme, "fiscal_period_format") else None
    if fmt == "YY/MM期":
        return f"{year % 100:02d}/{month:02d}期"
    # Default: free-form, no suffix.
    return f"{year}/{month:02d}"


def apply_line_spacing(pPr_or_p, theme) -> None:
    """Insert/replace <a:lnSpc><a:spcPts val="<lh*100>"/></a:lnSpc> in a paragraph.

    Args:
        pPr_or_p: either an <a:pPr> element or an <a:p> element. If <a:p>, this
                  function will look up or create the <a:pPr> child.
        theme: BrandTheme; ignored if theme.line_height_pt() returns None
               (caller is left untouched, preserving stella's existing behaviour
               of relying on default line spacing).
    """
    lh = theme.line_height_pt() if hasattr(theme, "line_height_pt") else None
    if lh is None:
        return

    # Resolve the pPr element.
    tag = pPr_or_p.tag
    if tag.endswith("}p"):
        # <a:p> — find or create <a:pPr> as the first child.
        pPr = pPr_or_p.find(_qn("a:pPr"))
        if pPr is None:
            pPr = etree.SubElement(pPr_or_p, _qn("a:pPr"))
            pPr_or_p.insert(0, pPr)
    elif tag.endswith("}pPr"):
        pPr = pPr_or_p
    else:
        raise ValueError(f"apply_line_spacing expects <a:p> or <a:pPr>, got {tag!r}")

    # Replace any existing <a:lnSpc>.
    old = pPr.find(_qn("a:lnSpc"))
    if old is not None:
        pPr.remove(old)
    lnSpc = etree.SubElement(pPr, _qn("a:lnSpc"))
    etree.SubElement(lnSpc, _qn("a:spcPts"), attrib={"val": str(int(lh * 100))})
    # OOXML schema requires lnSpc to be the first pPr child.
    pPr.insert(0, lnSpc)


def resolve_top_text(data: dict, theme) -> str:
    """Return the text written to the top (largest font) placeholder.

    Reads theme.top_placeholder_field() to decide which data field to use.
    Per skills_factory convention (2026-05-04):
      - stella: 'main_message' (結論文を最上部に)
      - roleup: 'chart_title'  (スライドタイトルを最上部に)

    Falls back to 'main_message' for schema 1.0 themes (= stella既存挙動).
    """
    field = theme.top_placeholder_field() if hasattr(theme, "top_placeholder_field") else "main_message"
    return data.get(field, "")


def resolve_subtitle_text(data: dict, theme) -> str:
    """Return the text written to the subtitle placeholder.

    Inverse of resolve_top_text. Falls back to 'chart_title' for schema 1.0.
    """
    field = theme.subtitle_placeholder_field() if hasattr(theme, "subtitle_placeholder_field") else "chart_title"
    return data.get(field, "")


def require_source(data: dict, theme, skill_id: Optional[str] = None) -> None:
    """Raise ValueError if theme requires a source field but data lacks one.

    Convention: data.get('source') or data.get('source_label') or
    data.get('source_text'). If the theme does not require a source, this is
    a no-op and stella's existing warning behaviour is preserved at call site.
    """
    if not theme.is_source_required():
        return
    src = data.get("source") or data.get("source_label") or data.get("source_text")
    if not src or (isinstance(src, str) and not src.strip()):
        ctx = f" (skill={skill_id})" if skill_id else ""
        raise ValueError(
            f"source field is required by brand={theme.id}{ctx} but the input "
            f"data has no 'source' / 'source_label' / 'source_text'."
        )
