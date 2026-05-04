"""Brand resolver for skills_factory PPTX skills.

Loads brand theme JSON (_common/brands/<id>/theme.json) and optional
per-skill layout overrides (<skill>/assets/<brand>/layout.json), then
exposes a BrandTheme object so fill_*.py can read fonts/colors/sizes/
coordinates without hardcoding per-brand values.

Usage in fill_*.py:

    import sys, os
    SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    sys.path.insert(0, os.path.join(SKILL_DIR, "..", "_common", "lib"))
    from brand_resolver import resolve_brand, add_brand_arg

    parser = argparse.ArgumentParser(...)
    add_brand_arg(parser)
    args = parser.parse_args()
    theme = resolve_brand(args.brand, SKILL_DIR)
    color = theme.color("text")          # RGBColor
    font  = theme.font_ea                  # str
    pt    = theme.pt("font_size_label_pt") # Pt
    panel = theme.layout("panel_y_in")     # EMU (Inches converted)
    tpl   = theme.template_path(SKILL_DIR, "customer-profile")

Schema 2.0 accessors (Phase 1, ISSUE-011):

    body_sz = theme.font_size_body_pt(skill_id="executive-summary-pptx")
    nf      = theme.number_format_excel()    # "_ * #,##0_ ;_ * (#,##0)_ ;_ \"-\"_ ;_ @_ "
    lh      = theme.line_height_pt()         # 12 for rollup, None for stella
    neg     = theme.negative_format()        # "paren" / "triangle" / "minus"
    z       = theme.zero_text()              # "-" for rollup, "0" for stella
    fp      = theme.fiscal_period_format()   # "YY/MM期" for rollup, None for stella
    guide_x = theme.layout_rule("left_align_guide_x_in")  # 0.41 for rollup, None for stella
    if theme.is_source_required() and not data.get("source"):
        raise ValueError("source required")
"""
from __future__ import annotations

import argparse
import json
import os
from dataclasses import dataclass, field
from typing import Optional

from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

VALID_BRANDS = ("stellar_aiz", "rollup")
DEFAULT_BRAND = "stellar_aiz"

# Resolve _common/brands/<id>/theme.json relative to this file.
_LIB_DIR = os.path.dirname(os.path.abspath(__file__))
_COMMON_DIR = os.path.dirname(_LIB_DIR)
_BRANDS_DIR = os.path.join(_COMMON_DIR, "brands")


def _hex_to_rgbcolor(hex_str: str) -> RGBColor:
    s = hex_str.lstrip("#")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


@dataclass(frozen=True)
class BrandTheme:
    id: str
    label: str
    slide_w: int   # EMU
    slide_h: int   # EMU
    font_latin: str
    font_ea: str
    fallback_ea: tuple
    chart_palette: tuple
    schema_version: str = "1.0"
    _colors: dict = field(default_factory=dict)
    _defaults: dict = field(default_factory=dict)
    _layout: dict = field(default_factory=dict)
    _layout_rules: dict = field(default_factory=dict)
    _executive_summary_skill_ids: tuple = field(default_factory=tuple)
    _skill_dir: Optional[str] = None

    def color(self, key: str) -> RGBColor:
        """Return color as RGBColor (for python-pptx font.color.rgb / fill.fore_color.rgb)."""
        if key not in self._colors:
            raise KeyError(f"theme color key not found: {key!r} (brand={self.id})")
        return _hex_to_rgbcolor(self._colors[key])

    def hex(self, key: str) -> str:
        """Return color as raw hex string '#RRGGBB' (for inline OOXML <a:srgbClr val='RRGGBB'>)."""
        if key not in self._colors:
            raise KeyError(f"theme color key not found: {key!r} (brand={self.id})")
        return self._colors[key]

    def hex_no_hash(self, key: str) -> str:
        """Return hex string 'RRGGBB' without leading '#' (for OOXML attribute val=)."""
        return self.hex(key).lstrip("#")

    def pt(self, key: str) -> "Pt":
        """Return font size as pptx Pt object."""
        if key not in self._defaults:
            raise KeyError(f"theme defaults key not found: {key!r} (brand={self.id})")
        return Pt(self._defaults[key])

    def pt_value(self, key: str) -> int:
        """Return font size as raw int (for places that need plain Pt int)."""
        if key not in self._defaults:
            raise KeyError(f"theme defaults key not found: {key!r} (brand={self.id})")
        return self._defaults[key]

    def layout(self, key: str) -> int:
        """Return layout coordinate as EMU (pptx Inches converted).

        Looks up `<key>` in skill-specific layout.json. Convention: keys end with
        '_in' for inches (e.g. 'panel_y_in'); the returned value is the EMU
        equivalent so it can be used directly as left/top/width/height.
        """
        if key not in self._layout:
            raise KeyError(f"theme layout key not found: {key!r} (brand={self.id}, skill_dir={self._skill_dir!r})")
        v = self._layout[key]
        # All current layout keys are inches; convert to EMU.
        return Inches(v)

    def layout_in(self, key: str) -> float:
        """Return layout coordinate as raw inches (for arithmetic, e.g. PANEL_Y + 0.55)."""
        if key not in self._layout:
            raise KeyError(f"theme layout key not found: {key!r} (brand={self.id})")
        return float(self._layout[key])

    # ---- schema 2.0 accessors (Phase 1, ISSUE-011) ------------------------
    # All accessors below are safe for schema 1.0 themes: missing keys return
    # the documented fallback so V1 (schema 1.0) brands continue to work even
    # when fill scripts are migrated to call these new accessors.

    def line_height_pt(self) -> Optional[int]:
        """Paragraph line spacing in pt for body/table cells.

        Used for OOXML `<a:lnSpc><a:spcPts val=...>`. Returns None when the
        brand does not specify a line height (caller should leave the existing
        OOXML untouched, preserving template default behaviour).
        """
        v = self._defaults.get("line_height_pt")
        return int(v) if v is not None else None

    def number_format_excel(self) -> Optional[str]:
        """Excel number format string for spreadsheet cells (e.g. '#,##0').

        Returns None when the brand does not pin a format (caller should fall
        back to the existing per-skill behaviour).
        """
        return self._defaults.get("number_format_excel")

    def zero_text(self) -> str:
        """Display text for cells whose value is exactly zero. Default '0'."""
        return self._defaults.get("zero_text", "0")

    def negative_format(self) -> str:
        """How to render negative numbers: 'paren' / 'triangle' / 'minus'."""
        return self._defaults.get("negative_format", "minus")

    def fiscal_period_format(self) -> Optional[str]:
        """Fiscal period label format (e.g. 'YY/MM期'). None if unspecified."""
        return self._defaults.get("fiscal_period_format")

    def font_size_body_pt(self, skill_id: Optional[str] = None) -> "Pt":
        """Body / table-cell font size as pptx Pt, with executive-summary auto-switch.

        When `skill_id` is in `executive_summary_skill_ids`, returns
        `font_size_executive_summary_body_pt`; otherwise `font_size_body_pt`.
        Falls back to legacy `font_size_label_pt` for schema 1.0 themes.
        """
        return Pt(self.font_size_body_pt_value(skill_id))

    def font_size_body_pt_value(self, skill_id: Optional[str] = None) -> int:
        """Same as font_size_body_pt() but returns raw int."""
        if skill_id and skill_id in self._executive_summary_skill_ids:
            v = self._defaults.get("font_size_executive_summary_body_pt")
            if v is not None:
                return int(v)
        v = self._defaults.get("font_size_body_pt")
        if v is not None:
            return int(v)
        # Schema 1.0 fallback: legacy label/value sizes (stella V1 = 14pt).
        return int(self._defaults.get("font_size_label_pt", 14))

    def layout_rule(self, key: str, default=None):
        """Read a value from theme.json `layout_rules` block.

        Common keys: 'left_align_guide_x_in' (float), 'uniform_table_width' (bool),
        'source_required' (bool).
        """
        return self._layout_rules.get(key, default)

    def is_source_required(self) -> bool:
        """True when 'source' field must be present (else fill_*.py raises)."""
        return bool(self._layout_rules.get("source_required", False))

    def is_executive_summary_skill(self, skill_id: str) -> bool:
        """True when the given skill_id should use executive-summary body size."""
        return skill_id in self._executive_summary_skill_ids

    # -----------------------------------------------------------------------

    def template_path(self, skill_dir: str, skill_name: str) -> str:
        """Resolve path to the brand-specific template pptx for this skill.

        Resolution order:
          1. `<skill_dir>/assets/<brand>/<skill_name>-template.pptx` (curated brand template)
          2. `<skill_dir>/assets/stellar_aiz/<skill_name>-template.pptx` (V1 placeholder
             fallback for non-stella brands that don't yet have a curated template;
             produces working output with stella slide structure + the requested brand's
             theme JSON applied)
          3. `<skill_dir>/assets/<skill_name>-template.pptx` (legacy single-template
             layout, pre-brand-aware skills)
        """
        branded = os.path.join(skill_dir, "assets", self.id, f"{skill_name}-template.pptx")
        if os.path.exists(branded):
            return branded
        stella_default = os.path.join(skill_dir, "assets", "stellar_aiz", f"{skill_name}-template.pptx")
        if os.path.exists(stella_default):
            return stella_default
        legacy = os.path.join(skill_dir, "assets", f"{skill_name}-template.pptx")
        if os.path.exists(legacy):
            return legacy
        raise FileNotFoundError(
            f"template not found for brand={self.id} skill={skill_name}: "
            f"tried {branded!r}, {stella_default!r}, {legacy!r}"
        )


def _load_theme_json(brand: str) -> dict:
    path = os.path.join(_BRANDS_DIR, brand, "theme.json")
    if not os.path.exists(path):
        raise FileNotFoundError(f"theme.json not found for brand={brand!r}: {path}")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _load_layout_json(skill_dir: str, brand: str) -> dict:
    """Load per-skill layout.json for the given brand. Returns {} if absent."""
    path = os.path.join(skill_dir, "assets", brand, "layout.json")
    if not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def resolve_brand(brand: str, skill_dir: Optional[str] = None) -> BrandTheme:
    """Resolve brand id to a BrandTheme.

    Args:
        brand: One of VALID_BRANDS. If not valid, raises ValueError.
        skill_dir: Absolute path to the skill directory (the dir containing
                   `assets/`, `scripts/`, etc.). Used to load per-skill
                   layout.json. Pass None if the script doesn't need layout
                   coordinates.

    Returns:
        A frozen BrandTheme with all values resolved.
    """
    if brand not in VALID_BRANDS:
        raise ValueError(f"invalid brand={brand!r}; must be one of {VALID_BRANDS}")

    theme_data = _load_theme_json(brand)

    required_keys = {"id", "slide_size", "fonts", "colors", "chart_palette", "defaults"}
    missing = required_keys - set(theme_data.keys())
    if missing:
        raise ValueError(f"theme.json for brand={brand!r} missing keys: {sorted(missing)}")

    if theme_data["id"] != brand:
        raise ValueError(
            f"theme.json id mismatch: file says id={theme_data['id']!r} "
            f"but loaded as brand={brand!r}"
        )

    slide_size = theme_data["slide_size"]
    fonts = theme_data["fonts"]

    layout_data = _load_layout_json(skill_dir, brand) if skill_dir else {}

    return BrandTheme(
        id=theme_data["id"],
        label=theme_data.get("label", theme_data["id"]),
        schema_version=str(theme_data.get("$schema_version", "1.0")),
        slide_w=Inches(slide_size["width_in"]),
        slide_h=Inches(slide_size["height_in"]),
        font_latin=fonts["latin"],
        font_ea=fonts["ea"],
        fallback_ea=tuple(fonts.get("fallback_ea", [])),
        chart_palette=tuple(theme_data["chart_palette"]),
        _colors=dict(theme_data["colors"]),
        _defaults=dict(theme_data["defaults"]),
        _layout=dict(layout_data),
        _layout_rules=dict(theme_data.get("layout_rules", {})),
        _executive_summary_skill_ids=tuple(theme_data.get("executive_summary_skill_ids", [])),
        _skill_dir=skill_dir,
    )


def add_brand_arg(parser: argparse.ArgumentParser) -> None:
    """Register --brand on the given argparse parser. Default = stellar_aiz."""
    parser.add_argument(
        "--brand",
        default=DEFAULT_BRAND,
        choices=VALID_BRANDS,
        help=f"Output brand (default: {DEFAULT_BRAND}). Options: {', '.join(VALID_BRANDS)}",
    )
