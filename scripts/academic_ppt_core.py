#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "python-pptx>=1.0.2",
#   "matplotlib>=3.9.0",
#   "pillow>=10.4.0",
#   "numpy>=2.0.0",
# ]
# ///
"""
Cross-Platform Academic Presentation Core Library (macOS Primary + Windows Compatible).

Key engineering guarantees:
1. Strict ECMA-376 OOXML element ordering for dual-script fonts (<a:latin>, <a:ea>, <a:cs>)
   so macOS (PowerPoint for Mac & Apple Keynote) and Windows PowerPoint never fall back to SimSun.
2. Authentic LaTeX `booktabs` three-line tables (no vertical rules, grouped cmidrules,
   automatic **best** bolding, __second-best__ underlining, muted n/a, and "Ours" highlight rows).
3. High-density academic components for showcasing experimental rigor & workload:
   KPI workload strips, assertion-evidence headers, callout insight boxes, aspect-preserving figure panels,
   and OS-aware 300-DPI Matplotlib styling.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any, Iterable, Sequence

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SKILL_ROOT = Path(__file__).resolve().parent.parent
THEMES_DIR = SKILL_ROOT / "resources" / "themes"
TEMPLATES_DIR = SKILL_ROOT / "resources" / "templates"

# OOXML "No Style, No Grid" table style GUID to prevent PowerPoint theme color overrides
NO_STYLE_TABLE_GUID = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"

TOKEN_PATTERN = re.compile(r"(\*\*[^*]+\*\*|__[^_]+__|`[^`]+`)")


def hex_to_rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def load_theme(theme_name_or_path: str = "classic-conference") -> dict[str, Any]:
    candidate = Path(theme_name_or_path)
    if candidate.is_file():
        return json.loads(candidate.read_text(encoding="utf-8"))
    theme_file = THEMES_DIR / f"{theme_name_or_path}.json"
    if not theme_file.is_file():
        available = sorted(p.stem for p in THEMES_DIR.glob("*.json"))
        raise FileNotFoundError(
            f"Theme '{theme_name_or_path}' not found. Available themes: {available}"
        )
    return json.loads(theme_file.read_text(encoding="utf-8"))


def setup_academic_mpl(theme: dict[str, Any]) -> Any:
    """Configure Matplotlib for 300-DPI cross-platform CJK + Latin academic figures."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager

    installed_fonts = {f.name for f in font_manager.fontManager.ttflist}
    cjk_chain = theme["fonts"].get("mpl_cjk_chain", [])
    chosen_cjk = [f for f in cjk_chain if f in installed_fonts]
    latin_body = theme["fonts"].get("latin_body", "Arial")

    # Put Latin font first ONLY if pure ASCII, otherwise CJK-capable sans-serif chain handles both cleanly.
    # On macOS PingFang SC / Hiragino Sans GB has clean Latin+CJK; on Windows Microsoft YaHei has clean Latin+CJK.
    font_stack = chosen_cjk + [latin_body, "DejaVu Sans"]

    plt.rcParams.update(
        {
            "font.sans-serif": font_stack,
            "font.family": "sans-serif",
            "axes.unicode_minus": False,
            "figure.dpi": 300,
            "savefig.dpi": 300,
            "pdf.fonttype": 42,
            "ps.fonttype": 42,
            "axes.edgecolor": theme["colors"]["border_light"],
            "axes.linewidth": 0.8,
            "axes.labelcolor": theme["colors"]["text_body"],
            "xtick.color": theme["colors"]["text_body"],
            "ytick.color": theme["colors"]["text_body"],
            "text.color": theme["colors"]["text_main"],
            "grid.color": theme["colors"]["border_subtle"],
            "grid.linestyle": "--",
            "grid.linewidth": 0.6,
            "axes.spines.top": False,
            "axes.spines.right": False,
        }
    )
    return plt


def set_run_dual_font(
    run: Any,
    latin_font: str,
    ea_font: str,
    size_pt: float,
    color_hex: str,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
) -> None:
    """
    Apply font properties in strict ECMA-376 OOXML child element order:
    <a:solidFill> -> <a:latin> -> <a:ea> -> <a:cs>
    Ensures identical rendering across macOS (PowerPoint & Keynote) and Windows PowerPoint.
    """
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    run.font.color.rgb = hex_to_rgb(color_hex)
    run.font.name = latin_font  # Creates <a:latin> in schema-valid position

    rPr = run._r.get_or_add_rPr()
    latin_el = rPr.find(qn("a:latin"))

    for tag in ("a:ea", "a:cs"):
        existing = rPr.find(qn(tag))
        if existing is not None:
            rPr.remove(existing)

    cs_el = etree.Element(qn("a:cs"))
    cs_el.set("typeface", latin_font)
    ea_el = etree.Element(qn("a:ea"))
    ea_el.set("typeface", ea_font)

    if latin_el is not None:
        latin_el.addnext(cs_el)
        latin_el.addnext(ea_el)  # Resulting order: a:latin -> a:ea -> a:cs
    else:
        rPr.append(ea_el)
        rPr.append(cs_el)


def populate_formatted_runs(
    paragraph: Any,
    text: str,
    latin_font: str,
    ea_font: str,
    mono_font: str,
    size_pt: float,
    color_hex: str,
    bold_color_hex: str | None = None,
    code_color_hex: str | None = None,
    base_bold: bool = False,
    base_italic: bool = False,
) -> None:
    """
    Parse lightweight inline markup:
      **bold**   -> bold (and bold_color_hex if specified)
      __under__  -> underlined (e.g. second-best score in ML tables)
      `code`     -> monospace font (and code_color_hex if specified)
    """
    paragraph.text = ""
    parts = TOKEN_PATTERN.split(str(text))
    for part in parts:
        if not part:
            continue
        run = paragraph.add_run()
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            run.text = part[2:-2]
            set_run_dual_font(
                run,
                latin_font=latin_font,
                ea_font=ea_font,
                size_pt=size_pt,
                color_hex=bold_color_hex or color_hex,
                bold=True,
                italic=base_italic,
            )
        elif part.startswith("__") and part.endswith("__") and len(part) > 4:
            run.text = part[2:-2]
            set_run_dual_font(
                run,
                latin_font=latin_font,
                ea_font=ea_font,
                size_pt=size_pt,
                color_hex=color_hex,
                bold=base_bold,
                italic=base_italic,
                underline=True,
            )
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            run.text = part[1:-1]
            set_run_dual_font(
                run,
                latin_font=mono_font,
                ea_font=ea_font,
                size_pt=max(size_pt - 0.5, 7.5),
                color_hex=code_color_hex or bold_color_hex or color_hex,
                bold=base_bold,
                italic=False,
            )
        else:
            run.text = part
            set_run_dual_font(
                run,
                latin_font=latin_font,
                ea_font=ea_font,
                size_pt=size_pt,
                color_hex=color_hex,
                bold=base_bold,
                italic=base_italic,
            )


def _build_ln_element(tag: str, width_pt: float | None, color_hex: str | None) -> etree._Element:
    ln = etree.Element(qn(tag))
    if width_pt is None or width_pt <= 0 or not color_hex:
        etree.SubElement(ln, qn("a:noFill"))
    else:
        ln.set("w", str(int(round(width_pt * 12700))))  # 1 pt = 12700 EMU
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        solid = etree.SubElement(ln, qn("a:solidFill"))
        srgb = etree.SubElement(solid, qn("a:srgbClr"))
        srgb.set("val", color_hex.lstrip("#").upper())
        prst = etree.SubElement(ln, qn("a:prstDash"))
        prst.set("val", "solid")
    return ln


def style_table_cell_box(
    cell: Any,
    top: tuple[float, str] | None = None,
    bottom: tuple[float, str] | None = None,
    left: tuple[float, str] | None = None,
    right: tuple[float, str] | None = None,
    fill_hex: str | None = None,
    pad_lr_in: float = 0.06,
    pad_tb_in: float = 0.035,
) -> None:
    """
    Configure table cell borders & background in strict ECMA-376 CT_TableCellProperties order:
    lnL -> lnR -> lnT -> lnB -> solidFill / noFill.
    """
    cell.margin_left = Inches(pad_lr_in)
    cell.margin_right = Inches(pad_lr_in)
    cell.margin_top = Inches(pad_tb_in)
    cell.margin_bottom = Inches(pad_tb_in)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    # Set fill via python-pptx first, then place lnL..lnB strictly BEFORE fill elements
    if fill_hex:
        cell.fill.solid()
        cell.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    else:
        cell.fill.background()

    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)

    borders = [
        _build_ln_element("a:lnL", *(left if left else (None, None))),
        _build_ln_element("a:lnR", *(right if right else (None, None))),
        _build_ln_element("a:lnT", *(top if top else (None, None))),
        _build_ln_element("a:lnB", *(bottom if bottom else (None, None))),
    ]
    for idx, b_el in enumerate(borders):
        tcPr.insert(idx, b_el)


def neutralize_table_style(table: Any) -> None:
    """Strip PowerPoint's default blue banded table style GUID."""
    tblPr = table._tbl.tblPr
    if tblPr is not None:
        for attr in ("firstRow", "firstCol", "lastRow", "lastCol", "bandRow", "bandCol"):
            tblPr.set(attr, "0")
        style_id = tblPr.find(qn("a:tableStyleId"))
        if style_id is not None:
            style_id.text = NO_STYLE_TABLE_GUID


class AcademicDeck:
    """High-level builder for cross-platform minimalist academic presentations."""

    def __init__(
        self,
        theme_name: str = "classic-conference",
        template_path: Path | None = None,
    ) -> None:
        self.theme = load_theme(theme_name)
        self.colors = self.theme["colors"]
        self.fonts = self.theme["fonts"]
        self.layout_cfg = self.theme["layout"]

        default_tpl = TEMPLATES_DIR / f"{self.theme['name']}.pptx"
        chosen_tpl = template_path or (default_tpl if default_tpl.is_file() else None)
        if chosen_tpl and chosen_tpl.is_file():
            self.prs = Presentation(str(chosen_tpl))
        else:
            self.prs = Presentation()
            self.prs.slide_width = Inches(self.layout_cfg["slide_width_in"])
            self.prs.slide_height = Inches(self.layout_cfg["slide_height_in"])

        self._slide_counter = 0

    @property
    def blank_layout(self) -> Any:
        # Layout 6 is Blank in standard PowerPoint templates; fall back to the cleanest layout
        for layout in self.prs.slide_layouts:
            if len(layout.placeholders) == 0:
                return layout
        return self.prs.slide_layouts[-1]

    def new_slide(self, bg_hex: str | None = None) -> Any:
        slide = self.prs.slides.add_slide(self.blank_layout)
        # Remove any stray inherited placeholders just in case
        for shp in list(slide.placeholders):
            sp = shp._element
            sp.getparent().remove(sp)
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(bg_hex or self.colors["bg_canvas"])
        self._slide_counter += 1
        return slide

    def add_rect(
        self,
        slide: Any,
        x: float,
        y: float,
        w: float,
        h: float,
        fill_hex: str | None = None,
        line_hex: str | None = None,
        line_pt: float = 0.75,
        rounded: bool = False,
        radius_in: float = 0.12,
    ) -> Any:
        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
        shp = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
        shp.shadow.inherit = False
        spPr = shp._element.spPr
        effect_lst = spPr.find(qn("a:effectLst"))
        if effect_lst is None:
            etree.SubElement(spPr, qn("a:effectLst"))
        if rounded and hasattr(shp, "adjustments") and len(shp.adjustments) > 0:
            min_dim = max(min(w, h), 0.05)
            shp.adjustments[0] = min(0.5, radius_in / min_dim)
        if fill_hex:
            shp.fill.solid()
            shp.fill.fore_color.rgb = hex_to_rgb(fill_hex)
        else:
            shp.fill.background()
        if line_hex and line_pt > 0:
            shp.line.color.rgb = hex_to_rgb(line_hex)
            shp.line.width = Pt(line_pt)
        else:
            shp.line.fill.background()
        return shp

    def add_line(
        self,
        slide: Any,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        color_hex: str,
        width_pt: float = 1.0,
    ) -> Any:
        conn = slide.shapes.add_connector(
            1, Inches(x1), Inches(y1), Inches(x2), Inches(y2)  # 1 = MSO_CONNECTOR.STRAIGHT
        )
        conn.line.color.rgb = hex_to_rgb(color_hex)
        conn.line.width = Pt(width_pt)
        return conn

    def add_textbox(
        self,
        slide: Any,
        x: float,
        y: float,
        w: float,
        h: float,
        lines: str | Sequence[str | tuple[str, dict[str, Any]]],
        size_pt: float = 13.0,
        color_hex: str | None = None,
        bold: bool = False,
        italic: bool = False,
        align: PP_ALIGN = PP_ALIGN.LEFT,
        anchor: MSO_ANCHOR = MSO_ANCHOR.TOP,
        line_spacing: float = 1.22,
        space_after_pt: float = 3.0,
        is_heading: bool = False,
        bold_color_hex: str | None = None,
        pad_in: float = 0.02,
    ) -> Any:
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = Inches(pad_in)
        tf.margin_right = Inches(pad_in)
        tf.margin_top = Inches(pad_in)
        tf.margin_bottom = Inches(pad_in)

        latin_font = self.fonts["latin_heading"] if is_heading else self.fonts["latin_body"]
        ea_font = self.fonts["ea_heading"] if is_heading else self.fonts["ea_body"]
        mono_font = self.fonts["latin_mono"]
        default_color = color_hex or self.colors["text_body"]

        if isinstance(lines, str):
            items: list[str | tuple[str, dict[str, Any]]] = [lines]
        else:
            items = list(lines)

        for idx, item in enumerate(items):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            if isinstance(item, tuple):
                txt, opts = item
            else:
                txt, opts = item, {}

            p.alignment = opts.get("align", align)
            p.line_spacing = opts.get("line_spacing", line_spacing)
            p.space_after = Pt(opts.get("space_after_pt", space_after_pt if idx < len(items) - 1 else 0))
            if "space_before_pt" in opts:
                p.space_before = Pt(opts["space_before_pt"])

            populate_formatted_runs(
                p,
                txt,
                latin_font=opts.get("latin_font", latin_font),
                ea_font=opts.get("ea_font", ea_font),
                mono_font=mono_font,
                size_pt=opts.get("size_pt", size_pt),
                color_hex=opts.get("color_hex", default_color),
                bold_color_hex=opts.get("bold_color_hex", bold_color_hex or self.colors["primary"]),
                code_color_hex=opts.get("code_color_hex", self.colors["secondary"]),
                base_bold=opts.get("bold", bold),
                base_italic=opts.get("italic", italic),
            )
        return tb

    def add_header(
        self,
        slide: Any,
        section_tag: str,
        title: str,
        subtitle: str = "",
        right_badge: str = "",
    ) -> float:
        """
        Add a clean minimalist slide header and return the Y coordinate
        where main slide content can safely begin.
        """
        import unicodedata

        mx = self.layout_cfg["margin_x_in"]
        sw = self.layout_cfg["slide_width_in"]
        cw = sw - 2 * mx

        top_y = 0.32
        badge_w = 0.0
        if right_badge:
            vis_w = sum(
                0.175 if unicodedata.east_asian_width(ch) in {"W", "F"} else 0.092
                for ch in right_badge
            )
            badge_w = max(1.8, min(4.8, vis_w + 0.42))
            bx = mx + cw - badge_w
            self.add_rect(
                slide,
                bx,
                0.28,
                badge_w,
                0.30,
                fill_hex=self.colors.get("bg_tint", self.colors["bg_highlight_row"]),
                line_hex=None,
                rounded=True,
                radius_in=0.15,
            )
            self.add_textbox(
                slide,
                bx + 0.06,
                0.28,
                badge_w - 0.12,
                0.30,
                right_badge,
                size_pt=9.6,
                color_hex=self.colors["primary"],
                bold=True,
                align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE,
                is_heading=True,
            )

        if section_tag:
            tag_w = (cw - badge_w - 0.20) if badge_w > 0 else cw
            self.add_textbox(
                slide,
                mx,
                top_y,
                tag_w,
                0.26,
                section_tag.upper(),
                size_pt=9.8,
                color_hex=self.colors["secondary"],
                bold=True,
                is_heading=True,
            )
            title_y = top_y + 0.30
        else:
            title_y = 0.42

        self.add_textbox(
            slide,
            mx,
            title_y,
            cw,
            0.50,
            title,
            size_pt=21.5,
            color_hex=self.colors["text_main"],
            bold=True,
            is_heading=True,
        )

        if subtitle:
            self.add_textbox(
                slide,
                mx,
                title_y + 0.48,
                cw,
                0.32,
                subtitle,
                size_pt=12.0,
                color_hex=self.colors["text_muted"],
            )
            rule_y = title_y + 0.86
        else:
            rule_y = title_y + 0.58

        self.add_line(slide, mx, rule_y, mx + cw, rule_y, self.colors["border_subtle"], width_pt=0.85)
        return rule_y + 0.22

    def add_kpi_strip(
        self,
        slide: Any,
        x: float,
        y: float,
        w: float,
        h: float,
        items: Sequence[tuple[str, str, str]],
        gap: float = 0.24,
        bg_hex: str | None = None,
    ) -> None:
        """Render N metric cards horizontally in soft light rounded rectangles (no vertical bar)."""
        n = len(items)
        card_w = (w - gap * (n - 1)) / n
        fill = bg_hex or self.colors["bg_subtle"]
        for i, (val, label, sub) in enumerate(items):
            cx = x + i * (card_w + gap)
            self.add_rect(
                slide,
                cx,
                y,
                card_w,
                h,
                fill_hex=fill,
                line_hex=None,
                rounded=True,
                radius_in=0.15,
            )
            val_pt = 20.0 if len(val) <= 15 else 16.5
            paragraphs: list[tuple[str, dict[str, Any]]] = [
                (
                    val,
                    {
                        "size_pt": val_pt,
                        "color_hex": self.colors["primary"],
                        "bold": True,
                        "space_after_pt": 4.0,
                        "line_spacing": 1.10,
                    },
                )
            ]
            if label:
                paragraphs.append(
                    (
                        label,
                        {
                            "size_pt": 11.5,
                            "color_hex": self.colors["text_main"],
                            "bold": True,
                            "space_after_pt": 3.0 if sub else 0.0,
                            "line_spacing": 1.16,
                        },
                    )
                )
            if sub:
                paragraphs.append(
                    (
                        sub,
                        {
                            "size_pt": 10.0,
                            "color_hex": self.colors["text_muted"],
                            "bold": False,
                            "space_after_pt": 0.0,
                            "line_spacing": 1.18,
                        },
                    )
                )
            self.add_textbox(
                slide,
                cx + 0.22,
                y + 0.12,
                card_w - 0.44,
                h - 0.24,
                paragraphs,
                anchor=MSO_ANCHOR.MIDDLE,
            )

    def add_callout_card(
        self,
        slide: Any,
        x: float,
        y: float,
        w: float,
        h: float,
        heading: str,
        bullets: Sequence[str],
        accent_hex: str | None = None,
        bg_hex: str | None = None,
        body_size_pt: float = 12.0,
        heading_size_pt: float = 13.5,
        bullet_prefix: str = "•  ",
    ) -> None:
        """Render a soft light-colored rounded card (zero stuck-on vertical bars)."""
        accent = accent_hex or self.colors["primary"]
        fill = bg_hex or self.colors["bg_subtle"]
        self.add_rect(
            slide,
            x,
            y,
            w,
            h,
            fill_hex=fill,
            line_hex=None,
            rounded=True,
            radius_in=0.16,
        )

        paragraphs: list[tuple[str, dict[str, Any]]] = []
        if heading:
            paragraphs.append(
                (
                    heading,
                    {
                        "size_pt": heading_size_pt,
                        "color_hex": accent,
                        "bold": True,
                        "space_after_pt": 7.0,
                        "line_spacing": 1.18,
                    },
                )
            )
        for idx, b in enumerate(bullets):
            paragraphs.append(
                (
                    f"{bullet_prefix}{b}" if bullet_prefix else b,
                    {
                        "size_pt": body_size_pt,
                        "color_hex": self.colors["text_body"],
                        "bold": False,
                        "space_after_pt": 6.0 if idx < len(bullets) - 1 else 0.0,
                        "line_spacing": 1.30,
                    },
                )
            )

        self.add_textbox(
            slide,
            x + 0.24,
            y + 0.18,
            w - 0.48,
            h - 0.36,
            paragraphs,
            size_pt=body_size_pt,
            color_hex=self.colors["text_body"],
            bold_color_hex=self.colors["primary"],
            anchor=MSO_ANCHOR.MIDDLE,
        )

    def add_image_panel(
        self,
        slide: Any,
        x: float,
        y: float,
        w: float,
        h: float,
        image_path: Path | str,
        caption: str = "",
        framed: bool = True,
        pad_in: float = 0.12,
        bg_hex: str | None = None,
    ) -> None:
        """Insert an image preserving exact aspect ratio inside a soft rounded container."""
        img_p = Path(image_path)
        panel_fill = bg_hex or self.colors.get("bg_card", "#F8FAFC")
        if framed:
            self.add_rect(
                slide,
                x,
                y,
                w,
                h,
                fill_hex=panel_fill,
                line_hex=None,
                rounded=True,
                radius_in=0.16,
            )

        cap_h = 0.34 if caption else 0.0
        avail_w = w - 2 * pad_in
        avail_h = h - 2 * pad_in - cap_h

        with Image.open(img_p) as im:
            iw, ih = im.size
        aspect = iw / ih
        box_aspect = avail_w / avail_h

        if aspect > box_aspect:
            draw_w = avail_w
            draw_h = avail_w / aspect
        else:
            draw_h = avail_h
            draw_w = avail_h * aspect

        ix = x + pad_in + (avail_w - draw_w) / 2
        iy = y + pad_in + (avail_h - draw_h) / 2
        slide.shapes.add_picture(str(img_p), Inches(ix), Inches(iy), Inches(draw_w), Inches(draw_h))

        if caption:
            self.add_textbox(
                slide,
                x + pad_in,
                y + h - cap_h - 0.03,
                w - 2 * pad_in,
                cap_h,
                caption,
                size_pt=9.0,
                color_hex=self.colors["text_muted"],
                align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE,
            )

    def add_booktabs_table(
        self,
        slide: Any,
        x: float,
        y: float,
        w: float,
        h: float,
        headers: Sequence[str],
        rows: Sequence[Sequence[str]],
        col_widths: Sequence[float] | None = None,
        col_aligns: Sequence[PP_ALIGN] | None = None,
        group_headers: Sequence[tuple[str, int, int]] | None = None,
        highlight_rows: set[int] | None = None,
        section_rows: dict[int, str] | None = None,
        midrule_after_rows: set[int] | None = None,
        font_size_pt: float = 10.0,
        header_font_size_pt: float | None = None,
        caption: str = "",
        pad_lr_in: float = 0.05,
        pad_tb_in: float = 0.03,
    ) -> Any:
        """
        Render an authentic LaTeX `booktabs` three-line table:
        - Optional 2-row grouped header (`group_headers`: list of (label, start_col, end_col_inclusive))
        - Optional `section_rows`: {row_idx_in_data: "Fine-tuned on HotPotQA"} full-width italic divider rows
        - `highlight_rows`: set of data row indices (0-based) shaded with `bg_highlight_row` (e.g. Ours)
        - `midrule_after_rows`: set of data row indices (0-based) with a subtle horizontal rule below
        - Supports `**best**`, `__second_best__`, and muted `n/a` or `--`.
        """
        cap_h = 0.28 if caption else 0.0
        if caption:
            self.add_textbox(
                slide,
                x,
                y,
                w,
                cap_h,
                caption,
                size_pt=9.8,
                color_hex=self.colors["text_main"],
                bold=False,
                bold_color_hex=self.colors["primary"],
            )

        table_y = y + cap_h
        table_h = h - cap_h
        n_cols = len(headers)
        has_group_hdr = bool(group_headers)
        hdr_rows = 2 if has_group_hdr else 1

        # Expand data rows with optional section banner rows
        section_rows = section_rows or {}
        highlight_rows = highlight_rows or set()
        midrule_after_rows = midrule_after_rows or set()

        expanded_rows: list[tuple[str, int | None, Sequence[str]]] = []
        for r_idx, r_vals in enumerate(rows):
            if r_idx in section_rows:
                expanded_rows.append(("section", None, [section_rows[r_idx]] + [""] * (n_cols - 1)))
            expanded_rows.append(("data", r_idx, r_vals))

        total_rows = hdr_rows + len(expanded_rows)
        shape = slide.shapes.add_table(total_rows, n_cols, Inches(x), Inches(table_y), Inches(w), Inches(table_h))
        table = shape.table
        neutralize_table_style(table)

        # Column widths
        if col_widths:
            total_weight = sum(col_widths)
            for c_idx, cw in enumerate(col_widths):
                table.columns[c_idx].width = Inches(w * (cw / total_weight))
        else:
            even_w = w / n_cols
            for c_idx in range(n_cols):
                table.columns[c_idx].width = Inches(even_w)

        # Row heights proportional: header slightly taller, section rows compact
        row_weights: list[float] = []
        for r in range(hdr_rows):
            row_weights.append(1.12 if r == hdr_rows - 1 else 1.02)
        for kind, _, _ in expanded_rows:
            row_weights.append(0.84 if kind == "section" else 1.0)
        w_sum = sum(row_weights)
        for r_i, rw in enumerate(row_weights):
            table.rows[r_i].height = Inches(table_h * (rw / w_sum))

        top_rule = (self.layout_cfg["table_top_rule_pt"], self.colors["rule_heavy"])
        mid_rule = (self.layout_cfg["table_mid_rule_pt"], self.colors["rule_heavy"])
        sub_rule = (self.layout_cfg["table_sub_rule_pt"], self.colors["rule_mid"])
        light_rule = (0.45, self.colors["border_light"])
        bot_rule = (self.layout_cfg["table_bottom_rule_pt"], self.colors["rule_heavy"])

        hdr_fs = header_font_size_pt or font_size_pt
        latin_b = self.fonts["latin_body"]
        ea_b = self.fonts["ea_body"]
        mono_b = self.fonts["latin_mono"]

        if not col_aligns:
            col_aligns = [PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * (n_cols - 1)

        # 1. Render Group Header Row (if present)
        grouped_cols: set[int] = set()
        if has_group_hdr and group_headers:
            for c_idx in range(n_cols):
                style_table_cell_box(
                    table.cell(0, c_idx),
                    top=top_rule,
                    bottom=None,
                    fill_hex=self.colors["bg_subtle"],
                    pad_lr_in=pad_lr_in,
                    pad_tb_in=pad_tb_in,
                )
            for label, c_start, c_end in group_headers:
                for c_k in range(c_start, c_end + 1):
                    grouped_cols.add(c_k)
                    style_table_cell_box(
                        table.cell(0, c_k),
                        top=top_rule,
                        bottom=sub_rule,
                        fill_hex=self.colors["bg_subtle"],
                        pad_lr_in=pad_lr_in,
                        pad_tb_in=pad_tb_in,
                    )
                if c_end > c_start:
                    table.cell(0, c_start).merge(table.cell(0, c_end))
                cell = table.cell(0, c_start)
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                populate_formatted_runs(
                    p,
                    label,
                    latin_font=latin_b,
                    ea_font=ea_b,
                    mono_font=mono_b,
                    size_pt=hdr_fs,
                    color_hex=self.colors["primary"],
                    bold_color_hex=self.colors["primary"],
                    base_bold=True,
                )

        # 2. Render Main Column Header Row
        h_row_idx = 1 if has_group_hdr else 0
        for c_idx, h_text in enumerate(headers):
            cell = table.cell(h_row_idx, c_idx)
            top_b = None if has_group_hdr and c_idx not in grouped_cols else (sub_rule if has_group_hdr else top_rule)
            if has_group_hdr and c_idx not in grouped_cols:
                top_b = None
            style_table_cell_box(
                cell,
                top=top_b if not has_group_hdr else (sub_rule if c_idx in grouped_cols else None),
                bottom=mid_rule,
                fill_hex=self.colors["bg_subtle"],
                pad_lr_in=pad_lr_in,
                pad_tb_in=pad_tb_in,
            )
            p = cell.text_frame.paragraphs[0]
            p.alignment = col_aligns[c_idx]
            populate_formatted_runs(
                p,
                h_text,
                latin_font=latin_b,
                ea_font=ea_b,
                mono_font=mono_b,
                size_pt=hdr_fs,
                color_hex=self.colors["text_main"],
                bold_color_hex=self.colors["primary"],
                base_bold=True,
            )

        # 3. Render Expanded Data & Section Rows
        n_exp = len(expanded_rows)
        for e_idx, (kind, orig_r_idx, row_vals) in enumerate(expanded_rows):
            t_row = hdr_rows + e_idx
            is_last = e_idx == n_exp - 1

            if kind == "section":
                for c_idx in range(n_cols):
                    style_table_cell_box(
                        table.cell(t_row, c_idx),
                        top=light_rule if e_idx > 0 else mid_rule,
                        bottom=light_rule,
                        fill_hex=self.colors["bg_group_row"],
                        pad_lr_in=pad_lr_in,
                        pad_tb_in=0.02,
                    )
                if n_cols > 1:
                    table.cell(t_row, 0).merge(table.cell(t_row, n_cols - 1))
                scell = table.cell(t_row, 0)
                p = scell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                populate_formatted_runs(
                    p,
                    row_vals[0],
                    latin_font=latin_b,
                    ea_font=ea_b,
                    mono_font=mono_b,
                    size_pt=max(font_size_pt - 0.8, 8.0),
                    color_hex=self.colors["secondary"],
                    base_bold=True,
                    base_italic=True,
                )
                continue

            is_hi = orig_r_idx in highlight_rows
            has_mid_after = orig_r_idx in midrule_after_rows
            row_fill = self.colors["bg_highlight_row"] if is_hi else None
            b_border = bot_rule if is_last else (light_rule if has_mid_after else None)

            for c_idx, val in enumerate(row_vals):
                cell = table.cell(t_row, c_idx)
                style_table_cell_box(
                    cell,
                    top=None,
                    bottom=b_border,
                    left=None,
                    fill_hex=row_fill,
                    pad_lr_in=pad_lr_in,
                    pad_tb_in=pad_tb_in,
                )
                p = cell.text_frame.paragraphs[0]
                p.alignment = col_aligns[c_idx]
                s_val = str(val).strip()
                is_na = s_val.lower() in {"n/a", "--", "—", "-"}
                cell_color = (
                    self.colors["text_muted"]
                    if is_na
                    else (self.colors["primary"] if (is_hi and c_idx == 0) else self.colors["text_body"])
                )
                populate_formatted_runs(
                    p,
                    s_val,
                    latin_font=latin_b,
                    ea_font=ea_b,
                    mono_font=mono_b,
                    size_pt=font_size_pt,
                    color_hex=cell_color,
                    bold_color_hex=self.colors["primary"],
                    base_bold=(is_hi and c_idx == 0),
                )

        return shape

    # =========================================================================
    # Canonical 5-Slide Short Deck Layout Archetypes (Group-Meeting / Idea Deck)
    # =========================================================================

    def add_background_slide(
        self,
        section_tag: str,
        title: str,
        bottlenecks: Sequence[tuple[str, str, str]],
        core_heading: str,
        core_points: Sequence[str],
        highlights_heading: str,
        highlights_points: Sequence[str],
        right_badge: str = "",
    ) -> Any:
        """
        Template 1 (Paper Background Topic Slide — replaces bulky title covers):
        - Top row: 3 equal soft neutral rounded cards showing prior paradigms & bottlenecks.
        - Bottom row: Wide soft-tinted Core Idea card (left) + concise Key Advantages card (right).
        """
        slide = self.new_slide()
        self.add_header(slide, section_tag, title, right_badge=right_badge)

        mx = self.layout_cfg["margin_x_in"]
        cw = self.layout_cfg["slide_width_in"] - 2 * mx
        tint_bg = self.colors.get("bg_tint", self.colors["bg_highlight_row"])
        subtle_bg = self.colors["bg_subtle"]

        # Upper row: 3 prior bottleneck cards
        n = len(bottlenecks)
        gap = 0.26
        card_w = (cw - gap * (n - 1)) / n
        top_y = 1.52
        top_h = 2.26

        for i, (route_title, sub_tag, desc) in enumerate(bottlenecks):
            cx = mx + i * (card_w + gap)
            self.add_rect(
                slide,
                cx,
                top_y,
                card_w,
                top_h,
                fill_hex=subtle_bg,
                line_hex=None,
                rounded=True,
                radius_in=0.16,
            )
            paragraphs: list[tuple[str, dict[str, Any]]] = [
                (
                    route_title,
                    {
                        "size_pt": 13.5,
                        "color_hex": self.colors["text_main"],
                        "bold": True,
                        "space_after_pt": 3.5,
                        "line_spacing": 1.15,
                    },
                ),
                (
                    sub_tag,
                    {
                        "size_pt": 10.8,
                        "color_hex": self.colors["secondary"],
                        "bold": True,
                        "space_after_pt": 7.0,
                        "line_spacing": 1.15,
                    },
                ),
                (
                    desc,
                    {
                        "size_pt": 11.6,
                        "color_hex": self.colors["text_body"],
                        "bold": False,
                        "space_after_pt": 0.0,
                        "line_spacing": 1.28,
                    },
                ),
            ]
            self.add_textbox(
                slide,
                cx + 0.24,
                top_y + 0.16,
                card_w - 0.48,
                top_h - 0.32,
                paragraphs,
                bold_color_hex=self.colors["primary"],
                anchor=MSO_ANCHOR.MIDDLE,
            )

        # Lower row: Core Idea synthesis card (left 63%) + Key Advantages card (right 37%)
        bot_y = 4.06
        bot_h = 2.58
        left_w = round(cw * 0.63 - gap / 2, 3)
        right_w = round(cw - left_w - gap, 3)

        self.add_callout_card(
            slide,
            mx,
            bot_y,
            left_w,
            bot_h,
            heading=core_heading,
            bullets=core_points,
            bg_hex=tint_bg,
            heading_size_pt=14.2,
            body_size_pt=12.0,
            bullet_prefix="•  " if len(core_points) > 1 else "",
        )
        self.add_callout_card(
            slide,
            mx + left_w + gap,
            bot_y,
            right_w,
            bot_h,
            heading=highlights_heading,
            bullets=highlights_points,
            bg_hex=subtle_bg,
            heading_size_pt=13.8,
            body_size_pt=11.8,
            bullet_prefix="•  ",
        )
        return slide

    def add_table_summary_slide(
        self,
        section_tag: str,
        title: str,
        headers: Sequence[str],
        rows: Sequence[Sequence[str]],
        kpi: tuple[str, str, str],
        summary_heading: str,
        summary_lines: Sequence[str],
        col_widths: Sequence[float] | None = None,
        col_aligns: Sequence[PP_ALIGN] | None = None,
        group_headers: Sequence[tuple[str, int, int]] | None = None,
        highlight_rows: set[int] | None = None,
        font_size_pt: float = 12.6,
        header_font_size_pt: float = 12.8,
        right_badge: str = "",
    ) -> Any:
        """
        Template 2 (Full-Width Benchmark Table + Bottom KPI & Summary Strip):
        - Upper zone: Full-width spacious `booktabs` three-line table.
        - Lower zone: Left headline metric card + Right concise 2–3 sentence summary card.
        """
        slide = self.new_slide()
        self.add_header(slide, section_tag, title, right_badge=right_badge)

        mx = self.layout_cfg["margin_x_in"]
        cw = self.layout_cfg["slide_width_in"] - 2 * mx
        tint_bg = self.colors.get("bg_tint", self.colors["bg_highlight_row"])
        subtle_bg = self.colors["bg_subtle"]

        self.add_booktabs_table(
            slide,
            mx,
            1.52,
            cw,
            3.28,
            headers=headers,
            rows=rows,
            col_widths=col_widths,
            col_aligns=col_aligns,
            group_headers=group_headers,
            highlight_rows=highlight_rows,
            font_size_pt=font_size_pt,
            header_font_size_pt=header_font_size_pt,
            pad_lr_in=0.10,
            pad_tb_in=0.07,
        )

        gap = 0.26
        kpi_w = 3.92
        sum_w = cw - kpi_w - gap
        bot_y = 5.06
        bot_h = 1.58

        self.add_kpi_strip(
            slide,
            mx,
            bot_y,
            kpi_w,
            bot_h,
            items=[kpi],
            bg_hex=tint_bg,
        )
        self.add_callout_card(
            slide,
            mx + kpi_w + gap,
            bot_y,
            sum_w,
            bot_h,
            heading=summary_heading,
            bullets=summary_lines,
            bg_hex=subtle_bg,
            heading_size_pt=13.2,
            body_size_pt=11.8,
            bullet_prefix="•  " if len(summary_lines) > 1 else "",
        )
        return slide

    def add_figure_summary_slide(
        self,
        section_tag: str,
        title: str,
        image_path: Path | str,
        kpi: tuple[str, str, str],
        summary_heading: str,
        summary_lines: Sequence[str],
        caption: str = "",
        right_badge: str = "",
    ) -> Any:
        """
        Template 3 (Left Flagship Figure + Right Stacked KPI & Summary Cards):
        - Left zone: High-DPI figure/curve in an aspect-locked soft rounded panel.
        - Right zone: Top headline KPI card + Bottom concise summary card.
        """
        slide = self.new_slide()
        self.add_header(slide, section_tag, title, right_badge=right_badge)

        mx = self.layout_cfg["margin_x_in"]
        cw = self.layout_cfg["slide_width_in"] - 2 * mx
        tint_bg = self.colors.get("bg_tint", self.colors["bg_highlight_row"])
        subtle_bg = self.colors["bg_subtle"]

        gap = 0.28
        fig_w = 7.10
        right_w = cw - fig_w - gap

        self.add_image_panel(
            slide,
            mx,
            1.52,
            fig_w,
            5.12,
            image_path=image_path,
            caption=caption,
            pad_in=0.16,
        )
        self.add_kpi_strip(
            slide,
            mx + fig_w + gap,
            1.52,
            right_w,
            1.96,
            items=[kpi],
            bg_hex=tint_bg,
        )
        self.add_callout_card(
            slide,
            mx + fig_w + gap,
            3.74,
            right_w,
            2.90,
            heading=summary_heading,
            bullets=summary_lines,
            bg_hex=subtle_bg,
            heading_size_pt=13.6,
            body_size_pt=12.0,
            bullet_prefix="•  " if len(summary_lines) > 1 else "",
        )
        return slide

    def add_hybrid_table_chart_slide(
        self,
        section_tag: str,
        title: str,
        headers: Sequence[str],
        rows: Sequence[Sequence[str]],
        summary_heading: str,
        summary_lines: Sequence[str],
        image_path: Path | str,
        col_widths: Sequence[float] | None = None,
        col_aligns: Sequence[PP_ALIGN] | None = None,
        highlight_rows: set[int] | None = None,
        font_size_pt: float = 12.2,
        header_font_size_pt: float = 12.2,
        caption: str = "",
        right_badge: str = "",
    ) -> Any:
        """
        Template 4 (Left Compact Table & Summary + Right Comparison Chart):
        - Left top: Compact `booktabs` numerical table.
        - Left bottom: Soft-tinted summary card explaining the key contrast.
        - Right: High-DPI comparison chart inside a soft rounded panel.
        """
        slide = self.new_slide()
        self.add_header(slide, section_tag, title, right_badge=right_badge)

        mx = self.layout_cfg["margin_x_in"]
        cw = self.layout_cfg["slide_width_in"] - 2 * mx
        tint_bg = self.colors.get("bg_tint", self.colors["bg_highlight_row"])

        gap = 0.26
        left_w = 6.32
        right_w = cw - left_w - gap

        self.add_booktabs_table(
            slide,
            mx,
            1.52,
            left_w,
            3.16,
            headers=headers,
            rows=rows,
            col_widths=col_widths,
            col_aligns=col_aligns,
            highlight_rows=highlight_rows,
            font_size_pt=font_size_pt,
            header_font_size_pt=header_font_size_pt,
            pad_lr_in=0.08,
            pad_tb_in=0.07,
        )
        self.add_callout_card(
            slide,
            mx,
            4.94,
            left_w,
            1.70,
            heading=summary_heading,
            bullets=summary_lines,
            bg_hex=tint_bg,
            heading_size_pt=13.2,
            body_size_pt=11.8,
            bullet_prefix="•  " if len(summary_lines) > 1 else "",
        )
        self.add_image_panel(
            slide,
            mx + left_w + gap,
            1.52,
            right_w,
            5.12,
            image_path=image_path,
            caption=caption,
            pad_in=0.14,
        )
        return slide

    def add_kpi_chart_summary_slide(
        self,
        section_tag: str,
        title: str,
        kpis: Sequence[tuple[str, str, str]],
        summary_heading: str,
        summary_lines: Sequence[str],
        image_path: Path | str,
        caption: str = "",
        right_badge: str = "",
    ) -> Any:
        """
        Template 5 (Top Multi-KPI Strip + Bottom Left Summary & Right Wide Chart):
        - Top row: 3-item horizontal KPI comparison strip in soft-tinted rounded cards.
        - Bottom left: Soft neutral summary card explaining why/how the deltas arise.
        - Bottom right: Wide-aspect high-DPI chart (e.g., ablation error bars or trade-off curve).
        """
        slide = self.new_slide()
        self.add_header(slide, section_tag, title, right_badge=right_badge)

        mx = self.layout_cfg["margin_x_in"]
        cw = self.layout_cfg["slide_width_in"] - 2 * mx
        tint_bg = self.colors.get("bg_tint", self.colors["bg_highlight_row"])
        subtle_bg = self.colors["bg_subtle"]

        gap = 0.26
        self.add_kpi_strip(
            slide,
            mx,
            1.52,
            cw,
            1.38,
            items=kpis,
            gap=gap,
            bg_hex=tint_bg,
        )

        bot_y = 3.16
        bot_h = 3.48
        left_w = 4.68
        right_w = cw - left_w - gap

        self.add_callout_card(
            slide,
            mx,
            bot_y,
            left_w,
            bot_h,
            heading=summary_heading,
            bullets=summary_lines,
            bg_hex=subtle_bg,
            heading_size_pt=13.6,
            body_size_pt=11.9,
            bullet_prefix="•  " if len(summary_lines) > 1 else "",
        )
        self.add_image_panel(
            slide,
            mx + left_w + gap,
            bot_y,
            right_w,
            bot_h,
            image_path=image_path,
            caption=caption,
            pad_in=0.14,
        )
        return slide

    def save(self, output_path: Path | str) -> Path:
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(out))
        return out
