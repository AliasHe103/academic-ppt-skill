#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "python-pptx>=1.0.2",
#   "lxml>=5.2.0",
# ]
# ///
"""
Build clean, reusable 16:9 widescreen (.pptx) master templates from theme JSON definitions.
Updates OOXML Slide Master `a:clrScheme` and `a:fontScheme` (both `a:latin` and `a:ea` tags)
so manual GUI edits in macOS PowerPoint / Apple Keynote / Windows PowerPoint also inherit
the exact cross-platform academic palette and dual-script typography.
"""

from __future__ import annotations

import sys
from pathlib import Path
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Inches

SKILL_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from academic_ppt_core import THEMES_DIR, TEMPLATES_DIR, hex_to_rgb, load_theme  # noqa: E402


def _patch_master_theme_xml(prs: Presentation, theme: dict) -> None:
    """Patch ppt/theme/theme1.xml color scheme and major/minor Latin+EA fonts."""
    slide_master = prs.slide_masters[0]
    theme_part = None
    for rel in slide_master.part.rels.values():
        if "theme" in rel.reltype:
            theme_part = rel.target_part
            break
    if theme_part is None:
        return

    root = etree.fromstring(theme_part.blob)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

    # 1. Patch color scheme (clrScheme)
    clr_scheme = root.find(".//a:clrScheme", namespaces=ns)
    if clr_scheme is not None:
        clr_scheme.set("name", theme["display_name"])
        c = theme["colors"]
        mapping = {
            "dk1": c["text_main"],
            "lt1": c["bg_canvas"],
            "dk2": c["primary"],
            "lt2": c["bg_subtle"],
            "accent1": c["primary"],
            "accent2": c["secondary"],
            "accent3": c["positive"],
            "accent4": c["accent"],
            "accent5": c["warning"],
            "accent6": c["text_muted"],
            "hlink": c["secondary"],
            "folHlink": c["text_muted"],
        }
        for slot, hex_val in mapping.items():
            slot_el = clr_scheme.find(f"a:{slot}", namespaces=ns)
            if slot_el is not None:
                for child in list(slot_el):
                    slot_el.remove(child)
                srgb = etree.SubElement(slot_el, qn("a:srgbClr"))
                srgb.set("val", hex_val.lstrip("#").upper())

    # 2. Patch font scheme (majorFont & minorFont for both Latin and East Asian)
    font_scheme = root.find(".//a:fontScheme", namespaces=ns)
    if font_scheme is not None:
        font_scheme.set("name", f"{theme['name']}-fonts")
        f = theme["fonts"]
        for group_tag, lat_face, ea_face in [
            ("majorFont", f["latin_heading"], f["ea_heading"]),
            ("minorFont", f["latin_body"], f["ea_body"]),
        ]:
            grp = font_scheme.find(f"a:{group_tag}", namespaces=ns)
            if grp is None:
                continue
            lat = grp.find("a:latin", namespaces=ns)
            if lat is not None:
                lat.set("typeface", lat_face)
            ea = grp.find("a:ea", namespaces=ns)
            if ea is not None:
                ea.set("typeface", ea_face)
            # Also update Hans/Hant supplemental font tags if present so Keynote/Mac PPT respects CJK
            for font_el in grp.findall("a:font", namespaces=ns):
                if font_el.get("script") in {"Hans", "Hant", "Jpan"}:
                    font_el.set("typeface", ea_face)

    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def build_all_templates() -> list[Path]:
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    built: list[Path] = []
    for theme_json in sorted(THEMES_DIR.glob("*.json")):
        theme = load_theme(theme_json.stem)
        prs = Presentation()
        prs.slide_width = Inches(theme["layout"]["slide_width_in"])
        prs.slide_height = Inches(theme["layout"]["slide_height_in"])

        # Set master background canvas color
        master_bg = prs.slide_masters[0].background.fill
        master_bg.solid()
        master_bg.fore_color.rgb = hex_to_rgb(theme["colors"]["bg_canvas"])

        _patch_master_theme_xml(prs, theme)

        out_path = TEMPLATES_DIR / f"{theme['name']}.pptx"
        prs.save(str(out_path))
        built.append(out_path)
        print(f"[OK] Built template: {out_path.name} ({theme['display_name']})")
    return built


if __name__ == "__main__":
    build_all_templates()
