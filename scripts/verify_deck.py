#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "python-pptx>=1.0.2",
#   "lxml>=5.2.0",
#   "pymupdf>=1.24.0",
#   "pillow>=10.4.0",
# ]
# ///
"""
Cross-Platform PPTX Quality Verifier & Slide Rasterizer (macOS + Windows + Linux).

Checks performed:
1. Dual-script OOXML font audit (<a:latin> + <a:ea>) across all text frames & table cells.
2. Canvas boundary audit (flags shapes overflowing 16:9 slide bounds).
3. Sibling text/table collision audit (detects unintended overlapping textboxes/tables).
4. CJK-aware text capacity audit (flags potential text overflow inside tight boxes).
5. Optional high-DPI PDF/PNG rendering (`--render-dir`):
   - Tries LibreOffice (`soffice --headless`) on macOS / Linux / Windows first.
   - Falls back to Microsoft PowerPoint automation on macOS (`osascript`) or Windows (`PowerPoint.Application` COM).
"""

from __future__ import annotations

import argparse
import json
import math
import shutil
import subprocess
import sys
import unicodedata
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.oxml.ns import qn

EMU_PER_INCH = 914400
EMU_PER_PT = 12700


def visual_char_units(text: str) -> float:
    """Estimate character width units (CJK wide = 1.85x, ASCII = 0.95x)."""
    units = 0.0
    for ch in text:
        if ch in "\r\n":
            continue
        eaw = unicodedata.east_asian_width(ch)
        units += 1.85 if eaw in {"W", "F"} else 0.95
    return units


def check_rpr_dual_fonts(rPr: Any) -> tuple[bool, bool]:
    if rPr is None:
        return False, False
    has_latin = rPr.find(qn("a:latin")) is not None
    has_ea = rPr.find(qn("a:ea")) is not None
    return has_latin, has_ea


def audit_pptx(pptx_path: Path, max_slides: int = 5) -> dict[str, Any]:
    prs = Presentation(str(pptx_path))
    sw_in = prs.slide_width / EMU_PER_INCH
    sh_in = prs.slide_height / EMU_PER_INCH

    issues: list[dict[str, Any]] = []
    slide_summaries: list[dict[str, Any]] = []

    if len(prs.slides) > max_slides:
        issues.append(
            {
                "slide": 0,
                "severity": "warning",
                "type": "slide_count_exceeds_short_deck",
                "detail": f"Deck has {len(prs.slides)} slides (short-deck default is <= {max_slides} slides).",
            }
        )

    for s_idx, slide in enumerate(prs.slides, start=1):
        text_boxes_meta: list[dict[str, Any]] = []
        missing_dual_font_runs = 0
        total_runs = 0

        for shp in slide.shapes:
            x = shp.left / EMU_PER_INCH
            y = shp.top / EMU_PER_INCH
            w = shp.width / EMU_PER_INCH
            h = shp.height / EMU_PER_INCH

            # 1. Canvas boundary & clean bottom margin check (no footers / page counters allowed)
            if x < -0.05 or y < -0.05 or (x + w) > (sw_in + 0.05) or (y + h) > (sh_in + 0.05):
                issues.append(
                    {
                        "slide": s_idx,
                        "severity": "error",
                        "type": "out_of_bounds",
                        "detail": f"Shape '{shp.name}' at ({x:.2f}, {y:.2f}, {w:.2f}x{h:.2f} in) exceeds ({sw_in:.2f}x{sh_in:.2f} in).",
                    }
                )
            elif (y + h) > (sh_in - 0.50):
                issues.append(
                    {
                        "slide": s_idx,
                        "severity": "error",
                        "type": "bottom_margin_violation",
                        "detail": f"Shape '{shp.name}' reaches y={y + h:.2f} in; bottom 0.50 in must stay completely empty (no footers or page counters).",
                    }
                )

            # 2. TextFrame audit
            if shp.has_text_frame:
                tf = shp.text_frame
                raw_text = tf.text.strip()
                if raw_text:
                    if "$" in raw_text and "\\" in raw_text:
                        issues.append(
                            {
                                "slide": s_idx,
                                "severity": "error",
                                "type": "raw_latex_in_textbox",
                                "detail": f"Raw LaTeX math syntax found in PPTX text '{raw_text[:45]}...'. Use clean Unicode math symbols instead.",
                            }
                        )
                    max_pt = 11.0
                    est_lines = 0
                    for p in tf.paragraphs:
                        p_size_pt = 10.5
                        p_units = 0.0
                        for r in p.runs:
                            if not r.text:
                                continue
                            total_runs += 1
                            has_lat, has_ea = check_rpr_dual_fonts(r._r.rPr)
                            if not (has_lat and has_ea):
                                missing_dual_font_runs += 1
                            if r.font.size:
                                r_pt = r.font.size / EMU_PER_PT
                                p_size_pt = max(p_size_pt, r_pt)
                                max_pt = max(max_pt, r_pt)
                            p_units += visual_char_units(r.text)
                        usable_w_pt = max((w - 0.06) * 72.0, p_size_pt)
                        chars_per_line = max(usable_w_pt / (p_size_pt * 0.56), 1.0)
                        p_lines = max(1, math.ceil(p_units / chars_per_line))
                        est_lines += p_lines * (p_size_pt / max_pt)

                    est_height_in = (est_lines * max_pt * 1.25) / 72.0
                    if est_height_in > h * 1.28 and h < 6.5:
                        issues.append(
                            {
                                "slide": s_idx,
                                "severity": "warning",
                                "type": "potential_text_overflow",
                                "detail": (
                                    f"Text '{raw_text[:34]}...' est height {est_height_in:.2f} in > box {h:.2f} in."
                                ),
                            }
                        )
                    text_boxes_meta.append(
                        {"kind": "text", "x": x, "y": y, "w": w, "h": h, "preview": raw_text[:28]}
                    )

            # 3. Table audit
            if shp.has_table:
                tbl = shp.table
                for r_i, row in enumerate(tbl.rows):
                    for c_i, cell in enumerate(row.cells):
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                if not r.text:
                                    continue
                                total_runs += 1
                                has_lat, has_ea = check_rpr_dual_fonts(r._r.rPr)
                                if not (has_lat and has_ea):
                                    missing_dual_font_runs += 1
                text_boxes_meta.append(
                    {
                        "kind": "table",
                        "x": x,
                        "y": y,
                        "w": w,
                        "h": h,
                        "preview": f"Table({len(tbl.rows)}x{len(tbl.columns)})",
                    }
                )

            # 4. Picture audit
            if shp.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                text_boxes_meta.append(
                    {"kind": "picture", "x": x, "y": y, "w": w, "h": h, "preview": shp.name}
                )

        if missing_dual_font_runs > 0:
            issues.append(
                {
                    "slide": s_idx,
                    "severity": "error",
                    "type": "missing_dual_font_tags",
                    "detail": f"{missing_dual_font_runs}/{total_runs} text runs lack explicit <a:latin> + <a:ea> font tags.",
                }
            )

        # 5. Pairwise overlap check between foreground content elements (text / table / picture)
        n_fg = len(text_boxes_meta)
        for i in range(n_fg):
            for j in range(i + 1, n_fg):
                a, b = text_boxes_meta[i], text_boxes_meta[j]
                ix1 = max(a["x"], b["x"])
                iy1 = max(a["y"], b["y"])
                ix2 = min(a["x"] + a["w"], b["x"] + b["w"])
                iy2 = min(a["y"] + a["h"], b["y"] + b["h"])
                if ix2 - ix1 > 0.08 and iy2 - iy1 > 0.08:
                    issues.append(
                        {
                            "slide": s_idx,
                            "severity": "warning",
                            "type": "foreground_overlap",
                            "detail": (
                                f"Overlap ({ix2 - ix1:.2f}x{iy2 - iy1:.2f} in) between "
                                f"[{a['kind']}: {a['preview']}] and [{b['kind']}: {b['preview']}]."
                            ),
                        }
                    )

        slide_summaries.append(
            {
                "slide": s_idx,
                "foreground_items": len(text_boxes_meta),
                "runs_checked": total_runs,
                "dual_font_ok": missing_dual_font_runs == 0,
            }
        )

    return {
        "file": str(pptx_path),
        "slide_count": len(prs.slides),
        "dimensions_in": [round(sw_in, 3), round(sh_in, 3)],
        "error_count": sum(1 for x in issues if x["severity"] == "error"),
        "warning_count": sum(1 for x in issues if x["severity"] == "warning"),
        "issues": issues,
        "slides": slide_summaries,
    }


def _convert_pptx_to_pdf(pptx_path: Path, pdf_path: Path) -> bool:
    pptx_abs = pptx_path.resolve()
    pdf_abs = pdf_path.resolve()
    pdf_abs.parent.mkdir(parents=True, exist_ok=True)

    # Strategy 1: Cross-platform LibreOffice (`soffice` on PATH or macOS /Applications)
    soffice_candidates = [
        shutil.which("soffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
    ]
    for cand in soffice_candidates:
        if cand and Path(cand).is_file():
            res = subprocess.run(
                [cand, "--headless", "--convert-to", "pdf", "--outdir", str(pdf_abs.parent), str(pptx_abs)],
                capture_output=True,
                text=True,
                timeout=90,
            )
            generated = pdf_abs.parent / f"{pptx_abs.stem}.pdf"
            if res.returncode == 0 and generated.is_file():
                if generated != pdf_abs:
                    generated.replace(pdf_abs)
                return True

    # Strategy 2: Windows PowerPoint COM via PowerShell
    if sys.platform == "win32":
        ps_script = f"""
$ErrorActionPreference = 'Stop'
$ppt = New-Object -ComObject PowerPoint.Application
try {{
    $pres = $ppt.Presentations.Open('{str(pptx_abs)}', $true, $false, $false)
    $pres.SaveAs('{str(pdf_abs)}', 32)
    $pres.Close()
}} finally {{
    if ($ppt.Presentations.Count -eq 0) {{
        $ppt.Quit()
    }}
}}
"""
        res = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=90,
        )
        if res.returncode == 0 and pdf_abs.is_file():
            return True

    # Strategy 3: macOS Microsoft PowerPoint or Keynote AppleScript
    if sys.platform == "darwin":
        osa = f'''
tell application "Microsoft PowerPoint"
    open POSIX file "{str(pptx_abs)}"
    save active presentation in POSIX file "{str(pdf_abs)}" as save as PDF
    close active presentation saving no
end tell
'''
        res = subprocess.run(["osascript", "-e", osa], capture_output=True, text=True, timeout=90)
        if res.returncode == 0 and pdf_abs.is_file():
            return True

    return False


def render_slides_to_png(pptx_path: Path, render_dir: Path, dpi: int = 180) -> list[Path]:
    import pymupdf as fitz

    render_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = render_dir / f"{pptx_path.stem}.pdf"
    if not _convert_pptx_to_pdf(pptx_path, pdf_path):
        print("[WARN] Neither LibreOffice nor PowerPoint automation was available for PDF rasterization.")
        return []

    doc = fitz.open(str(pdf_path))
    png_paths: list[Path] = []
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    for idx, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=mat, alpha=False)
        out_png = render_dir / f"slide_{idx:02d}.png"
        pix.save(str(out_png))
        png_paths.append(out_png)
    doc.close()
    return png_paths


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit and optionally rasterize an academic PPTX deck.")
    parser.add_argument("pptx", type=Path, help="Path to the .pptx file to verify.")
    parser.add_argument("--render-dir", type=Path, default=None, help="Directory to export PDF & slide PNGs.")
    parser.add_argument("--dpi", type=int, default=180, help="Rasterization DPI when --render-dir is set.")
    parser.add_argument("--max-slides", type=int, default=5, help="Max slides for short-deck check (default: 5).")
    args = parser.parse_args()

    report = audit_pptx(args.pptx, max_slides=args.max_slides)
    if args.render_dir:
        pngs = render_slides_to_png(args.pptx, args.render_dir, dpi=args.dpi)
        report["rendered_slides"] = [str(p) for p in pngs]

    print(json.dumps(report, ensure_ascii=False, indent=2))
    if report["error_count"] > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
