# Cross-Platform PPTX Guide (macOS Primary + Windows Compatible)

## 1. Why Academic PPTs Drift Between macOS and Windows

1. **Missing East Asian (`<a:ea>`) OOXML Tags**:
   Standard `python-pptx` calls like `run.font.name = "Arial"` write `<a:latin typeface="Arial"/>` only. When a text run contains Chinese characters without `<a:ea>`, **Windows PowerPoint falls back to `SimSun` (宋体)** while **macOS falls back to `PingFang SC`**, causing jarring visual mismatch and different line wraps.
2. **ECMA-376 XML Child Ordering**:
   In `<a:rPr>`, `<a:solidFill>` must precede `<a:latin>`, `<a:ea>`, `<a:cs>`. In `<a:tcPr>` (table cells), `<a:lnL>`, `<a:lnR>`, `<a:lnT>`, `<a:lnB>` must precede `<a:solidFill>` / `<a:noFill>`. Out-of-order tags cause Keynote or PowerPoint for Mac to drop borders or colors. `academic_ppt_core.py` enforces this exact order automatically.
3. **Default Table GUID Overrides**:
   PowerPoint's default table GUID (`Medium Style 2 - Accent 1`) injects opaque blue headers and alternating bands differently across versions. `neutralize_table_style()` switches tables to `{2D5ABB26-0587-4C30-8999-92F81FD0307C}` (*No Style, No Grid*).

## 2. Dual-Script Font Pairing Strategy

All built-in themes default to the universal high-consistency pair:
- **Latin / Numbers / Metrics (`<a:latin>`)**: `Arial`
  - Pre-installed natively on **100% of macOS and Windows systems** with identical glyph metric tables, keeping multi-column numerical tables pixel-aligned across OSes.
- **Chinese / East Asian (`<a:ea>`)**: `Microsoft YaHei`
  - Bundled inside **Microsoft PowerPoint for Mac**, pre-installed on **all Windows 10/11 machines**, and cleanly cascaded to **`PingFang SC`** (never `SimSun`) by Apple CoreText when opened in **Apple Keynote** on macOS.
- **Code / Model Identifiers**: `Courier New` (universal monospace on macOS & Windows).

### Matplotlib High-DPI Figures
Always initialize Matplotlib via `plt = setup_academic_mpl(deck.theme)`:
- Automatically detects installed CJK fonts in priority order: `PingFang SC` (macOS) → `Hiragino Sans GB` (macOS) → `Microsoft YaHei` (Mac Office / Windows) → `Noto Sans CJK SC` → `Source Han Sans SC`.
- Forces `dpi=300`, `axes.unicode_minus=False` (prevents missing minus signs), and `pdf.fonttype=42` (TrueType outlines).
- If converting vector `.pdf` figures from arXiv LaTeX sources for embedding into `.pptx`, rasterize them at **300 DPI PNG** via PyMuPDF (`fitz`) rather than inserting raw PDFs (Windows PowerPoint cannot natively place `.pdf` images, whereas 300-DPI PNG renders identically in macOS Keynote, Mac PowerPoint, and Windows PowerPoint).

## 3. Cross-Platform Command Execution (`uv`)

Avoid bare `python` (which triggers the Microsoft Store stub on Windows and is absent on clean macOS installs where only `python3` exists):
- Preferred on **both macOS and Windows**:
  ```bash
  uv run script.py
  ```
- Every script in `scripts/` includes PEP 723 inline dependency headers (`# /// script`), so `uv` creates and caches an isolated environment automatically on Mac (`~/.cache/uv`) and Windows (`%LOCALAPPDATA%\uv\cache`).
- Always pass `encoding="utf-8"` to `Path.read_text()`, `Path.write_text()`, and `open()` so Windows never defaults to `cp936`/`gbk`.
