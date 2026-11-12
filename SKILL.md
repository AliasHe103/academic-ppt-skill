---
name: academic-ppt
description: >-
  Creates cross-platform (macOS primary, Windows compatible) minimalist academic paper presentations (.pptx)
  using a concise 5-slide short-deck system (1 Paper Background topic slide + 4 distinct Data-Display &
  Few-Sentence-Summary templates) with soft rounded cards, publication-grade three-line tables (booktabs),
  and 300-DPI charts. Use this skill whenever the user asks to generate, design, or update an academic slide
  deck, group meeting paper share, idea pitch, benchmark report, or conference talk PPT.
---

# Academic Paper Presentation Skill (`academic-ppt`)

Generate clean, minimalist 16:9 `.pptx` slide decks tailored for **paper reading, group-meeting idea sharing, and experimental data reporting**, with identical rendering across **macOS (PowerPoint for Mac & Apple Keynote)** and **Windows PowerPoint**.

## Core Design Spec (5-Slide Short Deck Default)

Unless the user explicitly requests a longer formal defense/conference deck, always generate a **≤ 5-slide short presentation** following these strict visual and structural rules:

1. **No Start Cover or Ending Filler Slides**:
   - Do **not** generate a giant title cover slide or a "Thank You / Q&A / Conclusion" ending slide.
   - **Slide 1** uses the same minimalist header & card grammar as the rest of the deck, dedicated to **Paper Background & Core Positioning (`论文背景专题页`)** (with an optional compact `right_badge` pill in the top header row for venue/citation info).
2. **One Clear Data Display + Few-Sentence Summary (Slides 2–5)**:
   - Each subsequent slide presents **strictly one clear experiment/dataset result + a 2–3 sentence summary card**, so the audience grasps what was tested and why it works within 3 seconds.
3. **5 Cohesive Yet Distinct Slide Archetypes**:
   - Showcase visual variety without breaking style consistency by pairing Slide 1 (`add_background_slide`) with the 4 complementary data+summary layouts (`add_table_summary_slide`, `add_figure_summary_slide`, `add_hybrid_table_chart_slide`, `add_kpi_chart_summary_slide`). See [paper-deck-patterns.md](./references/paper-deck-patterns.md) for layout wireframes.
4. **Minimalist Canvas & Soft Rounded Cards**:
   - **Empty Bottom Margin**: Leave the bottom `~0.85 in` completely blank — zero bottom text, zero page counters, zero bottom rules, and zero top decorative bars (`verify_deck.py` flags any shape reaching `y > 7.00 in` as an error).
   - **Soft Light Rounded Rectangles Only**: Never attach dark vertical or top accent strips to cards, and disable shape drop shadows/hard borders. Use theme tokens `bg_tint` (`#EEF4FB`), `bg_subtle` (`#F4F6F9`), and `bg_card` (`#F8FAFC`).
   - **Bilingual (`zh-en`, default)**: Concise Chinese for assertion titles and summaries; crisp English for model names, benchmarks, metrics, and Unicode math (`Q(s_t, c^i) = ⟨E_s(s_t), E_a(c^i)⟩`, `ρ_t(i)`, `±`). Never place raw LaTeX `$...$` strings inside PowerPoint textboxes.

## Built-in Themes

Default to **`classic-conference`** unless another theme is requested:

| Theme ID | Palette & Surface Tokens | Best For |
| :--- | :--- | :--- |
| `classic-conference` *(default)* | Pure White (`#FFFFFF`) + Deep Navy (`#1B365D`) + Soft Ice (`#EEF4FB`) | Group meetings, ICLR/NeurIPS/ACL paper shares, clean benchmark reports |
| `slate-analytics` | Pure White (`#FFFFFF`) + Cool Slate (`#1E293B`) + Mint Tint (`#EBF6F4`) | Systems papers, throughput/latency dashboards, multi-metric comparisons |
| `warm-monograph` | Warm Ivory (`#FCFCFA`) + Oxford Burgundy (`#7A1C2E`) + Rose Tint (`#F9EFF1`) | Long-context NLP, theoretical + empirical analyses, seminar readings |

Theme tokens live in [resources/themes/](./resources/themes/) and master `.pptx` files in [resources/templates/](./resources/templates/). Run `uv run .agents/skills/academic-ppt/scripts/build_templates.py` after editing any theme JSON.

## Workflow

### 1. Extract Exact Paper Data First (Outside the Skill Directory)
- Extract exact numbers, `mean ± std`, sample sizes, hardware costs, and figures from the paper's `.tex`, CSVs, or PDF. Never fabricate or round metrics arbitrarily.
- Keep all paper PDFs, extracted CSVs, charts, and generated `.pptx` decks in the user's project workspace — **never store paper-specific data inside `.agents/skills/academic-ppt/`**.

### 2. Build the 5-Slide Deck via `academic_ppt_core.py`
Use [examples/build_5slide_template_demo.py](./examples/build_5slide_template_demo.py) (which produces [examples/5slide_template_demo.pptx](./examples/5slide_template_demo.pptx)) as the reference starter. Create a PEP 723 generator script importing `AcademicDeck` and `setup_academic_mpl` from [academic_ppt_core.py](./scripts/academic_ppt_core.py), and run with `uv` (identical on macOS & Windows):

```bash
uv run path/to/generate_deck.py --theme classic-conference
```

#### High-Level 5-Slide Template API (`AcademicDeck`)
1. **Slide 1 — Background Topic**: `deck.add_background_slide(section_tag, title, bottlenecks=[(route, tag, desc), ...], core_heading, core_points, highlights_heading, highlights_points, right_badge=...)`
2. **Slide 2 — Full-Width Table + Bottom Summary**: `deck.add_table_summary_slide(section_tag, title, headers, rows, kpi=(metric, label, sub), summary_heading, summary_lines, highlight_rows={...})`
3. **Slide 3 — Left Flagship Figure + Right Stacked Summary**: `deck.add_figure_summary_slide(section_tag, title, image_path, kpi=(metric, label, sub), summary_heading, summary_lines)`
4. **Slide 4 — Left Compact Table & Summary + Right Chart**: `deck.add_hybrid_table_chart_slide(section_tag, title, headers, rows, summary_heading, summary_lines, image_path, highlight_rows={...})`
5. **Slide 5 — Top 3-KPI Strip + Left Summary & Right Wide Chart**: `deck.add_kpi_chart_summary_slide(section_tag, title, kpis=[...], summary_heading, summary_lines, image_path)`

Low-level primitives (`add_header`, `add_booktabs_table`, `add_callout_card`, `add_kpi_strip`, `add_image_panel`) remain available for custom compositions. Read [cross-platform-guide.md](./references/cross-platform-guide.md) for OOXML dual-font (`Arial` + `Microsoft YaHei`) and 300-DPI figure rules.

### 3. Mandatory Verification (`verify_deck.py`)
Audit every generated `.pptx` and render slide PNGs before finishing:

```bash
uv run .agents/skills/academic-ppt/scripts/verify_deck.py path/to/output.pptx --render-dir path/to/preview_pngs --dpi 180
```

1. Confirm `error_count == 0` and `warning_count == 0` (dual-script `<a:latin>` + `<a:ea>` font tags, canvas bounds, zero sibling overlap, zero raw LaTeX).
2. Visually inspect every rendered PNG (`slide_01.png` .. `slide_05.png`) via `read_file` to verify balanced whitespace, clean line wrapping, non-colliding chart labels/error bars, and completely empty bottom margins.
