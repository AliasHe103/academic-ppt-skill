# 5-Slide Short Deck Patterns (Group-Meeting & Idea-Sharing Spec)

## 1. Visual & Structural Philosophy

1. **Zero Ceremonial Filler (`无封面页、无致谢结尾页`)**:
   - Group meetings value fast comprehension over ceremonial title or "Thank You" slides.
   - Start immediately on **Slide 1 (`Paper Background`)** using the exact same top header structure (`section_tag` + `21.5 pt` assertion `title` + subtle `0.85 pt` rule) as Slides 2–5. Place compact paper/venue metadata in `right_badge` (e.g., `"ICLR 2026 Oral · Q-RAG"`).
2. **One Slide = One Clear Experiment + Few-Sentence Summary (`一页一数据 + 几句话总结`)**:
   - Keep **4–6 foreground elements max** per slide and `11.8–13.0 pt` body/table font sizes.
   - Never cram full 15-row appendix tables onto a group-meeting slide; select **3–4 representative baselines + Ours** that isolate the exact scientific comparison.
3. **Clean Bottom Canvas & Soft Rounded Surfaces**:
   - Leave the bottom `~0.85 in` completely empty (no footer text, no page numbers, no bottom divider).
   - Use borderless, shadow-free rounded rectangles (`radius_in = 0.15–0.16 in` with aspect-corrected corner adjustment) across three subtle surface hierarchy levels:
     - `bg_tint` (`#EEF4FB` in `classic-conference`): Primary highlight card / key metric strip / core idea box.
     - `bg_subtle` (`#F4F6F9`): Secondary narrative summary card / prior bottleneck card.
     - `bg_card` (`#F8FAFC`): High-DPI figure / chart container background (match `facecolor` in `setup_academic_mpl` charts so plots blend seamlessly).

---

## 2. The 5 Canonical Slide Layout Archetypes

All 5 slides share the identical header (`y = 0.28..1.20 in`) and content safe zone (`x = 0.60..12.733 in`, `y = 1.52..6.64 in`), while providing **5 distinct layout geometries**:

### Template 1: Paper Background & Core Positioning (`add_background_slide`)
* **When to use**: Always **Slide 1**. Establishes why existing approaches struggle and how the paper reframes the problem.
* **Geometry**:
  ```text
  +-----------------------------------------------------------------------+
  | 01 · PAPER BACKGROUND                     [ ICLR 2026 Oral · Paper ]  |
  | Assertion Title Stating Domain Bottleneck & Core Shift                |
  | --------------------------------------------------------------------- |
  | +---------------------+ +---------------------+ +-------------------+ |
  | | Prior Route 1       | | Prior Route 2       | | Prior Route 3     | |
  | | Rep. Work · Cost    | | Rep. Work · Cost    | | Rep. Work · Cost  | |
  | | Key limitation...   | | Key limitation...   | | Key limitation... | |
  | +---------------------+ +---------------------+ +-------------------+ |
  | +-------------------------------------------+ +---------------------+ |
  | | Paper Core Idea (Soft Tinted Card)        | | Key Advantages      | |
  | | • Mechanism & formulation shift...        | | • Compute savings   | |
  | | • Why it overcomes the bottlenecks...     | | • Scale / Plug&Play | |
  | +-------------------------------------------+ +---------------------+ |
  +-----------------------------------------------------------------------+
  ```

### Template 2: Full-Width Benchmark Table + Bottom Summary (`add_table_summary_slide`)
* **When to use**: Multi-column main evaluation tables (`5–7` columns, `4–5` rows) such as length extrapolation or multi-task benchmarks.
* **Geometry**:
  ```text
  +-----------------------------------------------------------------------+
  | 02 · RESULT 1 : BENCHMARK TABLE                                       |
  | Assertion Title with Headline Quantitative Gain                       |
  | --------------------------------------------------------------------- |
  | ===================================================================== |
  | Method             Metric A    Metric B    Metric C    Metric D       |
  | --------------------------------------------------------------------- |
  | Baseline 1           ...         ...         ...         ...          |
  | Baseline 2           ...         ...         ...         ...          |
  | Ours (Highlighted) **Best**    **Best**    **Best**    **Best**       |
  | ===================================================================== |
  | +--------------------+ +--------------------------------------------+ |
  | | Headline KPI       | | Concise 2-3 Sentence Result Summary        | |
  | +--------------------+ +--------------------------------------------+ |
  +-----------------------------------------------------------------------+
  ```

### Template 3: Left Flagship Figure + Right Stacked Summary (`add_figure_summary_slide`)
* **When to use**: Scaling curves across context/compute, training dynamics, or high-resolution paper figures.
* **Geometry**:
  ```text
  +-----------------------------------------------------------------------+
  | 03 · RESULT 2 : SCALING CURVE                                         |
  | Assertion Title Describing Trend Across Scales                        |
  | --------------------------------------------------------------------- |
  | +----------------------------------------+ +------------------------+ |
  | |                                        | | Headline Metric Card   | |
  | |                                        | | (Soft Tinted)          | |
  | |     High-DPI Aspect-Locked Figure      | +------------------------+ |
  | |          (Soft Card Panel)             | +------------------------+ |
  | |                                        | | Few-Sentence Summary   | |
  | |                                        | | Explaining Mechanism   | |
  | +----------------------------------------+ +------------------------+ |
  +-----------------------------------------------------------------------+
  ```

### Template 4: Left Compact Table & Summary + Right Chart (`add_hybrid_table_chart_slide`)
* **When to use**: Pairing exact numerical comparisons (e.g., compute cost + ID vs. OOD scores) with a visual grouped bar chart.
* **Geometry**:
  ```text
  +-----------------------------------------------------------------------+
  | 04 · RESULT 3 : GENERALIZATION & EFFICIENCY                           |
  | Assertion Title Highlighting Cost-Performance / OOD Win               |
  | --------------------------------------------------------------------- |
  | +-----------------------------------+ +-----------------------------+ |
  | | === Compact 4x4 Booktabs Table == | |                             | |
  | | Method    Cost     ID      OOD    | |   300-DPI Grouped Bar or    | |
  | | ...       ...      ...     ...    | |   Scatter Comparison Plot   | |
  | +-----------------------------------+ |                             | |
  | +-----------------------------------+ |                             | |
  | | Soft-Tinted Summary Card          | |                             | |
  | +-----------------------------------+ +-----------------------------+ |
  +-----------------------------------------------------------------------+
  ```

### Template 5: Top 3-KPI Strip + Bottom Summary & Wide Chart (`add_kpi_chart_summary_slide`)
* **When to use**: Ablation studies (`mean ± std` error bars), component sensitivity, or inference efficiency trade-offs.
* **Geometry**:
  ```text
  +-----------------------------------------------------------------------+
  | 05 · RESULT 4 : ABLATION / INSIGHT                                    |
  | Assertion Title Isolating Critical Component Contribution             |
  | --------------------------------------------------------------------- |
  | +---------------------+ +---------------------+ +-------------------+ |
  | | KPI Delta 1         | | KPI Delta 2         | | KPI Delta 3       | |
  | +---------------------+ +---------------------+ +-------------------+ |
  | +---------------------------+ +-------------------------------------+ |
  | | Few-Sentence Ablation     | |   Wide-Aspect 300-DPI Chart         | |
  | | Summary Card              | |   (e.g., Horizontal Error-Bar Plot) | |
  | +---------------------------+ +-------------------------------------+ |
  +-----------------------------------------------------------------------+
  ```

---

## 3. Three-Line Table (`booktabs`) Formatting Rules

- **Zero vertical lines**: Separate columns via alignment (`PP_ALIGN.LEFT` for method names, `PP_ALIGN.CENTER` for metrics) and generous horizontal padding (`pad_lr_in = 0.08–0.10`).
- **Inline visual hierarchy**:
  - `**...**` renders **bold primary** for 1st-place scores.
  - `__...__` renders <u>underlined</u> for 2nd-place scores.
  - `highlight_rows={idx}` shades the proposed method's row with soft `bg_highlight_row` (no vertical left bar).
  - `--` or `n/a` automatically renders in muted gray (`text_muted`).
