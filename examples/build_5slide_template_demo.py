#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "python-pptx>=1.0.2",
#   "matplotlib>=3.9.0",
#   "pillow>=10.4.0",
#   "numpy>=2.0.0",
#   "lxml>=5.2.0",
# ]
# ///
"""
Minimal 5-Slide Short-Deck Template Demo for `academic-ppt`.

Demonstrates basic usage of all 5 canonical slide archetypes without any specific paper data:
  - Slide 1: `add_background_slide`        (Paper Background Topic Slide — no title cover)
  - Slide 2: `add_table_summary_slide`     (Full-Width Booktabs Table + Bottom KPI & Summary)
  - Slide 3: `add_figure_summary_slide`    (Left Scaling Figure + Right Stacked KPI & Summary)
  - Slide 4: `add_hybrid_table_chart_slide`(Left Compact Table & Summary + Right Bar Chart)
  - Slide 5: `add_kpi_chart_summary_slide` (Top 3-KPI Strip + Bottom Summary & Wide Ablation Chart)
"""

from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

import numpy as np
from pptx.enum.text import PP_ALIGN

EXAMPLE_DIR = Path(__file__).resolve().parent
SKILL_ROOT = EXAMPLE_DIR.parent
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from academic_ppt_core import AcademicDeck, setup_academic_mpl  # noqa: E402


def _render_demo_charts(deck: AcademicDeck, tmp_dir: Path) -> dict[str, Path]:
    plt = setup_academic_mpl(deck.theme)
    c = deck.colors
    pal = c["chart_palette"]
    bg_panel = c.get("bg_card", "#F8FAFC")
    paths: dict[str, Path] = {}

    # 1. Slide 3 Demo Figure: Scaling Curve Across Input Lengths
    fig, ax = plt.subplots(figsize=(6.4, 4.4), dpi=300, facecolor=bg_panel)
    ax.set_facecolor(bg_panel)
    scales = ["1K", "4K", "32K", "128K", "1M"]
    x = np.arange(len(scales))
    ours = [97.5, 97.2, 96.8, 96.4, 95.8]
    base_c = [94.0, 91.5, 84.0, 72.5, 48.0]
    base_b = [91.0, 82.0, 61.0, 39.0, 24.0]
    base_a = [68.0, 62.0, 49.0, 35.0, 22.0]

    ax.plot(x, ours, marker="o", lw=2.6, color=pal[0], label="Ours (本文方法)")
    ax.plot(x, base_c, marker="s", lw=2.1, color=pal[1], label="Baseline-C (大模型微调)")
    ax.plot(x, base_b, marker="^", lw=2.0, linestyle="--", color=pal[2], label="Baseline-B (重排基线)")
    ax.plot(x, base_a, marker="d", lw=1.8, linestyle=":", color="#94A3B8", label="Baseline-A (传统启发式)")

    ax.set_xticks(x)
    ax.set_xticklabels(scales, fontsize=9.8, fontweight="bold")
    ax.set_xlabel("Input Context / Problem Scale", fontsize=10, fontweight="bold")
    ax.set_ylabel("Accuracy / F1 (%)", fontsize=10, fontweight="bold")
    ax.set_title("跨尺度性能衰减对比曲线 (Scaling Curve)", fontsize=11, fontweight="bold", pad=10)
    ax.set_ylim(15, 105)
    ax.grid(True, alpha=0.5)
    ax.legend(frameon=False, fontsize=9.0, loc="lower left")
    fig.tight_layout()
    p_scale = tmp_dir / "demo_scaling.png"
    fig.savefig(p_scale, dpi=300, facecolor=bg_panel)
    plt.close(fig)
    paths["scaling"] = p_scale

    # 2. Slide 4 Demo Chart: In-Domain vs Out-of-Domain Grouped Bars
    fig, ax = plt.subplots(figsize=(5.3, 4.75), dpi=300, facecolor=bg_panel)
    ax.set_facecolor(bg_panel)
    methods = ["Baseline-A", "Baseline-B\n(1×GPU)", "Baseline-C\n(8×GPU)", "Ours\n(1×GPU)"]
    id_scores = [0.58, 0.79, 0.71, 0.78]
    ood_scores = [0.34, 0.45, 0.56, 0.64]

    x_m = np.arange(len(methods))
    w = 0.34
    b1 = ax.bar(x_m - w / 2, id_scores, width=w, label="In-Domain", color="#94A3B8", alpha=0.85)
    b2 = ax.bar(x_m + w / 2, ood_scores, width=w, label="Out-of-Domain (OOD)", color=pal[0], alpha=0.95)
    b2[-1].set_color(pal[1])

    for bars in (b1, b2):
        for bar in bars:
            h = bar.get_height()
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                h + 0.015,
                f"{h:.2f}",
                ha="center",
                va="bottom",
                fontsize=9.3,
                fontweight="bold",
                color=c["text_main"],
            )

    ax.set_ylabel("Evaluation Metric (F1)", fontsize=10, fontweight="bold")
    ax.set_title("域内表现 vs. 零样本跨分布泛化 (OOD)", fontsize=11, fontweight="bold", pad=12)
    ax.set_xticks(x_m)
    ax.set_xticklabels(methods, fontsize=9.2, fontweight="bold")
    ax.set_ylim(0, 0.98)
    ax.grid(axis="y", alpha=0.5)
    ax.legend(frameon=False, fontsize=9.0, loc="upper left")
    fig.tight_layout()
    p_hybrid = tmp_dir / "demo_hybrid.png"
    fig.savefig(p_hybrid, dpi=300, facecolor=bg_panel)
    plt.close(fig)
    paths["hybrid"] = p_hybrid

    # 3. Slide 5 Demo Chart: Wide Horizontal Ablation Error Bars
    fig, ax = plt.subplots(figsize=(6.9, 3.2), dpi=300, facecolor=bg_panel)
    ax.set_facecolor(bg_panel)
    variants = [
        "去掉核心模块 A (仅基础微调)",
        "去掉稳定器 B (高方差变体)",
        "去掉正则项 C (无探索奖励)",
        "Ours 完整方法 (A + B + C)",
    ]
    means = [42.5, 79.6, 91.2, 96.0]
    stds = [0.6, 18.4, 0.9, 0.2]
    colors = ["#CBD5E1", "#E2A03F", pal[1], pal[0]]

    y_pos = np.arange(len(variants))
    bars = ax.barh(
        y_pos,
        means,
        xerr=stds,
        height=0.55,
        color=colors,
        capsize=4.5,
        error_kw={"elinewidth": 1.5, "ecolor": "#475569"},
    )
    notes = ["42.5% ± 0.6%", "79.6% (±18.4)", "91.2% ± 0.9%", "96.0% ± 0.2%"]
    txt_colors = [c["text_main"], "#FFFFFF", "#FFFFFF", "#FFFFFF"]
    for idx, (bar, m) in enumerate(zip(bars, means)):
        x_txt = m + 2.5 if idx == 0 else 3.5
        ax.text(
            x_txt,
            bar.get_y() + bar.get_height() / 2,
            notes[idx],
            va="center",
            ha="left",
            fontsize=9.3,
            fontweight="bold",
            color=txt_colors[idx],
        )

    ax.set_yticks(y_pos)
    ax.set_yticklabels(variants, fontsize=9.2, fontweight="bold")
    ax.set_xlabel("评测指标得分 (%, 3 Seeds 均值 ± 标准差)", fontsize=9.5, fontweight="bold")
    ax.set_title("关键组件消融对比 (Ablation Study)", fontsize=10.8, fontweight="bold", pad=8)
    ax.set_xlim(0, 110)
    ax.grid(axis="x", alpha=0.5)
    fig.tight_layout()
    p_abl = tmp_dir / "demo_ablation.png"
    fig.savefig(p_abl, dpi=300, facecolor=bg_panel)
    plt.close(fig)
    paths["ablation"] = p_abl

    return paths


def build_template_demo(theme_name: str = "classic-conference", out_path: Path | None = None) -> Path:
    deck = AcademicDeck(theme_name=theme_name)

    with tempfile.TemporaryDirectory() as tmp:
        charts = _render_demo_charts(deck, Path(tmp))

        # Slide 1 — Template 1: Paper Background Topic Slide
        deck.add_background_slide(
            section_tag="01  ·  TEMPLATE 1 : PAPER BACKGROUND",
            right_badge="Venue 2026  ·  Paper ShortName",
            title="论文背景专题页：交代领域核心瓶颈与本文破局切入点（无独立封面页）",
            bottlenecks=[
                (
                    "既有路线 1 · 传统启发式",
                    "代表：Baseline-A  ·  免训练",
                    "推理阶段需遍历全局上下文，输入规模增长时延迟与计算开销急剧上升。",
                ),
                (
                    "既有路线 2 · 局部重排器",
                    "代表：Baseline-B  ·  1× GPU",
                    "每一步需重新执行跨块交互编码，长序列下显存与耗时难以扩展。",
                ),
                (
                    "既有路线 3 · 端到端训大模型",
                    "代表：Baseline-C  ·  ≥ 8× GPU",
                    "直接微调主干生成模型本体，算力门槛高且能力锁死在特定尺寸基座上。",
                ),
            ],
            core_heading="本文破局思路（Ours）：一句话点明核心方法范式转移",
            core_points=[
                "**核心机制设计**：冻结主干大模型，将复杂决策解耦至轻量表示空间（如双编码器内积 **f(s, a) = ⟨E_s(s), E_a(a)⟩**）。",
                "**关键结构创新**：引入平移不变的相对表征或正则机制，实现短序列训练向超大规模场景零损外推。",
            ],
            highlights_heading="核心特性速览",
            highlights_points=[
                "**1× GPU 轻量训练**：算力成本降低 80%+",
                "**10× 规模外推**：短序列训练零损泛化",
                "**即插即用解耦**：免重训适配任意主干模型",
            ],
        )

        # Slide 2 — Template 2: Full-Width Table + Bottom KPI & Summary
        deck.add_table_summary_slide(
            section_tag="02  ·  TEMPLATE 2 : FULL-WIDTH BENCHMARK TABLE",
            title="主实验结果（模板 2）：上方全宽标准三线表 + 下方核心指标与几句话总结",
            headers=["对比方法 (Method)", "Task A", "Task B", "Task C", "Extreme Scale", "Avg. Score"],
            rows=[
                ["Baseline-A (传统启发式)", "78.4", "74.2", "68.0", "--", "73.5"],
                ["Baseline-B (局部重排器)", "86.2", "82.5", "--", "--", "84.4"],
                ["Baseline-C (大模型微调)", "__91.5__", "__89.8__", "__86.4__", "79.2", "__86.7__"],
                ["**Ours (本文方法)**", "**96.8**", "**96.4**", "**95.9**", "**94.7**", "**96.0**"],
            ],
            col_widths=[3.35, 1.70, 1.70, 1.70, 1.85, 1.83],
            col_aligns=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 5,
            highlight_rows={3},
            kpi=("+9.3 pts 均分提升", "全尺度评测刷新 SOTA", "极端规模下仍保持 94.7%"),
            summary_heading="结果总结",
            summary_lines=[
                "适用于多列主实验对比表（精选 3–4 个代表性基线 + 本文方法）：表内自动将 **最优值深蓝加粗**、__次优值加下划线__、缺失值置灰，并在底部用 1–2 句话点明最关键的实验结论。"
            ],
        )

        # Slide 3 — Template 3: Left Flagship Figure + Right Stacked KPI & Summary
        deck.add_figure_summary_slide(
            section_tag="03  ·  TEMPLATE 3 : FIGURE SPOTLIGHT",
            title="趋势与扩展分析（模板 3）：左侧实验大图 + 右侧叠放关键数字与几句话总结",
            image_path=charts["scaling"],
            kpi=("95.8% @ 1M", "跨四个数量级近乎水平直线", "显著优于随规模退化的基线方法"),
            summary_heading="结果总结",
            summary_lines=[
                "适用于展示论文核心趋势折线图、Scaling 曲线或机制示意图，自动锁定原始宽高比居中嵌入浅色圆角容器。",
                "右上卡片提炼最醒目的单点指标，右下卡片用两句话解释曲线不衰减的内在原因。",
            ],
        )

        # Slide 4 — Template 4: Left Compact Table & Summary + Right Chart
        deck.add_hybrid_table_chart_slide(
            section_tag="04  ·  TEMPLATE 4 : TABLE & CHART HYBRID",
            title="泛化与开销对比（模板 4）：左上精简三线表 + 左下总结卡 + 右侧对比柱状图",
            headers=["对比方法", "训练算力", "In-Domain", "Out-of-Domain"],
            rows=[
                ["Baseline-A", "免训练", "0.58", "0.34"],
                ["Baseline-B", "1× GPU", "**0.79**", "0.45"],
                ["Baseline-C", "8× GPU", "0.71", "__0.56__"],
                ["**Ours (本文)**", "**1× GPU**", "__0.78__", "**0.64**"],
            ],
            col_widths=[2.10, 1.12, 1.55, 1.55],
            col_aligns=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.CENTER, PP_ALIGN.CENTER],
            highlight_rows={3},
            summary_heading="结果总结",
            summary_lines=[
                "适用于同时需要“精确数值查阅（左表）”与“直观视觉反差（右图）”的实验（如域内表现 vs. 零样本跨分布 OOD 泛化，或精度 vs. 训练算力对比）。"
            ],
            image_path=charts["hybrid"],
        )

        # Slide 5 — Template 5: Top 3-KPI Strip + Bottom Summary & Wide Chart
        deck.add_kpi_chart_summary_slide(
            section_tag="05  ·  TEMPLATE 5 : KPI STRIP & ABLATION",
            title="关键消融验证（模板 5）：顶部三联对比指标条 + 左下总结卡 + 右下宽幅图",
            kpis=[
                ("96.0% vs 42.5%", "核心模块 A 带来 2.3× 跃升", "去掉模块 A 后长程任务直接失效"),
                ("±0.2 vs ±18.4", "稳定器 B 消除训练崩溃", "跨随机种子标准差降低两个数量级"),
                ("+4.8 pts 增益", "探索正则项 C 进一步提分", "避免多步决策过早收敛至局部最优"),
            ],
            summary_heading="结果总结",
            summary_lines=[
                "适用于消融实验或敏感性分析：顶部三联卡片直接列出各关键组件的量化贡献。",
                "左下卡片用几句话回答“为什么每个组件不可或缺”，右下配带误差棒的宽幅对比图。",
            ],
            image_path=charts["ablation"],
        )

        target = out_path or (EXAMPLE_DIR / "5slide_template_demo.pptx")
        return deck.save(target)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build the generic 5-slide template demo deck.")
    parser.add_argument("--theme", default="classic-conference", help="Theme name.")
    parser.add_argument("--output", type=Path, default=None, help="Output .pptx path.")
    args = parser.parse_args()

    out = build_template_demo(args.theme, args.output)
    print(f"[OK] Built 5-slide template demo: {out}")


if __name__ == "__main__":
    main()
