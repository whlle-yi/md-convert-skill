#!/usr/bin/env python3
"""generate_chart.py — 重新生成 examples/demo-chart.png。

图表字体遵循项目排版规范：中文宋体、西文/数字 Times New Roman、
数学字体 STIX。黑白灰配色，与文档整体风格一致。

用法：python examples/generate_chart.py
"""

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

# matplotlib 对 .ttc 集合字体（SimSun）的字形加载不可靠；
# CJK 回退使用思源宋体（Noto Serif SC，宋体类衬线），西文/数字保持 Times New Roman
plt.rcParams["font.family"] = ["Times New Roman", "Noto Serif SC"]
plt.rcParams["mathtext.fontset"] = "stix"
plt.rcParams["axes.unicode_minus"] = False


def main() -> None:
    quarters = ["一季度", "二季度", "三季度", "四季度"]
    revenue = [1024, 1536, 2048, 2560]

    fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=200)
    bars = ax.bar(quarters, revenue, width=0.55, color="#404040", edgecolor="black", linewidth=0.6)

    for rect, value in zip(bars, revenue):
        ax.text(rect.get_x() + rect.get_width() / 2, value + 40, f"{value:,}",
                ha="center", va="bottom", fontsize=9)

    ax.set_title("示例：某产品 2026 年分季度营收（万元）", fontsize=12)
    ax.set_ylabel("营收（万元）", fontsize=10)
    ax.set_ylim(0, 3000)
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(labelsize=10)
    ax.yaxis.grid(True, linestyle=":", linewidth=0.6, color="#b0b0b0")
    ax.set_axisbelow(True)

    out = Path(__file__).resolve().parent / "demo-chart.png"
    fig.savefig(out, bbox_inches="tight")
    print(f"已生成 {out}")


if __name__ == "__main__":
    main()
