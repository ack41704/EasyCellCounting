"""
Regenerate the README figures from synthetic data (no real images are needed).

    pip install -e ".[figures]"
    python docs/make_figures.py

Writes docs/figures/method_overview.png and docs/figures/cluster_benchmark.png.
"""
from pathlib import Path
import sys

import cv2
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from easycellcounting import CountParams, annotate, count_cells
from easycellcounting.synthetic import synthetic_section

HERE = Path(__file__).resolve().parent
FIG = HERE / "figures"
sys.path.insert(0, str(HERE.parent / "examples"))
from benchmark_synthetic import run, summary  # noqa: E402

plt.rcParams.update({"font.family": "sans-serif", "font.size": 9, "axes.titlesize": 10,
                     "axes.titleweight": "bold", "axes.spines.top": False,
                     "axes.spines.right": False, "savefig.dpi": 200})


def rgb(bgr):
    return cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)


def method_overview():
    img, truth = synthetic_section(14, 11, clusters=(4, 6), size=(520, 760), seed=7)
    params = CountParams()
    res = count_cells(img, params)
    fig, ax = plt.subplots(1, 4, figsize=(11.5, 2.9), gridspec_kw={"width_ratios": [1, 1, 1, 0.9]})
    ax[0].imshow(rgb(img))
    ax[0].set_title("a  ROI (red channel)")
    ax[1].imshow(res.mask, cmap="gray")
    ax[1].set_title(f"b  threshold ≥ {params.threshold}")
    ax[2].imshow(rgb(annotate(img, res, label=False, thickness=2)))
    ax[2].set_title("c  singles / clusters")
    h, w = img.shape[:2]
    for text, x in ((f"L = {res.left}", w * 0.04), (f"R = {res.right}", w * 0.54)):
        ax[2].text(x, h * 0.08, text, color="white", fontsize=9, fontweight="bold")
    for a in ax[:3]:
        a.set_xticks([]), a.set_yticks([])
        for s in a.spines.values():
            s.set_visible(False)
    areas = np.array([o.area for o in res.objects if o.kind == "single"])
    cl = [o for o in res.objects if o.kind == "cluster"]
    top = max([o.area for o in cl] + [params.max_area]) * 1.12
    bins = np.arange(0, top + 50, 50)
    ax[3].hist(areas, bins=bins, color="#3a9d5d", label="single cells")
    ax[3].hist([o.area for o in cl], bins=bins, color="#2c6fbb", label="clusters")
    ymax = ax[3].get_ylim()[1]
    for x, name, c in ((params.min_area, "min_area", "0.4"), (params.max_area, "max_area", "0.4"),
                       (res.reference_area, "A_ref (p99)", "#c0392b")):
        ax[3].axvline(x, color=c, ls="--" if c == "0.4" else "-", lw=1)
        ax[3].text(x, ymax * 0.97, name + " ", fontsize=7, va="top", ha="right", rotation=90, color=c)
    for o in cl:
        ax[3].text(o.area, 1.4, f"{o.cells}", ha="center", fontsize=8, color="#2c6fbb", fontweight="bold")
    ax[3].set_xlim(0, top)
    ax[3].legend(frameon=False, fontsize=7, loc="center right")
    ax[3].set_xlabel("contour area (px²)")
    ax[3].set_ylabel("objects")
    ax[3].set_title("d  area classes")
    fig.text(0.01, -0.02, f"Synthetic section, truth L = {truth['left']}, R = {truth['right']}; "
             f"clusters ({', '.join(str(o.cells) for o in cl)} cells estimated) are outlined blue.",
             fontsize=7.5, color="0.35")
    fig.tight_layout()
    fig.savefig(FIG / "method_overview.png", bbox_inches="tight")
    plt.close(fig)


def cluster_benchmark():
    df = run(seeds=25)
    s = summary(df)
    fig, ax = plt.subplots(figsize=(3.6, 3.0))
    ks = np.array(sorted(df.cluster_size.unique()))
    jitter = np.random.default_rng(0).uniform(-0.15, 0.15, len(df))
    ax.scatter(df.cluster_size + jitter, df.cluster_estimate, s=8, alpha=0.35, color="#2c6fbb", lw=0)
    ax.plot(ks, s["cluster estimate, mean"], "o-", color="#2c6fbb", ms=4, label="estimate (mean)")
    ax.plot([1.5, 7.5], [1.5, 7.5], color="0.5", ls="--", lw=1, label="identity")
    ax.axvspan(1.5, 3.5, color="0.9", zorder=0)
    ax.text(2.5, 6.6, "group area\nnear / below\nmax_area:\ncounted as one", ha="center", fontsize=7, color="0.35")
    ax.set_xlabel("cells in touching group")
    ax.set_ylabel("cells estimated")
    ax.set_xticks(ks)
    ax.legend(frameon=False, fontsize=7, loc="lower right")
    ax.set_title("Cluster estimate (synthetic)")
    fig.tight_layout()
    fig.savefig(FIG / "cluster_benchmark.png", bbox_inches="tight")
    plt.close(fig)
    return s


if __name__ == "__main__":
    FIG.mkdir(parents=True, exist_ok=True)
    method_overview()
    print(cluster_benchmark().to_string())
    print(f"figures written to {FIG}")
