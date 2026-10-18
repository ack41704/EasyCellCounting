"""
Synthetic benchmark: how exactly does the algorithm recover a known number of cells?

Each image has a known number of isolated cells per side plus one touching group of
k cells (k = 2..7) on the left half. Writes benchmark_synthetic.csv and prints a summary table.
Synthetic data checks the implementation and shows the regime where the cluster rule
holds; it is not a substitute for validation against manual counts on real tissue.

    python examples/benchmark_synthetic.py [--seeds 25] [--out docs/benchmark]
"""
import argparse
from pathlib import Path

import pandas as pd

from easycellcounting import CountParams, count_cells
from easycellcounting.synthetic import synthetic_section


def run(seeds=25, ks=(2, 3, 4, 5, 6, 7), params=CountParams()):
    rows = []
    for k in ks:
        for seed in range(seeds):
            n_left, n_right = 15 + seed % 10, 25 - seed % 10
            img, truth = synthetic_section(n_left, n_right, clusters=(k,), seed=1000 * k + seed)
            res = count_cells(img, params)
            # the group sits on the left: the right half has isolated cells only
            rows.append({"cluster_size": k, "seed": seed,
                         "isolated_true": n_right, "isolated_found": res.right,
                         "cluster_estimate": res.left - n_left,
                         "reference_area_px": res.reference_area,
                         "left_true": truth["left"], "left_found": res.left,
                         "right_true": truth["right"], "right_found": res.right})
    return pd.DataFrame(rows)


def summary(df):
    d = df.assign(isolated_ok=df.isolated_found == df.isolated_true,
                  error=df.cluster_estimate - df.cluster_size)
    g = d.groupby("cluster_size")
    return pd.DataFrame({
        "images": g.size(),
        "isolated cells exact (% images)": 100 * g["isolated_ok"].mean(),
        "cluster estimate, mean": g["cluster_estimate"].mean().round(2),
        "exact (%)": 100 * g["error"].apply(lambda e: (e == 0).mean()),
        "within 1 cell (%)": 100 * g["error"].apply(lambda e: (e.abs() <= 1).mean()),
    })


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--seeds", type=int, default=25)
    ap.add_argument("--out", type=Path, default=Path("docs/benchmark"))
    a = ap.parse_args()
    df = run(a.seeds)
    a.out.mkdir(parents=True, exist_ok=True)
    df.to_csv(a.out / "benchmark_synthetic.csv", index=False)
    s = summary(df)
    s.to_csv(a.out / "benchmark_summary.csv")
    print(s.to_string())
