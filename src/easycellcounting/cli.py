"""Command line: ``easycellcount select | count | run | demo``."""
from __future__ import annotations

import argparse
from pathlib import Path

from . import __version__
from .counting import CountParams


def _params(a) -> CountParams:
    return CountParams(threshold=a.threshold, min_area=a.min_area, max_area=a.max_area,
                       cluster_max_area=a.cluster_max_area, channel=a.channel,
                       cluster_rule=a.cluster_rule, smooth=a.smooth)


def _box(text):
    try:
        x, y, w, h = (int(v) for v in text.split(","))
    except ValueError:
        raise argparse.ArgumentTypeError("box must be x,y,width,height (pixels)")
    return x, y, w, h


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(prog="easycellcount", description=__doc__)
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = ap.add_subparsers(dest="command", required=True)

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("images", type=Path, help="folder of full-resolution images")
    g = common.add_argument_group("counting parameters (defaults = original script)")
    g.add_argument("--threshold", type=int, default=160, help="channel value counted as signal, 0-255 (160)")
    g.add_argument("--min-area", type=float, default=40.0, help="smallest single cell, px^2 (40)")
    g.add_argument("--max-area", type=float, default=900.0, help="largest single cell, px^2 (900)")
    g.add_argument("--cluster-max-area", type=float, default=10000.0,
                   help="largest object treated as a cluster, px^2 (10000)")
    g.add_argument("--channel", default="red", choices=["red", "green", "blue", "gray"])
    g.add_argument("--cluster-rule", default="floor_plus_one", choices=["floor_plus_one", "round"],
                   help="cells per cluster: floor(A/Aref)+1 (original) or round(A/Aref)")
    g.add_argument("--smooth", action="store_true", help="3x3 morphological open + close of the mask")

    out = argparse.ArgumentParser(add_help=False)
    o = out.add_argument_group("output")
    o.add_argument("--out", type=Path, default=Path("easycellcount_results"), help="output folder")
    o.add_argument("--animal-regex", default=r"\d+\.\d+",
                   help=r"regex for the animal tag in file names (default \d+\.\d+, e.g. 3730.4)")
    o.add_argument("--animal-slice", default=None,
                   help="character slice start:stop instead of a regex (the original used 6:13)")
    o.add_argument("--no-images", action="store_true", help="do not write annotated crops and masks")
    o.add_argument("--ignore-roi-params", action="store_true",
                   help="use the command-line threshold/areas for every image, not each ROI's preset")

    s = sub.add_parser("select", parents=[common], help="draw one ROI per image (interactive)")
    s.add_argument("--rois", type=Path, default=Path("rois.csv"), help="ROI file to write (rois.csv)")
    s.add_argument("--preview-dir", type=Path, default=None,
                   help="folder of down-scaled copies with the same file names (optional)")
    s.add_argument("--restart", action="store_true", help="ignore ROIs already in the ROI file")

    c = sub.add_parser("count", parents=[common, out], help="count cells (batch, no interaction)")
    where = c.add_mutually_exclusive_group()
    where.add_argument("--rois", type=Path, help="ROI file from `easycellcount select`")
    where.add_argument("--box", type=_box, help="one x,y,width,height box for every image")
    where.add_argument("--whole-image", action="store_true", help="count the whole image (default)")

    r = sub.add_parser("run", parents=[common, out], help="select ROIs, then count")
    r.add_argument("--rois", type=Path, default=Path("rois.csv"))
    r.add_argument("--preview-dir", type=Path, default=None)
    r.add_argument("--restart", action="store_true")

    d = sub.add_parser("demo", help="make a synthetic dataset and count it")
    d.add_argument("--out", type=Path, default=Path("easycellcount_demo"))
    return ap


def main(argv=None) -> int:
    a = build_parser().parse_args(argv)
    from .pipeline import count_folder

    if a.command == "demo":
        import pandas as pd
        from .synthetic import write_demo_dataset
        images = a.out / "images"
        truth = pd.DataFrame(write_demo_dataset(images))
        df = count_folder(images, None, CountParams(), a.out / "results")
        cmp = df[["file", "left", "right"]].merge(truth, on="file")
        print(cmp.to_string(index=False))
        return 0

    params = _params(a)
    if a.command in ("select", "run"):
        from .interactive import select_rois
        select_rois(a.images, a.rois, a.preview_dir, params, resume=not a.restart)
        if a.command == "select":
            return 0
    rois = a.rois if a.command == "run" else (a.rois or a.box)
    count_folder(a.images, rois, params, a.out, save_images=not a.no_images,
                 animal_pattern=a.animal_regex, animal_slice=a.animal_slice,
                 use_roi_params=not a.ignore_roi_params)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
