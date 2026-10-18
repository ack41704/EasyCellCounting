"""Batch counting: images + ROIs -> counts table, annotated crops and an Excel workbook."""
from __future__ import annotations

from dataclasses import asdict
from pathlib import Path
import warnings

import cv2

from . import __version__
from .counting import CountParams, annotate, count_cells
from .io import list_images, load_image
from .report import DEFAULT_ANIMAL_PATTERN, results_table, write_csv, write_excel
from .roi import ROI, crop, load_rois, whole_image_roi


def _rois_for(paths, rois, params):
    """One ROI per image path (None = whole image)."""
    if rois is None:
        return {p.name: None for p in paths}
    if isinstance(rois, (str, Path)):
        rois = load_rois(rois)
    if isinstance(rois, tuple) and len(rois) == 4:
        box = tuple(int(v) for v in rois)
        return {p.name: ROI(p.name, *box, threshold=params.threshold, min_area=params.min_area,
                            max_area=params.max_area, cluster_max_area=params.cluster_max_area,
                            preset="fixed_box")
                for p in paths}
    by_name = {r.file: r for r in rois}
    unknown = sorted(set(by_name) - {p.name for p in paths})
    if unknown:
        raise FileNotFoundError(f"ROI file lists images that are not in the folder: {unknown[:10]}")
    skipped = [p.name for p in paths if p.name not in by_name]
    if skipped:
        warnings.warn(f"{len(skipped)} image(s) have no ROI and are skipped: {skipped[:5]}")
    return {name: r for name, r in by_name.items()}


def count_folder(image_dir, rois=None, params: CountParams = CountParams(), out_dir=None,
                 save_images: bool = True, animal_pattern: str | None = DEFAULT_ANIMAL_PATTERN,
                 animal_slice: str | None = None, use_roi_params: bool = True, verbose: bool = True):
    """Count every image in `image_dir`.

    rois:  None (whole image), a ROI CSV written by ``easycellcount select``, a list
           of ROI objects, or one (x, y, width, height) box used for every image.
    params: counting parameters; with `use_roi_params` each ROI's own threshold and
           area limits (chosen with the preset keys) take precedence.
    out_dir: when given, writes cell_counts.csv, cell_counts.xlsx and (with
           `save_images`) annotated/ and masks/ PNGs there.

    Returns the tidy results table (pandas.DataFrame)."""
    paths = list_images(image_dir)
    if not paths:
        raise FileNotFoundError(f"no images in {image_dir}")
    roi_map = _rois_for(paths, rois, params)
    out_dir = Path(out_dir) if out_dir else None
    if out_dir and save_images:
        (out_dir / "annotated").mkdir(parents=True, exist_ok=True)
        (out_dir / "masks").mkdir(parents=True, exist_ok=True)

    rows = []
    for p in paths:
        if p.name not in roi_map:
            continue
        image = load_image(p)
        roi = roi_map[p.name] or whole_image_roi(p.name, image.shape, params)
        pr = roi.params(params) if use_roi_params else params
        region = crop(image, roi)
        res = count_cells(region, pr)
        if out_dir and save_images:
            cv2.imwrite(str(out_dir / "annotated" / f"{p.stem}_annotated.png"), annotate(region, res))
            cv2.imwrite(str(out_dir / "masks" / f"{p.stem}_mask.png"), res.mask)
        r = asdict(roi)
        rows.append({"file": p.name, **res.summary(),
                     "roi_x": r["x"], "roi_y": r["y"], "roi_width": r["width"],
                     "roi_height": r["height"], "rotation_deg": r["rotation"],
                     "threshold": pr.threshold, "min_area": pr.min_area, "max_area": pr.max_area,
                     "cluster_max_area": pr.cluster_max_area, "channel": pr.channel,
                     "preset": r["preset"]})
        if verbose:
            print(f"{p.name}: left {res.left}  right {res.right}"
                  + (f"  ({res.n_clusters} cluster(s))" if res.n_clusters else ""))

    df = results_table(rows, animal_pattern, animal_slice)
    if out_dir:
        meta = {"software": f"EasyCellCounting {__version__}", "image_folder": str(image_dir),
                "rois": str(rois) if isinstance(rois, (str, Path, tuple)) else
                ("whole image" if rois is None else "list of ROIs"),
                **{f"default_{k}": v for k, v in asdict(params).items()},
                "per_image_roi_params": use_roi_params}
        write_csv(df, out_dir / "cell_counts.csv")
        write_excel(df, out_dir / "cell_counts.xlsx", meta)
        if verbose:
            print(f"wrote {out_dir / 'cell_counts.csv'} and cell_counts.xlsx")
    return df
