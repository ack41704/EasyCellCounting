# Changelog

## 1.0.0 (2026-10-18)

Packaged rewrite of the original script, for publication.

- `easycellcounting` Python package and `easycellcount` command (`select`, `count`, `run`, `demo`).
- Interactive ROI selection is a separate step. ROIs (box, rotation, per-image threshold
  and area limits) are saved to `rois.csv` after every image. Sessions can be resumed,
  and datasets recounted without redrawing.
- Same counting rules as the original by default. Equivalence is tested against a port of the original loop.
- Fixes:
  - the threshold now includes the saturated value 255;
  - the x/y ratios used to scale the ROI from preview to full image were swapped;
  - natural file sorting;
  - the `q` check no longer shifts later images' cluster caps.
- One input folder is enough; previews are made on the fly (`--preview-dir` still supported).
- Annotated ROI images and masks; tidy CSV; Excel workbook built without a template
  (left / right / total tables, region colours, parameters sheet).
- Options: colour channel, morphological smoothing, `round` cluster rule, animal-name regex or slice.
- Test suite, synthetic benchmark, figures, citation file.

## 0.1 (2025-05-30)

Original script `EasyCellCounting.py`, now in `legacy/`.
