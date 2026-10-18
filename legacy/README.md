# Original script (2025)

`EasyCellCounting.py` is the original script, exported from a Jupyter notebook. It is kept
unchanged so that analyses done with it can be reproduced exactly. For new work, use the
`easycellcounting` package. It applies the same counting rules by default, and
[`tests/test_legacy_equivalence.py`](../tests/test_legacy_equivalence.py) checks that its
counts match this script's.

## Running it

The script expects:

- full-resolution images in a folder (path set in `r_mypath`);
- down-scaled copies with the same names in a second folder (`mypath`);
- an empty workbook `input.xlsx` in the working directory.

It writes `cell_counts.xlsx`. Run it cell by cell in Jupyter, or with
`python EasyCellCounting.py` after editing the two folder paths. It needs `numpy`,
`pandas`, `opencv-python`, `Pillow` and `openpyxl`.

## Known issues (fixed in the package)

- Threshold `r in range(bright, 255)` treats saturated pixels (255) as background.
- The ROI is scaled from preview to full image with the height ratio applied to x and the
  width ratio applied to y.
- Files are sorted as plain strings, so `_s10` comes before `_s2`.
- Each `q` check appends an extra value to `cluster_max`, shifting the cluster cap used
  for later images. This only matters when `i` (cluster cap 10⁶) is also used.
- Pressing `u` on the first image raises an error. The final preview loop titles windows
  with the wrong file name.
