"""
Core counting algorithm (no user interaction, no file I/O).

For one region of interest (ROI) of a fluorescence image:

1. **Threshold** one colour channel: a pixel is signal when its value is
   ``>= threshold`` (8-bit scale, 0-255).
2. **Segment** the binary mask into objects with ``cv2.findContours``
   (outer contours only, so holes inside a cell do not create extra objects).
3. **Classify** each object by its contour area *A* (pixels^2):

   * single cell      ``min_area < A < max_area``
   * cell cluster     ``max_area <= A < cluster_max_area``
   * ignored          everything else (noise, large artefacts)

4. **Estimate cells per cluster** from a reference single-cell area, the
   ``reference_percentile`` (99th) percentile of the single-cell areas in the
   same ROI: ``n = floor(A / A_ref) + 1``. Clusters are only expanded when the
   ROI has at least ``min_singles_for_clusters`` single cells to estimate
   ``A_ref`` from; otherwise they are not counted.
5. **Split left / right** at the vertical midline of the ROI, using the centre
   of each object's minimum enclosing circle.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
import math

import cv2
import numpy as np

CHANNELS = {"blue": 0, "green": 1, "red": 2}

# drawing colours (BGR), as in the original script
SINGLE_BGR = (0, 250, 0)
CLUSTER_BGR = (250, 0, 0)
MIDLINE_BGR = (0, 255, 255)


@dataclass(frozen=True)
class CountParams:
    """Counting parameters. Defaults reproduce the original EasyCellCounting.py."""

    threshold: int = 160
    min_area: float = 40.0
    max_area: float = 900.0
    cluster_max_area: float = 10000.0
    min_singles_for_clusters: int = 4
    reference_percentile: float = 99.0
    channel: str = "red"
    cluster_rule: str = "floor_plus_one"
    smooth: bool = False

    def __post_init__(self):
        if self.channel not in (*CHANNELS, "gray"):
            raise ValueError(f"channel must be one of {sorted(CHANNELS)} or 'gray', not {self.channel!r}")
        if self.cluster_rule not in ("floor_plus_one", "round"):
            raise ValueError("cluster_rule must be 'floor_plus_one' or 'round'")
        if not 0 <= self.threshold <= 255:
            raise ValueError("threshold is on the 8-bit scale (0-255)")
        if not self.min_area < self.max_area <= self.cluster_max_area:
            raise ValueError("need min_area < max_area <= cluster_max_area")

    def replace(self, **changes) -> "CountParams":
        return CountParams(**{**asdict(self), **changes})


@dataclass
class CellObject:
    """One segmented object (a single cell or a cluster)."""

    contour: np.ndarray
    area: float
    x: float
    y: float
    side: str          # "left" | "right"
    kind: str          # "single" | "cluster"
    cells: int         # cells it contributes to the count (1 for a single cell)


@dataclass
class CountResult:
    left: int
    right: int
    single_left: int
    single_right: int
    cluster_cells_left: int
    cluster_cells_right: int
    n_clusters: int
    reference_area: float
    n_objects: int
    objects: list = field(default_factory=list, repr=False)
    mask: np.ndarray = field(default=None, repr=False)

    @property
    def total(self) -> int:
        return self.left + self.right

    def summary(self) -> dict:
        """Scalar fields only (one row of the results table)."""
        return {"left": self.left, "right": self.right, "total": self.total,
                "single_left": self.single_left, "single_right": self.single_right,
                "cluster_cells_left": self.cluster_cells_left,
                "cluster_cells_right": self.cluster_cells_right,
                "n_clusters": self.n_clusters, "reference_area_px": self.reference_area,
                "n_objects": self.n_objects}


def to_bgr8(image: np.ndarray) -> np.ndarray:
    """Any 2-D / 3-channel / 4-channel image as 8-bit BGR.

    16-bit data are scaled by 1/256, which is what ``cv2.imread`` does for 16-bit
    files, so thresholds mean the same thing as in the original script."""
    img = np.asarray(image)
    if img.dtype == np.uint16:
        img = (img >> 8).astype(np.uint8)
    elif img.dtype != np.uint8:
        img = np.clip(img, 0, 255).astype(np.uint8)
    if img.ndim == 2:
        return cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
    if img.shape[2] == 4:
        return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
    return img


def channel_of(image_bgr: np.ndarray, channel: str = "red") -> np.ndarray:
    img = to_bgr8(image_bgr)
    if channel == "gray":
        return cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    return np.ascontiguousarray(img[:, :, CHANNELS[channel]])


def threshold_mask(image_bgr: np.ndarray, threshold: int, channel: str = "red",
                   smooth: bool = False) -> np.ndarray:
    """uint8 mask, 255 where the channel is >= threshold."""
    mask = np.where(channel_of(image_bgr, channel) >= threshold, 255, 0).astype(np.uint8)
    if smooth:
        k = np.ones((3, 3), np.uint8)
        mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, k)
        mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, k)
    return mask


def cells_in_cluster(area: float, reference_area: float, rule: str = "floor_plus_one") -> int:
    """Estimated number of cells in a cluster of `area` given the single-cell area."""
    if rule == "round":
        return max(1, int(round(area / reference_area)))
    return int(area / reference_area) + 1


def count_cells(roi_bgr: np.ndarray, params: CountParams = CountParams()) -> CountResult:
    """Count cells left and right of the vertical midline of `roi_bgr`."""
    mask = threshold_mask(roi_bgr, params.threshold, params.channel, params.smooth)
    found = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = found[0] if len(found) == 2 else found[1]
    midline = mask.shape[1] / 2.0

    singles, clusters = [], []
    for c in contours:
        area = float(cv2.contourArea(c))
        if params.min_area < area < params.max_area:
            singles.append((c, area))
        elif params.max_area <= area < params.cluster_max_area:
            clusters.append((c, area))

    def obj(c, area, kind, cells):
        (x, y), _ = cv2.minEnclosingCircle(c)
        return CellObject(c, area, float(x), float(y), "left" if x < midline else "right", kind, cells)

    objects = [obj(c, a, "single", 1) for c, a in singles]
    ref = math.nan
    if len(singles) >= params.min_singles_for_clusters:
        ref = float(np.percentile([a for _, a in singles], params.reference_percentile))
        objects += [obj(c, a, "cluster", cells_in_cluster(a, ref, params.cluster_rule))
                    for c, a in clusters]

    def tally(side, kind):
        return sum(o.cells for o in objects if o.side == side and o.kind == kind)

    sl, sr = tally("left", "single"), tally("right", "single")
    cl, cr = tally("left", "cluster"), tally("right", "cluster")
    return CountResult(left=sl + cl, right=sr + cr, single_left=sl, single_right=sr,
                       cluster_cells_left=cl, cluster_cells_right=cr,
                       n_clusters=sum(o.kind == "cluster" for o in objects),
                       reference_area=ref, n_objects=len(contours), objects=objects, mask=mask)


def annotate(roi_bgr: np.ndarray, result: CountResult, label: bool = True,
             thickness: int = 2) -> np.ndarray:
    """Copy of the ROI with single cells outlined green, clusters blue (with their
    estimated cell number), the midline yellow and the counts written on top."""
    out = to_bgr8(roi_bgr).copy()
    h, w = out.shape[:2]
    for o in result.objects:
        colour = SINGLE_BGR if o.kind == "single" else CLUSTER_BGR
        cv2.drawContours(out, [o.contour], -1, colour, thickness)
        if o.kind == "cluster":
            cv2.putText(out, str(o.cells), (int(o.x), int(o.y)), cv2.FONT_HERSHEY_SIMPLEX,
                        0.6, CLUSTER_BGR, 2, cv2.LINE_AA)
    cv2.line(out, (w // 2, 0), (w // 2, h), MIDLINE_BGR, 1)
    if label:
        scale = max(0.5, min(w, h) / 900)
        for text, x in ((f"L {result.left}", 8), (f"R {result.right}", w // 2 + 8)):
            cv2.putText(out, text, (x, int(30 * scale)), cv2.FONT_HERSHEY_SIMPLEX, scale,
                        (255, 255, 255), 2, cv2.LINE_AA)
    return out
