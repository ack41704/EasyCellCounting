"""
Regions of interest: geometry, rotation, preview <-> full-resolution scaling and the
ROI file that makes a counting run reproducible.

An ROI is stored in **full-resolution pixel coordinates of the rotated image**:
the whole image is first rotated by ``rotation`` degrees about its centre
(counter-clockwise, canvas size kept, as ``cv2.warpAffine`` does), then the box
``(x, y, width, height)`` is cropped. Saving the ROIs to a CSV means a dataset can be
recounted with new parameters without drawing anything again.
"""
from __future__ import annotations

import csv
from dataclasses import asdict, dataclass, fields
from pathlib import Path

import cv2
import numpy as np

from .counting import CountParams

ROI_COLUMNS = ("file", "x", "y", "width", "height", "rotation", "threshold", "min_area",
               "max_area", "cluster_max_area", "preset")


@dataclass
class ROI:
    file: str
    x: int
    y: int
    width: int
    height: int
    rotation: float = 0.0
    threshold: int = 160
    min_area: float = 40.0
    max_area: float = 900.0
    cluster_max_area: float = 10000.0
    preset: str = "default"

    @property
    def box(self):
        return self.x, self.y, self.width, self.height

    def params(self, base: CountParams = CountParams()) -> CountParams:
        """`base` with this ROI's per-image threshold and area limits."""
        return base.replace(threshold=int(self.threshold), min_area=float(self.min_area),
                            max_area=float(self.max_area),
                            cluster_max_area=float(self.cluster_max_area))


def rotate_image(image: np.ndarray, angle: float) -> np.ndarray:
    """Rotate about the image centre by `angle` degrees (counter-clockwise), keeping
    the canvas size (corners are cropped), as in the original script."""
    if not angle % 360:
        return image
    centre = (image.shape[1] / 2.0, image.shape[0] / 2.0)
    m = cv2.getRotationMatrix2D(centre, angle, 1.0)
    return cv2.warpAffine(image, m, (image.shape[1], image.shape[0]), flags=cv2.INTER_LINEAR)


def scale_box(box, from_shape, to_shape):
    """Map an (x, y, w, h) box between two images of the same field of view
    (x scales with the width ratio, y with the height ratio)."""
    x, y, w, h = box
    sx = to_shape[1] / from_shape[1]
    sy = to_shape[0] / from_shape[0]
    x1, y1 = int(x * sx), int(y * sy)
    x2, y2 = int((x + w) * sx), int((y + h) * sy)
    return x1, y1, x2 - x1, y2 - y1


def clip_box(box, shape):
    x, y, w, h = (int(v) for v in box)
    x0, y0 = max(0, x), max(0, y)
    x1, y1 = min(shape[1], x + w), min(shape[0], y + h)
    if x1 <= x0 or y1 <= y0:
        raise ValueError(f"ROI {box} lies outside an image of size {shape[1]}x{shape[0]}")
    return x0, y0, x1 - x0, y1 - y0


def crop(image: np.ndarray, roi: ROI) -> np.ndarray:
    """The ROI's pixels from a full-resolution image."""
    img = rotate_image(image, roi.rotation)
    x, y, w, h = clip_box(roi.box, img.shape)
    return img[y:y + h, x:x + w]


def whole_image_roi(file: str, shape, params: CountParams = CountParams()) -> ROI:
    return ROI(file, 0, 0, int(shape[1]), int(shape[0]), 0.0, params.threshold, params.min_area,
               params.max_area, params.cluster_max_area, "whole_image")


def save_rois(rois, path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=ROI_COLUMNS)
        w.writeheader()
        for r in rois:
            w.writerow(asdict(r))
    return path


def load_rois(path) -> list:
    types = {f.name: f.type for f in fields(ROI)}
    conv = {"int": lambda v: int(float(v)), "float": float, "str": str}
    out = []
    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            out.append(ROI(**{k: conv[types[k]](v) for k, v in row.items()
                              if k in types and v not in (None, "")}))
    return out
