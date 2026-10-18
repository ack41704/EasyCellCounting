"""Finding, pairing and loading images."""
from __future__ import annotations

from pathlib import Path
import re

import cv2
import numpy as np

IMAGE_EXTENSIONS = (".tif", ".tiff", ".png", ".jpg", ".jpeg")


def natural_key(name: str):
    """Sort key that orders "s2" before "s10"."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", str(name))]


def list_images(folder) -> list:
    folder = Path(folder)
    if not folder.is_dir():
        raise FileNotFoundError(f"not a folder: {folder}")
    paths = [p for p in folder.iterdir()
             if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS and not p.name.startswith(".")]
    return sorted(paths, key=lambda p: natural_key(p.name))


def pair_images(full_dir, preview_dir=None) -> list:
    """[(full_res_path, preview_path or None)] in natural order.

    Previews are matched to full-resolution images by file stem, so both folders may
    use different formats (e.g. TIFF and PNG)."""
    full = list_images(full_dir)
    if not full:
        raise FileNotFoundError(f"no images ({', '.join(IMAGE_EXTENSIONS)}) in {full_dir}")
    if preview_dir is None:
        return [(p, None) for p in full]
    previews = {p.stem: p for p in list_images(preview_dir)}
    missing = [p.name for p in full if p.stem not in previews]
    if missing:
        raise FileNotFoundError(f"{len(missing)} image(s) have no preview in {preview_dir}: "
                                + ", ".join(missing[:10]))
    return [(p, previews[p.stem]) for p in full]


def load_image(path) -> np.ndarray:
    """8-bit BGR image (16-bit files are scaled to 8 bit by OpenCV)."""
    img = cv2.imread(str(path), cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError(f"could not read image: {path}")
    return img


def make_preview(image: np.ndarray, max_width: int = 1200, max_height: int = 900) -> np.ndarray:
    """Down-scaled copy that fits on screen (never up-scaled)."""
    h, w = image.shape[:2]
    s = min(max_width / w, max_height / h, 1.0)
    if s == 1.0:
        return image.copy()
    return cv2.resize(image, (max(1, int(w * s)), max(1, int(h * s))), interpolation=cv2.INTER_AREA)
