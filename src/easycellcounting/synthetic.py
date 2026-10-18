"""Synthetic fluorescence sections with a known number of cells (for tests and the demo)."""
from __future__ import annotations

from pathlib import Path

import cv2
import numpy as np


def synthetic_section(n_left: int = 25, n_right: int = 18, clusters=(2, 3),
                      size=(900, 1200), radius_range=(7, 11), intensity_range=(190, 245),
                      background: float = 25.0, noise: float = 12.0, seed: int = 0):
    """An 8-bit BGR image of red cells on a dark, noisy background.

    n_left / n_right: isolated cells on each half; clusters: sizes of touching groups
    (alternately placed left and right). Returns (image, truth) where truth gives the
    number of cells per side including those in clusters."""
    rng = np.random.default_rng(seed)
    h, w = size
    red = np.full((h, w), background, np.float32)
    taken = []
    margin = 2 * radius_range[1] + 4

    def free(x, y, r, gap):
        return all((x - a) ** 2 + (y - b) ** 2 > (r + c + gap) ** 2 for a, b, c in taken)

    def place(side, gap, chain=1, r=None, tries=5000):
        """Free position for `chain` cells in a vertical line (touching when chain > 1)."""
        x0, x1 = (margin, w // 2 - margin) if side == "left" else (w // 2 + margin, w - margin)
        for _ in range(tries):
            rr = r or rng.uniform(*radius_range)
            step = 1.6 * rr
            x = rng.uniform(x0, x1)
            y = rng.uniform(margin, h - margin - step * (chain - 1))
            spots = [(x, y + j * step, rr) for j in range(chain)]
            if all(free(a, b, c, gap) for a, b, c in spots):
                taken.extend(spots)
                return spots
        raise RuntimeError("image too crowded for the requested number of cells")

    def disk(x, y, r):
        cv2.circle(red, (int(round(x)), int(round(y))), int(round(r)),
                   float(rng.uniform(*intensity_range)), -1, cv2.LINE_AA)

    truth = {"left": n_left, "right": n_right}
    for side, n in (("left", n_left), ("right", n_right)):
        for _ in range(n):
            disk(*place(side, gap=8)[0])
    for i, k in enumerate(clusters):                          # touching groups of k cells
        side = "left" if i % 2 == 0 else "right"
        for spot in place(side, gap=10, chain=k, r=radius_range[1]):
            disk(*spot)
        truth[side] += k
    red = cv2.GaussianBlur(red, (0, 0), 1.2) + rng.normal(0, noise, red.shape)
    img = np.zeros((h, w, 3), np.uint8)
    img[:, :, 2] = np.clip(red, 0, 255).astype(np.uint8)
    img[:, :, 1] = np.clip(red * 0.15 + rng.normal(0, 4, red.shape), 0, 255).astype(np.uint8)
    return img, truth


def write_demo_dataset(folder, n_images: int = 6, seed: int = 1):
    """A small dataset (two animals, three sections each, region codes in the names)."""
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    rng = np.random.default_rng(seed)
    truth = []
    for i in range(n_images):
        animal = "3730.4" if i < n_images // 2 else "3731.2"
        section = i % (n_images // 2) + 1
        region = ("A1", "B2", "C3")[section % 3]
        name = f"Demo {animal} {region}_s{section}.tif"
        img, t = synthetic_section(int(rng.integers(10, 30)), int(rng.integers(10, 30)),
                                   clusters=(4, 5), seed=seed * 100 + i)
        cv2.imwrite(str(folder / name), img)
        truth.append({"file": name, "true_left": t["left"], "true_right": t["right"]})
    return truth
