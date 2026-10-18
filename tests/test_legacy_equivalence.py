"""The package reproduces the counting loop of legacy/EasyCellCounting.py.

`legacy_count` is a line-by-line port of the second loop of the original script
(threshold `r in range(bright, 255)`, singles 40 < A < 900, clusters 900 < A <
cluster_max, floor(A / p99) + 1 cells per cluster when there are > 3 singles,
left / right by the centre of the minimum enclosing circle). The per-pixel PIL loop
is replaced by the identical vectorised condition."""
import cv2
import numpy as np
import pytest

from easycellcounting.counting import CountParams, count_cells
from easycellcounting.synthetic import synthetic_section


def legacy_count(cropped_image, bright=160, min_area=40, cluster_max=10000):
    r = cropped_image[:, :, 2]
    img = np.where((r >= bright) & (r < 255), 76, 0).astype(np.uint8)   # gray of (255, 0, 0)
    cnts = cv2.findContours(img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]
    middle = cropped_image.shape[1] / 2
    max_area = 900
    left, right, area_list = [], [], []
    for c in cnts:
        area = cv2.contourArea(c)
        if area > min_area and area < max_area:
            (x, y), radius = cv2.minEnclosingCircle(c)
            if x < middle:
                left.append(c)
            elif x > middle:
                right.append(c)
            area_list.append(area)
    if len(area_list) > 3:
        biggest = np.percentile(area_list, 99)
        for c1 in cnts:
            area = cv2.contourArea(c1)
            if area > max_area and area < cluster_max:
                (x, y), radius = cv2.minEnclosingCircle(c1)
                check = int((area / biggest) + 1)
                if x < middle:
                    left += [c1] * check
                elif x > middle:
                    right += [c1] * check
    return len(left), len(right)


@pytest.mark.parametrize("seed", range(8))
@pytest.mark.parametrize("bright", [130, 160, 200])
def test_same_counts_as_original(seed, bright):
    # no saturated pixels: the original treated 255 as background (fixed here)
    img, _ = synthetic_section(15 + seed, 25 - seed, clusters=(3, 4, 6)[: 1 + seed % 3],
                               intensity_range=(170, 225), noise=8, seed=seed)
    assert img[:, :, 2].max() < 255
    res = count_cells(img, CountParams(threshold=bright))
    assert (res.left, res.right) == legacy_count(img, bright)


def test_original_drops_saturated_cores():
    img = np.zeros((200, 200, 3), np.uint8)
    cv2.circle(img, (60, 100), 14, (0, 0, 255), -1)     # fully saturated cell
    assert legacy_count(img) == (0, 0)
    assert count_cells(img).left == 1
