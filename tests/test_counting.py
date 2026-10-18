import cv2
import numpy as np
import pytest

from easycellcounting.counting import CountParams, cells_in_cluster, count_cells, to_bgr8
from easycellcounting.synthetic import synthetic_section


def blank(h=400, w=600):
    return np.zeros((h, w, 3), np.uint8)


def disk(img, x, y, r, value=220):
    cv2.circle(img, (x, y), r, (0, 0, value), -1)


@pytest.mark.parametrize("seed", range(5))
def test_isolated_cells_counted_exactly(seed):
    img, truth = synthetic_section(20 + seed, 15 + 2 * seed, clusters=(), seed=seed)
    res = count_cells(img)
    assert (res.left, res.right) == (truth["left"], truth["right"])
    assert res.n_clusters == 0


@pytest.mark.parametrize("seed", range(5))
def test_clusters_of_four_and_five_cells(seed):
    img, truth = synthetic_section(18, 22, clusters=(4, 5), seed=10 + seed)
    res = count_cells(img)
    assert res.n_clusters == 2
    # isolated cells are exact, each cluster estimate is within one cell
    assert (res.single_left, res.single_right) == (18, 22)
    assert abs(res.cluster_cells_left - 4) <= 1 and abs(res.cluster_cells_right - 5) <= 1


def test_area_classes():
    img = blank()
    disk(img, 100, 100, 2)        # ~13 px^2: below min_area -> ignored
    disk(img, 200, 100, 10)       # single
    disk(img, 450, 300, 70)       # ~15,000 px^2: above cluster_max_area -> ignored
    res = count_cells(img)
    assert (res.left, res.right) == (1, 0)
    assert res.n_objects == 3


def test_clusters_need_four_single_cells():
    img = blank()
    for x in (60, 120, 180):
        disk(img, x, 60, 10)
    disk(img, 450, 250, 25)       # ~1,900 px^2 cluster
    assert count_cells(img).right == 0            # 3 singles: cluster not expanded
    disk(img, 240, 60, 10)
    res = count_cells(img)
    ref = res.reference_area
    expected = int(cv2.contourArea(next(o.contour for o in res.objects if o.kind == "cluster")) / ref) + 1
    assert res.right == expected and res.left == 4


def test_cluster_rules():
    assert cells_in_cluster(1000, 300) == 4                     # floor(3.33) + 1
    assert cells_in_cluster(1000, 300, "round") == 3
    assert cells_in_cluster(910, 1000, "round") == 1


def test_saturated_pixels_are_signal():
    img = blank()
    disk(img, 100, 100, 10, value=255)
    assert count_cells(img).left == 1


def test_threshold_is_inclusive():
    img = blank()
    disk(img, 100, 100, 10, value=160)
    assert count_cells(img, CountParams(threshold=160)).left == 1
    assert count_cells(img, CountParams(threshold=161)).left == 0


def test_left_right_split_uses_roi_midline():
    img = blank(200, 400)
    disk(img, 190, 100, 8)
    disk(img, 215, 100, 8)
    res = count_cells(img)
    assert (res.left, res.right) == (1, 1)


def test_channel_selection():
    img = blank()
    cv2.circle(img, (100, 100), 10, (0, 220, 0), -1)   # green only
    assert count_cells(img).total == 0
    assert count_cells(img, CountParams(channel="green")).total == 1


def test_16_bit_input_is_scaled_like_opencv():
    img16 = np.zeros((100, 100), np.uint16)
    img16[40:60, 40:60] = 200 * 256
    assert to_bgr8(img16)[50, 50, 2] == 200
    assert count_cells(img16).total == 1


def test_parameter_validation():
    with pytest.raises(ValueError):
        CountParams(min_area=900, max_area=40)
    with pytest.raises(ValueError):
        CountParams(channel="purple")
    with pytest.raises(ValueError):
        CountParams(threshold=300)
