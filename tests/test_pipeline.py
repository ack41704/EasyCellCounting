import cv2
import pandas as pd
import pytest

from easycellcounting import cli, interactive
from easycellcounting.io import natural_key
from easycellcounting.pipeline import count_folder
from easycellcounting.roi import ROI, load_rois, save_rois
from easycellcounting.synthetic import write_demo_dataset


@pytest.fixture()
def dataset(tmp_path):
    truth = pd.DataFrame(write_demo_dataset(tmp_path / "images"))
    return tmp_path / "images", truth


def test_count_folder_matches_ground_truth(dataset, tmp_path):
    images, truth = dataset
    df = count_folder(images, out_dir=tmp_path / "out", verbose=False)
    m = df.merge(truth, on="file")
    assert len(m) == len(truth)
    # isolated cells are exact; each side has one cluster whose estimate may be off by one
    assert ((m["left"] - m["true_left"]).abs() <= 1).all()
    assert ((m["right"] - m["true_right"]).abs() <= 1).all()
    assert (m["n_clusters"] == 2).all()
    assert (tmp_path / "out" / "cell_counts.csv").exists()
    assert (tmp_path / "out" / "cell_counts.xlsx").exists()
    assert len(list((tmp_path / "out" / "annotated").glob("*.png"))) == len(truth)


def test_roi_file_restricts_and_sets_parameters(dataset, tmp_path):
    images, truth = dataset
    first = truth["file"].iloc[0]
    img = cv2.imread(str(images / first))
    h, w = img.shape[:2]
    save_rois([ROI(first, 0, 0, w // 2, h, threshold=160)], tmp_path / "rois.csv")
    with pytest.warns(UserWarning, match="no ROI"):
        df = count_folder(images, tmp_path / "rois.csv", verbose=False)
    assert list(df["file"]) == [first]
    # the left half of the image is the new ROI; its count is (about) the old left count
    assert abs(df["total"].iloc[0] - truth["true_left"].iloc[0]) <= 1


def test_unknown_roi_file_names_are_an_error(dataset, tmp_path):
    images, _ = dataset
    save_rois([ROI("not_there.tif", 0, 0, 5, 5)], tmp_path / "rois.csv")
    with pytest.raises(FileNotFoundError):
        count_folder(images, tmp_path / "rois.csv", verbose=False)


def test_interactive_selection_with_scripted_keys(dataset, tmp_path, monkeypatch):
    """Drive select_rois with fake key presses: rotate, preset, undo, presets, skip."""
    images, truth = dataset
    keys = iter("w" "t" "u" "e" "e" "v" "p" "s" "i" "m")
    boxes = iter([(10, 10, 100, 80)] * 6)
    monkeypatch.setattr(interactive.cv2, "imshow", lambda *a: None)
    monkeypatch.setattr(interactive.cv2, "destroyAllWindows", lambda: None)
    monkeypatch.setattr(interactive.cv2, "waitKey", lambda *a: ord(next(keys, "\x1b")))  # then ESC
    monkeypatch.setattr(interactive.cv2, "selectROI", lambda *a, **k: next(boxes))
    monkeypatch.setattr("builtins.input", lambda *_: "123")
    rois = interactive.select_rois(images, tmp_path / "rois.csv", resume=False)
    saved = load_rois(tmp_path / "rois.csv")
    names = sorted(truth["file"], key=natural_key)                # folder order
    assert [r.file for r in saved] == [names[0], names[1], names[3], names[4]]
    assert [r.preset for r in saved] == ["v200", "p250", "custom", "default"]
    assert saved[0].rotation == 353.0 and saved[0].threshold == 200
    assert saved[2].threshold == 123 and saved[2].cluster_max_area == 1e6
    assert rois == saved
    # the preview (<=1200 px wide) box maps onto the full image, here at the same scale
    assert saved[1].box == (10, 10, 100, 80)


def test_cli_count_and_demo(dataset, tmp_path, capsys):
    images, truth = dataset
    assert cli.main(["count", str(images), "--out", str(tmp_path / "o"), "--no-images",
                     "--box", "0,0,600,900"]) == 0
    df = pd.read_csv(tmp_path / "o" / "cell_counts.csv").merge(truth, on="file")
    assert (df["roi_width"] == 600).all()
    assert ((df["total"] - df["true_left"]).abs() <= 1).all()     # the left half only
    assert cli.main(["demo", "--out", str(tmp_path / "demo")]) == 0
    assert "true_left" in capsys.readouterr().out
