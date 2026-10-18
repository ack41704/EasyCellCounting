"""
Interactive ROI selection with OpenCV windows (same keys as the original script).

Each image is shown down-scaled. Choose a preset key, drag a box around the region to
count and press SPACE or ENTER to accept it (``c`` cancels the box). ROIs are written
to the ROI CSV after every image, so a session can be stopped and resumed.

    t   default threshold (160)            w / e   rotate +7 / -7 degrees
    v   threshold 200                      o       rotate 180 degrees
    p   threshold 250                      q       check: preview the count with a typed threshold
    i   typed threshold, no cluster cap    u       undo (go back one image)
    s   skip this image                    m       draw this ROI, then stop
    ESC stop now
"""
from __future__ import annotations

from pathlib import Path

import cv2

from .counting import CountParams, annotate, count_cells
from .io import load_image, make_preview, pair_images
from .roi import ROI, clip_box, load_rois, rotate_image, save_rois, scale_box

PRESETS = {"t": ("default", None), "v": ("v200", 200), "p": ("p250", 250), "i": ("custom", None)}
HELP = __doc__.split("\n\n", 2)[2]


def _ask_int(prompt):
    while True:
        try:
            return int(input(prompt))
        except ValueError:
            print("  please type a whole number (0-255)")


def _draw_box(window, image):
    x, y, w, h = cv2.selectROI(window, image, showCrosshair=True, fromCenter=False)
    return None if w == 0 or h == 0 else (x, y, w, h)


def select_rois(image_dir, rois_path="rois.csv", preview_dir=None,
                params: CountParams = CountParams(), resume: bool = True,
                max_display=(1200, 900), rotate_step: float = 7.0) -> list:
    """Walk through the images, record one ROI per image and save them to `rois_path`.

    preview_dir: optional folder of down-scaled copies with the same file stems (as in
                 the original two-folder workflow); otherwise previews are made on the fly.
    resume:      keep ROIs already in `rois_path` and start at the first image without one."""
    pairs = pair_images(image_dir, preview_dir)
    rois_path = Path(rois_path)
    done = {r.file: r for r in load_rois(rois_path)} if resume and rois_path.exists() else {}
    order = [p.name for p, _ in pairs]
    rois = [done[nm] for nm in order if nm in done]
    n = next((i for i, nm in enumerate(order) if nm not in done), len(order))
    if n == len(order):
        print(f"all {len(order)} images already have an ROI in {rois_path}")
        return rois
    print(HELP)

    angle, cache = 0.0, {}
    while n < len(pairs):
        full_path, prev_path = pairs[n]
        if n not in cache:
            full = load_image(full_path)
            prev = load_image(prev_path) if prev_path else make_preview(full, *max_display)
            cache = {n: (full, prev)}
        full, prev = cache[n]
        shown = rotate_image(prev, angle)
        window = f"{n + 1} / {len(pairs)}  {full_path.name}"
        cv2.imshow(window, shown)
        key = cv2.waitKey(0) & 0xFF
        cv2.destroyAllWindows()
        ch = chr(key) if key < 128 else ""

        if ch in ("w", "e", "o"):
            angle = (angle + {"w": rotate_step, "e": -rotate_step, "o": 180.0}[ch]) % 360
            continue
        if key == 27:
            break
        if ch == "s":
            n, angle = n + 1, 0.0
            continue
        if ch == "u":
            if rois:
                last = rois.pop()
                n = order.index(last.file)
                angle = last.rotation
                save_rois(sorted(rois, key=lambda r: order.index(r.file)), rois_path)
                print(f"undo: redo {last.file}")
            else:
                print("nothing to undo")
            continue
        if ch == "q":
            thr = _ask_int("Threshold to check: ")
            box = _draw_box("check: draw a box", shown)
            cv2.destroyAllWindows()
            if box:
                fb = clip_box(scale_box(box, shown.shape, full.shape), full.shape)
                x, y, w, h = fb
                region = rotate_image(full, angle)[y:y + h, x:x + w]
                res = count_cells(region, params.replace(threshold=thr))
                cv2.imshow(f"check (threshold {thr}): left {res.left}, right {res.right}"
                           " - any key to go back", annotate(region, res))
                cv2.waitKey(0)
                cv2.destroyAllWindows()
            continue
        if ch not in PRESETS and ch != "m":
            continue

        preset, thr = PRESETS.get(ch, ("default", None))
        cluster_max = params.cluster_max_area
        if ch == "i":
            thr = _ask_int("Threshold: ")
            cluster_max = 1e6
        box = _draw_box("draw the region to count, then SPACE / ENTER", shown)
        cv2.destroyAllWindows()
        if box is None:
            print("no box drawn - choose again")
            continue
        x, y, w, h = clip_box(scale_box(box, shown.shape, full.shape), full.shape)
        rois.append(ROI(full_path.name, x, y, w, h, rotation=angle,
                        threshold=int(thr if thr is not None else params.threshold),
                        min_area=params.min_area, max_area=params.max_area,
                        cluster_max_area=cluster_max, preset=preset))
        save_rois(sorted(rois, key=lambda r: order.index(r.file)), rois_path)
        n, angle = n + 1, 0.0
        have = {r.file for r in rois}
        while n < len(order) and order[n] in have:      # resumed sessions: skip finished images
            n += 1
        if ch == "m":
            break

    cv2.destroyAllWindows()
    cv2.waitKey(1)
    print(f"{len(rois)} ROI(s) saved to {rois_path}")
    return rois
