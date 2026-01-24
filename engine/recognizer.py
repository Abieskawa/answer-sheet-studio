from __future__ import annotations

import csv
import os
from pathlib import Path
from typing import List, Tuple, Dict, Optional, Any

import fitz  # PyMuPDF
import numpy as np
import cv2

from .config import (
    PAGE_W_PT, PAGE_H_PT,
    CORNER_MARK_MARGIN, CORNER_MARK_SIZE,
    GRADE_VALUES, GRADE_CIRCLE_XS, GRADE_CIRCLE_Y, GRADE_CIRCLE_RADIUS,
    CLASS_VALUES, CLASS_CIRCLE_XS, CLASS_CIRCLE_Y, CLASS_CIRCLE_RADIUS,
    SEAT_START_X, SEAT_STEP_X, SEAT_TOP_ROW_Y, SEAT_BOTTOM_ROW_Y, SEAT_CIRCLE_RADIUS,
    MAX_QUESTIONS, DEFAULT_CHOICES_COUNT, make_choices,
    BUBBLE_RADIUS, COL_COUNT,
    compute_answer_layout,
)


def _env_float(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if raw is None:
        return float(default)
    try:
        return float(raw)
    except ValueError:
        return float(default)


MIN_SCORE_GRADE = _env_float("ANSWER_SHEET_MIN_SCORE_GRADE", 0.05)
MIN_SCORE_CLASS = _env_float("ANSWER_SHEET_MIN_SCORE_CLASS", 0.05)
MIN_SCORE_SEAT = _env_float("ANSWER_SHEET_MIN_SCORE_SEAT", 0.06)
MIN_SCORE_CHOICE = _env_float("ANSWER_SHEET_MIN_SCORE_CHOICE", 0.025)


def render_page(page: fitz.Page, dpi: int = 200) -> Tuple[np.ndarray, float]:
    # scale points -> pixels
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    return img, zoom


def _estimate_zoom_from_image(gray: np.ndarray) -> float:
    h, w = gray.shape[:2]
    zx = float(w) / float(PAGE_W_PT)
    zy = float(h) / float(PAGE_H_PT)
    # Choose the smaller scale to stay conservative when the PDF has margins/crop.
    return float(min(zx, zy))


def _find_corner_mark_in_roi(
    roi_gray: np.ndarray,
    expected_px: float,
    target_corner_xy: Tuple[float, float],
) -> Optional[Tuple[float, float]]:
    if roi_gray.size == 0:
        return None

    blur = cv2.GaussianBlur(roi_gray, (5, 5), 0)
    # Otsu makes this robust to different scanners/brightness.
    _, bw = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    bw = cv2.morphologyEx(bw, cv2.MORPH_CLOSE, np.ones((5, 5), np.uint8), iterations=1)
    bw = cv2.medianBlur(bw, 5)

    contours, _ = cv2.findContours(bw, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None

    expected = float(max(8.0, expected_px))
    min_side = expected * 0.55
    max_side = expected * 2.0
    min_area = (expected * expected) * 0.15

    best: Optional[Tuple[float, float, float]] = None  # score, cx, cy
    for cnt in contours:
        area = float(cv2.contourArea(cnt))
        if area < min_area:
            continue
        x, y, ww, hh = cv2.boundingRect(cnt)
        ww_f = float(ww)
        hh_f = float(hh)
        ar = ww_f / (hh_f + 1e-6)
        if not (0.6 < ar < 1.4):
            continue
        if ww_f < min_side or hh_f < min_side:
            continue
        if ww_f > max_side or hh_f > max_side:
            continue
        fill = area / ((ww_f * hh_f) + 1e-6)
        if fill < 0.6:
            continue
        cx = float(x) + ww_f / 2.0
        cy = float(y) + hh_f / 2.0
        dx = cx - float(target_corner_xy[0])
        dy = cy - float(target_corner_xy[1])
        dist = float((dx * dx + dy * dy) ** 0.5)
        score = (area * fill) / (1.0 + dist)
        if best is None or score > best[0]:
            best = (score, cx, cy)

    if best is None:
        return None
    _, cx, cy = best
    return cx, cy


def _fill_missing_corner(corners: Dict[str, Tuple[float, float]]) -> Dict[str, Tuple[float, float]]:
    """
    Given up to 4 corners (tl/tr/br/bl), try to fill the missing one assuming a parallelogram:
      tl + br = tr + bl
    """
    required = {"tl", "tr", "br", "bl"}
    missing = list(required - set(corners))
    if len(missing) != 1:
        return corners

    m = missing[0]
    tl = corners.get("tl")
    tr = corners.get("tr")
    br = corners.get("br")
    bl = corners.get("bl")

    if m == "tl" and tr and bl and br:
        corners["tl"] = (tr[0] + bl[0] - br[0], tr[1] + bl[1] - br[1])
    elif m == "tr" and tl and br and bl:
        corners["tr"] = (tl[0] + br[0] - bl[0], tl[1] + br[1] - bl[1])
    elif m == "br" and tr and bl and tl:
        corners["br"] = (tr[0] + bl[0] - tl[0], tr[1] + bl[1] - tl[1])
    elif m == "bl" and tl and br and tr:
        corners["bl"] = (tl[0] + br[0] - tr[0], tl[1] + br[1] - tr[1])
    return corners


def find_corner_marks(img: np.ndarray) -> Optional[np.ndarray]:
    """Return 4 corner points in image pixels: TL, TR, BR, BL."""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape[:2]

    zoom_est = _estimate_zoom_from_image(gray)
    expected_px = float(CORNER_MARK_SIZE) * float(zoom_est)
    expected_center = float(CORNER_MARK_MARGIN + (CORNER_MARK_SIZE / 2.0)) * float(zoom_est)

    rx = int(w * 0.25)
    ry = int(h * 0.25)
    rois = {
        "tl": (0, 0, rx, ry),
        "tr": (w - rx, 0, w, ry),
        "bl": (0, h - ry, rx, h),
        "br": (w - rx, h - ry, w, h),
    }

    corners: Dict[str, Tuple[float, float]] = {}
    for key, (x0, y0, x1, y1) in rois.items():
        roi = gray[y0:y1, x0:x1]
        roi_h, roi_w = roi.shape[:2]
        off_x = float(min(max(expected_center, 0.0), float(max(0, roi_w - 1))))
        off_y = float(min(max(expected_center, 0.0), float(max(0, roi_h - 1))))
        if key == "tl":
            target = (off_x, off_y)
        elif key == "tr":
            target = (float(max(0, roi_w - 1)) - off_x, off_y)
        elif key == "bl":
            target = (off_x, float(max(0, roi_h - 1)) - off_y)
        else:
            target = (float(max(0, roi_w - 1)) - off_x, float(max(0, roi_h - 1)) - off_y)
        found = _find_corner_mark_in_roi(roi, expected_px=expected_px, target_corner_xy=target)
        if found is None:
            continue
        cx, cy = found
        corners[key] = (cx + float(x0), cy + float(y0))

    if len(corners) == 3:
        corners = _fill_missing_corner(corners)

    if len(corners) < 4:
        return None

    tl = corners["tl"]
    tr = corners["tr"]
    br = corners["br"]
    bl = corners["bl"]
    return np.array([tl, tr, br, bl], dtype=np.float32)


def warp_to_canonical(img: np.ndarray, zoom: float) -> Tuple[np.ndarray, np.ndarray]:
    """Warp the page to canonical A4 pixel size at the same zoom."""
    src = find_corner_marks(img)
    h, w = img.shape[:2]
    if src is None:
        # fallback: no warp
        dst_img = img.copy()
        M = np.eye(3, dtype=np.float32)
        return dst_img, M

    can_w = int(PAGE_W_PT * zoom)
    can_h = int(PAGE_H_PT * zoom)

    # destination points are slightly inside page (match where we draw marks)
    m = CORNER_MARK_MARGIN * zoom
    s = CORNER_MARK_SIZE * zoom
    dst = np.array([
        [m + s/2, m + s/2],
        [can_w - (m + s/2), m + s/2],
        [can_w - (m + s/2), can_h - (m + s/2)],
        [m + s/2, can_h - (m + s/2)],
    ], dtype=np.float32)

    M = cv2.getPerspectiveTransform(src, dst)
    warped = cv2.warpPerspective(img, M, (can_w, can_h), flags=cv2.INTER_LINEAR)
    return warped, M


def bubble_bbox_px(x_pt: float, y_pt: float, r_pt: float, zoom: float) -> Tuple[int,int,int,int]:
    x = x_pt * zoom
    y = (PAGE_H_PT - y_pt) * zoom  # convert PDF y-up to image y-down
    r = r_pt * zoom
    x0 = int(x - r - 2); y0 = int(y - r - 2)
    x1 = int(x + r + 2); y1 = int(y + r + 2)
    return x0, y0, x1, y1


def score_bubble(gray: np.ndarray, bbox: Tuple[int,int,int,int]) -> float:
    """
    Return a fill score in [0, 1], robust to dark/gray scans.

    We measure how much darker the *inner* region of the bubble is compared to the local
    background *outside* the bubble (annulus). This avoids treating the printed circle border
    (or uniformly dark scans) as "filled".
    """
    x0, y0, x1, y1 = bbox
    cx = (x0 + x1) / 2.0
    cy = (y0 + y1) / 2.0
    r = (min(x1 - x0, y1 - y0) / 2.0) - 2.0
    if r <= 1.0:
        return 0.0

    bg_r1 = r * 1.05
    bg_r2 = r * 1.45
    x0e = int(cx - bg_r2)
    y0e = int(cy - bg_r2)
    x1e = int(cx + bg_r2)
    y1e = int(cy + bg_r2)

    x0e = max(0, x0e)
    y0e = max(0, y0e)
    x1e = min(gray.shape[1], x1e)
    y1e = min(gray.shape[0], y1e)
    if x1e <= x0e or y1e <= y0e:
        return 0.0

    roi = gray[y0e:y1e, x0e:x1e]
    if roi.size == 0:
        return 0.0

    ccx = cx - x0e
    ccy = cy - y0e
    yy, xx = np.ogrid[: roi.shape[0], : roi.shape[1]]
    dx = xx - ccx
    dy = yy - ccy
    dist2 = dx * dx + dy * dy

    inner_r = max(1.0, r * 0.55)
    inner_mask = dist2 <= (inner_r * inner_r)
    bg_mask = (dist2 >= (bg_r1 * bg_r1)) & (dist2 <= (bg_r2 * bg_r2))

    if int(inner_mask.sum()) < 16 or int(bg_mask.sum()) < 16:
        return 0.0

    inner_mean = float(roi[inner_mask].mean())
    bg_mean = float(roi[bg_mask].mean())
    score = (bg_mean - inner_mean) / 255.0
    if score < 0.0:
        score = 0.0
    if score > 1.0:
        score = 1.0
    return float(score)


def pick_one(
    scores: List[float],
    labels: List[str],
    min_score: float = 0.08,
    amb_delta: float = 0.03,
) -> Tuple[Optional[str], str, float, int, Optional[str], float]:
    """Pick one label based on scores.
    - If max < min_score -> BLANK
    - If top1 - top2 < amb_delta -> AMBIGUOUS
    Returns: (value, status, best_score, best_idx, second_label, second_score)
    """
    idxs = np.argsort(scores)[::-1]
    top = int(idxs[0])
    best = scores[top]
    if best < min_score:
        second_label = labels[int(idxs[1])] if len(scores) >= 2 else None
        second_score = float(scores[int(idxs[1])]) if len(scores) >= 2 else 0.0
        return None, "BLANK", float(best), top, second_label, second_score
    if len(scores) >= 2:
        second = scores[int(idxs[1])]
        # For light scans, scores can be small; use a smaller ambiguity delta when best is low.
        delta = float(min(float(amb_delta), float(best) * 0.55))
        if float(second) >= float(min_score) * 0.85 and (best - second) < delta:
            second_label = labels[int(idxs[1])]
            second_score = float(second)
            return labels[top], "AMBIGUOUS", float(best), top, second_label, second_score
    second_label = labels[int(idxs[1])] if len(scores) >= 2 else None
    second_score = float(scores[int(idxs[1])]) if len(scores) >= 2 else 0.0
    return labels[top], "OK", float(best), top, second_label, second_score


def pick_choice_multi(
    scores: List[float],
    labels: List[str],
    min_score: float = 0.08,
    amb_delta: float = 0.03,
    multi_ratio: float = 0.65,
) -> Tuple[str, str, float, int, Optional[str], float, List[int]]:
    """
    Like pick_one, but supports multi-select:
      - If 2+ bubbles >= min_score => status=MULTI and value is concatenated labels (in label order)
      - Else fall back to pick_one (OK/AMBIGUOUS/BLANK)

    Returns: (value_str, status, best_score, best_idx, second_label, second_score, picked_indices)
    """
    best = float(max(scores) if scores else 0.0)
    multi_cut = max(float(min_score), best * float(multi_ratio))
    picked = [i for i, s in enumerate(scores) if float(s) >= multi_cut]
    if len(picked) >= 2:
        picked_sorted = sorted(int(i) for i in picked)
        value = "".join(labels[i] for i in picked_sorted)
        best_idx = int(max(picked_sorted, key=lambda i: scores[i]))
        best_score = float(scores[best_idx])
        # second best among picked (for diagnostics)
        picked_by_score = sorted(picked_sorted, key=lambda i: scores[i], reverse=True)
        second_idx = int(picked_by_score[1]) if len(picked_by_score) >= 2 else best_idx
        second_label = labels[second_idx] if len(labels) > second_idx else None
        second_score = float(scores[second_idx]) if len(scores) > second_idx else 0.0
        return value, "MULTI", best_score, best_idx, second_label, second_score, picked_sorted

    val, status, best_score, idx, second_label, second_score = pick_one(
        scores, labels, min_score=min_score, amb_delta=amb_delta
    )
    out = val if status == "OK" and val is not None else ""
    picked_indices = [int(idx)] if out else []
    return out, status, float(best_score), int(idx), second_label, float(second_score), picked_indices


def process_page(
    warped: np.ndarray,
    zoom: float,
    num_questions: int,
    choices_count: int,
) -> Tuple[Dict[str, Any], np.ndarray, List[Dict[str, Any]]]:
    can = warped.copy()
    gray_raw = cv2.cvtColor(can, cv2.COLOR_BGR2GRAY)
    # Boost local contrast so light marks are easier to detect.
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(16, 16))
    gray = clahe.apply(gray_raw)

    results: Dict[str, Any] = {}
    marks: List[Tuple[Tuple[int,int,int,int], str, str]] = []  # bbox, text, status
    flags: List[Dict[str, Any]] = []

    # Grade (7-12)
    grade_scores = []
    grade_bboxes = []
    for x in GRADE_CIRCLE_XS:
        bbox = bubble_bbox_px(x, GRADE_CIRCLE_Y, GRADE_CIRCLE_RADIUS, zoom)
        grade_bboxes.append(bbox)
        grade_scores.append(score_bubble(gray, bbox))
    g_label = [str(v) for v in GRADE_VALUES]
    g_val, g_status, _, g_idx, g_second_label, _ = pick_one(
        grade_scores, g_label, min_score=MIN_SCORE_GRADE, amb_delta=0.02
    )
    results["grade"] = g_val if g_status == "OK" else None
    results["grade_status"] = g_status
    marks.append((grade_bboxes[g_idx], f"G:{g_val or ''}", g_status))
    if g_status != "OK":
        flags.append(
            {
                "field": "grade",
                "question": "",
                "status": g_status,
                "best_label": g_label[g_idx],
                "second_label": g_second_label or "",
            }
        )

    # Class (single row digits 0-9; treat 0 as unused)
    class_scores = []
    class_bboxes = []
    for x in CLASS_CIRCLE_XS:
        bbox = bubble_bbox_px(x, CLASS_CIRCLE_Y, CLASS_CIRCLE_RADIUS, zoom)
        class_bboxes.append(bbox)
        class_scores.append(score_bubble(gray, bbox))
    c_label_all = [str(v) for v in CLASS_VALUES]
    c_scores_pick = class_scores
    c_labels_pick = c_label_all
    c_offset = 0
    if c_label_all and c_label_all[0] == "0":
        c_scores_pick = class_scores[1:]
        c_labels_pick = c_label_all[1:]
        c_offset = 1

    c_val, c_status, _, c_idx0, c_second_label, _ = pick_one(
        c_scores_pick, c_labels_pick, min_score=MIN_SCORE_CLASS, amb_delta=0.02
    )
    c_idx = int(c_idx0) + int(c_offset)
    results["class_no"] = c_val if c_status == "OK" else None
    results["class_status"] = c_status
    marks.append((class_bboxes[c_idx], f"C:{c_val or ''}", c_status))
    if c_status != "OK":
        flags.append(
            {
                "field": "class_no",
                "question": "",
                "status": c_status,
                "best_label": c_labels_pick[int(c_idx0)] if c_labels_pick else "",
                "second_label": c_second_label or "",
            }
        )

    # Seat number (00-99): support legacy digit order variants.
    # Default: 0-9 left-to-right; Legacy: 1-9 then 0.
    digit_labels_default = [str(d) for d in range(10)]
    digit_labels_legacy = [str(d) for d in range(1, 10)] + ["0"]

    def pick_digit_row(
        y_pt: float,
        digit_labels: List[str],
    ) -> Tuple[str, str, float, int, Optional[str], float, List[Tuple[int, int, int, int]], List[int], List[float]]:
        bboxes: List[Tuple[int, int, int, int]] = []
        scores: List[float] = []
        for i in range(10):
            x_pt = SEAT_START_X + i * SEAT_STEP_X
            bbox = bubble_bbox_px(x_pt, y_pt, SEAT_CIRCLE_RADIUS, zoom)
            bboxes.append(bbox)
            scores.append(score_bubble(gray, bbox))

        best = float(max(scores) if scores else 0.0)
        multi_cut = max(float(MIN_SCORE_SEAT), best * 0.65)
        picked = [i for i, s in enumerate(scores) if float(s) >= multi_cut]
        if len(picked) >= 2:
            picked_sorted = sorted(int(i) for i in picked)
            value = "".join(digit_labels[i] for i in picked_sorted)
            best_idx = int(max(picked_sorted, key=lambda i: scores[i]))
            best_score = float(scores[best_idx])
            picked_by_score = sorted(picked_sorted, key=lambda i: scores[i], reverse=True)
            second_idx = int(picked_by_score[1]) if len(picked_by_score) >= 2 else best_idx
            second_label = digit_labels[second_idx] if len(digit_labels) > second_idx else None
            second_score = float(scores[second_idx]) if len(scores) > second_idx else 0.0
            return value, "MULTI", best_score, best_idx, second_label, second_score, bboxes, picked_sorted, scores

        val, status, best_score, idx, second_label, second_score = pick_one(
            scores, digit_labels, min_score=MIN_SCORE_SEAT, amb_delta=0.02
        )
        val_out = str(val) if (val is not None) else ""
        return val_out, status, float(best_score), int(idx), second_label, float(second_score), bboxes, [int(idx)], scores

    def seat_candidate_from_rows(
        tens_val: str,
        tens_status: str,
        tens_best_score: float,
        ones_val: str,
        ones_status: str,
        ones_best_score: float,
    ) -> Tuple[Optional[str], str]:
        seat_no: Optional[str] = None
        if tens_status == "OK" and ones_status == "OK" and tens_val and ones_val:
            seat_no = f"{tens_val}{ones_val}"
        else:
            def candidate(value: str, status: str) -> Optional[str]:
                if status == "OK" and value:
                    return value
                if status == "MULTI" and len(value) == 2:
                    return value
                return None

            t_cand = candidate(tens_val, tens_status)
            o_cand = candidate(ones_val, ones_status)
            if ones_status == "BLANK" and t_cand is not None:
                seat_no = t_cand
            elif tens_status == "BLANK" and o_cand is not None:
                seat_no = o_cand

        seat_status = (
            "OK"
            if seat_no is not None
            else ("AMBIGUOUS" if ("AMBIGUOUS" in [tens_status, ones_status]) else "BLANK")
        )
        return seat_no, seat_status

    def score_seat_candidate(seat_no: Optional[str], seat_status: str, tens_best: float, ones_best: float) -> float:
        if seat_status == "OK" and seat_no is not None:
            return tens_best + ones_best
        if seat_status == "AMBIGUOUS":
            return 0.5 * (tens_best + ones_best)
        return 0.0

    # Try both digit label orders and also swapped rows (legacy templates sometimes swap tens/ones).
    best_variant = None
    best_variant_score = -1.0
    variants = [
        ("default", digit_labels_default),
        ("legacy", digit_labels_legacy),
    ]
    for label_mode, labels in variants:
        t_val, t_status, t_best, t_idx, t_second_label, _, t_bboxes, t_picked, _ = pick_digit_row(SEAT_TOP_ROW_Y, labels)
        o_val, o_status, o_best, o_idx, o_second_label, _, o_bboxes, o_picked, _ = pick_digit_row(SEAT_BOTTOM_ROW_Y, labels)

        seat_no_a, seat_status_a = seat_candidate_from_rows(t_val, t_status, t_best, o_val, o_status, o_best)
        score_a = score_seat_candidate(seat_no_a, seat_status_a, t_best, o_best)
        if score_a > best_variant_score:
            best_variant_score = score_a
            best_variant = ("normal", label_mode, labels, t_val, t_status, t_best, t_idx, t_second_label, t_bboxes, t_picked, o_val, o_status, o_best, o_idx, o_second_label, o_bboxes, o_picked, seat_no_a, seat_status_a)

        seat_no_b, seat_status_b = seat_candidate_from_rows(o_val, o_status, o_best, t_val, t_status, t_best)
        score_b = score_seat_candidate(seat_no_b, seat_status_b, o_best, t_best) - 0.001  # prefer non-swapped on ties
        if score_b > best_variant_score:
            best_variant_score = score_b
            best_variant = ("swapped", label_mode, labels, t_val, t_status, t_best, t_idx, t_second_label, t_bboxes, t_picked, o_val, o_status, o_best, o_idx, o_second_label, o_bboxes, o_picked, seat_no_b, seat_status_b)

    assert best_variant is not None
    row_mode, label_mode, digit_labels, t_val, t_status, _, t_idx, t_second_label, t_bboxes, t_picked, o_val, o_status, _, o_idx, o_second_label, o_bboxes, o_picked, seat_no, seat_status = best_variant

    results["seat_no"] = seat_no
    results["seat_status"] = seat_status

    # Marks + flags always annotate the physical top row as "T:" and bottom as "O:" to aid debugging,
    # even if we used swapped decoding for seat_no.
    if t_status == "MULTI":
        for j in t_picked:
            text = f"T:{t_val}" if j == t_idx else ""
            marks.append((t_bboxes[j], text, t_status))
    else:
        marks.append((t_bboxes[t_idx], f"T:{t_val or ''}", t_status))

    if o_status == "MULTI":
        for j in o_picked:
            text = f"O:{o_val}" if j == o_idx else ""
            marks.append((o_bboxes[j], text, o_status))
    else:
        marks.append((o_bboxes[o_idx], f"O:{o_val or ''}", o_status))

    if t_status != "OK":
        flags.append(
            {
                "field": "seat_tens",
                "question": "",
                "status": t_status,
                "best_label": digit_labels[t_idx],
                "second_label": t_second_label or "",
            }
        )
    if o_status != "OK":
        flags.append(
            {
                "field": "seat_ones",
                "question": "",
                "status": o_status,
                "best_label": digit_labels[o_idx],
                "second_label": o_second_label or "",
            }
        )

    # Questions
    answers = []
    choices = make_choices(choices_count)
    layout = compute_answer_layout(num_questions, choices_count=choices_count)
    rows = layout["rows_per_col"]
    first_y = layout["first_row_y"]
    row_step = layout["row_step"]
    bubble_xs = layout["bubble_xs"]

    q = 1
    for col in range(COL_COUNT):
        for row in range(rows):
            if q > num_questions:
                break
            y_pt = first_y - (row * row_step)
            bboxes = []
            scores = []
            xs = bubble_xs[col]
            for j, ch in enumerate(choices):
                x_pt = xs[j]
                bbox = bubble_bbox_px(x_pt, y_pt, BUBBLE_RADIUS, zoom)
                bboxes.append(bbox)
                scores.append(score_bubble(gray, bbox))
            val, status, _, idx, second_label, _, picked_idxs = pick_choice_multi(
                scores, choices, min_score=MIN_SCORE_CHOICE, amb_delta=0.03
            )
            answers.append(val if status in {"OK", "MULTI"} and val is not None else "")
            if status == "MULTI":
                for j in picked_idxs:
                    text = f"{q}:{val}" if j == idx else ""
                    marks.append((bboxes[j], text, status))
            else:
                marks.append((bboxes[idx], f"{q}:{val or ''}", status))
            if status != "OK":
                flags.append(
                    {
                        "field": f"Q{q}",
                        "question": q,
                        "status": status,
                        "best_label": val if status == "MULTI" else choices[idx],
                        "second_label": "" if status == "MULTI" else (second_label or ""),
                    }
                )
            q += 1

    results["answers"] = "".join(answers)
    # also keep per-question list
    for i in range(num_questions):
        results[f"Q{i+1}"] = answers[i] if i < len(answers) else ""

    # Annotate
    annotated = can.copy()
    for (x0,y0,x1,y1), text, status in marks:
        # color by status
        if status == "OK":
            color = (0,255,0)
        elif status == "AMBIGUOUS":
            color = (0,165,255)
        elif status == "MULTI":
            color = (255,0,255)
        else:
            color = (0,0,255)
        cv2.rectangle(annotated, (x0,y0), (x1,y1), color, 2)
        if text:
            cv2.putText(
                annotated,
                text,
                (x0, max(10, y0 - 4)),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.4,
                color,
                1,
                cv2.LINE_AA,
            )

    return results, annotated, flags


def images_to_pdf(images: List[np.ndarray], out_pdf_path: str, dpi: int = 200):
    doc = fitz.open()
    for img in images:
        h, w = img.shape[:2]
        # create page size in points
        # pixels -> inches -> points
        w_pt = w / dpi * 72.0
        h_pt = h / dpi * 72.0
        page = doc.new_page(width=w_pt, height=h_pt)
        # encode as png
        success, buf = cv2.imencode(".png", img)
        if not success:
            raise RuntimeError("Failed to encode annotated image")
        page.insert_image(page.rect, stream=buf.tobytes())
    doc.save(out_pdf_path)
    doc.close()


def process_pdf_to_csv_and_annotated_pdf(
    input_pdf_path: str,
    num_questions: int,
    out_csv_path: str,
    out_annotated_pdf_path: str,
    choices_count: int = DEFAULT_CHOICES_COUNT,
    dpi: int = 200,
    out_ambiguity_csv_path: Optional[str] = None,
):
    num_questions = max(1, min(MAX_QUESTIONS, int(num_questions)))
    choices_count = max(3, min(5, int(choices_count)))
    out_ambiguity_csv_path = out_ambiguity_csv_path or str(Path(out_csv_path).with_name("ambiguity.csv"))

    doc = fitz.open(input_pdf_path)
    people: List[Dict[str, Any]] = []
    flags: List[Dict[str, Any]] = []
    annotated_images = []

    for idx in range(doc.page_count):
        page = doc.load_page(idx)
        img, zoom = render_page(page, dpi=dpi)
        warped, _ = warp_to_canonical(img, zoom)
        result, annotated, page_flags = process_page(warped, zoom, num_questions=num_questions, choices_count=choices_count)
        result["page"] = idx + 1
        people.append(result)
        for flag in page_flags:
            flags.append({"page": idx + 1, **flag})
        annotated_images.append(annotated)

    def normalize_person_id(seat_no: Optional[str]) -> Optional[str]:
        if seat_no is None:
            return None
        seat = str(seat_no).strip()
        if not seat:
            return None
        if seat.isdigit():
            return str(int(seat))
        return seat

    used_person_ids: set[str] = set()
    for row in people:
        base_id = normalize_person_id(row.get("seat_no")) or f"page_{row['page']}"
        person_id = base_id
        suffix = 2
        while person_id in used_person_ids:
            person_id = f"{base_id}_{suffix}"
            suffix += 1
        used_person_ids.add(person_id)
        row["person_id"] = person_id

    def person_sort_key(person_id: str) -> Tuple[int, int, int, str]:
        def parse_numeric_with_suffix(value: str) -> Optional[Tuple[int, int]]:
            parts = value.split("_", 1)
            if not parts or not parts[0].isdigit():
                return None
            base = int(parts[0])
            suffix = int(parts[1]) if len(parts) == 2 and parts[1].isdigit() else 0
            return base, suffix

        parsed = parse_numeric_with_suffix(person_id)
        if parsed is not None:
            base, suffix = parsed
            return (0, base, suffix, person_id)

        if person_id.startswith("page_"):
            parsed = parse_numeric_with_suffix(person_id[5:])
            if parsed is not None:
                base, suffix = parsed
                return (1, base, suffix, person_id)

        return (2, 0, 0, person_id)

    people_sorted = sorted(people, key=lambda row: person_sort_key(str(row.get("person_id", ""))))
    person_ids = [str(row.get("person_id", "")) for row in people_sorted]

    # Transposed output: columns=people, rows=questions
    results_fields = ["number", *person_ids]
    with open(out_csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=results_fields)
        writer.writeheader()
        for q in range(1, num_questions + 1):
            out_row: Dict[str, Any] = {"number": q}
            for person in people_sorted:
                pid = str(person.get("person_id", ""))
                out_row[pid] = person.get(f"Q{q}", "") or ""
            writer.writerow(out_row)

    # Ambiguity/blank report
    people_by_page = {int(row["page"]): row for row in people if "page" in row}
    amb_fields = [
        "page",
        "person_id",
        "seat_no",
        "grade",
        "class_no",
        "field",
        "question",
        "status",
        "best_label",
        "second_label",
    ]
    with open(out_ambiguity_csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=amb_fields)
        writer.writeheader()
        for flag in flags:
            page_no = int(flag.get("page", 0) or 0)
            person = people_by_page.get(page_no, {})
            writer.writerow(
                {
                    "page": page_no or "",
                    "person_id": person.get("person_id", ""),
                    "seat_no": person.get("seat_no", "") or "",
                    "grade": person.get("grade", "") or "",
                    "class_no": person.get("class_no", "") or "",
                    "field": flag.get("field", ""),
                    "question": flag.get("question", ""),
                    "status": flag.get("status", ""),
                    "best_label": flag.get("best_label", ""),
                    "second_label": flag.get("second_label", ""),
                }
            )

    images_to_pdf(annotated_images, out_annotated_pdf_path, dpi=dpi)
    doc.close()
