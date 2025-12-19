from __future__ import annotations

import csv
from dataclasses import dataclass
from typing import List, Tuple, Dict, Optional

import fitz  # PyMuPDF
import numpy as np
import cv2

from .config import (
    PAGE_W_PT, PAGE_H_PT,
    CORNER_MARK_MARGIN, CORNER_MARK_SIZE,
    GRADE_VALUES, GRADE_CIRCLE_XS, GRADE_CIRCLE_Y, GRADE_CIRCLE_RADIUS,
    CLASS_VALUES, CLASS_CIRCLE_XS, CLASS_CIRCLE_Y, CLASS_CIRCLE_RADIUS,
    SEAT_START_X, SEAT_STEP_X, SEAT_TOP_ROW_Y, SEAT_BOTTOM_ROW_Y, SEAT_CIRCLE_RADIUS,
    MAX_QUESTIONS, CHOICES,
    BUBBLE_RADIUS, COL_COUNT,
    compute_answer_layout,
)

@dataclass
class PickResult:
    value: Optional[str]   # "A"/"B"/"C"/"D" or digit or grade, None if blank
    status: str            # "OK" | "BLANK" | "AMBIGUOUS"
    score: float
    bbox: Tuple[int,int,int,int]  # x0,y0,x1,y1 in pixels (canonical)


def render_page(page: fitz.Page, dpi: int = 200) -> Tuple[np.ndarray, float]:
    # scale points -> pixels
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    return img, zoom


def find_corner_marks(img: np.ndarray) -> Optional[np.ndarray]:
    """Return 4 corner points in image pixels: TL, TR, BR, BL."""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # strong threshold for black squares
    _, bw = cv2.threshold(gray, 60, 255, cv2.THRESH_BINARY_INV)
    bw = cv2.medianBlur(bw, 5)

    contours, _ = cv2.findContours(bw, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    h, w = gray.shape[:2]

    candidates = []
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area < (w*h)*0.0002:
            continue
        x,y,ww,hh = cv2.boundingRect(cnt)
        ar = ww / (hh + 1e-6)
        if 0.7 < ar < 1.3 and ww > 8 and hh > 8:
            # prefer near corners
            cx, cy = x + ww/2, y + hh/2
            dist_corner = min(
                (cx-0)**2 + (cy-0)**2,
                (cx-w)**2 + (cy-0)**2,
                (cx-w)**2 + (cy-h)**2,
                (cx-0)**2 + (cy-h)**2,
            )
            candidates.append((dist_corner, area, (cx, cy), (x,y,ww,hh)))

    if len(candidates) < 4:
        return None

    # pick 4 best by closeness to corners, then by area
    candidates.sort(key=lambda t: (t[0], -t[1]))
    pts = [c[2] for c in candidates[:8]]  # take more to be safe

    # cluster into 4 corners by quadrant
    pts = np.array(pts, dtype=np.float32)
    tl = pts[np.argmin(pts[:,0] + pts[:,1])]
    tr = pts[np.argmin((-pts[:,0]) + pts[:,1])]
    br = pts[np.argmin((-pts[:,0]) + (-pts[:,1]))]
    bl = pts[np.argmin((pts[:,0]) + (-pts[:,1]))]

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
    x0,y0,x1,y1 = bbox
    x0 = max(0, x0); y0 = max(0, y0)
    x1 = min(gray.shape[1]-1, x1); y1 = min(gray.shape[0]-1, y1)
    roi = gray[y0:y1, x0:x1]
    if roi.size == 0:
        return 0.0
    # invert: darker -> higher score
    return float(1.0 - (roi.mean() / 255.0))


def pick_one(scores: List[float], labels: List[str], min_score: float = 0.18, amb_delta: float = 0.03) -> Tuple[Optional[str], str, float, int]:
    """Pick one label based on scores.
    - If max < min_score -> BLANK
    - If top1 - top2 < amb_delta -> AMBIGUOUS
    """
    idxs = np.argsort(scores)[::-1]
    top = int(idxs[0])
    best = scores[top]
    if best < min_score:
        return None, "BLANK", best, top
    if len(scores) >= 2:
        second = scores[int(idxs[1])]
        if (best - second) < amb_delta:
            return labels[top], "AMBIGUOUS", best, top
    return labels[top], "OK", best, top


def process_page(warped: np.ndarray, zoom: float, num_questions: int) -> Tuple[Dict, np.ndarray]:
    can = warped.copy()
    gray = cv2.cvtColor(can, cv2.COLOR_BGR2GRAY)

    results: Dict = {}
    marks: List[Tuple[Tuple[int,int,int,int], str, str]] = []  # bbox, text, status

    # Grade (7-12)
    grade_scores = []
    grade_bboxes = []
    for x in GRADE_CIRCLE_XS:
        bbox = bubble_bbox_px(x, GRADE_CIRCLE_Y, GRADE_CIRCLE_RADIUS, zoom)
        grade_bboxes.append(bbox)
        grade_scores.append(score_bubble(gray, bbox))
    g_label = [str(v) for v in GRADE_VALUES]
    g_val, g_status, g_best, g_idx = pick_one(grade_scores, g_label, min_score=0.12, amb_delta=0.02)
    results["grade"] = g_val
    results["grade_status"] = g_status
    marks.append((grade_bboxes[g_idx], f"G:{g_val or ''}", g_status))

    # Class (single row digits 0-9)
    class_scores = []
    class_bboxes = []
    for x in CLASS_CIRCLE_XS:
        bbox = bubble_bbox_px(x, CLASS_CIRCLE_Y, CLASS_CIRCLE_RADIUS, zoom)
        class_bboxes.append(bbox)
        class_scores.append(score_bubble(gray, bbox))
    c_label = [str(v) for v in CLASS_VALUES]
    c_val, c_status, c_best, c_idx = pick_one(class_scores, c_label, min_score=0.12, amb_delta=0.02)
    results["class_no"] = c_val
    results["class_status"] = c_status
    marks.append((class_bboxes[c_idx], f"C:{c_val or ''}", c_status))

    # Seat number (00-99): top row = tens, bottom row = ones
    digit_labels = [str(d) for d in range(10)]
    def pick_digit(y_pt: float) -> Tuple[Optional[str], str, float, int, List[Tuple[int,int,int,int]]]:
        bboxes = []
        scores = []
        for i in range(10):
            x_pt = SEAT_START_X + i*SEAT_STEP_X
            bbox = bubble_bbox_px(x_pt, y_pt, SEAT_CIRCLE_RADIUS, zoom)
            bboxes.append(bbox)
            scores.append(score_bubble(gray, bbox))
        val, status, best, idx = pick_one(scores, digit_labels, min_score=0.14, amb_delta=0.02)
        return val, status, best, idx, bboxes

    tens, t_status, t_best, t_idx, t_bboxes = pick_digit(SEAT_TOP_ROW_Y)
    ones, o_status, o_best, o_idx, o_bboxes = pick_digit(SEAT_BOTTOM_ROW_Y)
    seat = (tens or "") + (ones or "")
    seat = seat if seat != "" else None
    results["seat_no"] = seat
    results["seat_status"] = "OK" if (t_status=="OK" and o_status=="OK") else ("AMBIGUOUS" if ("AMBIGUOUS" in [t_status,o_status]) else "BLANK")

    marks.append((t_bboxes[t_idx], f"T:{tens or ''}", t_status))
    marks.append((o_bboxes[o_idx], f"O:{ones or ''}", o_status))

    # Questions
    answers = []
    layout = compute_answer_layout(num_questions)
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
            for j, ch in enumerate(CHOICES):
                x_pt = xs[j]
                bbox = bubble_bbox_px(x_pt, y_pt, BUBBLE_RADIUS, zoom)
                bboxes.append(bbox)
                scores.append(score_bubble(gray, bbox))
            val, status, best, idx = pick_one(scores, CHOICES, min_score=0.18, amb_delta=0.03)
            answers.append(val if val is not None else "")
            marks.append((bboxes[idx], f"{q}:{val or ''}", status))
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
        else:
            color = (0,0,255)
        cv2.rectangle(annotated, (x0,y0), (x1,y1), color, 2)
        cv2.putText(annotated, text, (x0, max(10,y0-4)), cv2.FONT_HERSHEY_SIMPLEX, 0.4, color, 1, cv2.LINE_AA)

    return results, annotated


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
    dpi: int = 200,
):
    num_questions = max(1, min(MAX_QUESTIONS, int(num_questions)))

    doc = fitz.open(input_pdf_path)
    rows = []
    annotated_images = []

    for idx in range(doc.page_count):
        page = doc.load_page(idx)
        img, zoom = render_page(page, dpi=dpi)
        warped, _ = warp_to_canonical(img, zoom)
        result, annotated = process_page(warped, zoom, num_questions=num_questions)
        result["page"] = idx + 1
        rows.append(result)
        annotated_images.append(annotated)

    base_fields = [
        "page",
        "grade", "grade_status",
        "class_no", "class_status",
        "seat_no", "seat_status",
        "answers",
    ]
    q_fields = [f"Q{i+1}" for i in range(num_questions)]
    fieldnames = base_fields + q_fields
    extra_keys = sorted({k for row in rows for k in row.keys()} - set(fieldnames))
    fieldnames += extra_keys

    with open(out_csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: ("" if row.get(k) is None else row.get(k)) for k in fieldnames})

    images_to_pdf(annotated_images, out_annotated_pdf_path, dpi=dpi)
    doc.close()
