from __future__ import annotations

import csv
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
        if (best - second) < amb_delta:
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
    gray = cv2.cvtColor(can, cv2.COLOR_BGR2GRAY)

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
        grade_scores, g_label, min_score=0.05, amb_delta=0.02
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

    # Class (single row digits 0-9)
    class_scores = []
    class_bboxes = []
    for x in CLASS_CIRCLE_XS:
        bbox = bubble_bbox_px(x, CLASS_CIRCLE_Y, CLASS_CIRCLE_RADIUS, zoom)
        class_bboxes.append(bbox)
        class_scores.append(score_bubble(gray, bbox))
    c_label = [str(v) for v in CLASS_VALUES]
    c_val, c_status, _, c_idx, c_second_label, _ = pick_one(
        class_scores, c_label, min_score=0.05, amb_delta=0.02
    )
    results["class_no"] = c_val if c_status == "OK" else None
    results["class_status"] = c_status
    marks.append((class_bboxes[c_idx], f"C:{c_val or ''}", c_status))
    if c_status != "OK":
        flags.append(
            {
                "field": "class_no",
                "question": "",
                "status": c_status,
                "best_label": c_label[c_idx],
                "second_label": c_second_label or "",
            }
        )

    # Seat number (00-99): top row = tens, bottom row = ones
    digit_labels = [str(d) for d in range(10)]
    def pick_digit(
        y_pt: float,
    ) -> Tuple[Optional[str], str, float, int, Optional[str], float, List[Tuple[int, int, int, int]]]:
        bboxes = []
        scores = []
        for i in range(10):
            x_pt = SEAT_START_X + i*SEAT_STEP_X
            bbox = bubble_bbox_px(x_pt, y_pt, SEAT_CIRCLE_RADIUS, zoom)
            bboxes.append(bbox)
            scores.append(score_bubble(gray, bbox))
        val, status, best, idx, second_label, second_score = pick_one(
            scores, digit_labels, min_score=0.06, amb_delta=0.02
        )
        return val, status, best, idx, second_label, second_score, bboxes

    tens, t_status, _, t_idx, t_second_label, _, t_bboxes = pick_digit(SEAT_TOP_ROW_Y)
    ones, o_status, _, o_idx, o_second_label, _, o_bboxes = pick_digit(SEAT_BOTTOM_ROW_Y)

    tens_out = tens if t_status == "OK" else None
    ones_out = ones if o_status == "OK" else None
    if tens_out is not None and ones_out is not None:
        results["seat_no"] = f"{tens_out}{ones_out}"
    else:
        results["seat_no"] = None
    results["seat_status"] = "OK" if (t_status=="OK" and o_status=="OK") else ("AMBIGUOUS" if ("AMBIGUOUS" in [t_status,o_status]) else "BLANK")

    marks.append((t_bboxes[t_idx], f"T:{tens or ''}", t_status))
    marks.append((o_bboxes[o_idx], f"O:{ones or ''}", o_status))
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
                scores, choices, min_score=0.08, amb_delta=0.03
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
