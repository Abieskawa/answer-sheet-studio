from __future__ import annotations

import csv
import math
import statistics
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import cv2
import numpy as np


def _normalize_cell(value: Any) -> str:
    return ("" if value is None else str(value)).strip().upper()


def _write_excel_csv(path: Path, header: List[str], rows: List[List[Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(header)
        writer.writerows(rows)


def _score_histogram_png(scores: List[float], total_possible: float, out_path: Path, lang: str = "zh_TW") -> None:
    lang_norm = (lang or "").strip().lower().replace("-", "_")
    is_zh = lang_norm.startswith("zh")
    s = {
        "no_scores": "沒有成績資料" if is_zh else "No score data",
        "title": "成績分佈" if is_zh else "Score distribution",
        "x": "分數" if is_zh else "Score",
        "y": "學生人數" if is_zh else "Students",
        "mean": "平均" if is_zh else "Mean",
        "median": "中位數" if is_zh else "Median",
        "quantiles": "百分位數 (P88/P75/P25/P12)" if is_zh else "Percentiles (P88/P75/P25/P12)",
    }

    width, height = 800, 450
    left, right, top, bottom = 70, 30, 50, 60
    plot_w = width - left - right
    plot_h = height - top - bottom

    img = np.full((height, width, 3), 255, dtype=np.uint8)

    if not scores:
        cv2.putText(img, s["no_scores"], (left, top + 30), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 0, 0), 2, cv2.LINE_AA)
        cv2.imwrite(str(out_path), img)
        return

    max_score = float(max(scores))
    max_x = max(float(total_possible or 0.0), max_score, 1.0)
    binwidth = 1.0
    if total_possible and total_possible > 0:
        binwidth = max(1.0, float(int(round(float(total_possible) / 20.0)) or 1))
    max_x = math.ceil(max_x / binwidth) * binwidth

    edges = np.arange(0.0, max_x + binwidth, binwidth, dtype=float)
    counts, _ = np.histogram(np.array(scores, dtype=float), bins=edges)
    y_max = int(max(int(counts.max(initial=0)), 1))

    x_scale = plot_w / float(max_x)
    y_scale = plot_h / float(y_max)

    axis_color = (0, 0, 0)
    bar_color = (233, 165, 14)  # BGR for #0ea5e9
    mean_color = (68, 68, 239)  # BGR for #ef4444
    median_color = (246, 130, 59)  # BGR-ish blue

    # Axes
    cv2.line(img, (left, height - bottom), (left + plot_w, height - bottom), axis_color, 2)
    cv2.line(img, (left, top), (left, height - bottom), axis_color, 2)

    # Y ticks (integer students)
    tick_color = (160, 160, 160)
    for yv in range(0, y_max + 1):
        y = int(height - bottom - yv * y_scale)
        cv2.line(img, (left - 5, y), (left, y), axis_color, 2)
        cv2.line(img, (left, y), (left + plot_w, y), tick_color, 1)
        cv2.putText(img, str(yv), (left - 42, y + 6), cv2.FONT_HERSHEY_SIMPLEX, 0.55, axis_color, 2, cv2.LINE_AA)

    # Bars
    for i, c in enumerate(counts.tolist()):
        x0 = int(left + edges[i] * x_scale)
        x1 = int(left + edges[i + 1] * x_scale)
        if x1 <= x0:
            x1 = x0 + 1
        y1 = height - bottom
        y0 = int(y1 - (float(c) * y_scale))
        if c > 0:
            cv2.rectangle(img, (x0, y0), (x1 - 1, y1), bar_color, thickness=-1)
            cv2.rectangle(img, (x0, y0), (x1 - 1, y1), (255, 255, 255), thickness=1)

    scores_arr = np.array(scores, dtype=float)
    mean_v = float(scores_arr.mean())
    q88, q75, q50, q25, q12 = (float(v) for v in np.quantile(scores_arr, [0.88, 0.75, 0.5, 0.25, 0.12]))
    median_v = q50

    def x_of(v: float) -> int:
        return int(left + max(0.0, min(max_x, float(v))) * x_scale)

    q_color = (90, 90, 90)
    for v in (q88, q75, q25, q12):
        cv2.line(img, (x_of(v), top), (x_of(v), height - bottom), q_color, 1)

    cv2.line(img, (x_of(mean_v), top), (x_of(mean_v), height - bottom), mean_color, 2)
    cv2.line(img, (x_of(median_v), top), (x_of(median_v), height - bottom), median_color, 2)

    cv2.putText(img, s["title"], (left, 32), cv2.FONT_HERSHEY_SIMPLEX, 0.9, axis_color, 2, cv2.LINE_AA)
    cv2.putText(img, s["x"], (left + plot_w // 2 - 20, height - 18), cv2.FONT_HERSHEY_SIMPLEX, 0.7, axis_color, 2, cv2.LINE_AA)
    cv2.putText(img, s["y"], (12, top + plot_h // 2), cv2.FONT_HERSHEY_SIMPLEX, 0.7, axis_color, 2, cv2.LINE_AA)

    # Legend for lines
    legend_x = left + plot_w - 290
    legend_y = top + 10
    legend_w = 280
    legend_h = 86
    cv2.rectangle(img, (legend_x, legend_y), (legend_x + legend_w, legend_y + legend_h), (255, 255, 255), thickness=-1)
    cv2.rectangle(img, (legend_x, legend_y), (legend_x + legend_w, legend_y + legend_h), (210, 210, 210), thickness=1)
    items = [
        (mean_color, s["mean"]),
        (median_color, s["median"]),
        (q_color, s["quantiles"]),
    ]
    for i, (c, label) in enumerate(items):
        y = legend_y + 24 + i * 22
        cv2.line(img, (legend_x + 12, y - 6), (legend_x + 52, y - 6), c, 3)
        cv2.putText(img, label, (legend_x + 62, y), cv2.FONT_HERSHEY_SIMPLEX, 0.6, axis_color, 2, cv2.LINE_AA)

    cv2.imwrite(str(out_path), img)


def _item_plot_png(numbers: List[int], series: Dict[str, List[Optional[float]]], out_path: Path, lang: str = "zh_TW") -> None:
    lang_norm = (lang or "").strip().lower().replace("-", "_")
    is_zh = lang_norm.startswith("zh")
    s = {
        "no_items": "沒有題目資料" if is_zh else "No item data",
        "title": "題目分析" if is_zh else "Item analysis",
        "x": "題號" if is_zh else "Question",
        "difficulty": "難度（正答率）" if is_zh else "Difficulty (accuracy)",
        "discrimination": "鑑別度" if is_zh else "Discrimination",
        "blank_rate": "空白率" if is_zh else "Blank rate",
    }

    width, height = 800, 700
    left, right, top, bottom = 70, 30, 40, 50
    panel_gap = 22

    metrics = list(series.keys())
    panel_h = int((height - top - bottom - panel_gap * (len(metrics) - 1)) / max(len(metrics), 1))
    plot_w = width - left - right

    img = np.full((height, width, 3), 255, dtype=np.uint8)
    axis = (0, 0, 0)
    line = (180, 180, 180)
    point = (120, 120, 120)

    if not numbers:
        cv2.putText(img, s["no_items"], (left, top + 30), cv2.FONT_HERSHEY_SIMPLEX, 1.0, axis, 2, cv2.LINE_AA)
        cv2.imwrite(str(out_path), img)
        return

    xs = np.array(numbers, dtype=float)
    x_min = float(xs.min())
    x_max = float(xs.max())
    if x_max <= x_min:
        x_max = x_min + 1.0

    def x_of(x: float) -> int:
        return int(left + (float(x) - x_min) / (x_max - x_min) * plot_w)

    for idx, metric in enumerate(metrics):
        y0 = top + idx * (panel_h + panel_gap)
        y1 = y0 + panel_h

        label = {
            "difficulty": s["difficulty"],
            "discrimination": s["discrimination"],
            "blank_rate": s["blank_rate"],
        }.get(metric, metric)
        cv2.putText(img, label, (left, y0 - 8), cv2.FONT_HERSHEY_SIMPLEX, 0.8, axis, 2, cv2.LINE_AA)
        cv2.rectangle(img, (left, y0), (left + plot_w, y1), (245, 245, 245), thickness=-1)
        cv2.rectangle(img, (left, y0), (left + plot_w, y1), axis, thickness=1)

        values = series[metric]
        ys = [v for v in values if v is not None and not math.isnan(float(v))]
        if not ys:
            continue
        y_min = float(min(ys))
        y_max = float(max(ys))
        if y_max <= y_min:
            y_max = y_min + 1.0

        def y_of(v: float) -> int:
            return int(y1 - (float(v) - y_min) / (y_max - y_min) * (panel_h - 8) - 4)

        last_pt: Optional[Tuple[int, int]] = None
        for x, v in zip(numbers, values):
            if v is None:
                last_pt = None
                continue
            vf = float(v)
            if math.isnan(vf):
                last_pt = None
                continue
            pt = (x_of(float(x)), y_of(vf))
            if last_pt is not None:
                cv2.line(img, last_pt, pt, line, 2)
            cv2.circle(img, pt, 2, point, thickness=-1)
            last_pt = pt

    cv2.putText(img, s["title"], (left, 28), cv2.FONT_HERSHEY_SIMPLEX, 0.9, axis, 2, cv2.LINE_AA)
    cv2.putText(img, s["x"], (left + plot_w // 2 - 18, height - 18), cv2.FONT_HERSHEY_SIMPLEX, 0.7, axis, 2, cv2.LINE_AA)
    cv2.imwrite(str(out_path), img)


def run_analysis_template(template_csv_path: Path, outdir: Path, lang: str = "zh_TW") -> None:
    template_csv_path = Path(template_csv_path)
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    with open(template_csv_path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError("template.csv is empty")
        fieldnames = [str(x) for x in reader.fieldnames]

        required = ["number", "correct", "points"]
        for k in required:
            if k not in fieldnames:
                raise ValueError(f"Missing column: {k}")

        student_cols = [c for c in fieldnames if c not in required]
        if not student_cols:
            raise ValueError("No student columns found")

        rows = list(reader)

    q_numbers: List[int] = []
    corrects: List[Optional[str]] = []
    points: List[float] = []
    answers_by_student: Dict[str, List[Optional[str]]] = {s: [] for s in student_cols}

    for row in rows:
        raw_no = (row.get("number") or "").strip()
        try:
            qno = int(raw_no)
        except ValueError:
            continue

        c = _normalize_cell(row.get("correct"))
        c_val = c if c else None
        p_raw = (row.get("points") or "").strip()
        try:
            p_val = float(p_raw) if p_raw != "" else 0.0
        except ValueError:
            p_val = 0.0

        q_numbers.append(qno)
        corrects.append(c_val)
        points.append(p_val)

        for s in student_cols:
            a = _normalize_cell(row.get(s))
            answers_by_student[s].append(a if a else None)

    q_count = len(q_numbers)
    s_count = len(student_cols)
    if q_count == 0 or s_count == 0:
        raise ValueError("No data rows")

    points_arr = np.array(points, dtype=float)
    has_key = np.array([c is not None for c in corrects], dtype=bool)
    total_possible = float(points_arr[has_key].sum()) if bool(has_key.any()) else 0.0

    # Student scores
    student_scores: Dict[str, float] = {}
    blank_counts: Dict[str, int] = {}
    for s in student_cols:
        score = 0.0
        blanks = 0
        ans_list = answers_by_student[s]
        for i in range(q_count):
            a = ans_list[i]
            c = corrects[i]
            if a is None:
                blanks += 1
            if c is None or a is None:
                continue
            if a == c:
                score += float(points_arr[i])
        student_scores[s] = float(score)
        blank_counts[s] = int(blanks)

    scores_rows = []
    for s in student_cols:
        score = float(student_scores[s])
        percent = (score / total_possible) if total_possible > 0 else None
        scores_rows.append(
            [s, round(score, 2), int(blank_counts[s]), round(total_possible, 2), (round(percent, 2) if percent is not None else "")],
        )

    scores_rows.sort(key=lambda r: (-float(r[1]), str(r[0])))
    _write_excel_csv(
        outdir / "analysis_scores.csv",
        ["person_id", "score", "blank_count", "total_possible", "percent"],
        scores_rows,
    )

    # Per-item stats
    choices = ["A", "B", "C", "D", "E"]
    item_rows = []

    all_scores = np.array([student_scores[s] for s in student_cols], dtype=float)
    student_order = np.argsort(all_scores, kind="mergesort")
    half = int(math.floor(s_count / 2))
    low_idx = student_order[:half] if half > 0 else np.array([], dtype=int)
    high_idx = student_order[-half:] if half > 0 else np.array([], dtype=int)

    # Build answer matrix [q, s]
    ans_mat = np.empty((q_count, s_count), dtype=object)
    for j, s in enumerate(student_cols):
        for i, a in enumerate(answers_by_student[s]):
            ans_mat[i, j] = a

    for i, qno in enumerate(q_numbers):
        c = corrects[i]
        p = float(points_arr[i])
        row_ans = ans_mat[i, :]

        is_blank = np.array([a is None for a in row_ans], dtype=bool)
        blank_rate = float(is_blank.mean())
        multi_rate = float(np.mean([(a is not None and len(str(a)) > 1) for a in row_ans]))
        other_rate = float(
            np.mean(
                [
                    (a is not None and len(str(a)) <= 1 and str(a) not in choices)
                    for a in row_ans
                ]
            )
        )

        diff: Optional[float]
        disc: Optional[float]
        if c is None:
            diff = None
            disc = None
        else:
            correct_mask = np.array([(a == c) for a in row_ans], dtype=bool)
            diff = float(correct_mask.mean())
            if low_idx.size and high_idx.size:
                disc = float(correct_mask[high_idx].mean() - correct_mask[low_idx].mean())
            else:
                disc = None

        p_choice = {ch: float(np.mean([(a == ch) for a in row_ans])) for ch in choices}

        item_rows.append(
            [
                qno,
                c or "",
                (int(p) if abs(p - round(p)) < 1e-9 else round(p, 2)),
                (round(diff, 2) if diff is not None else ""),
                (round(disc, 2) if disc is not None else ""),
                round(blank_rate, 2),
                round(multi_rate, 2),
                round(other_rate, 2),
                *[round(p_choice[ch], 2) for ch in choices],
            ]
        )

    _write_excel_csv(
        outdir / "analysis_item.csv",
        ["number", "correct", "points", "difficulty", "discrimination", "blank_rate", "multi_rate", "other_rate", "p_A", "p_B", "p_C", "p_D", "p_E"],
        item_rows,
    )

    # Summary
    mean_v = float(statistics.mean(all_scores))
    sd_v = float(statistics.pstdev(all_scores) if len(all_scores) <= 1 else statistics.stdev(all_scores))
    q88, q75, q50, q25, q12 = np.quantile(all_scores, [0.88, 0.75, 0.5, 0.25, 0.12])

    _write_excel_csv(
        outdir / "analysis_summary.csv",
        ["students", "questions", "total_possible", "mean", "sd", "p88", "p75", "median", "p25", "p12"],
        [
            [
                int(s_count),
                int(q_count),
                round(total_possible, 2),
                round(mean_v, 2),
                round(sd_v, 2),
                round(float(q88), 2),
                round(float(q75), 2),
                round(float(q50), 2),
                round(float(q25), 2),
                round(float(q12), 2),
            ]
        ],
    )

    _score_histogram_png([float(x) for x in all_scores.tolist()], total_possible, outdir / "analysis_score_hist.png", lang=lang)
    _item_plot_png(
        q_numbers,
        {
            "difficulty": [None if v == "" else float(v) for v in [r[3] for r in item_rows]],
            "discrimination": [None if v == "" else float(v) for v in [r[4] for r in item_rows]],
        },
        outdir / "analysis_item_plot.png",
        lang=lang,
    )
