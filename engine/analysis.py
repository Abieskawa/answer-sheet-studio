from __future__ import annotations

import csv
import math
import re
import statistics
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
from PIL import Image as PILImage, ImageDraw, ImageFont
import os # Added for os.path.exists

from .xlsx import write_simple_xlsx_multi


def _normalize_cell(value: Any) -> str:
    return ("" if value is None else str(value)).strip().upper()


def _write_excel_csv(path: Path, header: List[str], rows: List[List[Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(header)
        writer.writerows(rows)


def _score_histogram_png(scores: List[float], total_possible: float, out_path: Path, lang: str = "zh_TW") -> None:
    try:
        import cv2  # type: ignore
    except (ImportError, OSError) as e:
        print(f"WARNING: Failed to import cv2: {e}")
        return

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
    left, right, top, bottom = 90, 30, 50, 60
    plot_w = width - left - right
    plot_h = height - top - bottom

    img = np.full((height, width, 3), 255, dtype=np.uint8)

    if not scores:
        cv2.putText(img, s["no_scores"], (left, top + 30), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 0, 0), 2, cv2.LINE_AA)
        # Use imencode to avoid Unicode path issues on Windows
        success, encoded = cv2.imencode('.png', img)
        if success:
            out_path.write_bytes(encoded.tobytes())
        else:
            raise RuntimeError("Failed to encode PNG image")
        return

    max_score = float(max(scores))
    max_x = max(float(total_possible or 0.0), max_score, 1.0)
    binwidth = 1.0
    if total_possible and total_possible > 0:
        binwidth = max(1.0, float(int(round(float(total_possible) / 20.0)) or 1))
    max_x = math.ceil(max_x / binwidth) * binwidth

    edges = np.arange(0.0, max_x + binwidth, binwidth, dtype=float)
    counts, _ = np.histogram(np.array(scores, dtype=float), bins=edges)
    # Give 25% room at the top for labels
    y_max_raw = int(counts.max(initial=0))
    y_max = int(max(int(y_max_raw * 1.25 + 1), 1))

    x_scale = plot_w / float(max_x)
    y_scale = plot_h / float(y_max)

    axis_color = (0, 0, 0)
    bar_color = (180, 180, 180)  # Gray
    mean_color = (68, 68, 239)  # BGR for #ef4444
    median_color = (246, 130, 59)  # BGR-ish blue

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

    # Axes
    cv2.line(img, (left, height - bottom), (left + plot_w, height - bottom), axis_color, 2)
    cv2.line(img, (left, top), (left, height - bottom), axis_color, 2)

    # Y ticks (integer students)
    tick_color = (160, 160, 160)
    for yv in range(0, y_max + 1):
        y = int(height - bottom - yv * y_scale)
        cv2.line(img, (left - 5, y), (left, y), axis_color, 2)
        cv2.line(img, (left, y), (left + plot_w, y), tick_color, 1)
        cv2.putText(img, str(yv), (left - 45, y + 6), cv2.FONT_HERSHEY_SIMPLEX, 0.55, axis_color, 2, cv2.LINE_AA)

    # --- 4. 新增：畫 X 軸刻度與分數數字 ---
    # 每 10 分標註一個刻度
    for xv in range(0, int(max_x) + 1, 10):
        x = int(left + xv * x_scale)
        if x > left + plot_w: break
        # 畫 X 軸小刻度線
        cv2.line(img, (x, height - bottom), (x, height - bottom + 5), axis_color, 2)
        # 標註分數數字
        cv2.putText(img, str(xv), (x - 10, height - bottom + 25), cv2.FONT_HERSHEY_SIMPLEX, 0.55, axis_color, 2, cv2.LINE_AA)

    

    scores_arr = np.array(scores, dtype=float)
    mean_v = float(scores_arr.mean())
    q88, q75, q50, q25, q12 = (float(v) for v in np.quantile(scores_arr, [0.88, 0.75, 0.5, 0.25, 0.12]))
    median_v = q50

    def x_of(v: float) -> int:
        return int(left + max(0.0, min(max_x, float(v))) * x_scale)

    q_color = (90, 90, 90)
    
    # Draw dashed lines for percentiles (Thicker line as requested: thickness 2)
    for v in (q88, q75, q25, q12):
        vx = x_of(v)
        # Draw dashed line manually
        dash_len = 5
        for y in range(top, height - bottom, dash_len * 2):
            cv2.line(img, (vx, y), (vx, min(y + dash_len, height - bottom)), q_color, 2)

    cv2.line(img, (x_of(mean_v), top), (x_of(mean_v), height - bottom), mean_color, 2)
    cv2.line(img, (x_of(median_v), top), (x_of(median_v), height - bottom), median_color, 2)

    # Use Pillow for all text in Histogram
    pil_img = PILImage.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(pil_img)
    
    font_paths = ["C:\\Windows\\Fonts\\msjh.ttc", "C:\\Windows\\Fonts\\simsun.ttc"]
    font = None
    for fp in font_paths:
        if os.path.exists(fp):
            try:
                font_title = ImageFont.truetype(fp, 22)
                font_axis = ImageFont.truetype(fp, 16)
                font = fp
                break
            except Exception: pass
            
    if font:
        font_label = ImageFont.truetype(font, 13)
        draw.text((left, 10), s["title"], font=font_title, fill=(0,0,0))
        draw.text((left + plot_w // 2, height - 25), s["x"], font=font_axis, fill=(0,0,0))
        # Move "Students" label to top left to avoid any overlap
        draw.text((10, top - 35), s["y"], font=font_axis, fill=(0,0,0))
        
        
        # Draw labels with background boxes for better aesthetics
        # --- 修正版：智慧避讓與精確對齊 ---
        labels = [
            (mean_v, s["mean"], (239, 68, 68), (255, 255, 255)),   # 紅色底
            (q50, s["median"], (59, 130, 246), (255, 255, 255)),   # 深灰底
            (q88, "P88", (130, 130, 130), (255, 255, 255)),      # 灰底
            (q75, "P75", (130, 130, 130), (255, 255, 255)),
            (q25, "P25", (130, 130, 130), (255, 255, 255)),
            (q12, "P12", (130, 130, 130), (255, 255, 255))
        ]
        labels.sort(key=lambda x: x[0])  # 依分數從低到高排序，方便避讓計算

        last_x1 = -999.0      # 紀錄前一個標籤框的右邊界
        current_layer = 0     # 當前使用的垂直層級索引
        layers = [2, 22, 42]  # 定義三層不同的 Y 軸偏移，間距加大以防垂直重疊

        for i, (val, txt, bg_clr, txt_clr) in enumerate(labels):
            lx = x_of(val)
            
            # 1. 精確計算文字尺寸與內部偏移
            bbox = draw.textbbox((0, 0), txt, font=font_label)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]
            
            # 2. 智慧避讓邏輯：如果當前 X 座標太靠近前一個標籤的右邊
            pad_x = 6  # 左右內邊距
            pad_y = 3  # 上下內邊距
            this_x0 = lx - (tw / 2) - pad_x
            
            if this_x0 < last_x1 + 8:  # 距離小於 8 像素就換層
                current_layer = (current_layer + 1) % len(layers)
            else:
                current_layer = 0 # 距離夠遠則回到第一層
            
            ly_base = top + layers[current_layer]
            
            # 3. 計算背景框位置（以 lx 為水平中心點）
            box_x0 = lx - (tw / 2) - pad_x
            box_y0 = ly_base - pad_y
            box_x1 = lx + (tw / 2) + pad_x
            box_y1 = ly_base + th + pad_y
            
            # 4. 繪製背景矩形
            draw.rectangle([box_x0, box_y0, box_x1, box_y1], fill=bg_clr)
            
            # 5. 繪製文字（修正 bbox 偏移，達到完美中心對齊）
            # text_draw_x = 中心點 - (文字寬度/2) - 文字本身的內部起始位移
            text_draw_x = lx - (tw / 2) - bbox[0]
            text_draw_y = ly_base - bbox[1]
            
            draw.text((text_draw_x, text_draw_y), txt, font=font_label, fill=txt_clr)
            
            # 更新最後右邊界，供下一個標籤判斷
            last_x1 = box_x1
        # --- 修正結束 ---



        img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    else:
        # Emergency fallback
        cv2.putText(img, "Scores", (left, 32), cv2.FONT_HERSHEY_SIMPLEX, 0.9, axis_color, 2, cv2.LINE_AA)

    # Use imencode to avoid Unicode path issues on Windows
    success, encoded = cv2.imencode('.png', img)
    if success:
        out_path.write_bytes(encoded.tobytes())
    else:
        raise RuntimeError("Failed to encode PNG image")


def _item_plot_png(numbers: List[int], series: Dict[str, List[Optional[float]]], out_path: Path, lang: str = "zh_TW") -> None:
    try:
        import cv2  # type: ignore
    except (ImportError, OSError):
        return

    lang_norm = (lang or "").strip().lower().replace("-", "_")
    # Fixed dimensions for A3 landscape alignment
    # Total width ~1130 (A3 landscape printable), Left margin 100 matches Table ID+Score (60+40)
    # ABSOLUTE CALIBRATION: Total width 1130. Left 100 aligns with Table Labels (60+40).
    # Question area is exactly 1030. Right margin 0 to prevent drift.
    width, height = 1130, 300
    left, right, top, bottom = 110, 0, 20, 20
    panel_gap = 22

    is_zh = lang_norm.startswith("zh")
    s = {
        "no_items": "沒有試題資料" if is_zh else "No item data",
        "title": "試題分析趨勢圖" if is_zh else "Item Analysis Trends",
        "difficulty": "正確率" if is_zh else "Correct Rate",
    }

    # Only keep Correct Rate (Difficulty) as requested for vertical space
    metrics = ["difficulty"] 
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
    q_count = len(numbers)
    # Unit width per question to match table: (Plot Width / Number of Questions)
    unit_w = plot_w / max(q_count, 1)

    def x_of(i):
        # i is 0-indexed. 
        # For precision alignment, we must use the EXACT SAME logic as the PDF table.
        # Table col widths: [60, 40, q_w, q_w, ...]
        # q_w = 1030 / q_count
        # The center of question N (0-indexed) starts at offset 100.
        return int(left + (i + 0.5) * unit_w)

    for idx, metric in enumerate(metrics):
        y0 = top + idx * (panel_h + panel_gap)
        y1 = y0 + panel_h
        
        # Background and Grid Lines for Vertical Correspondence
        cv2.rectangle(img, (left, y0), (left + plot_w, y1), (245, 245, 245), thickness=-1)
        
        # DRAW VERTICAL GRID LINES: High contrast (120) for visual guidance across the report
        grid_color = (220, 220, 220) 
        for i in range(q_count + 1):
            gx = int(left + i * unit_w)
            cv2.line(img, (gx, 0), (gx, height), grid_color, 1) # Full height
            
        cv2.rectangle(img, (left, y0), (left + plot_w, y1), axis, thickness=1)

        values = series[metric]
        ys = [v for v in values if v is not None and not math.isnan(float(v))]
        if not ys:
            continue
        if metric == "difficulty":
            y_min, y_max = 0.0, 1.0
        else:
            ys = [v for v in values if v is not None and not math.isnan(float(v))]
            if not ys:
                continue
            y_min = float(min(ys))
            y_max = float(max(ys))
            if y_max <= y_min:
                y_max = y_min + 1.0

        def y_of(v: float) -> int:
            return int(y1 - (float(v) - y_min) / (y_max - y_min) * (panel_h - 8) - 4)

        # DRAW HORIZONTAL GRID LINES & LABELS
        for level in [0, 0.25, 0.5, 0.75, 1.0]:
            ly = y_of(level)
            cv2.line(img, (left, ly), (left + plot_w, ly), (220, 220, 220), 1)
            # Add short tick line
            cv2.line(img, (left - 5, ly), (left, ly), (80, 80, 80), 1)
            # Add labels with OpenCV (Always visible, closer to axis)
            label = f"{level:.2f}"
            cv2.putText(img, label, (left - 35, ly + 4), cv2.FONT_HERSHEY_SIMPLEX, 0.35, (80, 80, 80), 1, cv2.LINE_AA)

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

    # Use Pillow to draw Chinese text on the OpenCV image
    pil_img = PILImage.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(pil_img)
    
    # Try to find a system Chinese font
    font_paths = [
        "C:\\Windows\\Fonts\\msjh.ttc",    # Microsoft JhengHei
        "C:\\Windows\\Fonts\\simsun.ttc",  # SimSun
        "C:\\Windows\\Fonts\\mingliu.ttc", # MingLiU
    ]
    font_path = None
    for fp in font_paths:
        if os.path.exists(fp):
            try:
                # Use a slightly smaller font for the axis labels
                font_axis = ImageFont.truetype(fp, 14)
                font_path = fp
                break
            except Exception: pass
            
    if font_path:
        # Draw Y-axis labels in the left margin area (difficulty only for now)
        for idx, metric in enumerate(metrics):
            if metric == "difficulty":
                # Ensure we use exactly the same scaling as the background grid
                y0_p = top + idx * (panel_h + panel_gap)
                y1_p = y0_p + panel_h
                def y_of_p(v: float) -> int:
                    return int(y1_p - (float(v) - 0.0) / (1.0 - 0.0) * (panel_h - 8) - 4)
                
                for level in [0, 0.25, 0.5, 0.75, 1.0]:
                    ly = y_of_p(level)
                    label = f"{level:.2f}"
                    bbox = draw.textbbox((0, 0), label, font=font_axis)
                    tw = bbox[2] - bbox[0]
                    th = bbox[3] - bbox[1]
                    draw.text((left - tw - 8, ly - th // 2), label, font=font_axis, fill=(80, 80, 80))

        img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    else:
        # Fallback to English if no font found (unlikely on Windows)
        cv2.putText(img, "Y-Axis labels added", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (120, 120, 120), 1, cv2.LINE_AA)


    # Use imencode to avoid Unicode path issues on Windows
    success, encoded = cv2.imencode('.png', img)
    if success:
        out_path.write_bytes(encoded.tobytes())
    else:
        raise RuntimeError("Failed to encode PNG image")


def _write_analysis_scores_by_class_xlsx(outdir: Path) -> None:
    """
    Reads analysis_scores.csv and roster.csv (if available) to group scores by class.
    Writes analysis_scores_by_class.xlsx with one sheet per class.
    """
    scores_path = outdir / "analysis_scores.csv"
    roster_path = outdir / "roster.csv"
    if not scores_path.exists():
        return

    # Read scores
    try:
        with open(scores_path, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            score_rows = list(reader)
    except Exception:
        return

    if not score_rows:
        return

    # Read roster to map person_id -> class_no
    class_map = {}
    if roster_path.exists():
        try:
            with open(roster_path, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    pid = (row.get("person_id") or "").strip()
                    c = (row.get("class_no") or "").strip()
                    if pid and c:
                        class_map[pid] = c
        except Exception:
            pass

    # Group by class
    # If not in roster, try to parse from ID (e.g. 70101 -> class 01? No, usually random)
    # Default to "Unknown" if not found.
    by_class: Dict[str, List[List[Any]]] = {}
    
    # Headers for Excel
    header = ["座號/ID", "分數", "空白數", "總分", "百分比"]

    for row in score_rows:
        pid = row.get("person_id", "")
        cls = class_map.get(pid, "Unknown")
        
        # Prepare row data
        # roster: person_id, grade, class_no, seat_no, page
        # We might want "seat_no" as the first column if available, else pid.
        # But analysis_scores.csv only has person_id.
        # Let's try to find seat_no from roster too.
        seat = pid
        # If we have roster data, maybe usage seat_no?
        # The user's request usually implies they want to see "Seat No" in the class sheet.
        # Let's just use person_id for now as it aligns with "座號/ID".
        
        data = [
            pid,
            float(row.get("得分", 0)),
            int(row.get("空白數", 0)),
            float(row.get("滿分", 0)),
            row.get("百分比", "")
        ]
        
        if cls not in by_class:
            by_class[cls] = [header]
        by_class[cls].append(data)

    # Sort keys (classes) naturally?
    sorted_classes = sorted(by_class.keys(), key=lambda k: (len(k), k))
    
    sheets = []
    for cls in sorted_classes:
        # Sort rows in sheet by seat/ID
        # rows[0] is header, rows[1:] are data
        rows = by_class[cls]
        header_row = rows[0]
        data_rows = rows[1:]
        # Sort by ID
        data_rows.sort(key=lambda r: str(r[0]))
        sheets.append((str(cls), [header_row] + data_rows))

    if sheets:
        try:
            write_simple_xlsx_multi(outdir / "analysis_scores_by_class.xlsx", sheets)
        except Exception:
            pass


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
        ["學號/ID", "得分", "空白數", "滿分", "百分比"],
        scores_rows,
    )
    _write_analysis_scores_by_class_xlsx(outdir)

    # Per-item stats
    choices = ["A", "B", "C", "D", "E"]
    item_rows = []

    all_scores = np.array([student_scores[s] for s in student_cols], dtype=float)
    student_order = np.argsort(all_scores, kind="mergesort")
    # Discrimination index: mean(correct_high) - mean(correct_low).
    # If students > 30, use top/bottom 27%; otherwise use top/bottom 50%.
    group_n = int(math.floor(s_count * 0.27)) if s_count > 30 else int(math.floor(s_count / 2.0))
    low_idx = student_order[:group_n] if group_n > 0 else np.array([], dtype=int)
    high_idx = student_order[-group_n:] if group_n > 0 else np.array([], dtype=int)

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
        ["題號", "正解", "配分", "正確率", "鑑別度", "空白率", "複選率", "其他錯誤率", "A百分比", "B百分比", "C百分比", "D百分比", "E百分比"],
        item_rows,
    )

    # Summary
    mean_v = float(statistics.mean(all_scores))
    sd_v = float(statistics.pstdev(all_scores) if len(all_scores) <= 1 else statistics.stdev(all_scores))
    q88, q75, q50, q25, q12 = np.quantile(all_scores, [0.88, 0.75, 0.5, 0.25, 0.12])

    _write_excel_csv(
        outdir / "analysis_summary.csv",
        ["學生人數", "總題數", "滿分", "平均分", "標準差", "P88", "P75", "中位數", "P25", "P12"],
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

    # Generate PNG charts with error handling
    try:
        print(f"DEBUG: Generating score histogram to {outdir / 'analysis_score_hist.png'}")
        _score_histogram_png([float(x) for x in all_scores.tolist()], total_possible, outdir / "analysis_score_hist.png", lang=lang)
        print(f"DEBUG: Score histogram generated successfully")
    except Exception as e:
        print(f"ERROR: Failed to generate score histogram: {e}")
        import traceback
        traceback.print_exc()
    
    try:
        print(f"DEBUG: Generating item plot to {outdir / 'analysis_item_plot.png'}")
        _item_plot_png(
            q_numbers,
            {
                "difficulty": [None if v == "" else float(v) for v in [r[3] for r in item_rows]],
                "discrimination": [None if v == "" else float(v) for v in [r[4] for r in item_rows]],
            },
            outdir / "analysis_item_plot.png",
            lang=lang,
        )
        print(f"DEBUG: Item plot generated successfully")
    except Exception as e:
        print(f"ERROR: Failed to generate item plot: {e}")
        import traceback
        traceback.print_exc()



def generate_integrated_report(job_dir: Path, lang: str = "zh_TW") -> None:
    """
    Generates an integrated Excel and PDF report with a horizontal layout:
    Students as rows, Questions as columns.
    Layout: [Student Answers] -> [Graphs] -> [Item Statistics]
    """
    item_csv = job_dir / "analysis_item.csv"
    scores_csv = job_dir / "analysis_scores.csv"
    template_csv = job_dir / "analysis_template.csv"
    roster_csv = job_dir / "roster.csv"

    if not (item_csv.exists() and scores_csv.exists() and template_csv.exists()):
        return

    lang_norm = (lang or "").strip().lower().replace("-", "_")
    is_zh = lang_norm.startswith("zh")

    # 1. Read Data
    try:
        with open(item_csv, "r", encoding="utf-8-sig") as f:
            item_stats_list = list(csv.DictReader(f))
        with open(scores_csv, "r", encoding="utf-8-sig") as f:
            student_scores_list = list(csv.DictReader(f))
        with open(template_csv, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            matrix_rows = list(reader)
            # Find student columns (anything after number, correct, points)
            all_cols = reader.fieldnames or []
            student_cols = [c for c in all_cols if c not in ("number", "correct", "points")]
    except Exception:
        return

    # Maps
    item_map = {int(r["題號"] if "題號" in r else r.get("number", 0)): r for r in item_stats_list}
    score_map = {r["學號/ID"] if "學號/ID" in r else r.get("person_id", ""): r for r in student_scores_list}
    q_numbers = sorted(item_map.keys())

    # 2. Build Horizontal Matrix for Excel
    # Rows:
    # Header: [ID, Score] + Questions...
    # Students: [PID, Score] + Answers...
    # Spacer
    # Metrics Header: ["分析數據", ""] + Questions...
    # Correct Ans: ["正解", ""] + ...
    # Difficulty: ["正確率", ""] + ...
    # Discrimination: ["鑑別度", ""] + ...

    l_id = "學號/ID" if is_zh else "ID"
    l_score = "總分" if is_zh else "Score"
    # Simplified: just use question numbers without prefix
    l_ans = "正解" if is_zh else "Ans"
    l_diff = "正確率" if is_zh else "Diff"
    l_disc = "鑑別度" if is_zh else "Disc"

    def natural_key(text):
        return [int(c) if c.isdigit() else c.lower() for c in re.split('([0-9]+)', str(text))]

    # Sort students by original order or naturally?
    # Use natural sort for student IDs to avoid "10" before "2"
    student_cols.sort(key=natural_key)

    excel_rows = []
    # Answers Section
    # Simplified headers: just question numbers
    header_ans = [l_id, l_score] + [str(q) for q in q_numbers]
    excel_rows.append(header_ans)

    for pid in student_cols:
        score = score_map.get(pid, {}).get("得分", score_map.get(pid, {}).get("score", ""))
        row = [pid, score]
        # Get answers for this student across all questions
        # We need a map from qno to answer for this student
        ans_map = {int(r["number"]): r.get(pid, "") for r in matrix_rows}
        for q in q_numbers:
            row.append(ans_map.get(q, ""))
        excel_rows.append(row)

    excel_rows.append([]) # Spacer

    # Metrics Section
    # Simplified headers: just question numbers to align with student answer columns
    m_header = ["試題屬性", ""] + [str(q) for q in q_numbers]
    excel_rows.append(m_header)
    
    ans_row = [l_ans, ""] + [item_map.get(q, {}).get("正解", item_map.get(q, {}).get("correct", "")) for q in q_numbers]
    diff_row = [l_diff, ""] + [item_map.get(q, {}).get("正確率", item_map.get(q, {}).get("difficulty", "")) for q in q_numbers]
    disc_row = [l_disc, ""] + [item_map.get(q, {}).get("鑑別度", item_map.get(q, {}).get("discrimination", "")) for q in q_numbers]
    
    excel_rows.append(ans_row)
    excel_rows.append(diff_row)
    excel_rows.append(disc_row)

    # Write Excel
    out_xlsx = job_dir / "試題分析整合檔.xlsx"
    try:
        from .xlsx import write_simple_xlsx
        write_simple_xlsx(out_xlsx, excel_rows)
    except Exception:
        pass

    # 3. Generate PDF (Standard Integrated Report)
    # [DELETE] Old version removed to unify output
    
    # 4. Generate PDF (Improved Precision Alignment Version)
    out_pdf_v2 = job_dir / "試題分析整合報表.pdf" # Renamed to main name
    _generate_integrated_pdf_v2(
        out_pdf_v2, 
        student_header=header_ans, 
        student_data=excel_rows[1:1+len(student_cols)],
        metric_header=m_header,
        metric_data=[ans_row, diff_row, disc_row],
        job_dir=job_dir, 
        lang=lang, 
        is_zh=is_zh,
        is_precision=True
    )


from reportlab.platypus import Flowable
class PrecisionGridChart(Flowable):
    def __init__(self, data_series, col_widths, height=200, font_name="Helvetica"):
        Flowable.__init__(self)
        processed = []
        for v in data_series:
            try:
                val = float(v)
                if math.isnan(val): processed.append(None)
                else: processed.append(val)
            except (ValueError, TypeError):
                processed.append(None)
        self.data = processed
        self.col_widths = col_widths
        # 這裡決定了畫布的總寬度（只包含問題區）
        self.q_ws = col_widths[2:] 
        self.width = sum(self.q_ws)
        self.height = height
        self.font_name = font_name

    def get_xy(self, i, val, y_min, y_max):
        """
        將原本在 draw 內部的 get_xy 移出來，
        這樣外部呼叫就不會報錯。
        """
        # 使用 q_ws 確保 x 軸相對於合併單元格的起點對齊
        x = sum(self.q_ws[:i]) + self.q_ws[i] / 2
        
        # Y 軸計算，增加安全除法防止 y_max == y_min
        denom = (y_max - y_min) if y_max != y_min else 1.0
        y = 10 + (val - y_min) / denom * (self.height - 20)
        return x, y

    def draw(self):
        canvas = self.canv
        q_ws = self.q_ws
        
        # 1. 繪製背景網格（垂直線）
        canvas.setStrokeColorRGB(0.9, 0.9, 0.9)
        canvas.setLineWidth(0.3)
        curr_x = 0
        for w in q_ws[:-1]:
            curr_x += w
            canvas.line(curr_x, 0, curr_x, self.height)
        
        # 1.1 繪製水平背景格線與 Y 軸標籤
        canvas.setStrokeColorRGB(0.85, 0.85, 0.85)
        canvas.setFont(self.font_name, 7)
        canvas.setFillColorRGB(0.4, 0.4, 0.4)
        for level in [0, 0.25, 0.5, 0.75, 1.0]:
            _, ly = self.get_xy(0, level, 0.0, 1.0)
            # 繪製水平線 (橫跨整個問題區)
            canvas.line(0, ly, self.width, ly)
            # 繪製短刻度線 (幫助視覺對準)
            canvas.line(-5, ly, 0, ly)
            # 文字標記 (距離縮近至 -20)
            canvas.drawString(-20, ly - 2, f"{level:.2f}")
        
        # 1. 數據繪製
        valid_vals = [v for v in self.data if v is not None]
        if not valid_vals: return
        
        # 強制使用 0-1 範圍以對齊標籤
        y_min, y_max = 0.0, 1.0

        points = []
        for i, val in enumerate(self.data):
            if val is not None:
                # 呼叫移到外部的 get_xy
                points.append(self.get_xy(i, val, y_min, y_max))
        
        # 繪製線條
        canvas.setStrokeColorRGB(0.2, 0.2, 0.2)
        canvas.setLineWidth(1.5)
        if points:
            p = canvas.beginPath()
            p.moveTo(points[0][0], points[0][1])
            for x, y in points[1:]:
                p.lineTo(x, y)
            canvas.drawPath(p)
            
        # 繪製藍點
        canvas.setFillColorRGB(0.3, 0.3, 0.8)
        for x, y in points:
            canvas.circle(x, y, 2.5, fill=1, stroke=0)


def _generate_integrated_pdf_v2(
    out_path: Path, 
    student_header: List[str], 
    student_data: List[List[Any]], 
    metric_header: List[str],
    metric_data: List[List[Any]],
    job_dir: Path, 
    lang: str, 
    is_zh: bool,
    is_precision: bool = False
):
    # Translations for PDF internal labels
    t = {
        "label_trend": "答對率趨勢線" if is_zh else "Correct Rate Trend",
        "label_metrics": "試題指標數據" if is_zh else "Item Metrics",
    }
    try:
        from reportlab.lib import pagesizes, colors
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Spacer, PageBreak, Paragraph
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    except ImportError:
        return

    # Register a reliable Windows Chinese font instead of CID fonts
    font_name = "Helvetica"
    if is_zh:
        # Priority list for Windows Chinese fonts
        sys_fonts = [
            ("MSJH", "C:\\Windows\\Fonts\\msjh.ttc"),    # Microsoft JhengHei
            ("MSJH_REG", "C:\\Windows\\Fonts\\msjh.ttf"),
            ("SIMSUN", "C:\\Windows\\Fonts\\simsun.ttc"),
        ]
        registered = False
        for name, path in sys_fonts:
            if os.path.exists(path):
                try:
                    from reportlab.pdfbase.ttfonts import TTFont
                    pdfmetrics.registerFont(TTFont(name, path))
                    font_name = name
                    registered = True
                    break
                except Exception: pass
        
        if not registered:
            # Fallback to CID if no TTF found
            try:
                pdfmetrics.registerFont(UnicodeCIDFont("MSung-Light"))
                font_name = "MSung-Light"
            except Exception: pass
    
    doc = SimpleDocTemplate(
        str(out_path), 
        pagesize=pagesizes.landscape(pagesizes.A3), 
        leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30
    )
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=18,
        alignment=1, # Center
        spaceAfter=15
    )
    
    title_text = "Analysis Integrated Report"
    if is_zh:
        title_text = "試題分析整合報表 (Integrated Analysis Report)"
    story.append(Paragraph(title_text, title_style))
    
    # 視覺提示說明
    guide_style = ParagraphStyle('Guide', parent=styles['Normal'], fontName=font_name, fontSize=9, textColor=colors.grey)
    if is_zh:
        story.append(Paragraph("※ 趨勢圖中的垂直格線已精確對齊上方學生的作答欄位與下方的試題分析指標。", guide_style))
    story.append(Spacer(1, 10))
    
    # Helper to build themed tables
    def create_horiz_table(header, data, highlight_errors=False):
        tbl_data = [header] + [[str(x) for x in r] for r in data]
        num_cols = len(header)
        col_w = []
        if num_cols > 2:
            fixed_w, score_w = 60, 40
            remain = 1030 
            q_w = remain / (num_cols - 2)
            col_w = [fixed_w, score_w] + [q_w] * (num_cols - 2)
        else:
            col_w = [100] * num_cols

        t = Table(tbl_data, repeatRows=1, colWidths=col_w, hAlign='LEFT')
        style = [
            ('FONT', (0, 0), (-1, -1), font_name),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTSIZE', (0, 0), (-1, -1), 7 if num_cols > 50 else 8),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ]
        if highlight_errors:
            correct_ans = metric_data[0][2:]
            for row_idx, r in enumerate(data):
                for col_idx in range(2, num_cols):
                    q_idx = col_idx - 2
                    val = str(r[col_idx])
                    target = str(correct_ans[q_idx]) if q_idx < len(correct_ans) else ""
                    if val and target and val != target:
                        style.append(('TEXTCOLOR', (col_idx, row_idx + 1), (col_idx, row_idx + 1), colors.red))
        t.setStyle(TableStyle(style))
        return t

    def build_unified_grid_section(header, data, chart_data=None, chart_height=220, highlight_errors=False):
        # 1. Prepare colWidths
        num_cols = len(header)
        fixed_w, score_w = 60, 40
        remain = 1030
        q_w = remain / (num_cols - 2) if num_cols > 2 else 0
        col_w = [fixed_w, score_w] + [q_w] * (num_cols - 2)

        # 2. Build rows
        rows = []
        # Header
        rows.append(header)
        # Data
        for r in data: rows.append([str(x) for x in r])
        
        # Chart Row
        chart_row_idx = -1
        if chart_data:
            chart_row_idx = len(rows)
            # Create a row where col 0-1 are labels, and col 2-end is SPANNED for the GridChart
            chart_row = [t.get("label_trend", "答對率趨勢線"), ""] + [""] * (num_cols - 2)
            rows.append(chart_row)

        t = Table(rows, repeatRows=1, colWidths=col_w, hAlign='LEFT')
        style = [
            ('FONT', (0, 0), (-1, -1), font_name),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTSIZE', (0, 0), (-1, -1), 7 if num_cols > 50 else 8),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ]

        if highlight_errors:
            correct_ans = metric_data[0][2:]
            for row_idx, r in enumerate(data):
                for col_idx in range(2, num_cols):
                    q_idx = col_idx - 2
                    val = str(r[col_idx])
                    target = str(correct_ans[q_idx]) if q_idx < len(correct_ans) else ""
                    if val and target and val != target:
                        style.append(('TEXTCOLOR', (col_idx, row_idx + 1), (col_idx, row_idx + 1), colors.red))

        if chart_row_idx != -1:
            # Span the chart area
            style.append(('SPAN', (2, chart_row_idx), (-1, chart_row_idx)))
            style.append(('BOTTOMPADDING', (2, chart_row_idx), (-1, chart_row_idx), 0))
            style.append(('TOPPADDING', (2, chart_row_idx), (-1, chart_row_idx), 0))
            style.append(('LEFTPADDING', (2, chart_row_idx), (-1, chart_row_idx), 0))
            style.append(('RIGHTPADDING', (2, chart_row_idx), (-1, chart_row_idx), 0))
            # Insert the Flowable
            inner_chart = PrecisionGridChart(chart_data, col_w, height=chart_height, font_name=font_name)
            t._cellvalues[chart_row_idx][2] = inner_chart

        t.setStyle(TableStyle(style))
        return t

    # 1. Styles
    section_style = ParagraphStyle('Sect', parent=styles['Heading2'], fontName=font_name, fontSize=11, spaceBefore=2, spaceAfter=2)

    # Layout Logic based on student count
    threshold = 25
    is_large_class = len(student_data) > threshold
    
    if is_large_class:
        # MODE B: FRONT / BACK
        story.append(Paragraph("一、學生作答矩陣 (Student Answer Matrix)", section_style))
        if is_precision:
            story.append(build_unified_grid_section(student_header, student_data, highlight_errors=True))
        else:
            story.append(Spacer(1, 2))
            story.append(create_horiz_table(student_header, student_data, highlight_errors=True))
        
        story.append(PageBreak())
        
        # Back: Analysis Core (Plot + Metrics)
        story.append(Paragraph("二、試題正確率趨勢圖與指標 (Integrated Analysis Report)", section_style))
        if is_precision:
            # Unified Analysis Table: Metrics Header row + Metric data rows with Chart merged in between
            diff_vals = metric_data[1][2:]
            # We treat the Metrics as the "data" and inject chart
            story.append(build_unified_grid_section(metric_header, metric_data, chart_data=diff_vals, chart_height=350))
        else:
            item_img = job_dir / "analysis_item_plot.png"
            if item_img.exists():
                img = Image(str(item_img), width=1130, height=350)
                img.hAlign = 'LEFT'
                story.append(img)
            story.append(Paragraph("三、題目統計指標數據", section_style))
            story.append(create_horiz_table(metric_header, metric_data))
    else:
        # MODE A: SINGLE PAGE DASHBOARD
        story.append(Paragraph("一、試題分析整合報表 (Integrated Analysis Dashboard)", section_style))
        if is_precision:
            # One GIANT table: Student Header -> Student Data -> Chart Row -> Metric Rows
            # We can't use build_unified_grid_section directly for all 3 parts easily without a complex merge
            # Let's build it manually here for perfect precision
            num_cols = len(student_header)
            fixed_w, score_w = 60, 40
            remain = 1030
            q_w = remain / (num_cols - 2)
            col_w = [fixed_w, score_w] + [q_w] * (num_cols - 2)
            
            all_rows = []
            all_rows.append(student_header)
            for r in student_data: all_rows.append([str(x) for x in r])
            
            # Chart Title Row (Label area)
            chart_row_idx = len(all_rows)
            all_rows.append([t.get("label_trend", "答對率趨勢線"), ""] + [""] * (num_cols - 2))
            
            # Metric Rows
            metric_start_idx = len(all_rows)
            for r in metric_data: all_rows.append([str(x) for x in r])
            
            u_t = Table(all_rows, repeatRows=1, colWidths=col_w, hAlign='LEFT')
            u_style = [
                ('FONT', (0, 0), (-1, -1), font_name),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTSIZE', (0, 0), (-1, -1), 7 if num_cols > 50 else 8),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue), # Student Header
                ('BACKGROUND', (0, metric_start_idx), (-1, metric_start_idx), colors.lightgrey), # Metrics Header
            ]
            
            # Highlights
            correct_ans = metric_data[0][2:]
            for row_idx, r in enumerate(student_data):
                for col_idx in range(2, num_cols):
                    q_idx = col_idx - 2
                    val = str(r[col_idx])
                    target = str(correct_ans[q_idx]) if q_idx < len(correct_ans) else ""
                    if val and target and val != target:
                        u_style.append(('TEXTCOLOR', (col_idx, row_idx + 1), (col_idx, row_idx + 1), colors.red))
            
            # Span Chart
            u_style.append(('SPAN', (2, chart_row_idx), (-1, chart_row_idx)))
            u_style.append(('BOTTOMPADDING', (2, chart_row_idx), (-1, chart_row_idx), 0))
            u_style.append(('TOPPADDING', (2, chart_row_idx), (-1, chart_row_idx), 0))
            
            diff_vals = metric_data[1][2:]
            inner_chart = PrecisionGridChart(diff_vals, col_w, height=220, font_name=font_name)
            u_t._cellvalues[chart_row_idx][2] = inner_chart
            
            u_t.setStyle(TableStyle(u_style))
            story.append(u_t)
        else:
            story.append(Spacer(1, 2))
            story.append(create_horiz_table(student_header, student_data, highlight_errors=True))
            story.append(Paragraph("二、試題正確率趨勢圖", section_style))
            item_img = job_dir / "analysis_item_plot.png"
            if item_img.exists():
                img = Image(str(item_img), width=1130, height=220)
                img.hAlign = 'LEFT'
                story.append(img)
            story.append(Paragraph("三、題目統計指標數據", section_style))
            story.append(create_horiz_table(metric_header, metric_data))

    try:
        doc.build(story)
    except Exception as e:
        print(f"PDF Build Error: {e}")
