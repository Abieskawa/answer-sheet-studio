#%%
from pathlib import Path
job_dir = Path(r"C:\Users\spook\Desktop\讀卡機測試\answer-sheet-studio-main_VP19_myversion\answer-sheet-studio-main\outputs\55255b74-7246-4a87-93cb-07b43ff88c4a")#test job dir
item_csv = job_dir / "analysis_item.csv"
scores_csv = job_dir / "analysis_scores.csv"
template_csv = job_dir / "analysis_template.csv"
#test job dir
#%%
from __future__ import annotations

import csv
import math
import re
import statistics
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
from PIL import Image as PILImage, ImageDraw, ImageFont
import os # Added for os.path.exists

from xlsx import write_simple_xlsx_multi
from IPython.display import display, Image

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

#%%
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


# %%

# --- 準備輸入參數 ---

# A. 隨機生成 20 個學生的分數 (平均 70, 標準差 15)
np.random.seed(42)
raw_data = np.random.normal(70, 15, 20)
mock_scores = [float(np.clip(s, 0, 100)) for s in raw_data]

# B. 設定總分為 100
total_score = 100.0

# C. 設定輸出路徑
test_out_path = Path("debug_preview.png")

# D. 語言設定
language = "zh_TW"

# --- 執行函數 ---

_score_histogram_png(
    scores=mock_scores,
    total_possible=total_score,
    out_path=test_out_path,
    lang="zh_TW"
)

# 4. 即時顯示圖片
if test_out_path.exists():
    display(Image(filename=test_out_path))
else:
    print("畫圖失敗，請檢查終端機報錯")
# %%
