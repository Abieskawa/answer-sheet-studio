"""
Answer Sheet PDF generator (alignment-safe v11)

Your requested layout tweaks:
1) 黑色 Answer Rectangle 對齊答題範圍邊緣
   - Within each column, A/B/C/D bubble X positions are now computed to "fill" the usable width:
     A is close to the left usable edge (after question-number area),
     D is close to the right usable edge (minus bubble radius).
   - Vertical placement also uses the box height tightly (smaller top offsets).

2) 身分資訊分隔線對齊答案區方框
   - Divider line now uses BOX_X0..BOX_X1 (same as the answer box).

3) 年級 / 班級整組再往左
   - First bubble center for 年級/班級 aligns with NAME underline start (NAME_LINE_X0).

Notes:
- Single-page output (generator no longer cares about people/pages).
- Bubble radius slightly smaller (6.0pt).
- Strict assertions ensure no bubble can ever be outside the inner padded box.

CLI:
  python generator.py --subject 國文 --questions 100 --out answer_sheet.pdf
"""

from __future__ import annotations

import argparse
import math
from pathlib import Path
from typing import Union

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

PAGE_W, PAGE_H = A4

DEFAULT_TITLE = "定期評量 答案卷"

# Corner alignment squares
CORNER_MARK_MARGIN = 18
CORNER_MARK_SIZE = 24

# CJK font
_CJK_FONT_NAME = "STSong-Light"
_CJK_READY = False

def _ensure_cjk_font():
    global _CJK_READY
    if _CJK_READY:
        return
    try:
        pdfmetrics.registerFont(UnicodeCIDFont(_CJK_FONT_NAME))
        _CJK_READY = True
    except Exception:
        _CJK_READY = False

def set_font_cjk(c: canvas.Canvas, size: int):
    _ensure_cjk_font()
    c.setFont(_CJK_FONT_NAME if _CJK_READY else "Helvetica", size)

def set_font_lat(c: canvas.Canvas, size: int, bold: bool=False):
    c.setFont("Helvetica-Bold" if bold else "Helvetica", size)


# -----------------------------
# Header
# -----------------------------
LEFT_X = 45
RIGHT_X = PAGE_W - 45

TITLE_Y = PAGE_H - 70
SUBTITLE_Y = PAGE_H - 95
HEADER_LINE_Y = PAGE_H - 110

# -----------------------------
# Identity area
# -----------------------------
NAME_Y = PAGE_H - 145
NAME_LINE_X0 = LEFT_X + 40
NAME_LINE_X1 = LEFT_X + 195

GRADE_Y = PAGE_H - 174
GRADE_VALUES = [7, 8, 9, 10, 11, 12]
GRADE_RADIUS = 6.2
GRADE_X0 = NAME_LINE_X0  # align with name underline start
GRADE_STEP = 26

CLASS_Y = PAGE_H - 199
CLASS_VALUES = list(range(10))  # 0..9
CLASS_RADIUS = 6.2
CLASS_X0 = NAME_LINE_X0  # align with name underline start
CLASS_STEP = 21

# numeric labels above bubbles (already moved down; keep a bit lower)
BUBBLE_LABEL_DY = 9

# Seat box (top-right block)
SEAT_BOX_W = 200
SEAT_BOX_H = 62
SEAT_BOX_X0 = RIGHT_X - SEAT_BOX_W
SEAT_BOX_Y0 = PAGE_H - 205
SEAT_RADIUS = 6.0
SEAT_DIGITS = list(range(10))
SEAT_STEP = 16
SEAT_X0 = SEAT_BOX_X0 + 26
SEAT_TOP_Y = SEAT_BOX_Y0 + 40
SEAT_BOTTOM_Y = SEAT_BOX_Y0 + 18

DIVIDER_Y = PAGE_H - 220


# -----------------------------
# Answer Region (BLACK reference box)
# -----------------------------
BOX_X0 = 25
BOX_X1 = PAGE_W - 25
BOX_Y1 = DIVIDER_Y - 12
BOX_Y0 = 55

BOX_PAD = 10

BUBBLE_RADIUS = 6.0
MIN_CHOICES_COUNT = 3
MAX_CHOICES_COUNT = 5
DEFAULT_CHOICES_COUNT = 4
COLS = 3
MAX_QUESTIONS = 100


def make_choices(choices_count: int):
    choices_count = int(choices_count)
    if not (MIN_CHOICES_COUNT <= choices_count <= MAX_CHOICES_COUNT):
        raise ValueError(f"choices_count must be {MIN_CHOICES_COUNT}–{MAX_CHOICES_COUNT}")
    return [chr(ord("A") + i) for i in range(choices_count)]


# -----------------------------
# Validation
# -----------------------------
def _assert_inside_box(x: float, y: float, r: float, x0: float, y0: float, x1: float, y1: float, label: str):
    if x - r < x0 or x + r > x1 or y - r < y0 or y + r > y1:
        raise ValueError(
            f"{label} is outside Answer Box: center=({x:.2f},{y:.2f}) r={r:.2f} "
            f"box=({x0:.2f},{y0:.2f})-({x1:.2f},{y1:.2f})"
        )


def compute_question_layout(num_questions: int, choices_count: int = DEFAULT_CHOICES_COUNT):
    """
    Compute positions so that the black rectangle "fits" the answer grid more tightly.

    Note: Question *positions* are fixed (based on MAX_QUESTIONS). Different `num_questions`
    simply render a prefix subset (e.g., Q1–Q10 positions are identical on 10Q vs 20Q sheets).
    """
    num_questions = int(max(1, min(MAX_QUESTIONS, num_questions)))
    choices = make_choices(choices_count)
    box_w = BOX_X1 - BOX_X0
    col_w = box_w / COLS
    rows_per_col = int(math.ceil(MAX_QUESTIONS / COLS))

    inner_x0_all = BOX_X0 + BOX_PAD
    inner_x1_all = BOX_X1 - BOX_PAD
    inner_y0_all = BOX_Y0 + BOX_PAD
    inner_y1_all = BOX_Y1 - BOX_PAD

    # tighter top offsets
    header_y = inner_y1_all - 10
    first_row_y = inner_y1_all - 30

    bottom_limit = inner_y0_all + BUBBLE_RADIUS
    usable = first_row_y - bottom_limit
    row_step = 0.0 if rows_per_col <= 1 else usable / (rows_per_col - 1)

    min_step = 2 * BUBBLE_RADIUS + 1.0
    if rows_per_col > 1 and row_step < min_step:
        raise ValueError(
            f"Row spacing too small: {row_step:.2f} < {min_step:.2f}. "
            "Increase BOX height, reduce bubble size, or increase columns."
        )

    # Per-column X: fill usable width
    q_num_area_w = 30
    gap_after_num = 10

    col_x0s = [BOX_X0 + i * col_w for i in range(COLS)]
    bubble_xs = []
    q_num_right_x = []

    for ci, cx0 in enumerate(col_x0s):
        inner_left = cx0 + BOX_PAD
        inner_right = cx0 + col_w - BOX_PAD

        q_right = inner_left + q_num_area_w

        xA = q_right + gap_after_num + BUBBLE_RADIUS
        x_last = inner_right - BUBBLE_RADIUS
        step = (x_last - xA) / float(len(choices) - 1)
        xs = [xA + i * step for i in range(len(choices))]

        # Validate X positions with first_row_y
        for j, x in enumerate(xs):
            _assert_inside_box(
                x, first_row_y, BUBBLE_RADIUS,
                inner_left, inner_y0_all, inner_right, inner_y1_all,
                label=f"col{ci} {choices[j]} (x-check)"
            )

        bubble_xs.append(xs)
        q_num_right_x.append(q_right)

    return {
        "rows_per_col": rows_per_col,
        "header_y": header_y,
        "first_row_y": first_row_y,
        "row_step": row_step,
        "bubble_xs": bubble_xs,
        "q_num_right_x": q_num_right_x,
        "inner_x0": inner_x0_all,
        "inner_x1": inner_x1_all,
        "inner_y0": inner_y0_all,
        "inner_y1": inner_y1_all,
    }


# -----------------------------
# Drawing
# -----------------------------
def draw_corner_marks(c: canvas.Canvas):
    c.setFillColor(colors.black)
    m = CORNER_MARK_MARGIN
    s = CORNER_MARK_SIZE
    c.rect(m, PAGE_H - m - s, s, s, fill=1, stroke=0)
    c.rect(PAGE_W - m - s, PAGE_H - m - s, s, s, fill=1, stroke=0)
    c.rect(m, m, s, s, fill=1, stroke=0)
    c.rect(PAGE_W - m - s, m, s, s, fill=1, stroke=0)

def draw_header(c: canvas.Canvas, title_text: str, subject: str, subject_label: str):
    c.setFillColor(colors.black)
    set_font_cjk(c, 20)
    c.drawCentredString(PAGE_W / 2, TITLE_Y, title_text)

    if subject:
        set_font_cjk(c, 13)
        c.drawCentredString(PAGE_W / 2, SUBTITLE_Y, f"{subject_label}{subject}")

    c.setStrokeColor(colors.black)
    c.setLineWidth(1.0)
    c.line(BOX_X0, HEADER_LINE_Y, BOX_X1, HEADER_LINE_Y)  # align with answer box too

def draw_identity(c: canvas.Canvas):
    set_font_cjk(c, 12)
    c.setFillColor(colors.black)

    # Name
    c.drawString(LEFT_X, NAME_Y, "姓名")
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.9)
    c.line(NAME_LINE_X0, NAME_Y - 2, NAME_LINE_X1, NAME_Y - 2)

    # Grade
    set_font_cjk(c, 12)
    c.drawString(LEFT_X, GRADE_Y - 2, "年級")
    set_font_lat(c, 9, bold=False)
    for i, v in enumerate(GRADE_VALUES):
        x = GRADE_X0 + i * GRADE_STEP
        c.drawCentredString(x, GRADE_Y + BUBBLE_LABEL_DY, str(v))
        c.circle(x, GRADE_Y, GRADE_RADIUS, stroke=1, fill=0)

    # Class
    set_font_cjk(c, 12)
    c.drawString(LEFT_X, CLASS_Y - 2, "班級")
    set_font_lat(c, 9, bold=False)
    for i, v in enumerate(CLASS_VALUES):
        x = CLASS_X0 + i * CLASS_STEP
        c.drawCentredString(x, CLASS_Y + BUBBLE_LABEL_DY, str(v))
        c.circle(x, CLASS_Y, CLASS_RADIUS, stroke=1, fill=0)

    # Seat box
    c.setStrokeColor(colors.black)
    c.setLineWidth(1.2)
    c.roundRect(SEAT_BOX_X0, SEAT_BOX_Y0, SEAT_BOX_W, SEAT_BOX_H, 8, stroke=1, fill=0)

    set_font_cjk(c, 11)
    c.drawString(SEAT_BOX_X0 - 26, SEAT_BOX_Y0 + 38, "座")
    c.drawString(SEAT_BOX_X0 - 26, SEAT_BOX_Y0 + 18, "號")

    # Labels for tens / ones
    set_font_cjk(c, 9)
    c.drawString(SEAT_BOX_X0 + 6, SEAT_TOP_Y - 4, "十位")
    c.drawString(SEAT_BOX_X0 + 6, SEAT_BOTTOM_Y - 4, "個位")

    set_font_lat(c, 8, bold=False)
    for i, d in enumerate(SEAT_DIGITS):
        x = SEAT_X0 + i * SEAT_STEP
        c.drawCentredString(x, SEAT_TOP_Y + 10, str(d))
        c.circle(x, SEAT_TOP_Y, SEAT_RADIUS, stroke=1, fill=0)
        c.drawCentredString(x, SEAT_BOTTOM_Y + 7, str(d))
        c.circle(x, SEAT_BOTTOM_Y, SEAT_RADIUS, stroke=1, fill=0)

    # Divider aligned with box
    c.setLineWidth(1.0)
    c.line(BOX_X0, DIVIDER_Y, BOX_X1, DIVIDER_Y)

def draw_answer_box(c: canvas.Canvas):
    c.setStrokeColor(colors.black)
    c.setLineWidth(2.0)
    c.rect(BOX_X0, BOX_Y0, BOX_X1 - BOX_X0, BOX_Y1 - BOX_Y0, stroke=1, fill=0)

def draw_questions(c: canvas.Canvas, num_questions: int, choices_count: int = DEFAULT_CHOICES_COUNT):
    layout = compute_question_layout(num_questions, choices_count=choices_count)
    choices = make_choices(choices_count)

    rows = layout["rows_per_col"]
    used_cols = min(COLS, int(math.ceil(num_questions / rows)))

    # Choice headers
    set_font_lat(c, 10, bold=False)
    for col in range(used_cols):
        xs = layout["bubble_xs"][col]
        for j, ch in enumerate(choices):
            c.drawCentredString(xs[j], layout["header_y"], ch)

    # Bubbles + numbers
    c.setLineWidth(1.1)
    c.setStrokeColor(colors.black)
    set_font_lat(c, 9, bold=False)

    first_y = layout["first_row_y"]
    step_y = layout["row_step"]

    inner_x0 = layout["inner_x0"]
    inner_x1 = layout["inner_x1"]
    inner_y0 = layout["inner_y0"]
    inner_y1 = layout["inner_y1"]

    q = 1
    for col in range(used_cols):
        for r in range(rows):
            if q > num_questions:
                break
            y = first_y - r * step_y

            # y validation
            _assert_inside_box((inner_x0 + inner_x1) / 2, y, BUBBLE_RADIUS, inner_x0, inner_y0, inner_x1, inner_y1, f"row-y q{q}")

            # question number (right-aligned inside the number area)
            qx = layout["q_num_right_x"][col]
            c.drawRightString(qx, y - 3.0, f"{q}.")

            # bubbles
            xs = layout["bubble_xs"][col]
            for j, x in enumerate(xs):
                _assert_inside_box(
                    x,
                    y,
                    BUBBLE_RADIUS,
                    inner_x0,
                    inner_y0,
                    inner_x1,
                    inner_y1,
                    f"bubble q{q} {choices[j]}",
                )
                c.circle(x, y, BUBBLE_RADIUS, stroke=1, fill=0)

            q += 1

def draw_footer(c: canvas.Canvas, footer_text: str):
    set_font_cjk(c, 11)
    c.setFillColor(colors.black)
    c.drawCentredString(PAGE_W / 2, 42, footer_text)


# -----------------------------
# Public API (single-page)
# -----------------------------
def generate_answer_sheet_pdf(
    subject: str,
    num_questions: int,
    out_pdf_path: Union[str, Path],
    choices_count: int = DEFAULT_CHOICES_COUNT,
    title_text: str = DEFAULT_TITLE,
    subject_label: str = "科目：",
    footer_text: str = "請使用 2B 鉛筆或深色原子筆將圓圈塗滿",
) -> Path:
    num_questions = int(max(1, min(MAX_QUESTIONS, num_questions)))
    choices_count = int(max(MIN_CHOICES_COUNT, min(MAX_CHOICES_COUNT, choices_count)))
    out_pdf_path = Path(out_pdf_path)

    c = canvas.Canvas(str(out_pdf_path), pagesize=A4)
    draw_corner_marks(c)
    draw_header(c, title_text=title_text, subject=subject, subject_label=subject_label)
    draw_identity(c)
    draw_answer_box(c)
    draw_questions(c, num_questions=num_questions, choices_count=choices_count)
    draw_footer(c, footer_text=footer_text)
    c.showPage()
    c.save()
    return out_pdf_path

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--title", default=DEFAULT_TITLE)
    ap.add_argument("--subject_label", default="科目：")
    ap.add_argument("--subject", default="國文")
    ap.add_argument("--questions", type=int, default=100)
    ap.add_argument("--choices", type=int, default=DEFAULT_CHOICES_COUNT, help="Choices per question (3–5)")
    ap.add_argument("--out", default="answer_sheet.pdf")
    ap.add_argument("--footer", default="請使用 2B 鉛筆或深色原子筆將圓圈塗滿")
    args = ap.parse_args()

    generate_answer_sheet_pdf(
        subject=args.subject,
        num_questions=args.questions,
        choices_count=args.choices,
        out_pdf_path=args.out,
        title_text=args.title,
        subject_label=args.subject_label,
        footer_text=args.footer,
    )
    print(f"Wrote: {args.out}")

if __name__ == "__main__":
    main()
