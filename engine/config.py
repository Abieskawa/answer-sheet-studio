import math
from typing import Dict, List

# PDF coordinates are in "points" (1/72 inch). All layout constants mirror engine.generator.

PAGE_W_PT = 595  # A4 width in points
PAGE_H_PT = 842  # A4 height in points

# Corner marks (alignment markers)
CORNER_MARK_MARGIN = 18
CORNER_MARK_SIZE = 24

# Identity area
LEFT_X = 45
RIGHT_X = PAGE_W_PT - 45

NAME_LINE_Y = PAGE_H_PT - 145
NAME_LINE_X0 = LEFT_X + 40
NAME_LINE_X1 = LEFT_X + 195

GRADE_VALUES = [7, 8, 9, 10, 11, 12]
GRADE_CIRCLE_RADIUS = 6.2
GRADE_CIRCLE_Y = PAGE_H_PT - 174
GRADE_STEP_X = 26
GRADE_CIRCLE_XS = [NAME_LINE_X0 + i * GRADE_STEP_X for i in range(len(GRADE_VALUES))]

CLASS_VALUES = list(range(10))  # 0-9
CLASS_CIRCLE_RADIUS = 6.2
CLASS_CIRCLE_Y = PAGE_H_PT - 199
CLASS_STEP_X = 21
CLASS_CIRCLE_XS = [NAME_LINE_X0 + i * CLASS_STEP_X for i in range(len(CLASS_VALUES))]

# Seat box
SEAT_BOX_W = 200
SEAT_BOX_H = 62
SEAT_BOX_X0 = RIGHT_X - SEAT_BOX_W
SEAT_BOX_Y0 = PAGE_H_PT - 205
SEAT_CIRCLE_RADIUS = 6.0
SEAT_DIGITS = list(range(10))
SEAT_STEP_X = 16
SEAT_START_X = SEAT_BOX_X0 + 26
SEAT_TOP_ROW_Y = SEAT_BOX_Y0 + 40
SEAT_BOTTOM_ROW_Y = SEAT_BOX_Y0 + 18

DIVIDER_Y = PAGE_H_PT - 220

# Answer box
BOX_X0 = 25
BOX_X1 = PAGE_W_PT - 25
BOX_Y0 = 55
BOX_Y1 = DIVIDER_Y - 12
BOX_PAD = 10

CHOICES = ["A", "B", "C", "D"]
MIN_CHOICES_COUNT = 3
MAX_CHOICES_COUNT = 5
DEFAULT_CHOICES_COUNT = 4
BUBBLE_RADIUS = 6.0
COL_COUNT = 3
MAX_QUESTIONS = 100

# Layout helpers (match engine.generator)
Q_NUM_AREA_W = 30
GAP_AFTER_NUM = 10


def make_choices(choices_count: int) -> List[str]:
    choices_count = int(choices_count)
    if not (MIN_CHOICES_COUNT <= choices_count <= MAX_CHOICES_COUNT):
        raise ValueError(f"choices_count must be {MIN_CHOICES_COUNT}–{MAX_CHOICES_COUNT}")
    return [chr(ord("A") + i) for i in range(choices_count)]


def compute_answer_layout(num_questions: int, choices_count: int = DEFAULT_CHOICES_COUNT) -> Dict:
    """
    Mirror engine.generator.compute_question_layout for recognizer.

    Note: Question *positions* are fixed (based on MAX_QUESTIONS). Different `num_questions`
    simply read a prefix subset (e.g., Q1–Q10 positions are identical on 10Q vs 20Q sheets).
    """
    num_questions = max(1, min(MAX_QUESTIONS, int(num_questions)))
    choices = make_choices(choices_count)
    box_w = BOX_X1 - BOX_X0
    col_w = box_w / COL_COUNT
    rows_per_col = int(math.ceil(MAX_QUESTIONS / COL_COUNT))

    inner_x0 = BOX_X0 + BOX_PAD
    inner_x1 = BOX_X1 - BOX_PAD
    inner_y0 = BOX_Y0 + BOX_PAD
    inner_y1 = BOX_Y1 - BOX_PAD

    header_y = inner_y1 - 10
    first_row_y = inner_y1 - 30

    bottom_limit = inner_y0 + BUBBLE_RADIUS
    usable_height = first_row_y - bottom_limit
    full_rows_per_col = rows_per_col
    if full_rows_per_col > 1:
        row_step = usable_height / (rows_per_col - 1)
    else:
        row_step = 0.0

    min_step = 2 * BUBBLE_RADIUS + 1.0
    if full_rows_per_col > 1 and row_step < min_step:
        raise ValueError(
            f"Row spacing too small: {row_step:.2f} < {min_step:.2f}. "
            "Adjust box height/cols."
        )

    col_x0s = [BOX_X0 + i * col_w for i in range(COL_COUNT)]
    bubble_xs: List[List[float]] = []
    q_num_right_x: List[float] = []

    for cx0 in col_x0s:
        inner_left = cx0 + BOX_PAD
        inner_right = cx0 + col_w - BOX_PAD

        q_right = inner_left + Q_NUM_AREA_W
        xA = q_right + GAP_AFTER_NUM + BUBBLE_RADIUS
        x_last = inner_right - BUBBLE_RADIUS
        step = (x_last - xA) / float(len(choices) - 1)
        xs = [xA + i * step for i in range(len(choices))]

        bubble_xs.append(xs)
        q_num_right_x.append(q_right)

    return {
        "rows_per_col": rows_per_col,
        "first_row_y": first_row_y,
        "row_step": row_step,
        "bubble_xs": bubble_xs,
        "q_num_right_x": q_num_right_x,
    }
