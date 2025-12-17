from dataclasses import dataclass
from typing import Dict, List, Tuple

# PDF coordinates are in "points" (1/72 inch). We render at DPI and scale points->pixels by dpi/72.

PAGE_W_PT = 595  # A4 width in points
PAGE_H_PT = 842  # A4 height in points

# Layout parameters (points)
CORNER_MARK_SIZE = 14
CORNER_MARK_MARGIN = 18

TITLE_Y = 805

NAME_LINE_Y = 735
NAME_LINE_X0 = 70
NAME_LINE_X1 = 250

GRADE_LABEL_X = 270
GRADE_CIRCLE_Y = 737
GRADE_CIRCLE_RADIUS = 6
# grade options (7-12)
GRADE_VALUES = [7, 8, 9, 10, 11, 12]
GRADE_CIRCLE_XS = [280, 305, 330, 360, 390, 420]

SEAT_BOX_X0 = 350
SEAT_BOX_Y0 = 705
SEAT_BOX_W = 220
SEAT_BOX_H = 60

SEAT_DIGITS = list(range(10))  # 0-9
SEAT_CIRCLE_RADIUS = 6
SEAT_TOP_ROW_Y = 744
SEAT_BOTTOM_ROW_Y = 724
SEAT_START_X = 405
SEAT_STEP_X = 16

# Questions
MAX_QUESTIONS = 50
CHOICES = ["A","B","C","D"]
BUBBLE_RADIUS = 7
ROW_STEP_Y = 20

# Column definitions: (start_question_index, count, base_x)
# base_x is where choice bubbles start
COL1 = (1, 17, 95)
COL2 = (18, 17, 285)
COL3 = (35, 16, 445)
COLS = [COL1, COL2, COL3]

Q_START_Y = 665
Q_NUM_X_OFF = -30  # relative to base_x
CHOICE_STEP_X = 40  # distance between A,B,C,D centers in points
