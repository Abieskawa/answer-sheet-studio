from __future__ import annotations

import argparse
import io
import random
import sys
import time
import uuid
from pathlib import Path

import fitz  # PyMuPDF
from reportlab.lib import colors
from reportlab.pdfgen import canvas

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from engine.analysis import run_analysis_template
from engine.generator import (
    CLASS_STEP,
    CLASS_VALUES,
    CLASS_X0,
    CLASS_Y,
    GRADE_STEP,
    GRADE_VALUES,
    GRADE_X0,
    GRADE_Y,
    SEAT_BOTTOM_Y,
    SEAT_DIGITS,
    SEAT_RADIUS,
    SEAT_STEP,
    SEAT_TOP_Y,
    SEAT_X0,
    compute_question_layout,
    generate_answer_sheet_pdf,
)
from engine.recognizer import process_pdf_to_csv_and_annotated_pdf

# Reuse the same helper functions used by the web app so the CLI demo matches the real pipeline.
from app.main import (  # noqa: E402
    _sanitize_download_component,
    _upload_base_name,
    _write_analysis_report_pdf,
    _write_analysis_template,
    _write_answer_key_files,
    _write_job_meta,
    _write_showwrong_xlsx,
)


DEFAULT_INPUT = Path("test") / "八年級期末掃描.pdf"


def _make_fake_answer_key(num_questions: int, choices_count: int, seed: int) -> dict[int, tuple[str, float]]:
    rng = random.Random(seed)
    choices = [chr(ord("A") + i) for i in range(int(choices_count))]
    return {qno: (rng.choice(choices), 1.0) for qno in range(1, int(num_questions) + 1)}


def _overlay_pdf_bytes(page_size: tuple[float, float], circles: list[tuple[float, float, float]]) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=page_size)
    c.setFillColor(colors.black)
    for x, y, r in circles:
        c.circle(float(x), float(y), float(r), stroke=0, fill=1)
    c.showPage()
    c.save()
    return buf.getvalue()


def _generate_synthetic_scan_pdf(
    *,
    base_sheet_pdf: Path,
    out_pdf: Path,
    num_questions: int,
    choices_count: int,
    pages: int,
    grade: int,
    class_no: int,
    seed: int,
) -> None:
    base_doc = fitz.open(str(base_sheet_pdf))
    if base_doc.page_count < 1:
        raise ValueError(f"Base sheet has no pages: {base_sheet_pdf}")

    base_page = base_doc[0]
    page_size = (float(base_page.rect.width), float(base_page.rect.height))

    layout = compute_question_layout(num_questions=num_questions, choices_count=choices_count)
    rows_per_col = int(layout["rows_per_col"])
    first_row_y = float(layout["first_row_y"])
    row_step = float(layout["row_step"])
    bubble_xs = layout["bubble_xs"]

    rng = random.Random(int(seed))
    choices = [chr(ord("A") + i) for i in range(int(choices_count))]

    grade = int(grade)
    class_no = int(class_no)
    grade_idx = GRADE_VALUES.index(grade) if grade in GRADE_VALUES else 0
    class_idx = CLASS_VALUES.index(class_no) if class_no in CLASS_VALUES else 0

    grade_x = float(GRADE_X0 + grade_idx * GRADE_STEP)
    class_x = float(CLASS_X0 + class_idx * CLASS_STEP)

    seat_xs = [float(SEAT_X0 + i * SEAT_STEP) for i in range(len(SEAT_DIGITS))]

    out_doc = fitz.open()
    for page_idx in range(int(pages)):
        out_doc.insert_pdf(base_doc, from_page=0, to_page=0)
        page = out_doc[-1]

        seat_no = page_idx + 1
        seat_str = f"{seat_no:02d}"
        tens = int(seat_str[0])
        ones = int(seat_str[1])

        marks: list[tuple[float, float, float]] = []

        # Identity marks.
        marks.append((grade_x, float(GRADE_Y), float(SEAT_RADIUS) * 0.85))
        marks.append((class_x, float(CLASS_Y), float(SEAT_RADIUS) * 0.85))
        marks.append((seat_xs[tens], float(SEAT_TOP_Y), float(SEAT_RADIUS) * 0.85))
        marks.append((seat_xs[ones], float(SEAT_BOTTOM_Y), float(SEAT_RADIUS) * 0.85))

        # Answer marks.
        for qno in range(1, int(num_questions) + 1):
            col = int((qno - 1) // rows_per_col)
            row = int((qno - 1) % rows_per_col)
            y = first_row_y - float(row) * row_step
            chosen = rng.choice(choices)
            choice_idx = int(ord(chosen) - ord("A"))
            x = float(bubble_xs[col][choice_idx])
            marks.append((x, y, float(SEAT_RADIUS) * 0.78))

        overlay = fitz.open("pdf", _overlay_pdf_bytes(page_size, marks))
        page.show_pdf_page(page.rect, overlay, 0)
        overlay.close()

    out_doc.save(str(out_pdf))
    out_doc.close()
    base_doc.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="Answer Sheet Studio end-to-end demo using a scan PDF (or a synthetic scan).")
    parser.add_argument(
        "--input",
        default=str(DEFAULT_INPUT),
        help="Input multi-page scan PDF (one student per page). Ignored when --synthetic-pages > 0.",
    )
    parser.add_argument("--num-questions", type=int, default=50, help="Number of questions (1–100).")
    parser.add_argument("--choices-count", type=int, default=4, choices=(3, 4, 5), help="Choices per question (3/4/5).")
    parser.add_argument("--lang", default="zh-Hant", choices=("zh-Hant", "zh-Hans", "en"), help="Report language (zh-Hant/zh-Hans/en).")
    parser.add_argument("--seed", type=int, default=7, help="RNG seed for the fake answer key.")
    parser.add_argument(
        "--synthetic-pages",
        type=int,
        default=0,
        help="Generate a synthetic scan PDF with N pages (students) instead of using --input. "
        "If omitted and the default sample input is missing, falls back to 8 pages.",
    )
    parser.add_argument("--synthetic-grade", type=int, default=8, help="Grade to mark for synthetic pages.")
    parser.add_argument("--synthetic-class", type=int, default=1, help="Class to mark for synthetic pages.")
    parser.add_argument("--job-id", default="", help="Optional job id (default: random UUID).")

    args = parser.parse_args()

    raw_input_pdf = Path(args.input).expanduser()
    input_pdf = (raw_input_pdf if raw_input_pdf.is_absolute() else (REPO_ROOT / raw_input_pdf)).resolve()

    num_questions = max(1, min(100, int(args.num_questions)))
    choices_count = max(3, min(5, int(args.choices_count)))

    job_id = (args.job_id or "").strip() or str(uuid.uuid4())
    out_dir = REPO_ROOT / "outputs" / job_id
    out_dir.mkdir(parents=True, exist_ok=True)

    default_input_pdf = (REPO_ROOT / DEFAULT_INPUT).resolve()
    use_synthetic = int(args.synthetic_pages) > 0
    if not use_synthetic and not input_pdf.exists():
        if input_pdf == default_input_pdf:
            use_synthetic = True
            args.synthetic_pages = 8
            print(f"⚠️ Sample input not found: {default_input_pdf}")
            print("   Generating a synthetic scan PDF instead (8 pages).")
            print("   Tip: provide a real scan via --input to test actual scan quality.\n")
        else:
            raise SystemExit(f"Input PDF not found: {input_pdf}")

    original_filename = "synthetic_scan.pdf" if use_synthetic else _sanitize_download_component(input_pdf.name, "upload.pdf")
    _write_job_meta(
        out_dir,
        {
            "original_filename": original_filename,
            "upload_base": _upload_base_name(original_filename),
            "created_at": int(time.time()),
            "num_questions": num_questions,
            "choices_count": choices_count,
        },
    )

    # 1) Generate the printable blank sheet (for students).
    sheet_pdf = out_dir / "sheet.pdf"
    generate_answer_sheet_pdf(
        subject="（Demo）",
        num_questions=num_questions,
        choices_count=choices_count,
        out_pdf_path=str(sheet_pdf),
        title_text="Answer Sheet Studio Demo",
    )

    # 2) Materialize the scan PDF into the job folder (matches the web app output structure).
    input_copy = out_dir / "input.pdf"
    if use_synthetic:
        _generate_synthetic_scan_pdf(
            base_sheet_pdf=sheet_pdf,
            out_pdf=input_copy,
            num_questions=num_questions,
            choices_count=choices_count,
            pages=int(args.synthetic_pages),
            grade=int(args.synthetic_grade),
            class_no=int(args.synthetic_class),
            seed=int(args.seed),
        )
    else:
        input_copy.write_bytes(input_pdf.read_bytes())

    # 3) Recognize answers -> results.csv + annotated.pdf + ambiguity.csv + roster.csv.
    results_csv = out_dir / "results.csv"
    ambiguity_csv = out_dir / "ambiguity.csv"
    roster_csv = out_dir / "roster.csv"
    annotated_pdf = out_dir / "annotated.pdf"
    process_pdf_to_csv_and_annotated_pdf(
        input_pdf_path=str(input_copy),
        num_questions=num_questions,
        choices_count=choices_count,
        out_csv_path=str(results_csv),
        out_ambiguity_csv_path=str(ambiguity_csv),
        out_roster_csv_path=str(roster_csv),
        out_annotated_pdf_path=str(annotated_pdf),
    )

    # 4) Create a fake answer key, then generate analysis outputs.
    answer_key = _make_fake_answer_key(num_questions=num_questions, choices_count=choices_count, seed=int(args.seed))
    answer_key_xlsx = out_dir / "answer_key.xlsx"
    _write_answer_key_files(
        num_questions=num_questions,
        answer_key=answer_key,
        out_xlsx_path=answer_key_xlsx,
        default_points=1.0,
    )

    showwrong_xlsx = out_dir / "showwrong.xlsx"
    _write_showwrong_xlsx(results_csv, answer_key, showwrong_xlsx)

    template_csv = out_dir / "analysis_template.csv"
    _write_analysis_template(results_csv, answer_key, template_csv, default_points=1.0)

    run_analysis_template(template_csv, out_dir, lang=str(args.lang))
    _write_analysis_report_pdf(out_dir, lang=str(args.lang))

    print(f"✅ Demo job created: {job_id}")
    print(f"📁 Output folder: {out_dir.resolve()}")
    print(f"🌐 If the server is running: http://127.0.0.1:8000/result/{job_id}/charts")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
