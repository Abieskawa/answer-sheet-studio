from __future__ import annotations

import argparse
import random
import sys
import time
import uuid
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from engine.analysis import run_analysis_template
from engine.generator import generate_answer_sheet_pdf
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


def _make_fake_answer_key(num_questions: int, choices_count: int, seed: int) -> dict[int, tuple[str, float]]:
    rng = random.Random(seed)
    choices = [chr(ord("A") + i) for i in range(int(choices_count))]
    return {qno: (rng.choice(choices), 1.0) for qno in range(1, int(num_questions) + 1)}


def main() -> int:
    parser = argparse.ArgumentParser(description="Answer Sheet Studio end-to-end demo using a sample scan PDF.")
    parser.add_argument(
        "--input",
        default=str(Path("test") / "八年級期末掃描.pdf"),
        help="Input multi-page scan PDF (one student per page).",
    )
    parser.add_argument("--num-questions", type=int, default=50, help="Number of questions (1–100).")
    parser.add_argument("--choices-count", type=int, default=4, choices=(3, 4, 5), help="Choices per question (3/4/5).")
    parser.add_argument("--lang", default="zh-Hant", help="Report language (zh-Hant/en).")
    parser.add_argument("--seed", type=int, default=7, help="RNG seed for the fake answer key.")
    parser.add_argument("--job-id", default="", help="Optional job id (default: random UUID).")

    args = parser.parse_args()

    input_pdf = Path(args.input).expanduser().resolve()
    if not input_pdf.exists():
        raise SystemExit(f"Input PDF not found: {input_pdf}")

    num_questions = max(1, min(100, int(args.num_questions)))
    choices_count = max(3, min(5, int(args.choices_count)))

    job_id = (args.job_id or "").strip() or str(uuid.uuid4())
    out_dir = Path("outputs") / job_id
    out_dir.mkdir(parents=True, exist_ok=True)

    original_filename = _sanitize_download_component(input_pdf.name, "upload.pdf")
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

    # 2) Copy the scan PDF into the job folder (matches the web app output structure).
    input_copy = out_dir / "input.pdf"
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

