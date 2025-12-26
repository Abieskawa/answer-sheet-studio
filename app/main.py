import os
import uuid
from pathlib import Path
from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from omr.generator import generate_answer_sheet_pdf, DEFAULT_TITLE as DEFAULT_SHEET_TITLE
from omr.recognizer import process_pdf_to_csv_and_annotated_pdf

APP_DIR = Path(__file__).resolve().parent
ROOT_DIR = APP_DIR.parent
OUTPUTS_DIR = ROOT_DIR / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)

templates = Jinja2Templates(directory=str(APP_DIR / "templates"))

app = FastAPI(title="Answer Sheet Studio")
app.mount("/static", StaticFiles(directory=str(APP_DIR / "static")), name="static")

@app.get("/favicon.ico", include_in_schema=False)
def favicon():
    icon = APP_DIR / "static" / "favicon.ico"
    if icon.exists():
        return FileResponse(path=str(icon), media_type="image/x-icon")
    return Response(status_code=204)


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("download.html", {"request": request})


@app.get("/upload", response_class=HTMLResponse)
def upload_page(request: Request):
    return templates.TemplateResponse("upload.html", {"request": request})


def _sanitize_token(value: str, fallback: str) -> str:
    token = "".join(ch if ch.isalnum() else "_" for ch in value).strip("_")
    return token or fallback


@app.post("/api/generate")
def api_generate(
    title_text: str = Form(DEFAULT_SHEET_TITLE),
    subject: str = Form(""),
    class_name: str = Form(""),
    num_questions: int = Form(50),
):
    num_questions = max(1, min(100, int(num_questions)))

    out_id = str(uuid.uuid4())[:8]
    out_path = OUTPUTS_DIR / f"answer_sheets_{out_id}.pdf"

    subject = subject.strip()
    class_name = class_name.strip()
    title_text = title_text.strip() or DEFAULT_SHEET_TITLE

    generate_answer_sheet_pdf(
        subject=subject,
        num_questions=num_questions,
        out_pdf_path=str(out_path),
        title_text=title_text,
    )

    subject_token = _sanitize_token(subject, "subject") if subject else "subject"
    class_token = _sanitize_token(class_name, "class") if class_name else None
    title_token = _sanitize_token(title_text, "exam")
    filename_bits = ["answer_sheets", subject_token]
    if class_token:
        filename_bits.append(class_token)
    filename_bits.append(title_token)
    filename_bits.append(f"{num_questions}q")
    filename = "_".join(filename_bits) + ".pdf"

    return FileResponse(
        path=str(out_path),
        media_type="application/pdf",
        filename=filename,
    )


@app.post("/api/process")
async def api_process(
    request: Request,
    pdf: UploadFile = File(...),
    num_questions: int = Form(50),
):
    num_questions = max(1, min(100, int(num_questions)))

    job_id = str(uuid.uuid4())
    job_dir = OUTPUTS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    input_pdf = job_dir / "input.pdf"
    with open(input_pdf, "wb") as f:
        f.write(await pdf.read())

    csv_path = job_dir / "results.csv"
    ambiguity_csv_path = job_dir / "ambiguity.csv"
    annotated_pdf_path = job_dir / "annotated.pdf"

    # Main processing
    process_pdf_to_csv_and_annotated_pdf(
        input_pdf_path=str(input_pdf),
        num_questions=num_questions,
        out_csv_path=str(csv_path),
        out_ambiguity_csv_path=str(ambiguity_csv_path),
        out_annotated_pdf_path=str(annotated_pdf_path),
    )

    return templates.TemplateResponse(
        "result.html",
        {
            "request": request,
            "job_id": job_id,
            "csv_url": f"/outputs/{job_id}/results.csv",
            "ambiguity_url": f"/outputs/{job_id}/ambiguity.csv",
            "pdf_url": f"/outputs/{job_id}/annotated.pdf",
        },
    )


@app.get("/outputs/{job_id}/{filename}")
def download_output(job_id: str, filename: str):
    job_dir = OUTPUTS_DIR / job_id
    file_path = job_dir / filename
    if not file_path.exists():
        return RedirectResponse(url="/upload", status_code=302)
    media = "application/octet-stream"
    if filename.lower().endswith(".pdf"):
        media = "application/pdf"
    if filename.lower().endswith(".csv"):
        media = "text/csv"
    return FileResponse(path=str(file_path), media_type=media, filename=filename)


def run():
    import uvicorn

    host = os.environ.get("ANSWER_SHEET_HOST", os.environ.get("HOST", "127.0.0.1"))
    port_str = os.environ.get("ANSWER_SHEET_PORT", os.environ.get("PORT", "8000"))
    try:
        port = int(port_str)
    except ValueError:
        port = 8000

    try:
        uvicorn.run("app.main:app", host=host, port=port, reload=False)
    except PermissionError as exc:
        print(
            f"\nERROR: Permission denied while binding to {host}:{port}.\n"
            "macOS or your security software blocked Python from accepting incoming connections.\n"
            "Allow Python in System Settings > Network > Firewall (or equivalent) and rerun.\n"
            "You can also choose another port by setting ANSWER_SHEET_PORT.\n"
        )
        raise exc
