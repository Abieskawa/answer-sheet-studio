import os
import uuid
import re
from typing import Optional
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

_JOB_ID_RE = re.compile(r"^[0-9a-fA-F-]{8,64}$")

LANG_COOKIE_NAME = "lang"
SUPPORTED_LANGS = ("zh-Hant", "en")
DEFAULT_LANG = "zh-Hant"

I18N = {
    "zh-Hant": {
        "lang_zh": "繁體中文",
        "lang_en": "English",
        "nav_download": "下載答案卡",
        "nav_upload": "上傳處理",
        "footer_local": "本機運行 / 不上傳到雲端",
        "download_title": "下載答案卡",
        "download_label_title": "標題（預設：定期評量 答案卷）",
        "download_ph_title": "例如：第一次段考",
        "download_label_subject": "科目（可留空）",
        "download_ph_subject": "例如：數學",
        "download_label_class": "班級（可留空）",
        "download_ph_class": "例如：701",
        "download_label_num_questions": "題數（1–100；上傳辨識時也要填一樣）",
        "download_btn_generate": "產生並下載 PDF",
        "download_hint_location": "下載後由瀏覽器決定儲存位置（會跳出「儲存」視窗或下載到預設資料夾）。",
        "download_hint_support": "年級欄位支援 7–12（含 10/11/12），座號支援 00–99，選擇題 A–D，最多 100 題。",
        "upload_title": "上傳處理",
        "upload_label_num_questions": "題數（1–100；請依每份掃描檔案輸入，需與答案卡一致）",
        "upload_ph_num_questions": "例如：50",
        "upload_label_pdf": "上傳多頁 PDF（每頁一位）",
        "upload_btn_process": "開始辨識",
        "upload_hint_output": "完成後會輸出 results.csv 與 annotated.pdf。",
        "result_title": "處理完成",
        "result_job_id": "Job ID：",
        "result_download_results": "下載 results.csv",
        "result_download_annotated": "下載 annotated.pdf",
        "result_hint_unstable": "如果結果不穩，通常是掃描歪斜或太淡；可以提高掃描解析度（建議 300dpi）或改用較深的筆。",
        "result_debug_hint": "需要回報問題時，可到 Debug Mode 下載診斷檔案（輸入 Job ID）。",
        "result_debug_open": "開啟 Debug Mode",
        "debug_title": "Debug Mode",
        "debug_hint": "一般使用者不需要這個頁面。若需要回報問題，請依指示下載檔案並提供 Job ID。",
        "debug_label_job_id": "Job ID",
        "debug_ph_job_id": "貼上處理完成頁面顯示的 Job ID",
        "debug_btn_open": "開啟 Debug 下載",
        "debug_error_invalid_job_id": "Job ID 格式不正確。請貼上處理完成頁面顯示的 Job ID。",
        "debug_error_not_found": "找不到此 Job ID 的輸出資料夾。請確認 Job ID 是否正確。",
        "debug_dl_results": "下載 results.csv",
        "debug_dl_ambiguity": "下載 ambiguity.csv",
        "debug_dl_annotated": "下載 annotated.pdf",
        "debug_dl_input": "下載 input.pdf（原始上傳檔）",
        "debug_report_hint": "回報時請提供：Job ID、results.csv、ambiguity.csv、annotated.pdf（必要時 input.pdf）。",
    },
    "en": {
        "lang_zh": "繁體中文",
        "lang_en": "English",
        "nav_download": "Download",
        "nav_upload": "Upload",
        "footer_local": "Runs locally / no cloud upload",
        "download_title": "Download Answer Sheet",
        "download_label_title": "Title (default: Exam Answer Sheet)",
        "download_ph_title": "e.g., Midterm 1",
        "download_label_subject": "Subject (optional)",
        "download_ph_subject": "e.g., Math",
        "download_label_class": "Class (optional)",
        "download_ph_class": "e.g., 701",
        "download_label_num_questions": "Number of questions (1–100; must match upload)",
        "download_btn_generate": "Generate & Download PDF",
        "download_hint_location": "Your browser decides where the file is saved (Save dialog or default Downloads folder).",
        "download_hint_support": "Grade supports 7–12 (including 10/11/12). Seat No supports 00–99. Choices A–D. Up to 100 questions.",
        "upload_title": "Upload for Recognition",
        "upload_label_num_questions": "Number of questions (1–100; enter per PDF and must match the sheet)",
        "upload_ph_num_questions": "e.g., 50",
        "upload_label_pdf": "Upload multi-page PDF (one student per page)",
        "upload_btn_process": "Start Recognition",
        "upload_hint_output": "Outputs results.csv and annotated.pdf.",
        "result_title": "Done",
        "result_job_id": "Job ID:",
        "result_download_results": "Download results.csv",
        "result_download_annotated": "Download annotated.pdf",
        "result_hint_unstable": "If results are unstable, scans may be skewed or too light. Try 300dpi or a darker pen.",
        "result_debug_hint": "For reporting/debugging, open Debug Mode and enter the Job ID to download diagnostic files.",
        "result_debug_open": "Open Debug Mode",
        "debug_title": "Debug Mode",
        "debug_hint": "Most users do not need this page. For reporting issues, download the files and provide the Job ID.",
        "debug_label_job_id": "Job ID",
        "debug_ph_job_id": "Paste the Job ID from the result page",
        "debug_btn_open": "Open Debug Downloads",
        "debug_error_invalid_job_id": "Invalid Job ID. Please paste the Job ID shown on the result page.",
        "debug_error_not_found": "No output folder found for this Job ID. Please check the Job ID.",
        "debug_dl_results": "Download results.csv",
        "debug_dl_ambiguity": "Download ambiguity.csv",
        "debug_dl_annotated": "Download annotated.pdf",
        "debug_dl_input": "Download input.pdf (original upload)",
        "debug_report_hint": "When reporting, include: Job ID, results.csv, ambiguity.csv, annotated.pdf (and input.pdf if needed).",
    },
}


def resolve_lang(request: Request) -> str:
    q = (request.query_params.get("lang") or "").strip()
    if q in SUPPORTED_LANGS:
        return q
    cookie = (request.cookies.get(LANG_COOKIE_NAME) or "").strip()
    if cookie in SUPPORTED_LANGS:
        return cookie
    return DEFAULT_LANG


def template_response(request: Request, name: str, ctx: Optional[dict] = None):
    lang = resolve_lang(request)
    t = I18N.get(lang, I18N[DEFAULT_LANG])
    merged = {"request": request, "lang": lang, "t": t}
    if ctx:
        merged.update(ctx)
    resp = templates.TemplateResponse(name, merged)
    q = (request.query_params.get("lang") or "").strip()
    if q in SUPPORTED_LANGS:
        resp.set_cookie(LANG_COOKIE_NAME, q, max_age=31536000, samesite="lax")
    return resp

@app.get("/favicon.ico", include_in_schema=False)
def favicon():
    icon = APP_DIR / "static" / "favicon.ico"
    if icon.exists():
        return FileResponse(path=str(icon), media_type="image/x-icon")
    return Response(status_code=204)


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return template_response(request, "download.html")


@app.get("/upload", response_class=HTMLResponse)
def upload_page(request: Request):
    return template_response(request, "upload.html")


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

    return template_response(
        request,
        "result.html",
        {
            "job_id": job_id,
            "csv_url": f"/outputs/{job_id}/results.csv",
            "pdf_url": f"/outputs/{job_id}/annotated.pdf",
        },
    )


@app.get("/debug", response_class=HTMLResponse)
def debug_page(request: Request, job_id: str = ""):
    lang = resolve_lang(request)
    t = I18N.get(lang, I18N[DEFAULT_LANG])
    job_id = (job_id or "").strip()
    ctx = {"job_id": job_id or None, "files": None, "error": None}

    if not job_id:
        return template_response(request, "debug.html", ctx)

    if not _JOB_ID_RE.match(job_id):
        ctx["error"] = t["debug_error_invalid_job_id"]
        return template_response(request, "debug.html", ctx)

    job_dir = OUTPUTS_DIR / job_id
    if not job_dir.exists():
        ctx["error"] = t["debug_error_not_found"]
        return template_response(request, "debug.html", ctx)

    def url_if_exists(filename: str) -> Optional[str]:
        p = job_dir / filename
        return f"/outputs/{job_id}/{filename}" if p.exists() else None

    ctx["files"] = {
        "results": url_if_exists("results.csv"),
        "ambiguity": url_if_exists("ambiguity.csv"),
        "annotated": url_if_exists("annotated.pdf"),
        "input": url_if_exists("input.pdf"),
    }
    return template_response(request, "debug.html", ctx)


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
