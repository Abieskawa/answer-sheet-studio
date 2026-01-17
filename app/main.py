import csv
import io
import os
import uuid
import re
import json
import time
import threading
import shutil
from typing import Optional
from pathlib import Path
from fastapi import FastAPI, Request, UploadFile, File, Form, BackgroundTasks
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from engine.generator import generate_answer_sheet_pdf, DEFAULT_TITLE as DEFAULT_SHEET_TITLE
from engine.recognizer import process_pdf_to_csv_and_annotated_pdf

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
        "nav_analysis": "答案分析",
        "nav_update": "更新",
        "footer_local": "本機運行 / 不上傳到雲端",
        "footer_docs": "說明文件",
        "footer_download": "下載",
        "download_title": "下載答案卡",
        "download_label_title": "標題（預設：定期評量 答案卷）",
        "download_ph_title": "例如：第一次段考",
        "download_label_subject": "科目（可留空）",
        "download_ph_subject": "例如：數學",
        "download_label_num_questions": "題數（1–100；上傳辨識時也要填一樣）",
        "download_label_choices_count": "每題選項（ABC/ABCD/ABCDE）",
        "download_btn_generate": "產生並下載 PDF",
        "download_hint_location": "下載後由瀏覽器決定儲存位置（會跳出「儲存」視窗或下載到預設資料夾）。",
        "download_hint_support": "年級欄位支援 7–12（含 10/11/12），座號支援 00–99，選擇題支援 ABC / ABCD / ABCDE，最多 100 題。",
        "upload_title": "上傳處理",
        "upload_label_num_questions": "題數（1–100；請依每份掃描檔案輸入，需與答案卡一致）",
        "upload_ph_num_questions": "例如：50",
        "upload_label_choices_count": "每題選項（ABC/ABCD/ABCDE；需與答案卡一致）",
        "upload_label_pdf": "上傳多頁 PDF（每頁一位）",
        "upload_btn_process": "開始辨識",
        "upload_hint_output": "完成後會輸出 results.csv 與 annotated.pdf。",
        "update_title": "更新",
        "update_hint": "下載最新 ZIP 後在此上傳套用更新。更新過程會短暫重新啟動。",
        "update_open_releases": "開啟下載頁（GitHub Releases）",
        "update_open_zip": "開啟 ZIP 下載（main.zip）",
        "update_btn_git": "用 Git 更新（git pull）",
        "update_git_hint": "（需要 Git，且此資料夾是 git clone）",
        "update_git_no_repo": "這個資料夾不是 git repository（沒有 .git），請改用 ZIP 更新。",
        "update_git_no_git": "找不到 Git，請先安裝 Git 或改用 ZIP 更新。",
        "update_git_dirty": "偵測到本機有未提交的修改，為避免衝突已停止 Git 更新；請先備份/提交/還原後再試。",
        "update_label_zip": "上傳更新 ZIP",
        "update_btn_apply": "套用更新並重新啟動",
        "update_started": "已開始更新，請稍候（會短暫重新啟動）。",
        "update_local_only": "只允許在本機（localhost）進行更新。",
        "update_invalid_zip": "請上傳有效的 ZIP 檔（.zip）。",
        "result_title": "處理完成",
        "result_file": "檔案：",
        "result_job_id": "Job ID：",
        "result_download_results": "下載 results.csv",
        "result_download_annotated": "下載 annotated.pdf",
        "result_open_analysis": "答案分析",
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
        "analysis_title": "答案分析",
        "analysis_hint": "流程：先下載 template.csv → 填入每題正確答案與分數（可留空白）→ 上傳並產生分析報表。",
        "analysis_job_id": "Job ID：",
        "analysis_label_default_points": "每題預設分數（可之後逐題微調）",
        "analysis_btn_download_template": "下載 template.csv",
        "analysis_hint_template": "template.csv 會包含學生作答；請在 correct/points 欄位補齊。",
        "analysis_label_job_id": "Job ID",
        "analysis_ph_job_id": "貼上處理完成頁面顯示的 Job ID",
        "analysis_label_template_file": "上傳已填好的 template.csv",
        "analysis_btn_run": "產生分析",
        "analysis_hint_upload": "如果你用 Excel 編輯，請存成 CSV 再上傳。",
        "analysis_error_missing_job": "找不到此 Job ID 的輸出資料夾。",
        "analysis_error_invalid_job_id": "Job ID 格式不正確。",
        "analysis_error_missing_results": "找不到 results.csv，請先完成一次辨識。",
        "analysis_error_missing_rscript": "找不到 Rscript。請先安裝 R（並確保 Rscript 在 PATH）。",
        "analysis_error_r_failed": "分析失敗：",
        "analysis_message_done": "分析完成，可下載報表與圖表。",
    },
    "en": {
        "lang_zh": "繁體中文",
        "lang_en": "English",
        "nav_download": "Download",
        "nav_upload": "Upload",
        "nav_analysis": "Analysis",
        "nav_update": "Update",
        "footer_local": "Runs locally / no cloud upload",
        "footer_docs": "Docs",
        "footer_download": "Download",
        "download_title": "Download Answer Sheet",
        "download_label_title": "Title (default: Exam Answer Sheet)",
        "download_ph_title": "e.g., Midterm 1",
        "download_label_subject": "Subject (optional)",
        "download_ph_subject": "e.g., Math",
        "download_label_num_questions": "Number of questions (1–100; must match upload)",
        "download_label_choices_count": "Choices per question (ABC/ABCD/ABCDE)",
        "download_btn_generate": "Generate & Download PDF",
        "download_hint_location": "Your browser decides where the file is saved (Save dialog or default Downloads folder).",
        "download_hint_support": "Grade supports 7–12 (including 10/11/12). Seat No supports 00–99. Choices support ABC / ABCD / ABCDE. Up to 100 questions.",
        "upload_title": "Upload for Recognition",
        "upload_label_num_questions": "Number of questions (1–100; enter per PDF and must match the sheet)",
        "upload_ph_num_questions": "e.g., 50",
        "upload_label_choices_count": "Choices per question (ABC/ABCD/ABCDE; must match the sheet)",
        "upload_label_pdf": "Upload multi-page PDF (one student per page)",
        "upload_btn_process": "Start Recognition",
        "upload_hint_output": "Outputs results.csv and annotated.pdf.",
        "update_title": "Update",
        "update_hint": "Download the latest ZIP and upload it here. The app will restart briefly.",
        "update_open_releases": "Open download page (GitHub Releases)",
        "update_open_zip": "Open ZIP download (main.zip)",
        "update_btn_git": "Update via Git (git pull)",
        "update_git_hint": "(Requires Git and a git clone)",
        "update_git_no_repo": "This folder is not a git repository (missing .git). Please use ZIP update instead.",
        "update_git_no_git": "Git not found. Please install Git or use ZIP update instead.",
        "update_git_dirty": "Uncommitted local changes detected. Git update is disabled to avoid conflicts; please commit/stash/revert and retry.",
        "update_label_zip": "Upload update ZIP",
        "update_btn_apply": "Apply update & restart",
        "update_started": "Update started. Please wait (the app will restart briefly).",
        "update_local_only": "Updates are allowed only from localhost.",
        "update_invalid_zip": "Please upload a valid .zip file.",
        "result_title": "Done",
        "result_file": "File:",
        "result_job_id": "Job ID:",
        "result_download_results": "Download results.csv",
        "result_download_annotated": "Download annotated.pdf",
        "result_open_analysis": "Answer analysis",
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
        "analysis_title": "Answer Analysis",
        "analysis_hint": "Workflow: download template.csv → fill correct answers and points (blanks supported) → upload to generate reports.",
        "analysis_job_id": "Job ID:",
        "analysis_label_default_points": "Default points per question (you can fine-tune per item later)",
        "analysis_btn_download_template": "Download template.csv",
        "analysis_hint_template": "template.csv includes student answers; fill the correct/points columns before uploading.",
        "analysis_label_job_id": "Job ID",
        "analysis_ph_job_id": "Paste the Job ID from the result page",
        "analysis_label_template_file": "Upload completed template.csv",
        "analysis_btn_run": "Run analysis",
        "analysis_hint_upload": "If you edit in Excel, save as CSV before uploading.",
        "analysis_error_missing_job": "Output folder not found for this Job ID.",
        "analysis_error_invalid_job_id": "Invalid Job ID.",
        "analysis_error_missing_results": "results.csv not found. Please run recognition first.",
        "analysis_error_missing_rscript": "Rscript not found. Please install R and ensure Rscript is on PATH.",
        "analysis_error_r_failed": "Analysis failed:",
        "analysis_message_done": "Analysis complete. Download reports and plots below.",
    },
}

_META_FILENAME = "meta.json"

_HEARTBEAT_INTERVAL_SEC = int(os.environ.get("ANSWER_SHEET_HEARTBEAT_SEC", "60"))
_IDLE_CHECK_INTERVAL_SEC = int(os.environ.get("ANSWER_SHEET_IDLE_CHECK_SEC", "300"))
_IDLE_TIMEOUT_SEC = int(os.environ.get("ANSWER_SHEET_IDLE_TIMEOUT_SEC", "600"))
_AUTO_EXIT_ENABLED = os.environ.get("ANSWER_SHEET_AUTO_EXIT", "1").strip().lower() not in {"0", "false", "no"}


def _sanitize_download_component(value: str, fallback: str) -> str:
    name = (value or "").strip().replace("\x00", "")
    if not name:
        return fallback
    name = name.replace("/", "_").replace("\\", "_")
    for ch in '<>:"|?*':
        name = name.replace(ch, "_")
    name = name.strip().strip(".")
    return name or fallback


def _upload_base_name(upload_filename: str) -> str:
    safe = (upload_filename or "").replace("\\", "/")
    name = Path(safe).name
    if not name:
        return "upload"
    base = name.rsplit(".", 1)[0] if "." in name else name
    return _sanitize_download_component(base, "upload")


def _read_job_meta(job_dir: Path) -> dict:
    meta_path = job_dir / _META_FILENAME
    if not meta_path.exists():
        return {}
    try:
        return json.loads(meta_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _write_job_meta(job_dir: Path, meta: dict) -> None:
    meta_path = job_dir / _META_FILENAME
    try:
        meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _touch_activity() -> None:
    app.state.last_heartbeat = time.monotonic()


def _shutdown_server() -> None:
    server = getattr(app.state, "uvicorn_server", None)
    if server is not None:
        server.should_exit = True
        return
    os._exit(0)


def _idle_shutdown_watcher() -> None:
    while True:
        time.sleep(max(10, int(_IDLE_CHECK_INTERVAL_SEC)))
        if not _AUTO_EXIT_ENABLED:
            continue
        active_jobs = int(getattr(app.state, "active_jobs", 0) or 0)
        if active_jobs > 0:
            continue
        last = float(getattr(app.state, "last_heartbeat", 0.0) or 0.0)
        now = time.monotonic()
        if last and (now - last) >= float(_IDLE_TIMEOUT_SEC):
            _shutdown_server()
            return


@app.on_event("startup")
def _startup_idle_shutdown():
    if not hasattr(app.state, "last_heartbeat"):
        app.state.last_heartbeat = time.monotonic()
    if not hasattr(app.state, "active_jobs"):
        app.state.active_jobs = 0
    if getattr(app.state, "idle_shutdown_started", False):
        return
    app.state.idle_shutdown_started = True
    threading.Thread(target=_idle_shutdown_watcher, daemon=True).start()


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
    merged = {"request": request, "lang": lang, "t": t, "heartbeat_interval_sec": _HEARTBEAT_INTERVAL_SEC}
    if ctx:
        merged.update(ctx)
    resp = templates.TemplateResponse(name, merged)
    q = (request.query_params.get("lang") or "").strip()
    if q in SUPPORTED_LANGS:
        resp.set_cookie(LANG_COOKIE_NAME, q, max_age=31536000, samesite="lax")
    return resp


def _is_local_request(request: Request) -> bool:
    host = (request.client.host if request.client else "") or ""
    return host in {"127.0.0.1", "::1"}

def _update_git_supported() -> bool:
    return (ROOT_DIR / ".git").exists() and (shutil.which("git") is not None)


def _update_page_ctx(extra: Optional[dict] = None) -> dict:
    ctx = {"git_supported": _update_git_supported()}
    if extra:
        ctx.update(extra)
    return ctx


def _get_bind_host_port() -> tuple[str, int]:
    host = os.environ.get("ANSWER_SHEET_HOST", os.environ.get("HOST", "127.0.0.1"))
    port_str = os.environ.get("ANSWER_SHEET_PORT", os.environ.get("PORT", "8000"))
    try:
        port = int(port_str)
    except ValueError:
        port = 8000
    return host, port


@app.middleware("http")
async def _activity_middleware(request: Request, call_next):
    _touch_activity()
    return await call_next(request)


@app.post("/api/heartbeat", include_in_schema=False)
def api_heartbeat():
    _touch_activity()
    return Response(status_code=204)

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


def _analysis_file_links(job_id: str) -> list[dict]:
    job_dir = OUTPUTS_DIR / job_id
    files: list[dict] = []
    label_by_name = {
        "analysis_template.csv": "template.csv",
        "analysis_scores.csv": "scores.csv",
        "analysis_item.csv": "item_analysis.csv",
        "analysis_summary.csv": "summary.csv",
        "analysis_score_hist.png": "score_hist.png",
        "analysis_item_plot.png": "item_plot.png",
    }
    for name, label in label_by_name.items():
        path = job_dir / name
        if path.exists():
            files.append({"url": f"/outputs/{job_id}/{name}", "label": label})
    return files


@app.get("/analysis", response_class=HTMLResponse)
def analysis_page(request: Request, job_id: str = ""):
    lang = resolve_lang(request)
    t = I18N.get(lang, I18N[DEFAULT_LANG])
    job_id = (job_id or "").strip()
    ctx: dict = {"job_id": job_id or None, "files": None, "error": None, "message": None, "default_points": 1}

    if job_id:
        if not _JOB_ID_RE.match(job_id):
            ctx["error"] = t["analysis_error_invalid_job_id"]
            return template_response(request, "analysis.html", ctx)
        if not (OUTPUTS_DIR / job_id).exists():
            ctx["error"] = t["analysis_error_missing_job"]
            return template_response(request, "analysis.html", ctx)
        ctx["files"] = _analysis_file_links(job_id)

    return template_response(request, "analysis.html", ctx)


@app.get("/analysis/template")
def analysis_template(request: Request, job_id: str, points: int = 1):
    lang = resolve_lang(request)
    t = I18N.get(lang, I18N[DEFAULT_LANG])
    job_id = (job_id or "").strip()
    if not _JOB_ID_RE.match(job_id):
        return template_response(
            request,
            "analysis.html",
            {"job_id": job_id or None, "error": t["analysis_error_invalid_job_id"], "default_points": 1},
        )

    job_dir = OUTPUTS_DIR / job_id
    if not job_dir.exists():
        return template_response(
            request,
            "analysis.html",
            {"job_id": job_id, "error": t["analysis_error_missing_job"], "default_points": 1},
        )

    results_path = job_dir / "results.csv"
    if not results_path.exists():
        return template_response(
            request,
            "analysis.html",
            {"job_id": job_id, "error": t["analysis_error_missing_results"], "default_points": 1},
        )

    points = max(0, min(100, int(points)))

    with open(results_path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows:
        return template_response(
            request,
            "analysis.html",
            {"job_id": job_id, "error": t["analysis_error_missing_results"], "default_points": points},
        )

    header = rows[0]
    if not header or header[0] != "number":
        return template_response(
            request,
            "analysis.html",
            {"job_id": job_id, "error": t["analysis_error_missing_results"], "default_points": points},
        )

    out = io.StringIO()
    writer = csv.writer(out, lineterminator="\n")
    writer.writerow(["number", "correct", "points", *header[1:]])
    for row in rows[1:]:
        if not row:
            continue
        qno = row[0]
        writer.writerow([qno, "", points, *row[1:]])

    content = out.getvalue().encode("utf-8-sig")
    filename = f"analysis_template_{job_id[:8]}.csv"
    return Response(
        content=content,
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename=\"{filename}\"'},
    )


@app.post("/api/analysis/run", response_class=HTMLResponse)
async def api_analysis_run(
    request: Request,
    job_id: str = Form(""),
    analysis_csv: UploadFile = File(...),
):
    lang = resolve_lang(request)
    t = I18N.get(lang, I18N[DEFAULT_LANG])
    job_id = (job_id or "").strip()
    ctx: dict = {"job_id": job_id or None, "files": None, "error": None, "message": None, "default_points": 1}

    if not _JOB_ID_RE.match(job_id):
        ctx["error"] = t["analysis_error_invalid_job_id"]
        return template_response(request, "analysis.html", ctx)

    job_dir = OUTPUTS_DIR / job_id
    if not job_dir.exists():
        ctx["error"] = t["analysis_error_missing_job"]
        return template_response(request, "analysis.html", ctx)

    template_path = job_dir / "analysis_template.csv"
    with open(template_path, "wb") as f:
        f.write(await analysis_csv.read())

    rscript = shutil.which("Rscript")
    if rscript is None:
        ctx["error"] = t["analysis_error_missing_rscript"]
        ctx["files"] = _analysis_file_links(job_id)
        return template_response(request, "analysis.html", ctx)

    script_path = ROOT_DIR / "engine" / "item_analysis_cli.R"
    if not script_path.exists():
        ctx["error"] = "item_analysis_cli.R not found."
        ctx["files"] = _analysis_file_links(job_id)
        return template_response(request, "analysis.html", ctx)

    import subprocess

    proc = subprocess.run(
        [rscript, str(script_path), "--input", str(template_path), "--outdir", str(job_dir)],
        cwd=str(ROOT_DIR),
        capture_output=True,
        text=True,
        errors="replace",
    )
    if proc.returncode != 0:
        out_text = ((proc.stdout or "") + (proc.stderr or "")).strip()
        if len(out_text) > 2000:
            out_text = out_text[-2000:]
        ctx["error"] = f"{t['analysis_error_r_failed']} {out_text or f'code {proc.returncode}'}"
        ctx["files"] = _analysis_file_links(job_id)
        return template_response(request, "analysis.html", ctx)

    ctx["message"] = t["analysis_message_done"]
    ctx["files"] = _analysis_file_links(job_id)
    return template_response(request, "analysis.html", ctx)


@app.get("/update", response_class=HTMLResponse)
def update_page(request: Request):
    return template_response(request, "update.html", _update_page_ctx())


def _sanitize_token(value: str, fallback: str) -> str:
    token = (value or "").strip().replace("\x00", "")
    token = re.sub(r"[^A-Za-z0-9-]+", "_", token)
    token = re.sub(r"_+", "_", token).strip("_.-")
    return token or fallback


@app.post("/api/generate")
def api_generate(
    title_text: str = Form(DEFAULT_SHEET_TITLE),
    subject: str = Form(""),
    num_questions: int = Form(50),
    choices_count: int = Form(4),
):
    num_questions = max(1, min(100, int(num_questions)))
    choices_count = max(3, min(5, int(choices_count)))

    out_id = str(uuid.uuid4())[:8]
    out_path = OUTPUTS_DIR / f"answer_sheets_{out_id}.pdf"

    subject = subject.strip()
    title_text = title_text.strip() or DEFAULT_SHEET_TITLE

    generate_answer_sheet_pdf(
        subject=subject,
        num_questions=num_questions,
        choices_count=choices_count,
        out_pdf_path=str(out_path),
        title_text=title_text,
    )

    subject_token = _sanitize_token(subject, "subject") if subject else "subject"
    title_token = _sanitize_token(title_text, "exam")
    choices_token = "".join(chr(ord("A") + i) for i in range(choices_count))
    filename_bits = ["answer_sheets", subject_token]
    filename_bits.append(title_token)
    filename_bits.append(choices_token)
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
    choices_count: int = Form(4),
):
    num_questions = max(1, min(100, int(num_questions)))
    choices_count = max(3, min(5, int(choices_count)))

    job_id = str(uuid.uuid4())
    job_dir = OUTPUTS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    original_filename = (pdf.filename or "").strip() or "upload.pdf"
    original_filename = original_filename.replace("\\", "/")
    original_filename = _sanitize_download_component(Path(original_filename).name, "upload.pdf")

    input_pdf = job_dir / "input.pdf"
    with open(input_pdf, "wb") as f:
        f.write(await pdf.read())

    _write_job_meta(
        job_dir,
        {
            "original_filename": original_filename,
            "upload_base": _upload_base_name(original_filename),
            "created_at": int(time.time()),
            "num_questions": num_questions,
            "choices_count": choices_count,
        },
    )

    csv_path = job_dir / "results.csv"
    ambiguity_csv_path = job_dir / "ambiguity.csv"
    annotated_pdf_path = job_dir / "annotated.pdf"

    # Main processing
    app.state.active_jobs = int(getattr(app.state, "active_jobs", 0) or 0) + 1
    try:
        process_pdf_to_csv_and_annotated_pdf(
            input_pdf_path=str(input_pdf),
            num_questions=num_questions,
            choices_count=choices_count,
            out_csv_path=str(csv_path),
            out_ambiguity_csv_path=str(ambiguity_csv_path),
            out_annotated_pdf_path=str(annotated_pdf_path),
        )
    finally:
        app.state.active_jobs = max(0, int(getattr(app.state, "active_jobs", 1) or 1) - 1)

    return template_response(
        request,
        "result.html",
        {
            "job_id": job_id,
            "display_filename": original_filename,
            "csv_url": f"/outputs/{job_id}/results.csv",
            "pdf_url": f"/outputs/{job_id}/annotated.pdf",
        },
    )


@app.post("/api/update/apply_zip", response_class=HTMLResponse)
async def api_update_apply_zip(
    request: Request,
    background_tasks: BackgroundTasks,
    zip_file: UploadFile = File(...),
):
    lang = resolve_lang(request)
    t = I18N.get(lang, I18N[DEFAULT_LANG])
    if not _is_local_request(request):
        return template_response(request, "update.html", _update_page_ctx({"error": t["update_local_only"]}))

    filename = (zip_file.filename or "").strip()
    if not filename.lower().endswith(".zip"):
        return template_response(request, "update.html", _update_page_ctx({"error": t["update_invalid_zip"]}))

    updates_dir = OUTPUTS_DIR / "_updates"
    updates_dir.mkdir(parents=True, exist_ok=True)
    zip_path = updates_dir / f"update_{int(time.time())}.zip"
    with open(zip_path, "wb") as f:
        f.write(await zip_file.read())

    host, port = _get_bind_host_port()
    worker = ROOT_DIR / "update_worker.py"
    if not worker.exists():
        return template_response(request, "update.html", _update_page_ctx({"error": "update_worker.py not found."}))

    import subprocess
    import sys

    exe = sys.executable
    if os.name == "nt":
        pyw = Path(sys.executable).with_name("pythonw.exe")
        if pyw.exists():
            exe = str(pyw)

    kwargs: dict = {"cwd": str(ROOT_DIR)}
    if os.name == "nt":
        creationflags = 0
        creationflags |= getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        creationflags |= getattr(subprocess, "DETACHED_PROCESS", 0)
        creationflags |= getattr(subprocess, "CREATE_NO_WINDOW", 0)
        kwargs["creationflags"] = creationflags
    else:
        kwargs["start_new_session"] = True

    subprocess.Popen(
        [exe, str(worker), "--zip", str(zip_path), "--host", host, "--port", str(port)],
        **kwargs,
    )

    background_tasks.add_task(_shutdown_server)
    return template_response(
        request,
        "update.html",
        _update_page_ctx({"updating": True, "message": t["update_started"]}),
    )


@app.post("/api/update/git_pull", response_class=HTMLResponse)
async def api_update_git_pull(
    request: Request,
    background_tasks: BackgroundTasks,
):
    lang = resolve_lang(request)
    t = I18N.get(lang, I18N[DEFAULT_LANG])
    if not _is_local_request(request):
        return template_response(request, "update.html", _update_page_ctx({"error": t["update_local_only"]}))

    if not (ROOT_DIR / ".git").exists():
        return template_response(request, "update.html", _update_page_ctx({"error": t["update_git_no_repo"]}))
    if shutil.which("git") is None:
        return template_response(request, "update.html", _update_page_ctx({"error": t["update_git_no_git"]}))

    import subprocess

    if os.name == "nt":
        status_kwargs = {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}
    else:
        status_kwargs = {}
    status = subprocess.run(
        ["git", "status", "--porcelain"],
        cwd=str(ROOT_DIR),
        capture_output=True,
        text=True,
        errors="replace",
        **status_kwargs,
    )
    if status.returncode == 0 and (status.stdout or "").strip():
        return template_response(request, "update.html", _update_page_ctx({"error": t["update_git_dirty"]}))

    host, port = _get_bind_host_port()
    worker = ROOT_DIR / "update_worker.py"
    if not worker.exists():
        return template_response(request, "update.html", _update_page_ctx({"error": "update_worker.py not found."}))

    import sys

    exe = sys.executable
    if os.name == "nt":
        pyw = Path(sys.executable).with_name("pythonw.exe")
        if pyw.exists():
            exe = str(pyw)

    kwargs: dict = {"cwd": str(ROOT_DIR)}
    if os.name == "nt":
        creationflags = 0
        creationflags |= getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        creationflags |= getattr(subprocess, "DETACHED_PROCESS", 0)
        creationflags |= getattr(subprocess, "CREATE_NO_WINDOW", 0)
        kwargs["creationflags"] = creationflags
    else:
        kwargs["start_new_session"] = True

    subprocess.Popen(
        [exe, str(worker), "--git", "--host", host, "--port", str(port)],
        **kwargs,
    )

    background_tasks.add_task(_shutdown_server)
    return template_response(
        request,
        "update.html",
        _update_page_ctx({"updating": True, "message": t["update_started"]}),
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

    job_tag = (str(job_id).split("-", 1)[0] or "")[:8] or "job"
    meta = _read_job_meta(job_dir)
    upload_base = _sanitize_download_component(str(meta.get("upload_base") or ""), "upload")
    upload_base_ascii = _sanitize_token(upload_base, "upload")

    download_name = filename
    if filename == "results.csv":
        download_name = f"{upload_base_ascii}_{job_tag}_results.csv"
    elif filename == "ambiguity.csv":
        download_name = f"{upload_base_ascii}_{job_tag}_ambiguity.csv"
    elif filename == "annotated.pdf":
        download_name = f"{upload_base_ascii}_{job_tag}_annotated.pdf"
    elif filename == "input.pdf":
        download_name = f"{upload_base_ascii}_{job_tag}_input.pdf"
    elif filename == "analysis_template.csv":
        download_name = f"{upload_base_ascii}_{job_tag}_analysis_template.csv"
    elif filename == "analysis_scores.csv":
        download_name = f"{upload_base_ascii}_{job_tag}_analysis_scores.csv"
    elif filename == "analysis_item.csv":
        download_name = f"{upload_base_ascii}_{job_tag}_analysis_item.csv"
    elif filename == "analysis_summary.csv":
        download_name = f"{upload_base_ascii}_{job_tag}_analysis_summary.csv"
    elif filename == "analysis_score_hist.png":
        download_name = f"{upload_base_ascii}_{job_tag}_analysis_score_hist.png"
    elif filename == "analysis_item_plot.png":
        download_name = f"{upload_base_ascii}_{job_tag}_analysis_item_plot.png"

    media = "application/octet-stream"
    if filename.lower().endswith(".pdf"):
        media = "application/pdf"
    if filename.lower().endswith(".csv"):
        media = "text/csv"
    if filename.lower().endswith(".png"):
        media = "image/png"
    return FileResponse(path=str(file_path), media_type=media, filename=download_name)


def run():
    import uvicorn

    host, port = _get_bind_host_port()

    try:
        config = uvicorn.Config(app, host=host, port=port, reload=False)
        server = uvicorn.Server(config)
        app.state.uvicorn_server = server
        server.run()
    except PermissionError as exc:
        print(
            f"\nERROR: Permission denied while binding to {host}:{port}.\n"
            "macOS or your security software blocked Python from accepting incoming connections.\n"
            "Allow Python in System Settings > Network > Firewall (or equivalent) and rerun.\n"
            "You can also choose another port by setting ANSWER_SHEET_PORT.\n"
        )
        raise exc
