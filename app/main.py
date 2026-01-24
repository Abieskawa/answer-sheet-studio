import csv
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

from engine.analysis import run_analysis_template
from engine.generator import generate_answer_sheet_pdf, DEFAULT_TITLE as DEFAULT_SHEET_TITLE
from engine.recognizer import process_pdf_to_csv_and_annotated_pdf
from engine.xlsx import read_simple_xlsx_table, write_simple_xlsx, write_simple_xlsx_multi

APP_DIR = Path(__file__).resolve().parent
ROOT_DIR = APP_DIR.parent
OUTPUTS_DIR = ROOT_DIR / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)

templates = Jinja2Templates(directory=str(APP_DIR / "templates"))

app = FastAPI(title="Answer Sheet Studio")
app.mount("/static", StaticFiles(directory=str(APP_DIR / "static")), name="static")

_JOB_ID_RE = re.compile(r"^[0-9a-fA-F-]{8,64}$")
_INLINE_OUTPUT_FILENAMES = {"analysis_score_hist.png", "analysis_item_plot.png"}

LANG_COOKIE_NAME = "lang"
SUPPORTED_LANGS = ("zh-Hant", "en")
DEFAULT_LANG = "zh-Hant"

DOCS_URL_BY_LANG = {
    "zh-Hant": "https://answer-sheet-studio.readthedocs.io/zh-tw/latest/",
    "en": "https://answer-sheet-studio.readthedocs.io/en/latest/",
}

I18N = {
    "zh-Hant": {
        "lang_zh": "繁體中文",
        "lang_en": "English",
        "nav_download": "下載答案卡",
        "nav_upload": "上傳處理",
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
        "download_btn_sheet": "下載答案卡 PDF",
        "download_btn_answer_key": "下載老師答案檔（Excel）",
        "download_hint_sheet": "先下載並列印答案卡 PDF（給學生填）。",
        "download_hint_answer_key": "需要編輯正確答案/配分時，再下載老師答案檔（Excel）填 correct/points，之後到「上傳處理」上傳即可。",
        "download_hint_support": "年級欄位支援 7–12（含 10/11/12），座號支援 00–99，選擇題支援 ABC / ABCD / ABCDE，最多 100 題。",
        "upload_title": "上傳處理",
        "upload_label_num_questions": "題數（1–100；請依每份掃描檔案輸入，需與答案卡一致）",
        "upload_ph_num_questions": "例如：50",
        "upload_label_choices_count": "每題選項（ABC/ABCD/ABCDE；需與答案卡一致）",
        "upload_label_pdf": "上傳多頁 PDF（每頁一位）",
        "upload_label_answer_key": "上傳老師答案檔（Excel .xlsx；correct/points）",
        "upload_btn_process": "開始辨識並分析",
        "upload_processing": "處理中，請稍候…",
        "upload_open_result": "開啟結果頁（含圖表）",
        "upload_open_result_hint": "處理完成後請點上方按鈕開啟結果頁。",
        "upload_error_generic": "處理失敗，請查看 outputs/launcher.log 或 outputs/server.log。",
        "upload_hint_output": "完成後會輸出 results.csv、annotated.pdf，以及答案分析報表/圖表（若已安裝 R 會使用 ggplot2 出圖）。",
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
        "result_download_showwrong": "下載 showwrong.xlsx（只顯示錯題）",
        "result_back_to_downloads": "查看試題分析原始數據",
        "result_view_item_analysis_data": "試題分析數據",
        "result_plots_title": "圖表",
        "result_plot_score_hist": "成績分佈",
        "result_plot_item_metrics": "題目分析",
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
        "analysis_error_missing_rscript": "找不到 Rscript（若要用 ggplot2 出圖，請先安裝 R）。",
        "analysis_error_r_failed": "R 分析失敗：",
        "analysis_error_builtin_failed": "內建分析失敗：",
        "analysis_message_done": "分析完成，可下載報表與圖表。",
        "analysis_message_done_fallback": "分析完成（已使用內建分析；若想用 ggplot2 出圖，請確認已安裝 R。程式會在需要時自動嘗試安裝 R 套件：readr、dplyr、tidyr、ggplot2）。",
        "analysis_note_discrimination_rule_27": "鑑別度：學生數 > 30 時採前後 27%（高分組答對率 − 低分組答對率）。",
        "analysis_note_discrimination_rule_50": "鑑別度：學生數 ≤ 30 時採前後 50%（高分組答對率 − 低分組答對率）。",
    },
    "en": {
        "lang_zh": "繁體中文",
        "lang_en": "English",
        "nav_download": "Download",
        "nav_upload": "Upload",
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
        "download_btn_sheet": "Download Answer Sheet PDF",
        "download_btn_answer_key": "Download Teacher Answer Key (Excel)",
        "download_hint_sheet": "Download and print the answer sheet PDF (for students).",
        "download_hint_answer_key": "When you need to edit correct answers/points, download the teacher answer key (Excel) and upload it on the Upload page.",
        "download_hint_support": "Grade supports 7–12 (including 10/11/12). Seat No supports 00–99. Choices support ABC / ABCD / ABCDE. Up to 100 questions.",
        "upload_title": "Upload for Recognition",
        "upload_label_num_questions": "Number of questions (1–100; enter per PDF and must match the sheet)",
        "upload_ph_num_questions": "e.g., 50",
        "upload_label_choices_count": "Choices per question (ABC/ABCD/ABCDE; must match the sheet)",
        "upload_label_pdf": "Upload multi-page PDF (one student per page)",
        "upload_label_answer_key": "Upload teacher answer key (Excel .xlsx; correct/points)",
        "upload_btn_process": "Run recognition + analysis",
        "upload_processing": "Processing…",
        "upload_open_result": "Open result page (with plots)",
        "upload_open_result_hint": "When processing finishes, click the button above to open the result page.",
        "upload_error_generic": "Processing failed. See outputs/launcher.log or outputs/server.log.",
        "upload_hint_output": "Outputs results.csv, annotated.pdf, and analysis reports/plots (uses ggplot2 plots if R is installed).",
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
        "result_download_showwrong": "Download showwrong.xlsx (wrong answers only)",
        "result_back_to_downloads": "View item analysis raw data",
        "result_view_item_analysis_data": "Item analysis data",
        "result_plots_title": "Plots",
        "result_plot_score_hist": "Score distribution",
        "result_plot_item_metrics": "Item analysis",
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
        "analysis_error_missing_rscript": "Rscript not found (install R to enable ggplot2 plots).",
        "analysis_error_r_failed": "R analysis failed:",
        "analysis_error_builtin_failed": "Built-in analysis failed:",
        "analysis_message_done": "Analysis complete. Download reports and plots below.",
        "analysis_message_done_fallback": "Analysis complete (built-in analysis used; to enable ggplot2 plots, ensure R is installed. The app will auto-attempt to install R packages when needed: readr, dplyr, tidyr, ggplot2).",
        "analysis_note_discrimination_rule_27": "Discrimination: if students > 30, uses top/bottom 27% (correct_high − correct_low).",
        "analysis_note_discrimination_rule_50": "Discrimination: if students ≤ 30, uses top/bottom 50% (correct_high − correct_low).",
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


def _safe_unlink(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
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
    merged = {
        "request": request,
        "lang": lang,
        "t": t,
        "docs_url": DOCS_URL_BY_LANG.get(lang, DOCS_URL_BY_LANG[DEFAULT_LANG]),
        "heartbeat_interval_sec": _HEARTBEAT_INTERVAL_SEC,
    }
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

def _safe_output_filename(filename: str) -> Optional[str]:
    name = (filename or "").strip().replace("\x00", "")
    if not name:
        return None
    if Path(name).name != name:
        return None
    if name in {".", ".."} or ".." in name:
        return None
    return name

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


@app.get("/result/{job_id}", response_class=HTMLResponse)
def result_page(request: Request, job_id: str):
    job_id = (job_id or "").strip()
    if not _JOB_ID_RE.match(job_id):
        return RedirectResponse(url="/upload", status_code=302)

    job_dir = OUTPUTS_DIR / job_id
    if not job_dir.exists():
        return RedirectResponse(url="/upload", status_code=302)

    meta = _read_job_meta(job_dir)
    display_filename = str(meta.get("original_filename") or "") or job_id

    score_hist = job_dir / "analysis_score_hist.png"
    item_plot = job_dir / "analysis_item_plot.png"
    discr_note_key = _discrimination_note_key(job_dir)

    return template_response(
        request,
        "result.html",
        {
            "body_class": "theme-light",
            "job_id": job_id,
            "display_filename": display_filename,
            "csv_url": f"/outputs/{job_id}/results.csv",
            "pdf_url": f"/outputs/{job_id}/annotated.pdf",
            "showwrong_url": (f"/outputs/{job_id}/showwrong.xlsx" if (job_dir / "showwrong.xlsx").exists() else None),
            "analysis_error": (str(meta.get("analysis_error") or "") or None),
            "analysis_message": (str(meta.get("analysis_message") or "") or None),
            "analysis_discrimination_note_key": discr_note_key,
            "analysis_score_hist_inline_url": (f"/outputs_inline/{job_id}/analysis_score_hist.png" if score_hist.exists() else None),
            "analysis_item_plot_inline_url": (f"/outputs_inline/{job_id}/analysis_item_plot.png" if item_plot.exists() else None),
        },
    )


@app.get("/result/{job_id}/charts", response_class=HTMLResponse)
def result_charts_page(request: Request, job_id: str):
    job_id = (job_id or "").strip()
    if not _JOB_ID_RE.match(job_id):
        return RedirectResponse(url="/upload", status_code=302)

    job_dir = OUTPUTS_DIR / job_id
    if not job_dir.exists():
        return RedirectResponse(url="/upload", status_code=302)

    meta = _read_job_meta(job_dir)
    display_filename = str(meta.get("original_filename") or "") or job_id

    score_hist = job_dir / "analysis_score_hist.png"
    item_plot = job_dir / "analysis_item_plot.png"
    scores_by_class = job_dir / "analysis_scores_by_class.xlsx"
    item_table = _read_analysis_item_table(job_dir)
    discr_note_key = _discrimination_note_key(job_dir)

    return template_response(
        request,
        "result_charts.html",
        {
            "job_id": job_id,
            "display_filename": display_filename,
            "csv_url": f"/outputs/{job_id}/results.csv",
            "pdf_url": f"/outputs/{job_id}/annotated.pdf",
            "showwrong_url": (f"/outputs/{job_id}/showwrong.xlsx" if (job_dir / "showwrong.xlsx").exists() else None),
            "analysis_error": (str(meta.get("analysis_error") or "") or None),
            "analysis_message": (str(meta.get("analysis_message") or "") or None),
            "analysis_discrimination_note_key": discr_note_key,
            "analysis_score_hist_inline_url": (f"/outputs_inline/{job_id}/analysis_score_hist.png" if score_hist.exists() else None),
            "analysis_item_plot_inline_url": (f"/outputs_inline/{job_id}/analysis_item_plot.png" if item_plot.exists() else None),
            "analysis_scores_by_class_url": (f"/outputs/{job_id}/analysis_scores_by_class.xlsx" if scores_by_class.exists() else None),
            "analysis_item_table": item_table,
        },
    )


def _discrimination_note_key(job_dir: Path) -> Optional[str]:
    path = Path(job_dir) / "analysis_summary.csv"
    try:
        if not path.exists():
            return None
    except Exception:
        return None

    try:
        with open(path, "r", newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
    except Exception:
        return None

    if len(rows) < 2 or not rows[1]:
        return None

    header = [str(x or "").strip().lower() for x in rows[0]]
    try:
        idx = header.index("students")
    except ValueError:
        return None

    try:
        students = int(float(str(rows[1][idx] or "").strip()))
    except Exception:
        return None

    return "analysis_note_discrimination_rule_27" if students > 30 else "analysis_note_discrimination_rule_50"


def _read_analysis_item_table(job_dir: Path, max_rows: int = 200) -> Optional[dict]:
    path = Path(job_dir) / "analysis_item.csv"
    try:
        if not path.exists():
            return None
    except Exception:
        return None

    try:
        with open(path, "r", newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
    except Exception:
        return None

    if not rows:
        return None

    header = [str(x or "").strip() for x in rows[0]]
    body = [[str(x or "") for x in r] for r in rows[1 : 1 + max_rows]]
    truncated = len(rows) - 1 > max_rows
    return {"header": header, "rows": body, "truncated": truncated, "total_rows": max(0, len(rows) - 1)}


def _analysis_file_links(job_id: str) -> list[dict]:
    job_dir = OUTPUTS_DIR / job_id
    files: list[dict] = []
    label_by_name = {
        "roster.csv": "roster.csv",
        "analysis_scores.csv": "analysis_scores.csv",
        "analysis_scores_by_class.xlsx": "analysis_scores_by_class.xlsx",
        "analysis_item.csv": "analysis_item.csv",
        "analysis_summary.csv": "analysis_summary.csv",
        "analysis_score_hist.png": "analysis_score_hist.png",
        "analysis_item_plot.png": "analysis_item_plot.png",
        "analysis_r.log": "analysis_r.log (R log)",
    }
    for name, label in label_by_name.items():
        path = job_dir / name
        if path.exists():
            files.append({"url": f"/outputs/{job_id}/{name}", "label": label})
    return files


def _find_rscript() -> Optional[str]:
    rscript = shutil.which("Rscript") or shutil.which("Rscript.exe")
    if rscript:
        return rscript

    candidates: list[Path] = []
    if os.name == "nt":
        roots = [os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")]
        for root in roots:
            if not root:
                continue
            r_dir = Path(root) / "R"
            if not r_dir.exists():
                continue
            for exe in sorted(r_dir.glob("R-*/bin/Rscript.exe")):
                candidates.append(exe)
    else:
        candidates.extend(
            [
                Path("/usr/local/bin/Rscript"),
                Path("/opt/homebrew/bin/Rscript"),
                Path("/Library/Frameworks/R.framework/Resources/bin/Rscript"),
            ]
        )

    for cand in candidates:
        try:
            if cand.exists():
                return str(cand)
        except Exception:
            continue
    return None


def _summarize_r_error(text: Optional[str], limit: int = 220) -> Optional[str]:
    if not text:
        return None
    msg = (text or "").strip().replace("\r", "\n")
    lines = [ln.strip() for ln in msg.split("\n") if ln.strip()]
    if not lines:
        return None

    m = re.search(r"Missing R packages and auto-install failed:\s*([^\n]+)", msg, flags=re.IGNORECASE)
    if m:
        reason = f"missing R packages: {m.group(1).strip()}"
    else:
        m = re.search(r"Missing required packages:\s*([^\n]+)", msg, flags=re.IGNORECASE)
        if m:
            reason = f"missing R packages: {m.group(1).strip()}"
        else:
            m = re.search(r"there is no package called ['\"]([^'\"]+)['\"]", msg, flags=re.IGNORECASE)
            if m:
                reason = f"missing R package: {m.group(1)}"
            else:
                m = re.search(r"Missing R packages and auto-install failed:.*?Error:\\s*([^\n]+)", msg, flags=re.IGNORECASE | re.DOTALL)
                if m:
                    reason = m.group(1).strip()
                else:
                    m = re.search(r"Error:\\s*([^\n]+)", msg, flags=re.IGNORECASE)
                    if m:
                        reason = m.group(1).strip()
                    else:
                        m = re.search(r"Error in [^:]+:\\s*([^\n]+)", msg, flags=re.IGNORECASE)
                        if m:
                            reason = m.group(1).strip()
                        else:
                            reason = lines[-1]

    reason = re.sub(r"\s+", " ", reason).strip()
    if len(reason) > limit:
        reason = reason[:limit].rstrip() + "…"
    return reason


def _normalize_answer_cell(value: object) -> str:
    s = ("" if value is None else str(value)).strip().upper()
    return s


def _read_answer_key_csv(path: Path) -> dict[int, tuple[str, float]]:
    with open(path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError("missing header (expected: number, correct, points)")

        field_map = {str(name).strip().lower(): str(name) for name in reader.fieldnames if name}
        required = ("number", "correct", "points")
        missing = [k for k in required if k not in field_map]
        if missing:
            raise ValueError(f"missing columns: {', '.join(missing)}")

        number_col = field_map["number"]
        correct_col = field_map["correct"]
        points_col = field_map["points"]

        out: dict[int, tuple[str, float]] = {}
        for row in reader:
            raw_no = (row.get(number_col) or "").strip()
            if not raw_no:
                continue
            try:
                qno = int(raw_no)
            except ValueError:
                continue

            correct = _normalize_answer_cell(row.get(correct_col, ""))
            raw_points = (row.get(points_col) or "").strip()
            try:
                points = float(raw_points) if raw_points != "" else 1.0
            except ValueError:
                points = 1.0
            out[qno] = (correct, points)
        return out


def _read_answer_key_xlsx(path: Path) -> dict[int, tuple[str, float]]:
    table = read_simple_xlsx_table(path)
    if not table:
        raise ValueError("empty xlsx")
    header = [str(x).strip() for x in table[0]]
    field_map = {str(name).strip().lower(): idx for idx, name in enumerate(header) if str(name).strip()}
    required = ("number", "correct", "points")
    missing = [k for k in required if k not in field_map]
    if missing:
        raise ValueError(f"missing columns: {', '.join(missing)}")

    idx_number = int(field_map["number"])
    idx_correct = int(field_map["correct"])
    idx_points = int(field_map["points"])

    out: dict[int, tuple[str, float]] = {}
    for row in table[1:]:
        raw_no = (row[idx_number] if idx_number < len(row) else "").strip()
        if not raw_no:
            continue
        try:
            qno = int(float(raw_no))  # Excel may store as "1" or "1.0"
        except ValueError:
            continue

        correct = _normalize_answer_cell(row[idx_correct] if idx_correct < len(row) else "")
        raw_points = (row[idx_points] if idx_points < len(row) else "").strip()
        try:
            points = float(raw_points) if raw_points != "" else 1.0
        except ValueError:
            points = 1.0
        out[qno] = (correct, points)
    return out


def _read_answer_key_file(path: Path) -> dict[int, tuple[str, float]]:
    suffix = (path.suffix or "").lower()
    if suffix == ".xlsx":
        return _read_answer_key_xlsx(path)
    return _read_answer_key_csv(path)


def _write_answer_key_files(
    num_questions: int,
    answer_key: dict[int, tuple[str, float]],
    out_csv_path: Optional[Path] = None,
    out_xlsx_path: Optional[Path] = None,
    default_points: float = 1.0,
) -> None:
    num_questions = max(1, min(100, int(num_questions)))
    rows: list[list[object]] = [["number", "correct", "points"]]
    for qno in range(1, num_questions + 1):
        correct = ""
        points = float(default_points)
        if qno in answer_key:
            c, p = answer_key[qno]
            correct = _normalize_answer_cell(c)
            try:
                points = float(p)
            except Exception:
                points = float(default_points)
        points_out: object = points
        if isinstance(points, float) and abs(points - round(points)) < 1e-9:
            points_out = int(round(points))
        rows.append([qno, correct, points_out])

    if out_csv_path is not None:
        out_csv_path = Path(out_csv_path)
        out_csv_path.parent.mkdir(parents=True, exist_ok=True)
        with open(out_csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, lineterminator="\n")
            writer.writerows(rows)

    if out_xlsx_path is not None:
        write_simple_xlsx(Path(out_xlsx_path), rows=rows, sheet_name="answer_key")


def _write_analysis_template(
    results_csv_path: Path,
    answer_key: dict[int, tuple[str, float]],
    out_csv_path: Path,
    default_points: float = 1.0,
) -> None:
    with open(results_csv_path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows:
        raise ValueError("results.csv is empty")

    header = rows[0]
    if not header or header[0] != "number":
        raise ValueError("unexpected results.csv header (expected first column: number)")

    with open(out_csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(["number", "correct", "points", *header[1:]])
        for row in rows[1:]:
            if not row:
                continue
            raw_no = (row[0] or "").strip()
            try:
                qno = int(raw_no)
            except ValueError:
                qno = None

            correct = ""
            points = float(default_points)
            if qno is not None and qno in answer_key:
                correct, points = answer_key[qno]
                if correct is None:
                    correct = ""

            points_out: object = points
            if isinstance(points, float) and abs(points - round(points)) < 1e-9:
                points_out = int(round(points))

            writer.writerow([row[0], _normalize_answer_cell(correct), points_out, *row[1:]])


def _write_showwrong_xlsx(
    results_csv_path: Path,
    answer_key: dict[int, tuple[str, float]],
    out_xlsx_path: Path,
    blank_label: str = "空白",
) -> None:
    with open(results_csv_path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows:
        raise ValueError("results.csv is empty")

    header = rows[0]
    if not header or header[0] != "number":
        raise ValueError("unexpected results.csv header (expected first column: number)")

    student_cols = [str(x or "").strip() for x in header[1:] if str(x or "").strip()]
    if not student_cols:
        raise ValueError("no students found in results.csv")

    q_numbers: list[int] = []
    corrects: list[str] = []
    points_list: list[float] = []
    answers_by_student: dict[str, list[str]] = {s: [] for s in student_cols}

    for row in rows[1:]:
        if not row:
            continue
        raw_no = (row[0] or "").strip()
        try:
            qno = int(raw_no)
        except ValueError:
            continue

        correct, points = answer_key.get(qno, ("", 1.0))
        correct_norm = _normalize_answer_cell(correct)
        try:
            points_f = float(points) if points is not None else 1.0
        except Exception:
            points_f = 1.0

        q_numbers.append(qno)
        corrects.append(correct_norm)
        points_list.append(points_f)

        for idx, student in enumerate(student_cols):
            val = row[idx + 1] if idx + 1 < len(row) else ""
            answers_by_student[student].append(_normalize_answer_cell(val))

    def score_out(v: float) -> object:
        if abs(float(v) - round(float(v))) < 1e-9:
            return int(round(float(v)))
        return round(float(v), 2)

    student_scores: dict[str, float] = {s: 0.0 for s in student_cols}

    out_rows: list[list[object]] = []
    out_rows.append(["number", "correct", *student_cols])

    for i, qno in enumerate(q_numbers):
        correct = corrects[i] if i < len(corrects) else ""
        row_out: list[object] = [qno, correct]
        if not correct:
            row_out.extend(["" for _ in student_cols])
            out_rows.append(row_out)
            continue

        points = points_list[i] if i < len(points_list) else 1.0
        for student in student_cols:
            answers = answers_by_student.get(student, [])
            ans = answers[i] if i < len(answers) else ""
            if ans == correct:
                student_scores[student] += float(points)
                row_out.append("")
            else:
                row_out.append(blank_label if ans == "" else ans)
        out_rows.append(row_out)

    out_rows.append(["score", "", *[score_out(student_scores[s]) for s in student_cols]])

    # If roster.csv exists, split into per-class worksheets (same class on same sheet).
    roster_path = Path(results_csv_path).with_name("roster.csv")
    roster_by_person: dict[str, dict[str, str]] = {}
    try:
        if roster_path.exists():
            with open(roster_path, "r", newline="", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    pid = str((r.get("person_id") or "")).strip()
                    if not pid:
                        continue
                    roster_by_person[pid] = {
                        "grade": str(r.get("grade") or "").strip(),
                        "class_no": str(r.get("class_no") or "").strip(),
                        "seat_no": str(r.get("seat_no") or "").strip(),
                    }
    except Exception:
        roster_by_person = {}

    def class_key(meta: dict[str, str]) -> str:
        g = (meta.get("grade") or "").strip()
        c = (meta.get("class_no") or "").strip()
        if g and c:
            return f"{g}-{c}"
        if g:
            return f"{g}-?"
        if c:
            return f"?-{c}"
        return "Unknown"

    grouped_students: dict[str, list[str]] = {}
    if roster_by_person:
        for s in student_cols:
            meta = roster_by_person.get(str(s).strip(), {})
            grouped_students.setdefault(class_key(meta), []).append(s)

    def student_sort_key(student_id: str) -> tuple[int, int, str]:
        meta = roster_by_person.get(str(student_id).strip(), {})
        seat_s = (meta.get("seat_no") or "").strip()
        try:
            seat_n = int(float(seat_s)) if seat_s != "" else 10**9
        except Exception:
            seat_n = 10**9
        return (0 if seat_n != 10**9 else 1, seat_n, str(student_id))

    if grouped_students:
        sheets: list[tuple[str, list[list[object]]]] = []
        for key in sorted(grouped_students.keys()):
            students = grouped_students[key]
            students.sort(key=student_sort_key)

            rows_for_sheet: list[list[object]] = []
            rows_for_sheet.append(["number", "correct", *students])

            for i, qno in enumerate(q_numbers):
                correct = corrects[i] if i < len(corrects) else ""
                row_out: list[object] = [qno, correct]
                if not correct:
                    row_out.extend(["" for _ in students])
                    rows_for_sheet.append(row_out)
                    continue

                for student in students:
                    answers = answers_by_student.get(student, [])
                    ans = answers[i] if i < len(answers) else ""
                    row_out.append("" if ans == correct else (blank_label if ans == "" else ans))
                rows_for_sheet.append(row_out)

            rows_for_sheet.append(["score", "", *[score_out(student_scores[s]) for s in students]])
            sheets.append((f"Class {key}", rows_for_sheet))

        write_simple_xlsx_multi(Path(out_xlsx_path), sheets=sheets)
    else:
        write_simple_xlsx(Path(out_xlsx_path), rows=out_rows, sheet_name="showwrong")


@app.get("/update", response_class=HTMLResponse)
def update_page(request: Request):
    return template_response(request, "update.html", _update_page_ctx())


def _sanitize_token(value: str, fallback: str) -> str:
    token = (value or "").strip().replace("\x00", "")
    token = re.sub(r"[^A-Za-z0-9-]+", "_", token)
    token = re.sub(r"_+", "_", token).strip("_.-")
    return token or fallback


@app.post("/api/generate_sheet")
def api_generate_sheet(
    background_tasks: BackgroundTasks,
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
    filename_bits = ["answer_sheets", subject_token, title_token, choices_token, f"{num_questions}q"]

    pdf_name = "_".join(filename_bits) + ".pdf"
    background_tasks.add_task(_safe_unlink, out_path)

    return FileResponse(path=str(out_path), media_type="application/pdf", filename=pdf_name)


@app.post("/api/generate_answer_key")
def api_generate_answer_key(
    background_tasks: BackgroundTasks,
    title_text: str = Form(DEFAULT_SHEET_TITLE),
    subject: str = Form(""),
    num_questions: int = Form(50),
    choices_count: int = Form(4),
):
    num_questions = max(1, min(100, int(num_questions)))
    choices_count = max(3, min(5, int(choices_count)))

    out_id = str(uuid.uuid4())[:8]

    subject = subject.strip()
    title_text = title_text.strip() or DEFAULT_SHEET_TITLE

    subject_token = _sanitize_token(subject, "subject") if subject else "subject"
    title_token = _sanitize_token(title_text, "exam")
    choices_token = "".join(chr(ord("A") + i) for i in range(choices_count))

    key_xlsx_name = "_".join(["answer_key", subject_token, title_token, choices_token, f"{num_questions}q"]) + ".xlsx"
    key_xlsx_path = OUTPUTS_DIR / f"answer_key_{out_id}.xlsx"
    _write_answer_key_files(
        num_questions=num_questions,
        answer_key={},
        out_xlsx_path=key_xlsx_path,
        default_points=1.0,
    )

    background_tasks.add_task(_safe_unlink, key_xlsx_path)

    return FileResponse(
        path=str(key_xlsx_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=key_xlsx_name,
    )


@app.post("/api/process")
async def api_process(
    request: Request,
    pdf: UploadFile = File(...),
    answer_key: UploadFile = File(...),
    num_questions: int = Form(50),
    choices_count: int = Form(4),
):
    lang = resolve_lang(request)
    t = I18N.get(lang, I18N[DEFAULT_LANG])
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

    answer_key_filename = (answer_key.filename or "").strip() or "answer_key.xlsx"
    answer_key_filename = answer_key_filename.replace("\\", "/")
    answer_key_filename = _sanitize_download_component(Path(answer_key_filename).name, "answer_key.xlsx")
    answer_key_suffix = (Path(answer_key_filename).suffix or "").lower()
    if answer_key_suffix not in {".csv", ".xlsx"}:
        answer_key_suffix = ".xlsx"
    answer_key_upload_path = job_dir / f"answer_key_upload{answer_key_suffix}"
    with open(answer_key_upload_path, "wb") as f:
        f.write(await answer_key.read())

    answer_key_xlsx_path = job_dir / "answer_key.xlsx"
    showwrong_xlsx_path = job_dir / "showwrong.xlsx"

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

    analysis_error: Optional[str] = None
    analysis_message: Optional[str] = None

    template_path = job_dir / "analysis_template.csv"
    try:
        key_map = _read_answer_key_file(answer_key_upload_path)
    except Exception as exc:
        key_map = None
        analysis_error = f"Answer key error: {exc}"
    else:
        try:
            _write_answer_key_files(
                num_questions=num_questions,
                answer_key=key_map,
                out_xlsx_path=answer_key_xlsx_path,
                default_points=1.0,
            )
        except Exception:
            pass
        try:
            _write_showwrong_xlsx(csv_path, key_map, showwrong_xlsx_path)
        except Exception as exc:
            if analysis_error is None:
                analysis_error = f"Showwrong error: {exc}"

    if key_map is not None:
        try:
            _write_analysis_template(csv_path, key_map, template_path, default_points=1.0)
        except Exception as exc:
            analysis_error = f"Analysis template error: {exc}"
        else:
            rscript = _find_rscript()
            script_path = ROOT_DIR / "engine" / "item_analysis_cli.R"
            r_failed_text: Optional[str] = None
            r_ok = False

            if rscript is not None and script_path.exists():
                import subprocess

                proc = subprocess.run(
                    [rscript, str(script_path), "--input", str(template_path), "--outdir", str(job_dir), "--lang", str(lang)],
                    cwd=str(ROOT_DIR),
                    capture_output=True,
                    text=True,
                    errors="replace",
                )
                if proc.returncode == 0:
                    r_ok = True
                else:
                    out_text = ((proc.stdout or "") + (proc.stderr or "")).strip()
                    if len(out_text) > 2000:
                        out_text = out_text[-2000:]
                    r_failed_text = out_text or f"code {proc.returncode}"
                    try:
                        (job_dir / "analysis_r.log").write_text(
                            ((proc.stdout or "") + "\n" + (proc.stderr or "")).strip() + "\n",
                            encoding="utf-8",
                            errors="replace",
                        )
                    except Exception:
                        pass

            if r_ok:
                analysis_message = t["analysis_message_done"]
            else:
                try:
                    run_analysis_template(template_path, job_dir, lang=lang)
                except Exception as exc:
                    analysis_error = f"{t['analysis_error_builtin_failed']} {exc}"
                    if rscript is not None:
                        analysis_error = f"{analysis_error} ({t['analysis_error_r_failed']} {r_failed_text or ''})"
                    else:
                        analysis_error = f"{analysis_error} ({t['analysis_error_missing_rscript']})"
                else:
                    analysis_message = t["analysis_message_done_fallback"]
                    if rscript is None:
                        analysis_message = f"{analysis_message} ({t['analysis_error_missing_rscript']})"
                    else:
                        reason = _summarize_r_error(r_failed_text)
                        if reason:
                            analysis_message = f"{analysis_message} ({t['analysis_error_r_failed']} {reason})"

    meta = _read_job_meta(job_dir)
    if analysis_error:
        meta["analysis_error"] = analysis_error
    else:
        meta.pop("analysis_error", None)
    if analysis_message:
        meta["analysis_message"] = analysis_message
    else:
        meta.pop("analysis_message", None)
    _write_job_meta(job_dir, meta)

    return RedirectResponse(url=f"/result/{job_id}/charts", status_code=303)


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
    job_id = (job_id or "").strip()
    filename = _safe_output_filename(filename) or ""
    if not _JOB_ID_RE.match(job_id) or not filename:
        return RedirectResponse(url="/upload", status_code=302)

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
    elif filename == "answer_key.xlsx":
        download_name = f"{upload_base_ascii}_{job_tag}_answer_key.xlsx"
    elif filename == "showwrong.xlsx":
        download_name = f"{upload_base_ascii}_{job_tag}_showwrong.xlsx"
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
    if filename.lower().endswith(".xlsx"):
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if filename.lower().endswith(".png"):
        media = "image/png"
    return FileResponse(path=str(file_path), media_type=media, filename=download_name)

@app.get("/outputs_inline/{job_id}/{filename}")
def view_output_inline(job_id: str, filename: str):
    job_id = (job_id or "").strip()
    filename = _safe_output_filename(filename) or ""
    if not _JOB_ID_RE.match(job_id) or filename not in _INLINE_OUTPUT_FILENAMES:
        return RedirectResponse(url="/upload", status_code=302)

    job_dir = OUTPUTS_DIR / job_id
    file_path = job_dir / filename
    if not file_path.exists():
        return RedirectResponse(url="/upload", status_code=302)

    return FileResponse(
        path=str(file_path),
        media_type="image/png",
        filename=filename,
        content_disposition_type="inline",
    )


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
